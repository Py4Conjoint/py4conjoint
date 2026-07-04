"""design とデータの対応ずれ対策（署名・構造チェック）の検証。

選択型では、アンケート作成に使った design と forms_to_data に渡す design が
1つでもずれる（水準順序違い・seed 違いなど）と、回答と代替案の対応が
**エラーなく食い違って結果が誤る**。この事故を「気づける」ようにするための

* design_signature : 内容ベースの署名（順序違いを区別する／同一なら一致する）
* forms_to_data    : 署名を出力に引き継ぐ／構造の崩れた design を弾く

を検証する。
"""
import sys
import warnings
from contextlib import contextmanager
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np
import pandas as pd
import pytest

import py4conjoint.choice as pcc

ATTRS = {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]}


# ---------------------------------------------------------------------------
# design_signature: 順序違いは別署名、完全に同一なら同じ署名
# ---------------------------------------------------------------------------

def test_signature_differs_when_level_order_changes():
    """水準の順序を変えると、同じ seed でも署名が変わる。"""
    d1 = pcc.design_choice_sets(
        {"price": [6, 10], "os": ["a", "b"]}, n_sets=4, n_alts=2, seed=1
    )
    d2 = pcc.design_choice_sets(
        {"price": [10, 6], "os": ["a", "b"]}, n_sets=4, n_alts=2, seed=1
    )
    assert pcc.design_signature(d1) != pcc.design_signature(d2)


def test_signature_same_for_identical_spec_and_seed():
    """属性・水準・順序・seed が完全に同一なら署名は一致する。"""
    d1 = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    d2 = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    assert pcc.design_signature(d1) == pcc.design_signature(d2)


def test_signature_attached_to_attrs():
    """design_choice_sets の出力 attrs に署名が入っている。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    assert "design_signature" in design.attrs
    assert design.attrs["design_signature"] == pcc.design_signature(design)


def test_signature_survives_csv_round_trip(tmp_path):
    """推奨ワークフロー（保存→読込）で署名が保たれる（内容から再計算できる）。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    csv = tmp_path / "design.csv"
    design.to_csv(csv, index=False)
    loaded = pd.read_csv(csv)
    assert pcc.design_signature(loaded) == pcc.design_signature(design)


def test_signature_changes_without_seed():
    """seed=None は呼ぶたびに中身が変わるため署名も変わる（正直な挙動）。"""
    a = pcc.design_signature(pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3))
    b = pcc.design_signature(pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3))
    assert a != b


def test_signature_ignores_attribute_column_order():
    """属性列の並び順だけが違う（中身は同じ）設計は同じ署名になる。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    swapped = design[["version", "choice_set_id", "alt_id", "brand", "price"]]
    assert pcc.design_signature(swapped) == pcc.design_signature(design)


def test_signature_requires_id_columns():
    """choice_set_id / alt_id を持たない DataFrame はエラー（日本語）。"""
    with pytest.raises(ValueError, match="choice_set_id"):
        pcc.design_signature(pd.DataFrame({"price": [1, 2]}))


def test_signature_ignores_pandas_index_column(tmp_path):
    """index=False を付け忘れて保存した CSV（Unnamed: 0 列入り）でも署名が一致する。

    行番号列は設計の中身ではないため、署名の計算から除外される。
    """
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    csv = tmp_path / "design_with_index.csv"
    design.to_csv(csv)                       # index=False を付け忘れたケース
    loaded = pd.read_csv(csv)
    assert "Unnamed: 0" in loaded.columns    # 行番号列が混入している
    assert pcc.design_signature(loaded) == pcc.design_signature(design)


def test_signature_independent_of_numpy_scalar_types():
    """値が同じなら、numpy スカラー列でも Python 値の列でも署名は一致する。

    回帰テスト：以前はセル値の repr をそのままハッシュしていたため、
    numpy 2.x（repr(np.int64(6)) == 'np.int64(6)'）と 1.x（'6'）で
    同じ設計の署名が食い違った。署名は時間・環境をまたいだ design の
    同一性確認に使うので、numpy のバージョンに依存してはならない。
    """
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    # 数値列を Python の int（object dtype）に変換した「値が同じ」設計
    as_python = design.copy()
    for c in ("version", "choice_set_id", "alt_id", "price"):
        as_python[c] = as_python[c].map(int).astype(object)
    assert pcc.design_signature(as_python) == pcc.design_signature(design)


# ---------------------------------------------------------------------------
# forms_to_data: 署名を出力に引き継ぐ
# ---------------------------------------------------------------------------

def _make_responses_xlsx(path, n_resp=3, n_sets=4, seed=0):
    rng = np.random.default_rng(seed)
    data = {
        "ID": range(1, n_resp + 1),
        "Start time": ["2026-06-01 10:00"] * n_resp,
        "Completion time": ["2026-06-01 10:05"] * n_resp,
        "Email": ["anonymous"] * n_resp,
        "Name": [""] * n_resp,
    }
    for q in range(n_sets):
        data[f"Q{q+1}. どの製品を選びますか？"] = list(
            rng.choice(["製品A", "製品B", "製品C"], size=n_resp)
        )
    pd.DataFrame(data).to_excel(path, index=False)


def test_forms_to_data_carries_design_signature(tmp_path):
    """forms_to_data の出力 attrs に、使った design の署名が入る。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_sets=4)
    df = pcc.forms_to_data(str(f), design, ["A", "B", "C"])
    assert df.attrs.get("design_signature") == pcc.design_signature(design)


def test_forms_to_data_signature_matches_survey_design(tmp_path):
    """アンケート作成時と同じ design（CSV 保存→読込）なら署名が一致する。"""
    survey_design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    csv = tmp_path / "design.csv"
    survey_design.to_csv(csv, index=False)

    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_sets=4)

    analysis_design = pd.read_csv(csv)            # 分析時は読み込むだけ
    df = pcc.forms_to_data(str(f), analysis_design, ["A", "B", "C"])
    # 出力の署名 == アンケート作成に使った design の署名
    assert df.attrs["design_signature"] == pcc.design_signature(survey_design)


def test_forms_to_data_warns_and_drops_index_column(tmp_path):
    """index=False を付け忘れた design CSV でも、警告のうえ正しく動く。

    行番号列（Unnamed: 0）は属性でないため出力から除外され、
    出力に引き継がれる署名も元の設計と一致する。
    """
    survey_design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    csv = tmp_path / "design.csv"
    survey_design.to_csv(csv)                     # index=False を付け忘れたケース

    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_sets=4)

    analysis_design = pd.read_csv(csv)
    with pytest.warns(UserWarning, match="index=False"):
        df = pcc.forms_to_data(str(f), analysis_design, ["A", "B", "C"])
    # 行番号列は属性として混入しない
    assert "Unnamed: 0" not in df.columns
    # 属性列は本来のものだけ
    assert set(df.columns) >= {"price", "brand"}
    # 署名は元の設計と一致する（行番号列は無視される）
    assert df.attrs["design_signature"] == pcc.design_signature(survey_design)


# ---------------------------------------------------------------------------
# 構造チェック（案B）: 崩れた design を日本語エラーで弾く
# ---------------------------------------------------------------------------

def test_forms_to_data_rejects_uneven_alt_counts(tmp_path):
    """選択セットごとの代替案数が揃っていない design はエラー（日本語）。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    # choice_set_id=1 の alt_id=3 を1行だけ削って代替案数を不揃いにする
    broken = design.drop(
        design[(design["choice_set_id"] == 1) & (design["alt_id"] == 3)].index
    )
    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_sets=4)
    with pytest.raises(ValueError, match="代替案数"):
        pcc.forms_to_data(str(f), broken, ["A", "B", "C"])


def test_forms_to_data_rejects_non_contiguous_alt_ids(tmp_path):
    """alt_id が 1..n_alts を網羅しない design はエラー（日本語）。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    broken = design.copy()
    # alt_id=3 をすべて 99 に置き換える（数は揃うが 1,2,3 を網羅しない）
    broken.loc[broken["alt_id"] == 3, "alt_id"] = 99
    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_sets=4)
    with pytest.raises(ValueError, match="網羅"):
        pcc.forms_to_data(str(f), broken, ["A", "B", "C"])


# ---------------------------------------------------------------------------
# 正常ケースでは警告を一切出さない（過剰検出しない）
# ---------------------------------------------------------------------------

@contextmanager
def warnings_as_errors():
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        yield


def test_normal_workflow_emits_no_warning(tmp_path):
    """同一 design を使う正常なワークフローでは警告が出ない。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=3, seed=42)
    f = tmp_path / "responses.xlsx"
    _make_responses_xlsx(f, n_resp=5, n_sets=4)
    with warnings_as_errors():
        pcc.forms_to_data(str(f), design, ["A", "B", "C"])
