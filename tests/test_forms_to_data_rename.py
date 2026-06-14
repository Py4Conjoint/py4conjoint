"""Forms 変換関数の統一改名（forms_to_data）の検証。

- rating / choice の両サブパッケージが **同じ関数名** ``forms_to_data`` を
  公開していること。
- 旧名（``forms_to_conjoint_data`` / ``cbc_forms_to_data``）が
  サブパッケージから消えていること。
- rating の第2引数が ``profiles`` であり、旧名 ``attributes`` で渡すと
  エラーになること。
- rating の出力 DataFrame の列名が英語（respondent_id, profile_id, rating）で
  あること。
- choice の出力列名（respondent_id, choice_set_id, choice, alt）は不変であること。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np
import pandas as pd
import pytest

import py4conjoint.choice as pcc
import py4conjoint.rating as pcr

PROFILES = {
    "price":  [6, 10, 6, 10],
    "os":     ["android", "apple", "apple", "android"],
    "camera": ["標準", "標準", "高性能", "高性能"],
}


def _make_rating_microsoft_xlsx(path, n_resp=6):
    """Microsoft Forms 形式の回答ファイル（4プロファイル評点）を作る。"""
    data = {
        "ID": range(1, n_resp + 1),
        "Start time": ["2026-06-01 10:00"] * n_resp,
        "Completion time": ["2026-06-01 10:05"] * n_resp,
        "Email": ["anonymous"] * n_resp,
        "Name": [""] * n_resp,
    }
    for p in range(4):
        data[f"製品案{p+1}を何点で評価しますか？"] = [
            ((p + r) % 7) + 1 for r in range(n_resp)
        ]
    pd.DataFrame(data).to_excel(path, index=False)


# ---------------------------------------------------------------------------
# 関数名の統一（forms_to_data）と旧名の削除
# ---------------------------------------------------------------------------

def test_rating_forms_to_data_is_public():
    assert callable(pcr.forms_to_data)
    assert "forms_to_data" in pcr.__all__


def test_choice_forms_to_data_is_public():
    assert callable(pcc.forms_to_data)
    assert "forms_to_data" in pcc.__all__


def test_old_rating_function_name_removed():
    # 旧名 forms_to_conjoint_data はサブパッケージから削除済み
    assert not hasattr(pcr, "forms_to_conjoint_data")
    assert "forms_to_conjoint_data" not in pcr.__all__


def test_old_choice_function_name_removed():
    # 旧名 cbc_forms_to_data はサブパッケージから削除済み
    assert not hasattr(pcc, "cbc_forms_to_data")
    assert "cbc_forms_to_data" not in pcc.__all__


# ---------------------------------------------------------------------------
# rating: 第2引数は profiles（旧名 attributes は不可）
# ---------------------------------------------------------------------------

def test_rating_attributes_kwarg_rejected(tmp_path):
    f = tmp_path / "responses.xlsx"
    _make_rating_microsoft_xlsx(f)
    # 旧キーワード attributes は廃止されており TypeError になる
    with pytest.raises(TypeError):
        pcr.forms_to_data(str(f), attributes=PROFILES)


def test_rating_profiles_kwarg_works(tmp_path):
    f = tmp_path / "responses.xlsx"
    _make_rating_microsoft_xlsx(f)
    df = pcr.forms_to_data(str(f), profiles=PROFILES)
    assert len(df) > 0


# ---------------------------------------------------------------------------
# rating の出力列名は英語（respondent_id, profile_id, rating）
# ---------------------------------------------------------------------------

def test_rating_output_columns_are_english(tmp_path):
    f = tmp_path / "responses.xlsx"
    _make_rating_microsoft_xlsx(f)
    df = pcr.forms_to_data(str(f), PROFILES)
    assert "respondent_id" in df.columns
    assert "profile_id" in df.columns
    assert "rating" in df.columns
    # 旧い日本語の列名は残っていない
    assert "回答者ID" not in df.columns
    assert "プロファイルID" not in df.columns


def test_rating_fit_uses_english_default(tmp_path):
    """英語列のまま fit() を呼ぶと、デフォルトで回答者IDが認識される。"""
    f = tmp_path / "responses.xlsx"
    _make_rating_microsoft_xlsx(f, n_resp=12)
    df = pcr.forms_to_data(str(f), PROFILES)
    coded = pcr.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pcr.fit(coded)
    # respondent_id 列が既定で認識され、クラスタロバストSEになる
    assert result.respondent_id_col == "respondent_id"
    assert result.se_type == "cluster"


# ---------------------------------------------------------------------------
# choice の出力列名は不変（respondent_id, choice_set_id, choice, alt）
# ---------------------------------------------------------------------------

def test_choice_output_columns_unchanged(tmp_path):
    design = pcc.design_choice_sets(
        {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
        n_sets=4, n_alts=3, seed=42,
    )
    n = 3
    data = {
        "ID": range(1, n + 1),
        "Start time": ["2026-06-01 10:00"] * n,
        "Completion time": ["2026-06-01 10:05"] * n,
        "Email": ["anonymous"] * n,
        "Name": [""] * n,
    }
    for q in range(4):
        data[f"Q{q+1}. どの製品を選びますか？"] = ["製品A", "製品B", "製品C"]
    f = tmp_path / "responses.xlsx"
    pd.DataFrame(data).to_excel(f, index=False)
    df = pcc.forms_to_data(str(f), design, ["A", "B", "C"])
    for col in ("respondent_id", "choice_set_id", "choice", "alt"):
        assert col in df.columns, f"{col} が出力に無い"


# ---------------------------------------------------------------------------
# choice: forms_to_data → fit を引数省略で通すと既定列名が一致して動く
# ---------------------------------------------------------------------------

def test_choice_fit_defaults_match_forms_output(tmp_path):
    """forms_to_data() → fit() を列名引数なしで通すと、出力列名（respondent_id /
    choice_set_id）が fit() の既定と一致し、クラスタロバストSEが適用される
    （independence_assumed 警告が出ない）。"""
    design = pcc.design_choice_sets(
        {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
        n_sets=4, n_alts=3, seed=42,
    )
    rng = np.random.default_rng(0)
    n_resp = 15
    data = {
        "ID": range(1, n_resp + 1),
        "Start time": ["2026-06-01 10:00"] * n_resp,
        "Completion time": ["2026-06-01 10:05"] * n_resp,
        "Email": ["anonymous"] * n_resp,
        "Name": [""] * n_resp,
    }
    for q in range(4):
        data[f"Q{q+1}. どの製品を選びますか？"] = list(
            rng.choice(["製品A", "製品B", "製品C"], size=n_resp)
        )
    f = tmp_path / "responses.xlsx"
    pd.DataFrame(data).to_excel(f, index=False)

    df = pcc.forms_to_data(str(f), design, ["A", "B", "C"])
    coded = pcc.encode(df, reference_levels={"brand": "A社"})

    # 列名引数（choice_set_id_col / respondent_id_col）を一切渡さない
    result = pcc.fit(coded)

    # 既定値が forms_to_data の出力列名と一致している
    assert result.choice_set_id_col == "choice_set_id"
    assert result.respondent_id_col == "respondent_id"
    # 回答者15人 → 回答者IDが既定で認識され、クラスタロバストSEが適用される
    assert result.se_type == "cluster"
    # 回答者ID列が認識されたので独立性仮定の警告は出ない
    cats = list(result.warnings()["category"].values)
    assert "independence_assumed" not in cats
