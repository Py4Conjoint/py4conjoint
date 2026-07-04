"""choice/design.py（選択セット設計）のテスト。

* design_choice_sets : 形状・セット内重複なし・再現性・エラー処理
* check_design       : 診断結果の構造・警告検出
* suggest_n_respondents : Johnson-Orme の経験則の計算
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd
import pytest

import py4conjoint.choice as pcc

ATTRS = {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]}


# ---------------------------------------------------------------------------
# design_choice_sets
# ---------------------------------------------------------------------------

def test_design_shape_and_columns():
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    assert list(design.columns) == ["version", "choice_set_id", "alt_id", "price", "brand"]
    assert len(design) == 1 * 8 * 3
    assert design["version"].unique().tolist() == [1]
    assert sorted(design["choice_set_id"].unique()) == list(range(1, 9))
    assert sorted(design["alt_id"].unique()) == [1, 2, 3]
    assert design.attrs["n_candidates"] == 9  # 3水準 × 3水準


def test_design_versions():
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=2,
                                    n_versions=3, seed=0)
    assert len(design) == 3 * 4 * 2
    assert sorted(design["version"].unique()) == [1, 2, 3]


def test_design_no_duplicate_profiles_within_set():
    design = pcc.design_choice_sets(ATTRS, n_sets=50, n_alts=3, seed=1)
    for _, block in design.groupby(["version", "choice_set_id"]):
        profiles = block[["price", "brand"]].apply(tuple, axis=1)
        assert profiles.nunique() == len(block), \
            "同一選択セット内にプロファイルの重複があります"


def test_design_reproducible_with_seed():
    d1 = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    d2 = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    pd.testing.assert_frame_equal(d1, d2)


def test_design_validation_errors():
    with pytest.raises(ValueError, match="attributes"):
        pcc.design_choice_sets({}, n_sets=8, n_alts=3)
    with pytest.raises(ValueError, match="水準数は 2 以上"):
        pcc.design_choice_sets({"price": [100]}, n_sets=8, n_alts=3)
    with pytest.raises(ValueError, match="n_alts"):
        pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=1)
    with pytest.raises(ValueError, match="n_sets"):
        pcc.design_choice_sets(ATTRS, n_sets=0, n_alts=3)
    with pytest.raises(ValueError, match="n_versions"):
        pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, n_versions=0)
    # n_alts > 完全交差の候補数（2×2=4）
    with pytest.raises(ValueError, match="完全交差の候補数"):
        pcc.design_choice_sets(
            {"a": [1, 2], "b": [1, 2]}, n_sets=4, n_alts=5
        )
    # auto_balance=True で n_candidates < 1
    with pytest.raises(ValueError, match="n_candidates"):
        pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3,
                               auto_balance=True, n_candidates=0)


# ---------------------------------------------------------------------------
# design_choice_sets: auto_balance（バランスの良い設計を自動選択）
# ---------------------------------------------------------------------------

# 警告が出やすい小さめのスマホ設計（auto_balance の効果が見える）
SMARTPHONE = {"price": [6, 10], "os": ["apple", "android"],
              "camera": ["標準", "高性能"]}


def _cv_sum_and_warnings(design):
    chk = pcc.check_design(design)
    return float(chk.balance["CV"].sum()), len(chk.diagnostics)


def test_auto_balance_default_is_false_and_unchanged():
    """auto_balance を指定しなければ従来と完全に同一の設計（後方互換）。"""
    d_plain = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=0)
    d_default = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=0,
                                       auto_balance=False)
    pd.testing.assert_frame_equal(d_plain, d_default)
    # 従来の attrs はそのまま（auto_balance の来歴は付かない）
    assert "auto_balance" not in d_plain.attrs
    assert d_plain.attrs["n_candidates"] == 8  # 2×2×2


def test_auto_balance_not_worse_than_single_seed():
    """auto_balance=True は、典型的な単一 seed より悪くならない。"""
    single = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=0)
    auto = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=0,
                                  auto_balance=True, n_candidates=200)
    cv_s, w_s = _cv_sum_and_warnings(single)
    cv_a, w_a = _cv_sum_and_warnings(auto)
    assert w_a <= w_s        # 警告数は同等以下
    assert cv_a <= cv_s + 1e-9  # CV 合計も同等以下


def test_auto_balance_reproducible_with_seed():
    """同じ seed・引数なら auto_balance でも毎回同じ設計（同じ署名）。"""
    d1 = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=7,
                                auto_balance=True, n_candidates=50)
    d2 = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=7,
                                auto_balance=True, n_candidates=50)
    pd.testing.assert_frame_equal(d1, d2)
    assert d1.attrs["design_signature"] == d2.attrs["design_signature"]


def test_auto_balance_records_provenance():
    """選定の来歴が attrs に入る（候補数・警告数・CV 合計）。"""
    d = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=1,
                               auto_balance=True, n_candidates=100)
    prov = d.attrs["auto_balance"]
    assert prov["n_candidates"] == 100
    cv_sum, n_warn = _cv_sum_and_warnings(d)
    # 来歴の値が実際の設計の診断と一致する（表示と実態の一致）
    assert prov["n_warnings"] == n_warn
    assert abs(prov["cv_sum"] - cv_sum) < 1e-6


def test_auto_balance_small_n_candidates_ok():
    """小さい n_candidates でもエラーなく動く。"""
    d = pcc.design_choice_sets(SMARTPHONE, n_sets=6, n_alts=3, seed=2,
                               auto_balance=True, n_candidates=1)
    assert len(d) == 6 * 3
    assert d.attrs["auto_balance"]["n_candidates"] == 1


# ---------------------------------------------------------------------------
# check_design
# ---------------------------------------------------------------------------

def test_check_design_structure():
    design = pcc.design_choice_sets(ATTRS, n_sets=30, n_alts=3,
                                    n_versions=2, seed=42)
    res = pcc.check_design(design)
    assert isinstance(res, pcc.ChoiceDesignCheckResult)
    # balance: 属性ごとに1行
    assert sorted(res.balance.index) == ["brand", "price"]
    assert "CV" in res.balance.columns
    # chi2: 属性ペア1組
    assert len(res.chi2) == 1
    assert "χ²/自由度" in res.chi2.columns
    # overlap: 属性ごとに1行、率は 0〜1
    assert sorted(res.overlap.index) == ["brand", "price"]
    assert ((res.overlap["オーバーラップ率"] >= 0)
            & (res.overlap["オーバーラップ率"] <= 1)).all()
    # 和文サマリー
    text = res.summary()
    assert "選択セット設計チェック" in text
    assert "水準バランス" in text
    assert "セット内オーバーラップ" in text
    # warnings() は DataFrame
    w = res.warnings()
    assert list(w.columns) == ["severity", "category", "message",
                               "recommendation"]


def test_check_design_detects_overlap():
    """全代替案が同じ水準を持つ設計ではオーバーラップ警告が出る。"""
    # brand が全セットで同一水準になるよう手作りする
    design = pd.DataFrame({
        "version": [1] * 8,
        "choice_set_id":  [1, 1, 2, 2, 3, 3, 4, 4],
        "alt_id":  [1, 2] * 4,
        "price":   [100, 150, 150, 200, 100, 200, 150, 100],
        "brand":   ["A社", "A社", "B社", "B社", "A社", "A社", "B社", "B社"],
    })
    res = pcc.check_design(design)
    cats = [d.category for d in res.diagnostics]
    assert "overlap_brand" in cats
    overlap_diags = [d for d in res.diagnostics
                     if d.category == "overlap_brand"]
    assert overlap_diags[0].severity == "大"


def test_check_design_errors():
    with pytest.raises(TypeError, match="DataFrame"):
        pcc.check_design([1, 2, 3])
    design = pcc.design_choice_sets(ATTRS, n_sets=4, n_alts=2, seed=0)
    with pytest.raises(ValueError, match="存在しません"):
        pcc.check_design(design, attributes=["存在しない属性"])


def test_check_design_ignores_pandas_index_column(tmp_path):
    """index=False を付け忘れた design CSV でも、行番号列を属性として診断しない。"""
    design = pcc.design_choice_sets(ATTRS, n_sets=8, n_alts=3, seed=42)
    csv = tmp_path / "design.csv"
    design.to_csv(csv)                        # index=False を付け忘れたケース
    loaded = pd.read_csv(csv)
    assert "Unnamed: 0" in loaded.columns
    res = pcc.check_design(loaded)
    # 行番号列は診断対象にならず、本来の属性だけが並ぶ
    assert sorted(res.balance.index) == ["brand", "price"]
    assert "Unnamed: 0" not in res.balance.index


# ---------------------------------------------------------------------------
# suggest_n_respondents
# ---------------------------------------------------------------------------

def test_suggest_n_respondents_johnson_orme(capsys):
    # c=3, t=8, a=3 → n ≥ 500*3/(8*3) = 62.5 → 63
    res = pcc.suggest_n_respondents(ATTRS, n_sets=8, n_alts=3)
    assert res.attrs["c_max"] == 3
    assert res.attrs["n_required"] == 63
    assert res.loc["price", "必要回答者数（目安）"] == 63
    assert res.loc["brand", "必要回答者数（目安）"] == 63
    # 和文の説明メッセージが表示される
    out = capsys.readouterr().out
    assert "Johnson-Orme" in out
    assert "63 人以上" in out


def test_suggest_n_respondents_uses_max_levels():
    # 水準数が異なる場合、必要回答者数は最大水準数 c で決まる
    attrs = {"price": [100, 150], "brand": ["A", "B", "C", "D"]}
    # c=4, t=10, a=2 → 500*4/20 = 100
    res = pcc.suggest_n_respondents(attrs, n_sets=10, n_alts=2)
    assert res.attrs["n_required"] == 100
    assert res.loc["price", "必要回答者数（目安）"] == 50
    assert res.loc["brand", "必要回答者数（目安）"] == 100


def test_suggest_n_respondents_errors():
    with pytest.raises(ValueError, match="attributes"):
        pcc.suggest_n_respondents({}, n_sets=8, n_alts=3)
    with pytest.raises(ValueError, match="n_sets"):
        pcc.suggest_n_respondents(ATTRS, n_sets=0, n_alts=3)
    with pytest.raises(ValueError, match="n_alts"):
        pcc.suggest_n_respondents(ATTRS, n_sets=8, n_alts=1)


# ---------------------------------------------------------------------------
# 統合：design → check_design がきれいな設計で警告を出さない
# ---------------------------------------------------------------------------

def test_large_random_design_is_clean():
    design = pcc.design_choice_sets(ATTRS, n_sets=100, n_alts=3,
                                    n_versions=5, seed=42)
    res = pcc.check_design(design)
    severities = [d.severity for d in res.diagnostics]
    assert "大" not in severities, \
        f"十分大きいランダム設計で重大警告が出ました: {res.warnings()}"
