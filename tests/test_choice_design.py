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
