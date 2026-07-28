"""suggest_n_profiles の印字と、check_design の案内文（auto_balance の存在を伝える）。

判定そのものの正しさは、末尾の一致テスト（内部の高速判定と check_design を
全設計で突き合わせる）で担保している。
"""

import itertools
import math
import time
import warnings

import numpy as np
import pandas as pd
import pytest

import py4conjoint.rating as pcr
from py4conjoint.rating import _feasibility
from py4conjoint.rating import analysis as analysis_module

ATTRS_2x2x2 = {
    "price": [6, 10],
    "os": ["android", "apple"],
    "camera": ["標準", "高性能"],
}
ATTRS_3x2x2 = {
    "price": [6, 8, 10],
    "os": ["android", "apple"],
    "camera": ["標準", "高性能"],
}
ATTRS_2x2x2x2 = {
    "price": [6, 10],
    "os": ["android", "apple"],
    "camera": ["標準", "高性能"],
    "battery": ["標準", "大容量"],
}
ATTRS_3x3x3 = {
    "price": [6, 8, 10],
    "camera": ["標準", "高性能", "超高性能"],
    "storage": [64, 128, 256],
}
# 候補 24 通り。n=7 なら C(24,7) = 346,104 で、報告用の上限（100,000）は
# 超えるが design_profiles の総当たり上限（1,000,000）には収まる。
ATTRS_4x3x2 = {
    "price": [6, 8, 10, 12],
    "camera": ["標準", "高性能", "超高性能"],
    "os": ["android", "apple"],
}


# ---------------------------------------------------------------------------
# 1. 返り値の DataFrame は変わっていない（印字を足しただけ）
# ---------------------------------------------------------------------------


def test_suggest_n_profiles_dataframe_is_unchanged():
    """列・行・値・attrs のいずれも従来どおりであること。"""
    out = pcr.suggest_n_profiles(ATTRS_2x2x2, n_respondents=30)

    assert list(out.columns) == [
        "回答者数",
        "obs/pred≥10（最低限）",
        "推奨 n_profiles",
        "obs/pred（達成）",
        "観測数 obs",
    ]
    assert len(out) == 1
    row = out.iloc[0]
    assert row["回答者数"] == 30
    assert row["obs/pred≥10（最低限）"] == 4
    assert row["推奨 n_profiles"] == 6
    assert row["obs/pred（達成）"] == 60.0
    assert row["観測数 obs"] == 180
    assert out.attrs == {
        "n_params": 4,
        "n_encoded": 3,
        "n_candidates": 8,
        "m_min": 4,
        "m_orme": 6,
    }


def test_suggest_n_profiles_default_table_is_unchanged():
    """回答者数を省略したときの行数・回答者数の並びも従来どおり。"""
    out = pcr.suggest_n_profiles(ATTRS_2x2x2)
    assert list(out["回答者数"]) == [5, 10, 20, 30, 50, 100]
    assert out.attrs["n_candidates"] == 8


# ---------------------------------------------------------------------------
# 2〜5. 印字される「バランスと D 最適性が両立するか」
# ---------------------------------------------------------------------------


def test_prints_incompatible_with_ratio_2x2x2(capsys):
    """2×2×2 の推奨 n（6）は両立せず、比 75% が印字される。"""
    pcr.suggest_n_profiles(ATTRS_2x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "推奨 n_profiles: 6" in printed
    assert "n = 6 では、D 最適な設計は必ずどこかの水準が不均衡になります。" in printed
    assert "auto_balance=True" in printed
    assert "75%" in printed
    assert "n_profiles を変えると両立する場合があります。" in printed


def test_prints_incompatible_with_ratio_3x2x2(capsys):
    """3×2×2 の推奨 n（8）は両立せず、比 93.75% が印字される。"""
    pcr.suggest_n_profiles(ATTRS_3x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "推奨 n_profiles: 8" in printed
    assert "n = 8 では、D 最適な設計は必ずどこかの水準が不均衡になります。" in printed
    assert "93.75%" in printed


def test_prints_compatible_case(capsys):
    """2×2×2×2 の推奨 n（8）は両立する（auto_balance は不要）。"""
    pcr.suggest_n_profiles(ATTRS_2x2x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "推奨 n_profiles: 8" in printed
    assert "n = 8 なら、D 最適な設計がそのまま水準バランスも満たします" in printed
    assert "auto_balance=True を指定する必要はありません" in printed
    # 「両立しない」側の文言が混ざっていないこと
    assert "不均衡になります" not in printed


def test_does_not_judge_when_too_many_combinations(capsys):
    """候補が上限を超える構成では判定せず、しかも即座に返る。"""
    start = time.perf_counter()
    pcr.suggest_n_profiles(ATTRS_3x3x3, n_respondents=30)
    elapsed = time.perf_counter() - start
    printed = capsys.readouterr().out

    assert "確認していません" in printed
    assert "17,383,860 通り" in printed
    assert 'df.attrs["auto_balance"]["det_ratio"]' in printed
    # 判定していないことの担保（17,383,860 通りを列挙すれば十数秒かかる）
    assert elapsed < 2.0


# ---------------------------------------------------------------------------
# 6〜8. check_design の案内文
# ---------------------------------------------------------------------------


def test_balance_warning_mentions_auto_balance():
    """2×2×2 の既定 6 プロファイル設計：バランス警告に auto_balance の案内が付く。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, seed=1)
    diags = pcr.check_design(design).diagnostics
    balance = [d for d in diags if d.category.startswith("balance_")]

    assert balance, "この設計ではバランス警告が出るはず"
    for d in balance:
        assert "auto_balance=True" in d.hint
        assert "バランスの取れた設計が存在します" in d.hint


def test_correlation_warning_is_unavoidable_at_this_n():
    """auto_balance=True の設計：バランス警告は消え、相関は回避不可能と案内される。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, auto_balance=True, seed=1)
    diags = pcr.check_design(design).diagnostics

    assert not [d for d in diags if d.category.startswith("balance_")]
    corr = [d for d in diags if d.category.startswith("correlation_")]
    assert corr, "相関の指摘は残るはず"
    for d in corr:
        assert "どのプロファイルの組み合わせを選んでも相関の指摘は残ります" in d.hint
        assert "auto_balance=True は水準バランス専用" in d.hint


def test_severity_is_unchanged():
    """案内文が変わっても重大度は従来どおり（[大] / [中]）。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, seed=1)
    by_category = {d.category: d.severity for d in pcr.check_design(design).diagnostics}

    assert by_category["balance_os"] == "大"
    assert by_category["correlation_camera_price"] == "中"

    balanced = pcr.design_profiles(ATTRS_2x2x2, 6, auto_balance=True, seed=1)
    for d in pcr.check_design(balanced).diagnostics:
        if d.category.startswith("correlation_"):
            assert d.severity == "中"


# ---------------------------------------------------------------------------
# 9. 「回避できるが、バランスとは両立しない」状態（4つ目の文言）
# ---------------------------------------------------------------------------


def test_correlation_avoidable_but_incompatible_with_balance():
    """3×2×2 の n=6：相関だけなら避けられるが、バランスと同時には満たせない。

    2×2×2 の n=6（相関の指摘が出ない設計が0件）とは別の状態であり、
    両者を取り違えていないことの担保でもある。
    """
    design = pcr.design_profiles(ATTRS_3x2x2, 6, auto_balance=True, seed=1)
    judgment = _feasibility.judge_avoidability(design, list(ATTRS_3x2x2))

    # 判定の中身：相関だけなら避けられる／バランスとの両立はできない
    assert judgment["correlation_avoidable"] is True
    assert judgment["both_avoidable"] is False

    corr = [
        d
        for d in pcr.check_design(design).diagnostics
        if d.category.startswith("correlation_")
    ]
    assert corr, "この設計では相関の指摘が残るはず"
    for d in corr:
        assert "相関の指摘が出ない組み合わせは存在しますが" in d.hint
        assert "その設計は水準バランスを満たしません" in d.hint
        assert "両方を満たす設計が存在しないためです" in d.hint
        # 「この n では相関を避けられない」という別状態の文言が出ていないこと
        assert (
            "どのプロファイルの組み合わせを選んでも相関の指摘は残ります" not in d.hint
        )


def test_2x2x2_n6_is_the_other_state():
    """比較対象：2×2×2 の n=6 は「相関の指摘が出ない設計が存在しない」側。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, auto_balance=True, seed=1)
    judgment = _feasibility.judge_avoidability(design, list(ATTRS_2x2x2))

    assert judgment["correlation_avoidable"] is False
    assert judgment["both_avoidable"] is False


# ---------------------------------------------------------------------------
# 10. 一致テスト：ベクトル化した判定と check_design が全設計で一致すること
# ---------------------------------------------------------------------------


def test_fast_judgment_agrees_with_check_design():
    """2×2×2 の全 C(8, 6) = 28 設計で、内部の高速判定と check_design を突き合わせる。

    存在判定は check_design を呼ばずにベクトル化した式で警告の有無を求めている。
    その式が正典（check_design）と一致していることを担保する。閾値を変えたときに、
    両者がずれればこのテストが落ちる。
    """
    names = list(ATTRS_2x2x2)
    full = pd.DataFrame(
        [dict(zip(names, combo)) for combo in itertools.product(*ATTRS_2x2x2.values())]
    )
    indicator, spans = _feasibility._design._level_indicator(full, ATTRS_2x2x2)
    _attr_of_col, pair_mask = _feasibility._encoded_layout(ATTRS_2x2x2)

    combos = list(itertools.combinations(range(len(full)), 6))
    idx = np.array(combos, dtype=np.int64)
    levels_ok, balance_ok, correlation_ok = _feasibility._batch_flags(
        idx, indicator, spans, pair_mask, 6
    )

    for i, combo in enumerate(combos):
        design = full.iloc[list(combo)].copy()
        design.index = [f"P{j + 1}" for j in range(len(design))]
        diags = pcr.check_design(design).diagnostics

        assert bool(levels_ok[i]) is all(
            design[a].nunique() == len(lv) for a, lv in ATTRS_2x2x2.items()
        )
        for a, attr in enumerate(names):
            warned = any(d.category == f"balance_{attr}" for d in diags)
            assert bool(balance_ok[i, a]) is (not warned), (combo, attr)
        warned_corr = any(d.category.startswith("correlation_") for d in diags)
        assert bool(correlation_ok[i]) is (not warned_corr), combo


def test_fast_judgment_agrees_with_check_design_three_levels():
    """3水準を含む構成でも一致すること（基準水準＝最頻値の再現の担保）。

    基準水準を固定水準で代用すると、3水準を含む構成では判定がずれる
    （実測で数%の設計が食い違った）。このテストがその代用を検出する。

    比較するのは水準がすべて現れる設計だけ。水準が欠けた設計では
    check_design が観測された水準だけで CV を計算するのに対し、高速判定は
    出現回数 0 も数に入れるため一致しない（そのような設計は候補から
    除くので、一致している必要がない）。
    """
    names = list(ATTRS_3x2x2)
    full = pd.DataFrame(
        [dict(zip(names, combo)) for combo in itertools.product(*ATTRS_3x2x2.values())]
    )
    indicator, spans = _feasibility._design._level_indicator(full, ATTRS_3x2x2)
    _attr_of_col, pair_mask = _feasibility._encoded_layout(ATTRS_3x2x2)

    combos = list(itertools.combinations(range(len(full)), 6))
    idx = np.array(combos, dtype=np.int64)
    levels_ok, balance_ok, correlation_ok = _feasibility._batch_flags(
        idx, indicator, spans, pair_mask, 6
    )

    compared = 0
    for i, combo in enumerate(combos):
        design = full.iloc[list(combo)].copy()
        design.index = [f"P{j + 1}" for j in range(len(design))]
        complete = all(design[a].nunique() == len(lv) for a, lv in ATTRS_3x2x2.items())
        assert bool(levels_ok[i]) is complete, combo
        if not complete:
            continue
        diags = pcr.check_design(design).diagnostics
        for a, attr in enumerate(names):
            warned = any(d.category == f"balance_{attr}" for d in diags)
            assert bool(balance_ok[i, a]) is (not warned), (combo, attr)
        warned_corr = any(d.category.startswith("correlation_") for d in diags)
        assert bool(correlation_ok[i]) is (not warned_corr), combo
        compared += 1

    # 3水準の属性があるので、基準水準（最頻値）の再現が効いている設計が含まれる
    assert compared >= 300


# ---------------------------------------------------------------------------
# 判定しない場合（候補数が大きい）は従来の文言のまま
# ---------------------------------------------------------------------------


def test_check_design_keeps_original_wording_when_not_judged(monkeypatch):
    """判定しないときは案内文を足さない（従来どおり）。"""
    monkeypatch.setattr(_feasibility._design, "_REPORT_MAX_COMBINATIONS", 0)
    design = pcr.design_profiles(ATTRS_2x2x2, 6, seed=1)
    diags = pcr.check_design(design).diagnostics

    balance = [d for d in diags if d.category.startswith("balance_")]
    assert balance
    for d in balance:
        assert d.recommendation == (
            "各水準の出現回数を均等にしてください（バランスの良いデザイン）。"
        )
        assert d.hint == ""
    for d in diags:
        if d.category.startswith("correlation_"):
            assert d.recommendation == (
                "可能であればプロファイルの組み合わせを調整してください。"
            )
            assert d.hint == ""


def test_judgment_is_skipped_for_single_level_attribute():
    """水準が1つしかない属性がある設計は判定しない（候補空間を復元できない）。"""
    design = pd.DataFrame(
        {
            "price": [6, 6, 6, 6],
            "os": ["android", "android", "apple", "apple"],
            "camera": ["標準", "高性能", "標準", "高性能"],
        }
    )
    assert _feasibility.judge_avoidability(design, list(design.columns)) is None
    # 例外にはならず、警告は従来どおり出る
    assert pcr.check_design(design).diagnostics is not None


@pytest.mark.parametrize("n_profiles", [4, 5, 6, 7])
def test_judgment_runs_for_all_reasonable_sizes(n_profiles):
    """判定が例外を出さずに返ること（n の大小によらず）。"""
    design = pcr.design_profiles(ATTRS_2x2x2, n_profiles, seed=1)
    judgment = _feasibility.judge_avoidability(design, list(ATTRS_2x2x2))
    assert judgment is not None
    assert judgment["n_profiles"] == n_profiles
    assert set(judgment["balance_avoidable"]) == set(ATTRS_2x2x2)


def test_hint_wording_for_unreached_states():
    """走査では現れなかった分岐の文言（防御的に残してある経路）。

    属性構成5通り × 妥当な n をすべて走査したが、「バランスと相関の両方を
    満たす設計が存在するのに auto_balance の設計には相関が残る」状態と、
    「バランス警告を避けられるのに auto_balance の設計でも残る」状態は
    観測されなかった。到達しにくいだけで不可能とは限らないため分岐は残し、
    文言だけをここで確認する。
    """
    judgment = {
        "n_profiles": 6,
        "balance_avoidable": {"price": True},
        "correlation_avoidable": True,
        "both_avoidable": True,
        "auto_balance_balance": {"price": True},
        "auto_balance_correlation": True,
    }
    corr_hint = analysis_module._correlation_hint(judgment)
    assert "両方を満たす設計が存在しますが" in corr_hint
    assert "det(X'X) の最大化を優先するためそれを選びません" in corr_hint
    # 「両立する設計が存在しない」という別状態の断定が混ざっていないこと
    assert "存在しないためです" not in corr_hint

    balance_hint = analysis_module._balance_hint("price", judgment)
    assert "auto_balance=True の設計でもこの偏りは残ります" in balance_hint
    assert "n_profiles を変えると解消する場合があります" in balance_hint


def test_analysis_module_has_single_source_of_thresholds():
    """存在判定は analysis の定数を参照している（数値の写し取りがないこと）。"""
    assert _feasibility._analysis is analysis_module
    assert analysis_module._BALANCE_CV_MINOR == 0.15
    assert analysis_module._CORRELATION_ABS_MINOR == 0.3


# ---------------------------------------------------------------------------
# 上限は2つある（報告する関数は低いほうを使う）
# ---------------------------------------------------------------------------


def test_report_limit_is_lower_than_exhaustive_limit():
    """報告用の上限は design_profiles の総当たり上限より低い。"""
    assert _feasibility._design._REPORT_MAX_COMBINATIONS == 100_000
    assert _feasibility._design._EXHAUSTIVE_MAX_COMBINATIONS == 1_000_000


def test_check_design_does_not_judge_between_the_two_limits():
    """C が 2つの上限の間にある構成では、check_design は判定せず即座に返る。"""
    n_profiles = 7
    assert (
        _feasibility._design._REPORT_MAX_COMBINATIONS
        < math.comb(24, n_profiles)
        <= _feasibility._design._EXHAUSTIVE_MAX_COMBINATIONS
    )

    design = pcr.design_profiles(ATTRS_4x3x2, n_profiles, seed=1)
    start = time.perf_counter()
    diags = pcr.check_design(design).diagnostics
    elapsed = time.perf_counter() - start

    assert _feasibility.judge_avoidability(design, list(ATTRS_4x3x2)) is None
    judged = [d for d in diags if d.category.startswith(("balance_", "correlation_"))]
    assert judged, "この設計では警告が出るはず"
    for d in judged:
        assert d.hint == ""
    # 列挙していれば数秒かかる規模なので、短時間で返ること自体が担保になる
    assert elapsed < 1.0


def test_design_profiles_still_enumerates_between_the_two_limits():
    """同じ構成でも design_profiles(auto_balance=True) は従来どおり総当たりする。"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        design = pcr.design_profiles(ATTRS_4x3x2, 7, auto_balance=True, seed=1)
    assert design.attrs["auto_balance"]["method"] == "exhaustive"


# ---------------------------------------------------------------------------
# 案内文は recommendation ではなく hint に入る（warnings() が読める形になる）
# ---------------------------------------------------------------------------


def test_warnings_has_hint_column_and_recommendation_is_unchanged():
    """warnings() は hint を独立した列で返し、recommendation は従来のまま。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, seed=1)
    w = pcr.check_design(design).warnings()

    assert list(w.columns) == [
        "severity",
        "category",
        "message",
        "recommendation",
        "hint",
    ]

    balance = w[w["category"].str.startswith("balance_")]
    assert not balance.empty
    for _, row in balance.iterrows():
        assert row["recommendation"] == (
            "各水準の出現回数を均等にしてください（バランスの良いデザイン）。"
        )
        assert "auto_balance=True" in row["hint"]

    correlation = w[w["category"].str.startswith("correlation_")]
    assert not correlation.empty
    for _, row in correlation.iterrows():
        assert row["recommendation"] == (
            "可能であればプロファイルの組み合わせを調整してください。"
        )
        assert row["hint"]

    # セルは1行に連結されていて、改行も字下げも混ざらない
    for value in w["hint"]:
        assert "\n" not in value
        assert "  " not in value


def test_hint_is_indented_in_text_display():
    """テキスト表示では従来どおり recommendation の下に字下げして続く。"""
    design = pcr.design_profiles(ATTRS_2x2x2, 6, seed=1)
    text = pcr.check_design(design).summary()

    assert (
        "      → 各水準の出現回数を均等にしてください（バランスの良いデザイン）。"
        in text
    )
    assert "\n        この n（6）ならバランスの取れた設計が存在します。" in text
    # エスケープされた改行がそのまま出ていないこと
    assert "\\n" not in text
