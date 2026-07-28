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
    # 具体的な n を列挙するので、方向を示さない一文は出さない
    assert "n_profiles を変えると両立する場合があります。" not in printed


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
# 印字される「どの n なら指摘が出ないか」
# ---------------------------------------------------------------------------


def test_prints_which_n_is_clean_2x2x2(capsys):
    """2×2×2：指摘が出ないのは n = 4, 8（推奨の 6 ではない）。

    答えが推奨値より**小さい側**にもあることを、実際に走査して示す。
    """
    pcr.suggest_n_profiles(ATTRS_2x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "水準バランスと属性間相関の指摘がどちらも出ないのは n = 4, 8 です" in printed
    assert "design_profiles(..., auto_balance=True) で作った場合" in printed


def test_prints_what_remains_at_the_smallest_clean_n(capsys):
    """n = 4 で残るのは別カテゴリ（プロファイル数）の指摘だと書いてあること。

    「この n にすれば警告が消えます」とだけ書くと、消えない警告が出たときに
    読んだ人が混乱する。
    """
    pcr.suggest_n_profiles(ATTRS_2x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "ただし n = 4 では、プロファイル数がパラメータ数（4）に対して" in printed
    assert "少ないという指摘は残ります" in printed
    assert "バランス・相関とは別のカテゴリの警告です" in printed


def test_prints_which_n_is_clean_3x2x2(capsys):
    """3×2×2：指摘が出ないのは n = 12（＝全候補）だけ。残る指摘の注記は付かない。"""
    pcr.suggest_n_profiles(ATTRS_3x2x2, n_respondents=30)
    printed = capsys.readouterr().out

    assert "水準バランスと属性間相関の指摘がどちらも出ないのは n = 12 です" in printed
    # n = 12 は p + 2 以上なので、プロファイル数の指摘は残らない
    assert "少ないという指摘は残ります" not in printed


def test_clean_n_scan_skips_combinations_over_the_report_limit(capsys):
    """C(N, n) が報告用の上限を超える n は調べず、その旨を印字する。"""
    start = time.perf_counter()
    pcr.suggest_n_profiles(ATTRS_4x3x2, n_respondents=30)
    elapsed = time.perf_counter() - start
    printed = capsys.readouterr().out

    assert "候補の組み合わせが 100,000 通りを超えるため調べていません" in printed
    # 上限を超える n（C(24, 12) = 2,704,156 など）を列挙していないことの担保
    assert elapsed < 2.0


def test_clean_n_scan_prints_nothing_when_every_n_is_over_the_limit(capsys):
    """すべての n が上限を超える構成では、この行を印字せず従来の文言のまま。"""
    pcr.suggest_n_profiles(ATTRS_3x3x3, n_respondents=30)
    printed = capsys.readouterr().out

    assert "どちらも出ないのは" not in printed
    assert "見つかりませんでした" not in printed
    # 走査していないので、方向を示さない従来の一文が残る
    assert "確認していません" in printed


def test_clean_n_scan_does_not_hang_on_a_huge_candidate_space(capsys):
    """候補数が 65,536 の構成でも即座に返る（C を全部計算していないこと）。"""
    big = {f"a{i}": [1, 2, 3, 4] for i in range(8)}  # N = 4^8 = 65,536, p = 25
    start = time.perf_counter()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        pcr.suggest_n_profiles(big, n_respondents=30)
    elapsed = time.perf_counter() - start
    capsys.readouterr()

    assert elapsed < 2.0


def test_huge_combination_counts_are_printed_in_exponent_form(capsys):
    """桁が大きい組み合わせ数は指数表記にする（教材規模はカンマ区切りのまま）。"""
    design_module = _feasibility._design
    assert design_module._format_count(17_383_860) == "17,383,860"
    assert design_module._format_count(2_704_156) == "2,704,156"
    assert design_module._format_count(10**12) == "約 1.00×10^12"

    big = {f"a{i}": [1, 2, 3, 4] for i in range(8)}  # N = 65,536, p = 25
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        pcr.suggest_n_profiles(big, n_respondents=30)
    printed = capsys.readouterr().out

    assert "×10^" in printed
    # 96桁の整数がそのまま出ていないこと（カンマ区切りだと127文字になる）
    assert not any(len(line) > 100 for line in printed.splitlines())


def test_join_ints_compresses_consecutive_runs():
    """列挙が長くなる場合は「12〜18」とまとめる（3つ以上連続するとき）。"""
    assert _feasibility._design._join_ints([4, 8]) == "4, 8"
    assert _feasibility._design._join_ints([12]) == "12"
    assert _feasibility._design._join_ints([19, 20]) == "19, 20"
    assert _feasibility._design._join_ints([12, 13, 14, 15]) == "12〜15"
    assert _feasibility._design._join_ints([4, 6, 7, 8, 11]) == "4, 6〜8, 11"


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


def test_no_other_pair_wording_when_there_is_only_one_pair():
    """評価対象のペアが1組しかない構成では「別のペアに移ります」を出さない。

    属性2つ × 2水準なら符号化列は2本で、属性をまたぐペアは1組だけ。
    移る先が存在しないので、この一文は成り立たない。
    """
    two_attrs = {"price": [6, 10], "os": ["android", "apple"]}
    design = pcr.design_profiles(two_attrs, 3, seed=1)
    judgment = _feasibility.judge_avoidability(design, list(two_attrs))

    # 前提：この n では相関の指摘を避けられない（状態B に入る）
    assert judgment["correlation_avoidable"] is False

    corr = [
        d
        for d in pcr.check_design(design).diagnostics
        if d.category.startswith("correlation_")
    ]
    assert corr, "この設計では相関の指摘が出るはず"
    for d in corr:
        assert "どのプロファイルの組み合わせを選んでも相関の指摘は残ります" in d.hint
        assert "別のペアに移ります" not in d.hint

    # ペアが2組以上ある構成（2×2×2）では従来どおり出る
    three = pcr.design_profiles(ATTRS_2x2x2, 6, auto_balance=True, seed=1)
    corr3 = [
        d
        for d in pcr.check_design(three).diagnostics
        if d.category.startswith("correlation_")
    ]
    assert corr3
    for d in corr3:
        assert "指摘は別のペアに移ります" in d.hint


def test_saturated_design_wording_differs_from_one_more_profile():
    """n = p（飽和設計）と n = p+1 では文言が違うこと。

    条件は n < p + 2 で共通だが、n = p はちょうど最小限であり、回答者を
    増やしても当てはまりの悪さを検出する自由度は生まれない。教材の
    2水準3属性（p = 4, n = 4）は必ずここに当たる。
    """
    p = 4  # 切片 + 符号化列 3（2×2×2）
    saturated = pcr.design_profiles(ATTRS_2x2x2, p, seed=1)
    one_more = pcr.design_profiles(ATTRS_2x2x2, p + 1, seed=1)

    sat = [
        d
        for d in pcr.check_design(saturated).diagnostics
        if d.category == "few_profiles"
    ]
    more = [
        d
        for d in pcr.check_design(one_more).diagnostics
        if d.category == "few_profiles"
    ]
    assert len(sat) == 1 and len(more) == 1

    assert "ちょうど最小限です（飽和設計）" in sat[0].message
    assert "回答者を何人増やしても" in sat[0].hint
    assert "検出する自由度は生まれません" in sat[0].hint

    assert "ほぼ最小限です" in more[0].message
    assert "飽和設計" not in more[0].message
    assert more[0].hint == ""
    assert "推定の安定性が上がります" in more[0].recommendation

    # 重大度とカテゴリは共通のまま
    assert sat[0].severity == more[0].severity == "中"


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
        assert "相関の指摘が出ない組み合わせも存在しますが" in d.hint
        assert "それらはいずれも水準バランスを満たしません" in d.hint
        assert "両方を満たす設計が存在しないためです" in d.hint
        # 「この n では相関を避けられない」という別状態の文言が出ていないこと
        assert (
            "どのプロファイルの組み合わせを選んでも相関の指摘は残ります" not in d.hint
        )


def test_correlation_hint_distinguishes_this_design_from_other_designs():
    """4つ目の文言が、いまの設計と「存在しうる別の設計」を区別していること。

    バランス表に ◎ が並んだ直後に「その設計は水準バランスを満たしません」と
    だけ書くと、主語がいまの設計に読めて自己矛盾に見える（教材データで確認）。
    """
    design = pcr.design_profiles(ATTRS_3x2x2, 6, auto_balance=True, seed=1)
    result = pcr.check_design(design)

    # 前提：いまの設計はバランスの指摘が出ていない（表では ◎ が並ぶ）
    assert not [d for d in result.diagnostics if d.category.startswith("balance_")], (
        "この設計はバランスの指摘が出ないはず（文言の前提）"
    )

    corr = [d for d in result.diagnostics if d.category.startswith("correlation_")]
    assert corr
    for d in corr:
        # 別の設計の話であることが主語で分かる
        assert "それらはいずれも" in d.hint
        # いまの設計がどちら側なのかも書いてある
        assert "いまの設計は水準バランスのほうを満たしていて" in d.hint

    # いまの設計にバランスの指摘も出ている場合は、そちらの文言になる
    judgment = {
        "n_profiles": 6,
        "balance_avoidable": {"price": True},
        "correlation_avoidable": True,
        "both_avoidable": False,
        "auto_balance_balance": {"price": True},
        "auto_balance_correlation": True,
    }
    hint_warned = analysis_module._correlation_hint(judgment, balance_warned=True)
    assert "いまの設計は、どちらの指摘も出ている状態です。" in hint_warned
    assert "いまの設計は水準バランスのほうを満たしていて" not in hint_warned


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
