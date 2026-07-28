"""この n では警告を避けられるのか、を全数列挙で判定する（rating）。

``check_design()`` が出す警告に「auto_balance=True で解消できます」
「この n では回避できません」といった案内を添えるために、同じ属性・水準で
作れる n プロファイルの設計を **すべて** 調べ、警告の出ない設計が存在するかを返す。

判定基準（どの CV・どの |r| から警告を出すのか、どの列ペアを評価するのか）は
:mod:`py4conjoint.rating.analysis` の定数とヘルパーが唯一の定義であり、
このモジュールは数値を1つも持たない。ベクトル化のためにしきい値の**定数**を
参照している箇所はあるが、値を書き写してはいない。両者が一致していることは
``tests/test_design_feasibility.py`` の一致テスト（全設計を列挙して
``check_design()`` の結果と突き合わせる）で担保する。

**計算量の上限**：候補の選び方の総数 C(N, n) が
``design._REPORT_MAX_COMBINATIONS``（= 100,000）を超える場合は判定しない
（``None`` を返す）。``check_design()`` は即座に返る関数であり、待たされる
理由が利用者に見えないため。``design_profiles(auto_balance=True)`` が総当たりに
入る上限（``design._EXHAUSTIVE_MAX_COMBINATIONS`` = 1,000,000）とは別の、
より低い値である。

**実装上の要点**：属性の水準を 0/1 で表した指標行列 L について、選んだ行の
グラム行列 ``G = L_sub' L_sub`` を計算すると、対角に各水準の出現回数が、
非対角に属性ペアのクロス集計が入る。効果コーディングされた列は
「ある水準の指標 − 基準水準の指標」なので、相関に必要な内積・合計はすべて
G から引ける。列そのものを設計ごとに作り直す必要はない。

基準水準は :func:`analysis._check_correlation` と同じく **設計ごとの最頻値**
（同数なら小さい水準）である。ここを固定水準で代用すると結果が変わる
（3水準を含む構成で実測したところ、警告の有無の判定が 3〜9% の設計でずれた）
ため、設計ごとの最頻値を G の添字として再現している。
"""

from __future__ import annotations

import math
import warnings as _warnings
from itertools import combinations as _itertools_combinations
from itertools import product as _itertools_product
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

from . import analysis as _analysis
from . import design as _design

# 1回の einsum に流す設計の数（メモリと速度の兼ね合い）
_BATCH = 5_000


def _observed_levels(
    profiles: pd.DataFrame, attrs: List[str]
) -> Optional[Dict[str, List[Any]]]:
    """設計から属性ごとの水準を推定する。判定できない形なら None。

    check_design は設計の DataFrame しか受け取らないので、水準はそこから
    推定するしかない。**設計に1度も現れない水準は原理的に検出できない**ため、
    観測された水準がすべてであると仮定している。
    """
    levels: Dict[str, List[Any]] = {}
    for attr in attrs:
        vals = list(pd.unique(profiles[attr]))
        if len(vals) < 2:
            # 水準が1つしかない属性がある設計は、候補空間を復元できない
            return None
        try:
            # 基準水準（最頻値、同数なら小さい水準）を再現するため昇順に並べる
            vals = sorted(vals)
        except TypeError:
            # 大小比較できない水準が混ざっている場合は判定しない
            return None
        levels[attr] = vals
    return levels


def _encoded_layout(
    levels: Dict[str, List[Any]],
) -> "tuple[np.ndarray, np.ndarray]":
    """効果コーディング列の並び（属性の対応）と、評価するペアの位置を返す。

    列の位置と属性の対応は基準水準の選び方によらず固定なので、
    「同一属性内のペアを除く」マスクは設計ごとに作り直さなくてよい。
    """
    attr_of_col = []
    for a, lvs in enumerate(levels.values()):
        attr_of_col += [a] * (len(lvs) - 1)
    attr_of_col = np.array(attr_of_col, dtype=np.int64)
    m = len(attr_of_col)
    cross = attr_of_col[:, None] != attr_of_col[None, :]
    upper = np.triu(np.ones((m, m), dtype=bool), k=1)
    return attr_of_col, cross & upper


def _batch_flags(
    idx: np.ndarray,
    L: np.ndarray,
    spans: List["tuple[int, int]"],
    pair_mask: np.ndarray,
    n_profiles: int,
) -> "tuple[np.ndarray, np.ndarray, np.ndarray]":
    """設計のかたまりについて (水準がすべて現れるか, 属性ごとにバランス警告なしか,
    相関警告なしか) を返す。

    **バランスと相関のフラグは、水準がすべて現れる設計についてのみ
    check_design と一致する。** 水準が欠けた設計では
    :func:`analysis._check_balance` が観測された水準だけで CV を計算し、
    :func:`analysis._check_correlation` も列を作らないのに対し、ここでは
    出現回数 0 の水準も数に入れるため、両者は一致しない。呼び出し側は
    ``levels_ok`` との論理積だけを使うこと（欠けた水準がある設計は、
    そもそも候補として数えてはいけない）。
    """
    n_designs = len(idx)
    sub = L[idx]  # (B, M, T)
    counts = sub.sum(axis=1).astype(np.float64)  # (B, T)
    gram = np.einsum("bmt,bmu->btu", sub, sub).astype(np.float64)  # (B, T, T)

    levels_ok = np.all(counts > 0, axis=1)

    # ---- 水準バランス（_check_balance と同じ CV：標本標準偏差 ÷ 平均、小数4桁） ----
    balance_ok = np.ones((n_designs, len(spans)), dtype=bool)
    for a, (s, e) in enumerate(spans):
        block = counts[:, s:e]
        mean = block.mean(axis=1)
        std = block.std(axis=1, ddof=1)
        cv = np.round(np.divide(std, mean, out=np.zeros_like(std), where=mean > 0), 4)
        # 警告が出ない ＝ _balance_severity() が None を返す範囲
        balance_ok[:, a] = cv <= _analysis._BALANCE_CV_MINOR

    # ---- 基準水準（設計ごとの最頻値。同数なら小さい水準＝ブロック内の先頭） ----
    ref_cols = []
    for s, e in spans:
        ref_cols.append(s + np.argmax(counts[:, s:e], axis=1))

    # ---- 効果コーディング列（非基準水準）と、その基準水準の添字 ----
    u_parts, r_parts = [], []
    for (s, e), ref in zip(spans, ref_cols):
        cand = np.arange(s, e)
        is_ref = cand[None, :] == ref[:, None]
        order = np.argsort(is_ref, axis=1, kind="stable")  # 非基準が先に来る
        u_parts.append(cand[order[:, : (e - s) - 1]])
        r_parts.append(np.repeat(ref[:, None], (e - s) - 1, axis=1))
    u = np.concatenate(u_parts, axis=1)  # (B, m)
    r = np.concatenate(r_parts, axis=1)  # (B, m)

    b = np.arange(n_designs)[:, None, None]
    # 効果コーディング列同士の内積：(d_u - d_r)'(d_v - d_s)
    cross = (
        gram[b, u[:, :, None], u[:, None, :]]
        - gram[b, u[:, :, None], r[:, None, :]]
        - gram[b, r[:, :, None], u[:, None, :]]
        + gram[b, r[:, :, None], r[:, None, :]]
    )
    cnt_u = np.take_along_axis(counts, u, axis=1)
    cnt_r = np.take_along_axis(counts, r, axis=1)
    col_sum = cnt_u - cnt_r  # 各列の合計
    col_sq = cnt_u + cnt_r  # 各列の二乗和（指標は排反なので出現回数の和）

    cov = cross - col_sum[:, :, None] * col_sum[:, None, :] / n_profiles
    var = col_sq - col_sum**2 / n_profiles
    denom = np.sqrt(var[:, :, None] * var[:, None, :])
    with np.errstate(divide="ignore", invalid="ignore"):
        corr = np.where(denom > 0, cov / denom, np.nan)
    # pandas の .corr().round(4) と同じ丸め。定数列は NaN になり、
    # _correlation_severity(NaN) は None（警告なし）なので 0 と同じ扱いにする。
    abs_r = np.abs(np.round(corr, 4))
    abs_r = np.where(np.isnan(abs_r), 0.0, abs_r)
    max_abs_r = (
        abs_r[:, pair_mask].max(axis=1) if pair_mask.any() else np.zeros(n_designs)
    )
    # 警告が出ない ＝ _correlation_severity() が None を返す範囲
    correlation_ok = max_abs_r <= _analysis._CORRELATION_ABS_MINOR

    return levels_ok, balance_ok, correlation_ok


def _auto_balance_warnings(
    levels: Dict[str, List[Any]], n_profiles: int
) -> Optional["tuple[Dict[str, bool], bool]"]:
    """``design_profiles(auto_balance=True)`` が返す設計に、どの警告が残るかを調べる。

    案内文で「auto_balance=True で解消できます」と言えるかどうかは、実際に
    その関数が返す設計を ``check_design`` と同じ基準で評価して決める。
    """
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")
        try:
            design = _design.design_profiles(levels, n_profiles, auto_balance=True)
        except ValueError:
            # n_profiles がパラメータ数を下回る場合など。判定しない。
            return None

    attrs = list(levels)
    balance_df = _analysis._check_balance(design, attrs)
    corr_df = _analysis._check_correlation(design, attrs)

    warned_balance = {
        attr: _analysis._balance_severity(balance_df.loc[attr, "CV"]) is not None
        for attr in attrs
    }
    warned_correlation = False
    if not corr_df.empty:
        for c1, c2 in _analysis._cross_attribute_pairs(list(corr_df.columns)):
            if (
                _analysis._correlation_severity(abs(float(corr_df.loc[c1, c2])))
                is not None
            ):
                warned_correlation = True
                break
    return warned_balance, warned_correlation


def judge_avoidability(
    profiles: pd.DataFrame, attrs: List[str]
) -> Optional[Dict[str, Any]]:
    """この n で警告を避けられるかを判定する。判定しない場合は None を返す。

    Returns
    -------
    dict または None
        ``"balance_avoidable"``     — 属性ごとに、その属性のバランス警告が
        出ない設計が存在するか

        ``"correlation_avoidable"`` — 相関の指摘が1つも出ない設計が存在するか

        ``"both_avoidable"``        — バランスと相関の警告がどちらも出ない
        設計が存在するか

        ``"auto_balance_balance"``  — 属性ごとに、``auto_balance=True`` の
        設計でバランス警告が残るか

        ``"auto_balance_correlation"`` — ``auto_balance=True`` の設計で
        相関の指摘が残るか

        ``"n_profiles"`` / ``"n_candidates"`` / ``"n_combinations"``

    Notes
    -----
    相関を「そのペアだけ避けられるか」で判定してはいけない。2×2×2 を6プロファイル
    にする設計では、特定のペアの |r| を 0 にできるが、必ず別のペアに移るだけで、
    相関の指摘が出ない設計は 28 通り中1つも存在しない。そのため判定は
    「**相関の指摘が1つも出ない設計が存在するか**」で行う。

    候補からは「1度も現れない水準がある設計」を除く。除かないと、ある属性を
    1水準に固定した（＝その属性を推定できない）設計が「相関の指摘なし」として
    数えられてしまう。
    """
    n_profiles = len(profiles)
    levels = _observed_levels(profiles, attrs)
    if levels is None:
        return None

    n_candidates = 1
    for lvs in levels.values():
        n_candidates *= len(lvs)
    if n_profiles < 2 or n_profiles > n_candidates:
        return None

    n_combinations = math.comb(n_candidates, n_profiles)
    if n_combinations > _design._REPORT_MAX_COMBINATIONS:
        return None

    names = list(levels)
    df_full = pd.DataFrame(
        [dict(zip(names, combo)) for combo in _itertools_product(*levels.values())]
    )
    indicator, spans = _design._level_indicator(df_full, levels)
    _attr_of_col, pair_mask = _encoded_layout(levels)

    balance_avoidable = np.zeros(len(names), dtype=bool)
    correlation_avoidable = False
    both_avoidable = False

    buf: List["tuple[int, ...]"] = []

    def flush(buf_local: List["tuple[int, ...]"]) -> None:
        nonlocal correlation_avoidable, both_avoidable
        idx = np.array(buf_local, dtype=np.int64)
        levels_ok, balance_ok, correlation_ok = _batch_flags(
            idx, indicator, spans, pair_mask, n_profiles
        )
        if not levels_ok.any():
            return
        balance_avoidable[:] |= (balance_ok & levels_ok[:, None]).any(axis=0)
        correlation_avoidable = bool(correlation_avoidable) or bool(
            (correlation_ok & levels_ok).any()
        )
        both_avoidable = bool(both_avoidable) or bool(
            (correlation_ok & levels_ok & balance_ok.all(axis=1)).any()
        )

    for combo in _itertools_combinations(range(n_candidates), n_profiles):
        buf.append(combo)
        if len(buf) == _BATCH:
            flush(buf)
            buf = []
    if buf:
        flush(buf)

    auto_balance = _auto_balance_warnings(levels, n_profiles)
    if auto_balance is None:
        return None
    warned_balance, warned_correlation = auto_balance

    return {
        "balance_avoidable": dict(zip(names, balance_avoidable.tolist())),
        "correlation_avoidable": correlation_avoidable,
        "both_avoidable": both_avoidable,
        "auto_balance_balance": warned_balance,
        "auto_balance_correlation": warned_correlation,
        "n_profiles": n_profiles,
        "n_candidates": n_candidates,
        "n_combinations": n_combinations,
    }
