"""
design.py
=========
D 最適計画法によるコンジョイント調査の **プロファイル設計** を担当するモジュール。

全属性水準の完全交差（Full factorial）N 個の候補から、
効果コーディング後の情報行列 X'X の行列式 det(X'X) を最大化する
M 個のプロファイルを選ぶ。

>>> profiles = pc.design_profiles(
...     {"price": [6, 8, 10], "os": ["android", "apple"],
...      "camera": ["標準", "高性能", "超高性能"]},
...     n_profiles=12,
...     seed=42,
... )
>>> pc.check_design(profiles)
"""
from __future__ import annotations

import math
import warnings
from itertools import product as _itertools_product
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# 公開 API
# ---------------------------------------------------------------------------

def design_profiles(
    attribute_levels: Dict[str, List[Any]],
    n_profiles: int,
    *,
    reference_levels: Optional[Dict[str, object]] = None,
    n_starts: int = 10,
    seed: Optional[int] = None,
    profile_id_prefix: str = "P",
) -> pd.DataFrame:
    """
    D 最適計画法（Fedorov 交換アルゴリズム）で n_profiles 個のプロファイルを選択する。

    全属性水準の完全交差 N = ∏(水準数) 個の候補から、
    効果コーディング後の **情報行列 X'X の行列式 det(X'X) を最大化** する
    M 個のプロファイルを選ぶ。

    Parameters
    ----------
    attribute_levels : dict
        ``{"属性名": [水準1, 水準2, ...]}`` の辞書。
        辞書のキー順が列の順序になる。

        例::

            {"price": [6, 8, 10], "os": ["android", "apple"],
             "camera": ["標準", "高性能", "超高性能"]}

    n_profiles : int
        選択するプロファイル数 M。
        最低でもパラメータ数（切片 + Σ(水準数 - 1)）以上が必要。

    reference_levels : dict, optional
        効果コーディングの基準水準を指定する辞書。
        省略時は各属性の先頭水準を基準とする。
        ``encode()`` に渡す値と同じ辞書を渡すと、設計の最適化基準と
        符号化基準が一致する。

    n_starts : int, default 10
        ランダム初期化の試行回数。
        多いほど良い解を見つけやすいが実行時間も増える。

    seed : int, optional
        乱数シード（再現性のため）。

    profile_id_prefix : str, default "P"
        プロファイル ID の接頭辞（"P1", "P2", ...）。

    Returns
    -------
    pd.DataFrame
        選択された n_profiles 行・属性列の DataFrame。
        インデックスは ``"P1"``, ``"P2"``, ... になる。

        ``df.attrs["d_efficiency"]`` — D 相対効率（0〜1）。
        完全交差と比べてどれだけ情報を保っているかの指標。
        1.0 に近いほど良い設計。

        ``df.attrs["n_candidates"]`` — 完全交差の候補数 N。

        ``df.attrs["det_xpx"]`` — 選択された設計の det(X'X)。

    Raises
    ------
    ValueError
        ``attribute_levels`` が空または辞書でない場合。
        いずれかの属性の水準数が 2 未満、または水準リストに重複がある場合。
        ``n_starts`` が 1 未満の場合。
        ``n_profiles`` が完全交差の候補数 N を超える場合。
        ``n_profiles`` がパラメータ数より少ない場合。
        ``reference_levels`` に存在しない水準が指定された場合。

    Notes
    -----
    **D 最適計画（D-optimal design）とは**

    情報行列 X'X の行列式 det(X'X) を最大化するプロファイルの組み合わせを選ぶ設計法。
    det(X'X) が大きいほどパラメータ推定量の一般化分散が小さくなり
    （推定精度が上がる）、直交性に近い設計が得られる。

    **Fedorov 交換アルゴリズム**

    1. N 個の候補からランダムに M 個を選んで初期設計とする。
    2. 選択済みの各行と未選択の各行の **交換** を試みる。
    3. 行列式補題（matrix determinant lemma）を使い det(X'X) の変化を
       O(p²) で計算する（行列の再計算不要）。
    4. det(X'X) が増加する交換を実施する。改善がなくなるまで繰り返す。
    5. ``n_starts`` 回の試行のうち最良の設計を返す。

    **D 相対効率**

    D-efficiency = (det(X_selected'X_selected) / det(X_full'X_full))^(1/p)

    ここで p はパラメータ数（切片含む）。完全交差を選んだ場合は 1.0 になる。
    n_profiles が少ないほど小さくなる。

    Examples
    --------
    >>> import py4conjoint as pc
    >>> profiles = pc.design_profiles(
    ...     {"price": [6, 8, 10], "os": ["android", "apple"],
    ...      "camera": ["標準", "高性能", "超高性能"]},
    ...     n_profiles=12,
    ...     reference_levels={"price": 10, "os": "android", "camera": "標準"},
    ...     seed=42,
    ... )
    >>> pc.check_design(profiles)
    """
    if not isinstance(attribute_levels, dict) or len(attribute_levels) == 0:
        raise ValueError(
            "attribute_levels は空でない辞書を指定してください。\n"
            "  例: {'price': [6, 8, 10], 'os': ['android', 'apple']}"
        )
    for attr, levels in attribute_levels.items():
        if len(levels) < 2:
            raise ValueError(
                f"属性 '{attr}' の水準数は 2 以上にしてください（現在: {len(levels)}）。"
            )
        # 重複した水準は候補数 N・パラメータ数 p を架空に膨らませ、
        # D 相対効率などが誤った値になるため、ここで弾く。
        if len(set(levels)) != len(levels):
            raise ValueError(
                f"属性 '{attr}' の水準リストに重複があります: {list(levels)}\n"
                "  水準は重複なく指定してください。"
            )

    attrs = list(attribute_levels.keys())
    levels_list = [list(attribute_levels[a]) for a in attrs]

    # 完全交差の候補数
    N = 1
    for levels in levels_list:
        N *= len(levels)

    # パラメータ数（切片 + Σ(水準数 - 1)）
    p = 1  # 切片
    for levels in levels_list:
        p += len(levels) - 1

    if n_starts < 1:
        raise ValueError(
            f"n_starts は 1 以上の整数を指定してください（指定値: {n_starts}）。"
        )
    if n_profiles > N:
        raise ValueError(
            f"n_profiles ({n_profiles}) が完全交差の候補数 N ({N}) を超えています。\n"
            f"  n_profiles を {N} 以下にしてください。"
        )
    if n_profiles < p:
        raise ValueError(
            f"n_profiles ({n_profiles}) がパラメータ数 ({p}) より少ないため、\n"
            f"回帰分析を実行できません。\n"
            f"  n_profiles を {p} 以上にしてください。\n"
            f"  パラメータ数 = 切片(1) + Σ(水準数 - 1) = {p}"
        )

    # 基準水準の設定（省略時は各属性の先頭水準）
    ref_lvls: Dict[str, object] = {}
    for attr, levels in attribute_levels.items():
        ref_lvls[attr] = (reference_levels or {}).get(attr, levels[0])
        if ref_lvls[attr] not in levels:
            raise ValueError(
                f"属性 '{attr}' に基準水準 '{ref_lvls[attr]}' が存在しません。\n"
                f"  存在する水準: {levels}"
            )

    # 完全交差の生成
    df_full = pd.DataFrame(
        [dict(zip(attrs, combo)) for combo in _itertools_product(*levels_list)]
    )

    # 完全交差をそのまま返す場合
    if n_profiles == N:
        out = df_full.copy()
        out.index = [f"{profile_id_prefix}{i+1}" for i in range(N)]
        out.attrs["d_efficiency"] = 1.0
        out.attrs["n_candidates"] = N
        X_full = _build_effect_matrix(df_full, attribute_levels, ref_lvls)
        out.attrs["det_xpx"] = float(np.linalg.det(X_full.T @ X_full))
        return out

    # 効果コーディング設計行列（N × p）
    X_full = _build_effect_matrix(df_full, attribute_levels, ref_lvls)

    # D 相対効率の基準値: det(X_full'X_full)
    det_full = float(np.linalg.det(X_full.T @ X_full))

    # D 最適交換アルゴリズムを n_starts 回実行
    rng = np.random.default_rng(seed)
    best_indices: List[int] = []
    best_det = -np.inf

    for _ in range(n_starts):
        indices, det_val = _d_exchange_run(X_full, n_profiles, rng)
        if det_val > best_det:
            best_det = det_val
            best_indices = indices

    # 結果の整形（行インデックス順でソート）
    sorted_idx = sorted(best_indices)
    out = df_full.iloc[sorted_idx].copy()
    out.index = [f"{profile_id_prefix}{i+1}" for i in range(n_profiles)]
    out.index.name = None

    # D 相対効率: (det(X_sel'X_sel) / det(X_full'X_full))^(1/p)
    X_sel = X_full[sorted_idx]
    det_sel = float(np.linalg.det(X_sel.T @ X_sel))
    if det_full > 0 and det_sel > 0:
        d_eff = (det_sel / det_full) ** (1.0 / p)
    else:
        d_eff = 0.0

    out.attrs["d_efficiency"] = float(d_eff)
    out.attrs["n_candidates"] = N
    out.attrs["det_xpx"] = det_sel

    return out


# ---------------------------------------------------------------------------
# 公開 API: suggest_n_profiles 関数
# ---------------------------------------------------------------------------

def suggest_n_profiles(
    attribute_levels: Dict[str, List[Any]],
    *,
    n_respondents: Optional[int] = None,
    obs_per_predictor: int = 10,
    max_burden: int = 20,
) -> pd.DataFrame:
    """
    ``design_profiles()`` の ``n_profiles`` 引数に設定する値の目安を返す。

    属性数・水準数・予定回答者数から「推奨プロファイル数」を計算し、
    その根拠とともに DataFrame で返す。

    Parameters
    ----------
    attribute_levels : dict
        ``design_profiles()`` と同じ形式の辞書。

    n_respondents : int, optional
        予定回答者数。指定すると特定の回答者数に対する推奨値のみ返す。
        省略時は代表的な回答者数（5, 10, 20, 30, 50, 100 人）の一覧を返す。

    obs_per_predictor : int, default 10
        目標とする「観測数 ÷ 符号化列数」の比率。
        ``fit()`` の中警告（``obs_per_predictor``）の閾値と同じ値がデフォルト。

    max_burden : int, default 20
        1 回答者に提示するプロファイル数の上限目安（アンケート負担の観点）。
        ただし統計的最低限（パラメータ数 p）を下回る場合は p が優先され、
        ``UserWarning`` が出る。

    Returns
    -------
    pd.DataFrame
        列：``"回答者数"``, ``"obs/pred≥{obs_per_predictor}（最低限）"``,
        ``"推奨 n_profiles"``, ``"obs/pred（達成）"``, ``"観測数 obs"``

        ``df.attrs["n_params"]``     — パラメータ数 p（切片含む）

        ``df.attrs["n_encoded"]``    — 符号化列数 = p - 1 = Σ(水準数 - 1)

        ``df.attrs["n_candidates"]`` — 完全交差の候補数 N

        ``df.attrs["m_min"]``        — 統計的最低限 M（= p）

        ``df.attrs["m_orme"]``       — Orme の経験則による目安（= 符号化列数 × 2、N で上限制限済み）

    Raises
    ------
    ValueError
        ``attribute_levels`` が空または辞書でない場合。
        いずれかの属性の水準数が 2 未満、または水準リストに重複がある場合。
        ``n_respondents`` が指定されており 1 未満の場合。
        ``obs_per_predictor`` または ``max_burden`` が 1 未満の場合。

    Notes
    -----
    **n_profiles を決める 3 つの基準**

    1. **統計的最低限** （= p = 切片 + 符号化列数）

       回帰分析が実行できる最小値。このとき自由度 = M × 回答者数 - p であり、
       M = p だと自由度がゼロに近く推定が非常に不安定。

    2. **Orme (2010) の経験則** （≈ 符号化列数 × 2）

       コンジョイント分析の実務でよく使われる経験則。
       部分効用の推定が実用的に安定する最小プロファイル数の目安。

    3. **観測数条件** （M × 回答者数 ≥ ``obs_per_predictor`` × 符号化列数）

       ``fit()`` の中警告（``obs_per_predictor``、重大度：中）を出さない水準。
       M ≥ ceil(obs_per_predictor × 符号化列数 / 回答者数) で計算。

    **推奨値の決定方法**

    上記 3 基準の最大値を取り、``min(N, max_burden)`` で上限制限する。
    ただし上限制限後の値が統計的最低限 p を下回る場合
    （``max_burden < p`` のとき）は p まで引き上げる。
    プロファイル数が p 未満だと、回答者を何人集めても設計行列が
    ランク落ちして回帰分析が実行できないため。

    Examples
    --------
    >>> import py4conjoint as pc
    >>> pc.suggest_n_profiles(
    ...     {"price": [6, 8, 10], "os": ["android", "apple"],
    ...      "camera": ["標準", "高性能", "超高性能"]},
    ...     n_respondents=30,
    ... )
    """
    # ---------- 入力バリデーション ----------
    if not isinstance(attribute_levels, dict) or len(attribute_levels) == 0:
        raise ValueError(
            "attribute_levels は空でない辞書を指定してください。\n"
            "  例: {'price': [6, 8, 10], 'os': ['android', 'apple']}"
        )
    for attr, lvs in attribute_levels.items():
        if len(lvs) < 2:
            raise ValueError(
                f"属性 '{attr}' の水準数は 2 以上にしてください（現在: {len(lvs)}）。"
            )
        # 重複した水準は符号化列数・候補数 N の計算を狂わせるため弾く
        if len(set(lvs)) != len(lvs):
            raise ValueError(
                f"属性 '{attr}' の水準リストに重複があります: {list(lvs)}\n"
                "  水準は重複なく指定してください。"
            )
    if n_respondents is not None and n_respondents < 1:
        raise ValueError(
            f"n_respondents は 1 以上の整数を指定してください（指定値: {n_respondents}）。"
        )
    if obs_per_predictor < 1:
        raise ValueError(
            f"obs_per_predictor は 1 以上の整数を指定してください（指定値: {obs_per_predictor}）。"
        )
    if max_burden < 1:
        raise ValueError(
            f"max_burden は 1 以上の整数を指定してください（指定値: {max_burden}）。"
        )

    # ---------- 基本統計 ----------
    n_encoded = sum(len(lvs) - 1 for lvs in attribute_levels.values())
    p = n_encoded + 1
    N = 1
    for lvs in attribute_levels.values():
        N *= len(lvs)

    m_min = p
    m_orme = 2 * n_encoded

    # max_burden がパラメータ数を下回ると回帰分析が実行できないため、
    # 推奨値は m_min（= p）を下限として引き上げる
    if max_burden < m_min:
        warnings.warn(
            f"max_burden ({max_burden}) がパラメータ数 p ({m_min}) を下回っています。\n"
            f"プロファイル数がパラメータ数未満だと回帰分析を実行できないため、\n"
            f"推奨 n_profiles は統計的最低限の {m_min} に引き上げられます。\n"
            "アンケート負担を抑えたい場合は属性数・水準数を減らすことを検討してください。",
            UserWarning,
            stacklevel=2,
        )

    resp_list = (
        [n_respondents] if n_respondents is not None
        else [5, 10, 20, 30, 50, 100]
    )

    rows = []
    for n_resp in resp_list:
        # obs 条件のみから計算した最小 M
        m_obs_only = math.ceil(obs_per_predictor * n_encoded / n_resp)
        # 3 基準の最大値をとり、上限制限（ただし統計的最低限 m_min は下回らない）
        m_rec_raw = max(m_min, m_orme, m_obs_only)
        m_rec = max(min(m_rec_raw, max_burden, N), m_min)

        actual_obs = m_rec * n_resp
        actual_ratio = (
            round(actual_obs / n_encoded, 1) if n_encoded > 0 else float("inf")
        )

        rows.append({
            "回答者数": n_resp,
            # obs/pred 条件と統計的最低限の両方を満たす最小 M
            f"obs/pred≥{obs_per_predictor}（最低限）": min(max(m_obs_only, m_min), N),
            "推奨 n_profiles": m_rec,
            "obs/pred（達成）": actual_ratio,
            "観測数 obs": actual_obs,
        })

    result = pd.DataFrame(rows)
    result.attrs.update({
        "n_params": p,
        "n_encoded": n_encoded,
        "n_candidates": N,
        "m_min": m_min,
        "m_orme": min(m_orme, N),
    })
    return result


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------

def _build_effect_matrix(
    df: pd.DataFrame,
    attribute_levels: Dict[str, List[Any]],
    reference_levels: Dict[str, object],
) -> np.ndarray:
    """
    効果コーディング済み設計行列（先頭列 = 切片）を構築する。

    K 水準属性には K-1 列を生成する。
    基準水準の行 = すべての K-1 列が -1。
    """
    n = len(df)
    cols = [np.ones(n, dtype=float)]
    for attr, levels in attribute_levels.items():
        ref = reference_levels[attr]
        non_ref = [lv for lv in levels if lv != ref]
        for lv in non_ref:
            col = np.array(
                [1.0 if v == lv else (-1.0 if v == ref else 0.0)
                 for v in df[attr]],
                dtype=float,
            )
            cols.append(col)
    return np.column_stack(cols)


def _d_exchange_run(
    X_full: np.ndarray,
    M: int,
    rng: np.random.Generator,
) -> Tuple[List[int], float]:
    """
    Fedorov 交換アルゴリズムの 1 試行。

    行列式補題を用いた高速更新:
        swap(行 x_i → x_j) 後の det(X'X_new) / det(X'X_old) =
            (1 - d_i)(1 + d_j) + c_ij²

        d_i = x_i' (X'X)^{-1} x_i
        d_j = x_j' (X'X)^{-1} x_j
        c_ij = x_i' (X'X)^{-1} x_j

    比率 > 1 なら det が増加 → 交換を採用する。

    Returns
    -------
    (selected_indices, det_value)
    """
    N, p = X_full.shape

    # ランダム初期選択（not_sel をソートして順序を決定的にする）
    sel: List[int] = list(rng.choice(N, M, replace=False))
    not_sel: List[int] = sorted(set(range(N)) - set(sel))

    improved = True
    while improved:
        improved = False

        X_sel_arr = X_full[sel]      # shape (M, p)
        X_not_arr = X_full[not_sel]  # shape (N-M, p)
        XtX = X_sel_arr.T @ X_sel_arr

        # 正則でない場合は擬似逆行列にフォールバック
        if np.linalg.matrix_rank(XtX) < p:
            XtX_inv = np.linalg.pinv(XtX)
        else:
            XtX_inv = np.linalg.inv(XtX)

        # 全行の d 値を一括計算
        # d_i = diag(X_sel @ XtX_inv @ X_sel.T)
        tmp_sel = X_sel_arr @ XtX_inv  # (M, p)
        d_sel = (tmp_sel * X_sel_arr).sum(axis=1)  # (M,)

        tmp_not = X_not_arr @ XtX_inv  # (N-M, p)
        d_not = (tmp_not * X_not_arr).sum(axis=1)  # (N-M,)

        best_ratio = 1.0 + 1e-10  # 改善の閾値
        best_swap: Optional[Tuple[int, int]] = None

        for i_pos in range(M):
            d_i = d_sel[i_pos]
            # c_ij = x_j' (XtX_inv x_i) を N-M 個まとめて計算
            XtX_inv_xi = XtX_inv @ X_sel_arr[i_pos]  # (p,)
            c_all = X_not_arr @ XtX_inv_xi            # (N-M,)

            # 交換後の行列式比
            ratios = (1.0 - d_i) * (1.0 + d_not) + c_all ** 2
            j_best = int(np.argmax(ratios))

            if ratios[j_best] > best_ratio:
                best_ratio = ratios[j_best]
                best_swap = (i_pos, j_best)

        if best_swap is not None:
            i_pos, j_pos = best_swap
            sel[i_pos], not_sel[j_pos] = not_sel[j_pos], sel[i_pos]
            improved = True

    # 最終 det(X'X)
    X_final = X_full[sel]
    det = float(np.linalg.det(X_final.T @ X_final))
    return sel, det
