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
from itertools import combinations as _itertools_combinations
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
    auto_balance: bool = False,
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

    auto_balance : bool, default False
        ``True`` にすると、**水準バランスを満たす設計の中で** det(X'X) を
        最大化する。既定の ``False`` では従来どおり、バランスを考慮せずに
        det(X'X) を最大化する（挙動は完全に同じ）。

        ここでいうバランスとは「すべての属性について、水準の出現回数の
        最大と最小の差が 1 以下」であること。``suggest_n_profiles()`` が
        推奨する n_profiles で設計を作ると ``check_design()`` が [大] の
        バランス警告を出すことがあり、その逃げ道として使う。

        バランスを優先する分、det(X'X) は制約なしの最良解より小さくなり
        うる（どれだけ失ったかは ``df.attrs["auto_balance"]`` で確認できる。
        後述の Notes を参照）。

    n_starts : int, default 10
        ランダム初期化の試行回数。
        多いほど良い解を見つけやすいが実行時間も増える。
        ``auto_balance=True`` で総当たり経路に入る場合（Notes 参照）は
        結果に影響しない。

    seed : int, optional
        乱数シード（再現性のため）。
        ``n_starts`` と同じく、総当たり経路では結果に影響しない
        （総当たりは決定的で、同点は候補の並び順で最初のものを採る）。

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

        ``auto_balance=True`` のときは、さらに探索の来歴が入る::

            df.attrs["auto_balance"] = {
                "method": "exhaustive",          # 探索の方法（後述）
                "balanced": True,                # バランス制約を満たせたか
                "det_xpx": 768.0,                # 返した設計の det(X'X)
                "det_xpx_unconstrained": 1024.0, # 制約なしで最適化した場合
                "det_ratio": 0.75,               # 前者 ÷ 後者
            }

        ``"det_ratio"`` は「バランスを取ったことで失う精度の割合」を表す。
        1.0 なら、バランスを満たしたまま制約なしの最良解に到達している。
        バランスを満たせなかった場合（``"balanced"`` が ``False``）は
        比を定義できないので ``None`` になる。

        ``"method"`` が ``"exchange"`` のときは、比が 1.0 をわずかに
        超えることがある。この経路では ``"det_xpx_unconstrained"`` 自体も
        発見的探索の到達点であり、真の最適値ではないため、制約付きの探索が
        そちらを上回ることがあるからである。

        ``"method"`` は次の3つのいずれか。

        * ``"exhaustive"``   … 総当たり。厳密解。
        * ``"exchange"``     … 制約付き交換アルゴリズム。**発見的探索**で、
          最良解である保証はない。
        * ``"full_factorial"`` … n_profiles が完全交差の候補数 N と等しく、
          完全交差をそのまま返した場合（定義上つねに完全バランス）。

        ``"det_ratio"`` を読むときは ``"method"`` も見ること。厳密解の比か、
        探索の到達点の比かで意味が変わるため。

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

    **auto_balance の探索方法（厳密解と発見的探索の境界）**

    候補の選び方の総数 C(N, n_profiles) が
    ``_EXHAUSTIVE_MAX_COMBINATIONS``（= 1,000,000）以下なら **総当たり** で
    厳密解を求め、それを超えるときは制約付き交換アルゴリズム（発見的探索）に
    切り替える。この定数は速度の閾値であると同時に、**返る解が厳密解か
    発見的探索かの境界**でもある。どちらだったかは
    ``df.attrs["auto_balance"]["method"]`` で判別できる。

    値は実測に基づく（バッチ化した numpy 実装）::

        C(N, n) = 184,756  （N=20, n=10）→ 約 0.25 秒
        C(N, n) = 1,000,000            → 約 1.5 秒
        C(N, n) = 2,704,156（N=24, n=12）→ 約 4.0 秒

    授業で使う規模（N ≤ 20 程度）はほぼ総当たり側に入る。実行時間の許容度が
    変わったら、この定数を見直せばよい。

    交換アルゴリズム側（``"exchange"``）の実測は次のとおり（既定の
    ``n_starts=10``）。総当たりと違って C(N, n) ではなく、候補数 N と
    n_profiles でおおよそ決まる::

        N =  72, n = 18（4×3×3×2）      → 約 0.3 秒
        N = 288, n = 30（4×4×3×3×2）    → 約 4 秒
        N = 360, n = 40（5×4×3×3×2）    → 約 7 秒
        N = 720, n = 60（5×4×3×3×2×2）  → 約 28 秒

    所要時間は ``n_starts`` にほぼ比例する。時間がかかりすぎる場合は
    ``n_starts`` を下げればよい（そのぶん解の質は落ちうる）。

    **auto_balance が保証するのは水準バランスだけ（直交性は別）**

    ``auto_balance=True`` の契約は「水準の出現回数の最大と最小の差が 1 以下」
    のみで、**属性間の相関（直交性）については何も約束しない**。両者は別の
    基準であり、バランスを満たしていても属性間に相関が残ることがある。
    たとえば 3水準を含む属性構成で 6 プロファイルを選ぶ設計では、
    構造的に |r| = 1/3 程度の相関が避けられない。

    そのため ``auto_balance=True`` で作った設計に対しても、
    ``check_design()`` が相関について警告を出すことはある。バランス警告が
    消えたのに相関の指摘が残るのは矛盾ではない。

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
        out.index = [f"{profile_id_prefix}{i + 1}" for i in range(N)]
        out.attrs["d_efficiency"] = 1.0
        out.attrs["n_candidates"] = N
        X_full = _build_effect_matrix(df_full, attribute_levels, ref_lvls)
        det_all = float(np.linalg.det(X_full.T @ X_full))
        out.attrs["det_xpx"] = det_all
        if auto_balance:
            # 完全交差はすべての水準が同じ回数だけ現れるので、定義上つねに
            # バランスを満たす。制約なしの最良解とも一致する。
            out.attrs["auto_balance"] = {
                "method": "full_factorial",
                "balanced": True,
                "det_xpx": det_all,
                "det_xpx_unconstrained": det_all,
                "det_ratio": 1.0,
            }
        return out

    # 効果コーディング設計行列（N × p）
    X_full = _build_effect_matrix(df_full, attribute_levels, ref_lvls)

    # D 相対効率の基準値: det(X_full'X_full)
    det_full = float(np.linalg.det(X_full.T @ X_full))

    balance_info: Optional[Dict[str, Any]] = None

    if auto_balance:
        best_indices, balance_info = _search_balanced_design(
            X_full, df_full, attribute_levels, n_profiles, n_starts, seed
        )
    else:
        # D 最適交換アルゴリズムを n_starts 回実行
        rng = np.random.default_rng(seed)
        best_indices = []
        best_det = -np.inf

        for _ in range(n_starts):
            indices, det_val = _d_exchange_run(X_full, n_profiles, rng)
            if det_val > best_det:
                best_det = det_val
                best_indices = indices

    # 結果の整形（行インデックス順でソート）
    sorted_idx = sorted(best_indices)
    out = df_full.iloc[sorted_idx].copy()
    out.index = [f"{profile_id_prefix}{i + 1}" for i in range(n_profiles)]
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

    if balance_info is not None:
        out.attrs["auto_balance"] = balance_info

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

    あわせて **要約を印字する**（choice の ``suggest_n_respondents()`` と
    揃えた振る舞い）。印字するのは (1) 推奨値の根拠となる3基準、(2) 推奨された
    n_profiles で「水準バランスと D 最適性が両立するか」、(3) 水準バランスと
    属性間相関の指摘がどちらも出ない n はどれか、の3点。**返り値の
    DataFrame は印字の有無にかかわらず同じ**である。

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

    **バランスと D 最適性が両立するかの判定**

    推奨された n_profiles について、``design_profiles(auto_balance=True)`` を
    実際に実行して次のどれかを印字する。

    * 両立する   … D 最適な設計がそのまま水準バランスも満たす
      （``auto_balance`` は不要）
    * 両立しない … その n では D 最適な設計が必ず不均衡になる。
      ``auto_balance=True`` でバランスは取れるが det(X'X) が何 % に下がるかを示す
    * 判定せず   … 候補が多く判定に時間がかかるため確認していない

    判定するのは、候補の選び方の総数 C(N, n) が ``_REPORT_MAX_COMBINATIONS``
    （= 100,000、実測で約0.4秒）以下の場合だけである。この関数は即座に返る
    関数であり、待たされる理由が利用者に見えないため。``design_profiles`` が
    総当たりに入る上限（``_EXHAUSTIVE_MAX_COMBINATIONS`` = 1,000,000）とは
    別の、より低い値を使う。授業で使う規模はこの上限に収まる。
    判定するのは表に現れる**相異なる推奨値**についてのみで、
    表の全行について計算はしない（同じ n が繰り返されるため）。

    **どの n なら指摘が出ないか**

    続けて、``p`` から ``min(N, max_burden)`` までの n を実際に走査し、
    **水準バランスと属性間相関の指摘がどちらも出ない n** を列挙する
    （``design_profiles(..., auto_balance=True)`` で作った場合）。
    「n_profiles を変えると両立する場合があります」だけでは増やせばよいのか
    減らせばよいのかが分からないため、実際に調べて示す。答えが推奨値より
    小さい側にあることもある。

    走査するのも C(N, n) が ``_REPORT_MAX_COMBINATIONS`` 以下の n だけで、
    超える n は調べずにその旨を印字する（1つも調べられない場合はこの行を
    印字しない）。なお、プロファイル数が少ないときの警告
    （``check_design`` の ``few_profiles``）はバランス・相関とは
    **別のカテゴリ**なので、列挙した n に残る場合はそれを明示する。

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
        [n_respondents] if n_respondents is not None else [5, 10, 20, 30, 50, 100]
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

        rows.append(
            {
                "回答者数": n_resp,
                # obs/pred 条件と統計的最低限の両方を満たす最小 M
                f"obs/pred≥{obs_per_predictor}（最低限）": min(
                    max(m_obs_only, m_min), N
                ),
                "推奨 n_profiles": m_rec,
                "obs/pred（達成）": actual_ratio,
                "観測数 obs": actual_obs,
            }
        )

    result = pd.DataFrame(rows)
    result.attrs.update(
        {
            "n_params": p,
            "n_encoded": n_encoded,
            "n_candidates": N,
            "m_min": m_min,
            "m_orme": min(m_orme, N),
        }
    )

    # 要約を印字する（choice の suggest_n_respondents と揃えた振る舞い）。
    # 返り値の DataFrame には手を加えない。
    print(
        "n_profiles の 3 基準: 統計的最低限 p / "
        "Orme (2010) の経験則（符号化列数 × 2）/ 観測数条件"
    )
    print(
        f"  p = {p}, Orme の目安 = {n_encoded} × 2 = {m_orme}, "
        f"観測数条件 obs/pred ≥ {obs_per_predictor}（回答者数しだい・表の列を参照）"
    )
    recommended = sorted({int(row["推奨 n_profiles"]) for row in rows})
    joined = ", ".join(str(v) for v in recommended)
    print(
        f"  → 推奨 n_profiles: {joined}"
        f"（3基準の最大値を min(N={N}, max_burden={max_burden}) で上限制限）"
    )
    # 「指摘が出ないのはどの n か」を先に調べる。これを印字するなら、
    # _balance_note の「n_profiles を変えると両立する場合があります」は
    # 方向を示さないぶん誤解を招くので出さない。
    clean_note = _clean_n_note(attribute_levels, m_min, N, max_burden)
    for m_rec in recommended:
        print(
            _balance_note(attribute_levels, m_rec, N, show_change_hint=not clean_note)
        )
    if clean_note:
        print(clean_note)

    return result


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _format_ratio(ratio: float) -> str:
    """det の比を「75%」「93.75%」のように、余分な 0 を付けずに表す。"""
    return f"{ratio * 100:.4g}%"


# 候補の組み合わせ数をカンマ区切りで書く上限。これを超えたら指数表記にする。
# C(N, n) は桁数が急に増える（8属性×4水準の N = 65,536 から 25 個選ぶ場合は
# 96桁、カンマ区切りだと127文字になり画面を1行で埋める）。授業で使う規模は
# 10桁程度までなので、そこはカンマ区切りのまま読める。
_COUNT_COMMA_MAX = 10**12


def _format_count(count: int) -> str:
    """組み合わせ数を「17,383,860」または「約 1.66×10^95」（96桁）の形で表す。"""
    if count < _COUNT_COMMA_MAX:
        return f"{count:,}"
    exponent = len(str(count)) - 1
    mantissa = count / 10**exponent
    return f"約 {mantissa:.2f}×10^{exponent}"


def _balance_note(
    attribute_levels: Dict[str, List[Any]],
    n_profiles: int,
    n_candidates: int,
    *,
    show_change_hint: bool = True,
) -> str:
    """推奨 n_profiles で「バランスと D 最適性が両立するか」の注記を作る。

    判定には ``design_profiles(auto_balance=True)`` をそのまま使う。ただし
    C(N, n) が ``_REPORT_MAX_COMBINATIONS`` 以下のときだけで、それを超える
    ときは判定しない。``suggest_n_profiles()`` は即座に返る関数であり、
    待たされる理由が利用者に見えないため。``design_profiles`` 自身が総当たり
    に入る上限（``_EXHAUSTIVE_MAX_COMBINATIONS``）とは別の、より低い値である。

    ``show_change_hint=False`` にすると「n_profiles を変えると両立する場合が
    あります」の行を落とす。どの n なら指摘が出ないかを :func:`_clean_n_note`
    が具体的に列挙するときは、方向を示さないこの行は不要（かえって
    「増やせばよい」と読まれる）ため。
    """
    n_comb = math.comb(n_candidates, n_profiles)
    if n_comb > _REPORT_MAX_COMBINATIONS:
        return (
            f"    ※ n = {n_profiles} でバランスと D 最適性が両立するかは"
            "確認していません\n"
            f"       （候補の組み合わせが {_format_count(n_comb)} 通りあり、"
            "判定に時間がかかるため）。\n"
            "       design_profiles(..., auto_balance=True) を実行し、\n"
            '       df.attrs["auto_balance"]["det_ratio"] で確認できます。'
        )

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            design = design_profiles(attribute_levels, n_profiles, auto_balance=True)
        except ValueError:
            return (
                f"    ※ n = {n_profiles} ではバランスと D 最適性の両立を"
                "判定できませんでした。"
            )
    info = design.attrs["auto_balance"]

    if not info["balanced"]:
        return (
            f"    ※ n = {n_profiles} では、水準を均等にできる設計が存在しません。\n"
            "       n_profiles を変えてください"
            "（各属性の水準数の公倍数に近い値が候補です）。"
        )
    if info["det_ratio"] == 1.0:
        return (
            f"    ※ n = {n_profiles} なら、D 最適な設計がそのまま水準バランスも"
            "満たします\n"
            "       （auto_balance=True を指定する必要はありません）。"
        )
    note = (
        f"    ※ n = {n_profiles} では、D 最適な設計は必ずどこかの水準が"
        "不均衡になります。\n"
        "       design_profiles(..., auto_balance=True) ならバランスは取れますが、\n"
        f"       det(X'X) は制約なしの {_format_ratio(info['det_ratio'])} に下がります。"
    )
    if show_change_hint:
        note += "\n       n_profiles を変えると両立する場合があります。"
    return note


def _scan_range(
    m_min: int, n_upper: int, n_candidates: int
) -> "tuple[List[int], List[int]]":
    """走査する n と、候補が多すぎて調べない n を返す（どちらも m_min 以上 n_upper 以下）。

    C(N, n) は n = N/2 まで増え、そこから減る。増える側では上限を超えた時点で
    それより大きい n（N/2 まで）もすべて超えるので、そこで打ち切る。減る側は
    大きいほうから見て同様に打ち切る。**全部の n について C を計算しない**の
    が要点で、N が大きいと C の計算自体が重い（例：8属性×4水準なら
    N = 65,536）。
    """
    mid = n_candidates // 2
    scanned: List[int] = []
    skipped: List[int] = []

    # 増える側（m_min 〜 N/2）
    low_end = min(n_upper, mid)
    n = m_min
    while n <= low_end:
        if math.comb(n_candidates, n) > _REPORT_MAX_COMBINATIONS:
            skipped.extend(range(n, low_end + 1))
            break
        scanned.append(n)
        n += 1

    # 減る側（N/2 より大きい側を、大きいほうから）
    high_end = max(mid, m_min - 1)
    n = n_upper
    while n > high_end:
        if math.comb(n_candidates, n) > _REPORT_MAX_COMBINATIONS:
            skipped.extend(range(high_end + 1, n + 1))
            break
        scanned.append(n)
        n -= 1

    return sorted(scanned), sorted(skipped)


def _clean_n_note(
    attribute_levels: Dict[str, List[Any]],
    m_min: int,
    n_candidates: int,
    max_burden: int,
) -> str:
    """水準バランスと属性間相関の指摘がどちらも出ない n を、走査して列挙する。

    「n_profiles を変えると両立する場合があります」だけでは、増やせばよいのか
    減らせばよいのかが分からない。判定の仕組みはすでにあるので、実際に調べて
    どの n かを示す（答えが推奨値より小さい側にあることもある）。

    走査する範囲は m_min（= p）から ``min(N, max_burden)`` まで。推奨値と同じ
    上限制限をかけるのは、そこを超える n は推奨されず、示しても使えないため。
    さらに C(N, n) が ``_REPORT_MAX_COMBINATIONS`` を超える n は調べない
    （:func:`_scan_range`）。1つも調べられなければ空文字列を返す
    （呼び出し側は何も印字しない）。

    判定は ``design_profiles(..., auto_balance=True)`` が実際に返す設計を
    ``check_design`` と同じ基準で評価して行う。「存在する」ではなく
    「その関数で作れば出ない」と言えることを確かめるため。

    プロファイル数についての警告（``check_design`` の ``few_profiles``）は
    バランス・相関とは別のカテゴリなので、「この n なら警告が消えます」と
    まとめずに、残るものとして分けて示す。
    """
    # 循環 import（design → _feasibility → design）を避けるため呼び出し時に読み込む。
    from . import _feasibility

    n_upper = max(m_min, min(n_candidates, max_burden))
    scanned, skipped = _scan_range(m_min, n_upper, n_candidates)
    if not scanned:
        return ""

    clean: List[int] = []
    for n in scanned:
        judged = _feasibility._auto_balance_warnings(attribute_levels, n)
        if judged is None:
            continue
        warned_balance, warned_correlation = judged
        if not any(warned_balance.values()) and not warned_correlation:
            clean.append(n)

    skip_note = ""
    if skipped:
        skip_note = (
            f"\n       （n = {_join_ints(skipped)} は候補の組み合わせが "
            f"{_REPORT_MAX_COMBINATIONS:,} 通りを超えるため調べていません。）"
        )

    if not clean:
        return (
            f"    ※ n = {scanned[0]}〜{scanned[-1]} のどれで作っても、水準バランスと"
            "属性間相関の指摘が\n"
            "       どちらも出ない設計は見つかりませんでした。" + skip_note
        )

    note = (
        "    ※ 水準バランスと属性間相関の指摘がどちらも出ないのは "
        f"n = {_join_ints(clean)} です\n"
        "       （design_profiles(..., auto_balance=True) で作った場合）。"
    )
    few = [n for n in clean if n < m_min + 2]
    if few:
        note += (
            f"\n       ただし n = {_join_ints(few)} では、プロファイル数が"
            f"パラメータ数（{m_min}）に対して\n"
            "       少ないという指摘は残ります"
            "（バランス・相関とは別のカテゴリの警告です）。"
        )
    return note + skip_note


def _join_ints(values: List[int]) -> str:
    """整数の並びを「4, 8」の形にする。3つ以上連続する部分は「12〜18」とまとめる。"""
    if not values:
        return ""
    # 連続する部分（run）に分ける
    runs: List[List[int]] = [[values[0]]]
    for v in values[1:]:
        if v == runs[-1][-1] + 1:
            runs[-1].append(v)
        else:
            runs.append([v])

    parts: List[str] = []
    for run in runs:
        if len(run) >= 3:
            parts.append(f"{run[0]}〜{run[-1]}")
        else:
            parts.extend(str(v) for v in run)
    return ", ".join(parts)


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
                [1.0 if v == lv else (-1.0 if v == ref else 0.0) for v in df[attr]],
                dtype=float,
            )
            cols.append(col)
    return np.column_stack(cols)


# ---------------------------------------------------------------------------
# 内部ヘルパー（auto_balance）
# ---------------------------------------------------------------------------

# 総当たりに切り替える上限（候補の選び方の総数 C(N, n_profiles)）。
# 速度の閾値であると同時に、厳密解か発見的探索かの境界でもある。
# 根拠となる実測値は design_profiles の Notes を参照。
_EXHAUSTIVE_MAX_COMBINATIONS = 1_000_000

# 報告する関数（check_design / suggest_n_profiles）が「この n では回避できるか」
# を判定する上限。上の _EXHAUSTIVE_MAX_COMBINATIONS より低くしてある。
#
# design_profiles(auto_balance=True) は利用者が明示的に最適化を依頼した関数
# なので、1.5秒程度は許容される。一方 check_design と suggest_n_profiles は
# 報告する関数で、即座に返ることが期待されている。しかも所要時間が設計の
# 大きさで変わり、利用者には理由が見えない。
#
# 100,000 は実測で約0.4秒。授業規模は最大でも C = 43,758（3×3×2 の n=10）
# なので、この上限でも教材での価値は失われない。
_REPORT_MAX_COMBINATIONS = 100_000

# 総当たりを何件ずつまとめて numpy に流すか（メモリと速度の兼ね合い）
_EXHAUSTIVE_BATCH = 20_000

# 制約付き探索で、バランスを満たす初期解を作り直す上限回数
_BALANCED_INIT_TRIES = 50


def _level_indicator(
    df_full: pd.DataFrame,
    attribute_levels: Dict[str, List[Any]],
) -> "Tuple[np.ndarray, List[Tuple[int, int]]]":
    """
    水準の出現回数を数えるための 0/1 行列（N × 全水準数）と、属性ごとの列範囲を返す。

    選んだ行の和をとれば、各水準が何回現れたかがそのまま得られる。
    """
    cols: List[np.ndarray] = []
    spans: List[Tuple[int, int]] = []
    start = 0
    for attr, levels in attribute_levels.items():
        for lv in levels:
            cols.append((df_full[attr] == lv).to_numpy(dtype=np.int32))
        spans.append((start, start + len(levels)))
        start += len(levels)
    return np.column_stack(cols), spans


def _level_codes(
    df_full: pd.DataFrame,
    attribute_levels: Dict[str, List[Any]],
) -> "Tuple[np.ndarray, Dict[Tuple[int, ...], int]]":
    """
    各候補プロファイルを「属性ごとの水準番号」の配列にし、その逆引きも返す。

    属性の水準だけを入れ替える操作（:func:`_balanced_exchange_run` の操作2）で、
    入れ替えた結果のプロファイルが候補の何番目かを引くために使う。
    """
    cols = []
    for attr, levels in attribute_levels.items():
        pos = {lv: i for i, lv in enumerate(levels)}
        cols.append(np.array([pos[v] for v in df_full[attr]], dtype=np.int64))
    codes = np.column_stack(cols)
    index_of = {tuple(code): i for i, code in enumerate(codes)}
    return codes, index_of


def _is_balanced(counts: np.ndarray, spans: "List[Tuple[int, int]]") -> bool:
    """水準の出現回数（1次元）がすべての属性で「最大 − 最小 ≤ 1」か。"""
    return all(counts[s:e].max() - counts[s:e].min() <= 1 for s, e in spans)


def _unbalanced_attrs(
    counts: np.ndarray,
    spans: "List[Tuple[int, int]]",
    attribute_levels: Dict[str, List[Any]],
) -> List[str]:
    """均等にできなかった属性の名前を返す（警告文に使う）。"""
    names = list(attribute_levels.keys())
    return [
        names[a]
        for a, (s, e) in enumerate(spans)
        if counts[s:e].max() - counts[s:e].min() > 1
    ]


def _exhaustive_search(
    X_full: np.ndarray,
    L: np.ndarray,
    spans: "List[Tuple[int, int]]",
    M: int,
) -> "Tuple[Optional[List[int]], float, List[int], float]":
    """
    C(N, M) 通りをすべて調べ、「制約なしの最良」と「バランス制約下の最良」を返す。

    全解を漏れなく走査するので、返るのは **厳密解** である。
    同点のときは候補の並び順（辞書順）で最初のものを採るため、結果は決定的で
    seed に依存しない。

    Returns
    -------
    (balanced_indices, balanced_det, best_indices, best_det)
        balanced_indices は制約を満たす設計が1つも無ければ None。
    """
    N = X_full.shape[0]
    best_idx: List[int] = []
    best_det = -np.inf
    bal_idx: Optional[List[int]] = None
    bal_det = -np.inf

    buf: List[Tuple[int, ...]] = []

    def flush(buf_local: "List[Tuple[int, ...]]") -> None:
        nonlocal best_idx, best_det, bal_idx, bal_det
        idx = np.array(buf_local, dtype=np.int64)  # (B, M)
        rows = X_full[idx]  # (B, M, p)
        dets = np.linalg.det(np.einsum("bmp,bmq->bpq", rows, rows))  # (B,)

        k = int(np.argmax(dets))
        if float(dets[k]) > best_det:
            best_det = float(dets[k])
            best_idx = list(buf_local[k])

        counts = L[idx].sum(axis=1)  # (B, 全水準数)
        ok = np.ones(len(idx), dtype=bool)
        for s, e in spans:
            seg = counts[:, s:e]
            ok &= (seg.max(axis=1) - seg.min(axis=1)) <= 1
        if ok.any():
            dets_bal = np.where(ok, dets, -np.inf)
            kb = int(np.argmax(dets_bal))
            if float(dets_bal[kb]) > bal_det:
                bal_det = float(dets_bal[kb])
                bal_idx = list(buf_local[kb])

    for combo in _itertools_combinations(range(N), M):
        buf.append(combo)
        if len(buf) == _EXHAUSTIVE_BATCH:
            flush(buf)
            buf = []
    if buf:
        flush(buf)

    return bal_idx, bal_det, best_idx, best_det


def _balanced_initial(
    L: np.ndarray,
    spans: "List[Tuple[int, int]]",
    M: int,
    rng: np.random.Generator,
) -> Optional[List[int]]:
    """
    バランスを満たす初期解を作る（貪欲法）。

    候補をランダムな順に見て、「出現回数が上限 ceil(M / 水準数) を超えない」
    ものの中から、現在の出現回数の合計が最も小さいものを選ぶ。
    作れなければ ``_BALANCED_INIT_TRIES`` 回まで作り直し、それでも駄目なら
    None を返す。
    """
    N = L.shape[0]
    caps = np.zeros(L.shape[1], dtype=np.int64)
    for s, e in spans:
        caps[s:e] = -(-M // (e - s))  # ceil

    for _ in range(_BALANCED_INIT_TRIES):
        order = rng.permutation(N)
        counts = np.zeros(L.shape[1], dtype=np.int64)
        used = np.zeros(N, dtype=bool)
        sel: List[int] = []
        for _step in range(M):
            best_i, best_key = None, None
            for i in order:
                if used[i] or np.any(counts + L[i] > caps):
                    continue
                key = int((counts * L[i]).sum())
                if best_key is None or key < best_key:
                    best_i, best_key = int(i), key
            if best_i is None:
                break
            counts += L[best_i]
            used[best_i] = True
            sel.append(best_i)
        if len(sel) == M and _is_balanced(counts, spans):
            return sel
    return None


def _balanced_exchange_run(
    X_full: np.ndarray,
    L: np.ndarray,
    spans: "List[Tuple[int, int]]",
    codes: np.ndarray,
    index_of: Dict[Tuple[int, ...], int],
    M: int,
    rng: np.random.Generator,
) -> "Optional[Tuple[List[int], float]]":
    """
    バランスを保ったまま局所改善する 1 試行（発見的探索）。

    バランスを満たす初期解から出発し、**バランスを保つ操作だけ** で
    det(X'X) を改善していく。使う操作は次の2種類。

    1. **入れ替え**：選択済みの1行を未選択の1行と交換する。
       交換後もバランスを満たす組み合わせだけを候補にする。
    2. **属性の交換**：選択済みの2行のあいだで、ある属性の水準だけを
       入れ替える（例：(6, apple) と (10, android) → (10, apple) と
       (6, android)）。どの属性の出現回数も変わらないので、バランスは
       つねに保たれる。

    操作2が必要なのは、水準数が n_profiles を割り切る「完全に均等な」
    設計では、操作1がまったく使えなくなるためである（1行だけ入れ替えると
    必ずどこかの水準の回数が ±1 ずれ、最大と最小の差が 2 になる）。
    操作1だけだと、その場合は初期解から1歩も動けない。

    最良解である保証はない（バランスを満たす解の近傍で局所最適）。
    """
    sel = _balanced_initial(L, spans, M, rng)
    if sel is None:
        return None

    N = X_full.shape[0]
    in_sel = np.zeros(N, dtype=bool)
    in_sel[sel] = True
    counts = L[sel].sum(axis=0)

    def det_of(indices: List[int]) -> float:
        rows = X_full[indices]
        return float(np.linalg.det(rows.T @ rows))

    current = det_of(sel)
    improved = True
    while improved:
        improved = False
        best_det = current * (1.0 + 1e-10)
        best_sel: Optional[List[int]] = None

        # 操作1：選択済みの1行 ↔ 未選択の1行
        not_sel = [j for j in range(N) if not in_sel[j]]
        for i_pos in range(M):
            base_counts = counts - L[sel[i_pos]]
            for j in not_sel:
                if not _is_balanced(base_counts + L[j], spans):
                    continue
                cand = list(sel)
                cand[i_pos] = j
                d = det_of(cand)
                if d > best_det:
                    best_det, best_sel = d, cand

        # 操作2：選択済みの2行のあいだで、1属性の水準だけを入れ替える
        n_attrs = len(spans)
        for a in range(M):
            for b in range(a + 1, M):
                code_a = codes[sel[a]]
                code_b = codes[sel[b]]
                for k in range(n_attrs):
                    if code_a[k] == code_b[k]:
                        continue  # 同じ水準なら何も変わらない
                    new_a = code_a.copy()
                    new_b = code_b.copy()
                    new_a[k], new_b[k] = code_b[k], code_a[k]
                    i_new = index_of[tuple(new_a)]
                    j_new = index_of[tuple(new_b)]
                    # 入れ替えた結果、他の選択済み行と重複してはいけない
                    if in_sel[i_new] or in_sel[j_new]:
                        continue
                    cand = list(sel)
                    cand[a], cand[b] = i_new, j_new
                    d = det_of(cand)
                    if d > best_det:
                        best_det, best_sel = d, cand

        if best_sel is not None:
            in_sel[:] = False
            in_sel[best_sel] = True
            counts = L[best_sel].sum(axis=0)
            sel = best_sel
            current = best_det
            improved = True

    return sel, current


def _search_balanced_design(
    X_full: np.ndarray,
    df_full: pd.DataFrame,
    attribute_levels: Dict[str, List[Any]],
    M: int,
    n_starts: int,
    seed: Optional[int],
) -> "Tuple[List[int], Dict[str, Any]]":
    """
    バランス制約下で det(X'X) を最大化する設計を探し、来歴とともに返す。

    候補の選び方の総数が ``_EXHAUSTIVE_MAX_COMBINATIONS`` 以下なら総当たり
    （厳密解）、超えるなら制約付き交換アルゴリズム（発見的探索）を使う。

    バランスを満たす設計が得られなかった場合は、警告のうえ制約なしの
    最良解を返す（例外にはしない）。
    """
    N = X_full.shape[0]
    L, spans = _level_indicator(df_full, attribute_levels)

    if math.comb(N, M) <= _EXHAUSTIVE_MAX_COMBINATIONS:
        method = "exhaustive"
        bal_idx, bal_det, best_idx, best_det = _exhaustive_search(X_full, L, spans, M)
    else:
        method = "exchange"
        rng = np.random.default_rng(seed)
        best_idx, best_det = [], -np.inf
        for _ in range(n_starts):
            indices, det_val = _d_exchange_run(X_full, M, rng)
            if det_val > best_det:
                best_det, best_idx = det_val, indices

        codes, index_of = _level_codes(df_full, attribute_levels)
        bal_idx, bal_det = None, -np.inf
        for _ in range(n_starts):
            found = _balanced_exchange_run(X_full, L, spans, codes, index_of, M, rng)
            if found is not None and found[1] > bal_det:
                bal_idx, bal_det = found[0], found[1]

    if bal_idx is None:
        # 到達しにくい経路だが、防御的に残してある（例外にはしない）。
        bad = _unbalanced_attrs(L[best_idx].sum(axis=0), spans, attribute_levels)
        if method == "exhaustive":
            reason = (
                f"属性 {bad} で水準を均等にできる設計が存在しません"
                "（すべての組み合わせを調べました）。"
            )
        else:
            reason = (
                f"属性 {bad} で水準を均等にできる設計を見つけられませんでした"
                "（探索は網羅的ではないため、存在しないとは限りません）。"
            )
        warnings.warn(
            f"auto_balance=True が指定されましたが、{reason}\n"
            "  バランス制約なしの最良解を返します。\n"
            "  n_profiles を変えると均等にできる場合があります"
            "（各属性の水準数の公倍数に近い値が候補です）。",
            UserWarning,
            stacklevel=3,
        )
        info = {
            "method": method,
            "balanced": False,
            "det_xpx": float(best_det),
            "det_xpx_unconstrained": float(best_det),
            # バランスを満たせなかったので比は定義しない。同じ設計を指す
            # 2つの det の比（= 1.0）を入れると「精度の損失なし」と読めてしまい、
            # 制約を満たせなかったことが伝わらないため None にする。
            "det_ratio": None,
        }
        return best_idx, info

    ratio = float(bal_det / best_det) if best_det > 0 else 0.0
    info = {
        "method": method,
        "balanced": True,
        "det_xpx": float(bal_det),
        "det_xpx_unconstrained": float(best_det),
        "det_ratio": ratio,
    }
    return bal_idx, info


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

        X_sel_arr = X_full[sel]  # shape (M, p)
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
            c_all = X_not_arr @ XtX_inv_xi  # (N-M,)

            # 交換後の行列式比
            ratios = (1.0 - d_i) * (1.0 + d_not) + c_all**2
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
