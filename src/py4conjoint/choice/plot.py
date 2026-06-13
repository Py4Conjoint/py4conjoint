"""
plot.py（choice 版）
====================
選択型コンジョイント分析（CBC）結果の **可視化** を担うモジュール。

rating 版（:mod:`py4conjoint.rating.plot`）と対称的な3種類のグラフを提供：

* :func:`plot_importance` … 属性の重要度（合計100%の棒グラフ）
* :func:`plot_partworth` … 各水準の部分効用（パートワース）
* :func:`plot_wtp` … 各属性のWTP（限界支払意思額）

すべて matplotlib で描画する。日本語フォントは自動設定を試みる。

rating 版との違い
-----------------
choice 版は **ダミーコーディング（0/1）** を使うため、基準水準の部分効用は
常に 0 になる。:func:`plot_partworth` では基準水準（係数0）も明示的に
表示し、各係数が「基準水準との差」であることを視覚的に確認できる。
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import pandas as pd

# 日本語フォント設定は rating 版の仕組みをそのまま使う
from ..rating.plot import _ensure_japanese_font

if TYPE_CHECKING:
    from .analysis import ChoiceConjointResult


# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------

def plot_importance(
    result: "ChoiceConjointResult",
    *,
    ax=None,
    title: str = "属性の重要度",
    color: str = "#4C78A8",
    show_values: bool = True,
    sort: bool = True,
):
    """
    各属性の **重要度** を棒グラフで描画する。

    Parameters
    ----------
    result : ChoiceConjointResult
        :func:`py4conjoint.choice.fit` で得られた結果オブジェクト。
    ax : matplotlib.axes.Axes, optional
        既存の Axes に描画したい場合に指定。省略時は新規作成。
    title : str
        グラフのタイトル。
    color : str
        棒の色（matplotlibのカラー指定）。
    show_values : bool, default True
        各棒に数値ラベル（％）を表示するか。
    sort : bool, default True
        重要度の高い順にソートするか。

    Returns
    -------
    matplotlib.axes.Axes

    Examples
    --------
    >>> result.plot_importance()
    """
    _ensure_japanese_font()
    imp = result.importance(as_percent=True)
    if sort:
        imp = imp.sort_values("重要度", ascending=True)

    if ax is None:
        fig, ax = plt.subplots(figsize=(7, max(2.5, 0.6 * len(imp) + 1.5)))

    ax.barh(imp.index, imp["重要度"], color=color)
    ax.set_xlabel("重要度（%）")
    ax.set_title(title)
    ax.set_xlim(0, max(100, imp["重要度"].max() * 1.15))

    if show_values:
        for y, v in enumerate(imp["重要度"]):
            ax.text(v + 1.0, y, f"{v:.1f}%", va="center", fontsize=10)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return ax


def plot_partworth(
    result: "ChoiceConjointResult",
    *,
    ax=None,
    title: str = "部分効用（パートワース）",
    show_zero_line: bool = True,
):
    """
    各属性の各水準の **部分効用** を棒グラフで描画する。

    ダミーコーディング（0/1）のもとでは：

    * 基準水準の部分効用は **常に 0**。
    * 各非基準水準の係数 ``b_k`` が、その水準の
      「基準水準と比べた効用の差」をそのまま表す。

    基準水準（係数0）も「{水準名}（基準）」のラベルで明示的に表示するので、
    各棒が **基準との差** であることをグラフ上で確認できる。

    Parameters
    ----------
    result : ChoiceConjointResult
    ax : matplotlib.axes.Axes, optional
    title : str
    show_zero_line : bool, default True
        ゼロ（基準水準の効用）を示す垂直線を描画するか。

    Returns
    -------
    matplotlib.axes.Axes

    Notes
    -----
    水準名は符号化列名（例：``brand_hiland``）から復元する。
    基準水準は ``{水準名}（基準）`` のように表示される。
    価格などの数値（連続）変数は「{列名}（1単位あたり）」のラベルで
    係数をそのまま表示する。
    """
    _ensure_japanese_font()

    groups, numeric_cols = _group_columns(result)

    rows = []
    for attr, cols in groups.items():
        # 各非基準水準（係数 = 基準との効用差）
        for c in cols:
            level = c[len(attr) + 1:] if c.startswith(f"{attr}_") else c
            rows.append({
                "attribute": attr,
                "level": level,
                "partworth": float(result.params[c]),
            })
        # 基準水準（ダミーコーディングでは効用 0）
        rows.append({
            "attribute": attr,
            "level": _reference_level_label(result, attr),
            "partworth": 0.0,
        })
    # 数値（連続）変数は係数をそのまま表示
    for c in numeric_cols:
        rows.append({
            "attribute": c,
            "level": "（1単位あたり）",
            "partworth": float(result.params[c]),
        })

    df_pw = pd.DataFrame(rows)
    # 表示順：属性ごとに固める。属性内では値の小さい順
    df_pw = df_pw.sort_values(["attribute", "partworth"]).reset_index(drop=True)

    if ax is None:
        fig, ax = plt.subplots(figsize=(8, max(3, 0.5 * len(df_pw) + 1)))

    # 属性ごとに色を変える
    attrs = df_pw["attribute"].unique().tolist()
    cmap = plt.get_cmap("tab10")
    color_map = {a: cmap(i % 10) for i, a in enumerate(attrs)}
    colors = [color_map[a] for a in df_pw["attribute"]]

    labels = [
        f"{a} = {l}" if not l.startswith("（") else f"{a}{l}"
        for a, l in zip(df_pw["attribute"], df_pw["level"])
    ]
    ax.barh(labels, df_pw["partworth"], color=colors)
    if show_zero_line:
        ax.axvline(0, color="gray", linewidth=0.8)
    ax.set_xlabel("部分効用（基準水準との差）")
    ax.set_title(title)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return ax


def plot_wtp(
    result: "ChoiceConjointResult",
    *,
    ax=None,
    title: Optional[str] = None,
    color: str = "#E45756",
    show_values: bool = True,
    sort: bool = True,
    price_unit: Optional[str] = None,
):
    """
    各非価格変数の **WTP（限界支払意思額）** を棒グラフで描画する。

    ここでの WTP は厳密には限界支払意思額（MWTP）であり、
    製品全体に対する支払上限額ではない。
    詳細は :meth:`py4conjoint.choice.ChoiceConjointResult.wtp` の「定義」を参照。

    Parameters
    ----------
    result : ChoiceConjointResult
    ax : matplotlib.axes.Axes, optional
    title : str, optional
        省略時は自動生成。
    color : str
    show_values : bool, default True
    sort : bool, default True
    price_unit : str, optional
        棒に表示する単位ラベル（例：``"万円"``）。
        省略時は単位なし（数値のみ）。

    Returns
    -------
    matplotlib.axes.Axes
    """
    _ensure_japanese_font()
    # 棒グラフは属性ごとに1本にするため、線形近似（単一値）の WTP を使う。
    # 価格区間ごとの WTP は result.wtp() で表として確認できる。
    wtp = result.wtp(method="linear")
    if sort:
        wtp = wtp.sort_values("限界支払意思額")

    if ax is None:
        fig, ax = plt.subplots(figsize=(7, max(2.5, 0.5 * len(wtp) + 1.5)))

    ax.barh(wtp.index, wtp["限界支払意思額"], color=color)
    ax.axvline(0, color="gray", linewidth=0.8)

    label_unit = f"（{price_unit}）" if price_unit else ""
    ax.set_xlabel(f"限界支払意思額{label_unit}")
    ax.set_title(title or "属性の限界支払意思額")

    if show_values:
        max_abs = max(abs(wtp["限界支払意思額"]).max(), 1e-9)
        for y, v in enumerate(wtp["限界支払意思額"]):
            offset = 0.02 * max_abs
            x = v + offset if v >= 0 else v - offset
            ha = "left" if v >= 0 else "right"
            label = f"{v:.2f}{price_unit or ''}"
            ax.text(x, y, label, va="center", ha=ha, fontsize=10)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return ax


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------

def _group_columns(
    result: "ChoiceConjointResult",
) -> Tuple[Dict[str, List[str]], List[str]]:
    """
    結果オブジェクトの説明変数を
    「ダミーコーディングした属性 → 符号化列のリスト」と
    「数値（連続）変数のリスト」に分ける。

    分類ロジックは :meth:`ChoiceConjointResult._attribute_ranges` と同じ。
    """
    known = sorted(result.reference_levels.keys(), key=len, reverse=True)
    groups: Dict[str, List[str]] = {}
    numeric_cols: List[str] = []
    for c in result.encoded_columns:
        matched = None
        for a in known:
            if c.startswith(f"{a}_"):
                matched = a
                break
        if matched is None:
            numeric_cols.append(c)
        else:
            groups.setdefault(matched, []).append(c)
    return groups, numeric_cols


def _reference_level_label(result: "ChoiceConjointResult", attr: str) -> str:
    """
    属性 ``attr`` の基準水準を「{値}（基準）」のラベルにして返す。
    reference_levels に登録されていれば値を、なければ「基準」とだけ表示。
    """
    if attr in result.reference_levels:
        return f"{result.reference_levels[attr]}（基準）"
    return "（基準）"
