"""
plot.py
========
コンジョイント分析結果の **可視化** を担うモジュール。

3種類のグラフを提供：

* :func:`plot_importance` … 属性の相対重要度（合計100%の棒グラフ）
* :func:`plot_partworth` … 各水準の部分効用（パートワース）
* :func:`plot_wtp` … 各属性のWTP（支払意思額）

すべて matplotlib で描画する。日本語フォントは自動設定を試みる。
"""
from __future__ import annotations

import warnings
from typing import TYPE_CHECKING, Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

if TYPE_CHECKING:
    from .analysis import ConjointResult


# ---------------------------------------------------------------------------
# 日本語フォント設定（ベストエフォート）
# ---------------------------------------------------------------------------

_FONT_INITIALIZED = False


def _ensure_japanese_font() -> None:
    """
    日本語が描画可能なフォントを matplotlib に設定する。
    存在しなければ警告を出すだけで処理は続行する。
    """
    global _FONT_INITIALIZED
    if _FONT_INITIALIZED:
        return

    candidates = [
        "IPAexGothic",
        "IPAGothic",
        "Noto Sans CJK JP",
        "Noto Sans JP",
        "Hiragino Sans",
        "Hiragino Maru Gothic Pro",
        "Yu Gothic",
        "Meiryo",
        "TakaoGothic",
        "VL Gothic",
        "MS Gothic",
    ]
    try:
        from matplotlib import font_manager as fm
        installed = {f.name for f in fm.fontManager.ttflist}
        for name in candidates:
            if name in installed:
                plt.rcParams["font.family"] = name
                plt.rcParams["axes.unicode_minus"] = False
                _FONT_INITIALIZED = True
                return
        warnings.warn(
            "日本語フォントが見つかりませんでした。グラフの日本語が "
            "□ になる場合があります。\n"
            "  Linux の例: sudo apt-get install -y fonts-ipaexfont\n"
            "  Colab の例: !apt-get -y install fonts-ipafont-gothic",
            UserWarning,
            stacklevel=2,
        )
    finally:
        _FONT_INITIALIZED = True


# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------

def plot_importance(
    result: "ConjointResult",
    *,
    ax=None,
    title: str = "属性の相対重要度",
    color: str = "#4C78A8",
    show_values: bool = True,
    sort: bool = True,
):
    """
    各属性の **相対重要度** を棒グラフで描画する。

    Parameters
    ----------
    result : ConjointResult
        :func:`fit` で得られた結果オブジェクト。
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
        imp = imp.sort_values("importance", ascending=True)

    if ax is None:
        fig, ax = plt.subplots(figsize=(7, max(2.5, 0.6 * len(imp) + 1.5)))

    ax.barh(imp.index, imp["importance"], color=color)
    ax.set_xlabel("相対重要度（%）")
    ax.set_title(title)
    ax.set_xlim(0, max(100, imp["importance"].max() * 1.15))

    if show_values:
        for y, v in enumerate(imp["importance"]):
            ax.text(v + 1.0, y, f"{v:.1f}%", va="center", fontsize=10)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return ax


def plot_partworth(
    result: "ConjointResult",
    *,
    ax=None,
    title: str = "部分効用（パートワース）",
    show_zero_line: bool = True,
):
    """
    各属性の各水準の **部分効用** を棒グラフで描画する。

    効果コーディングのもとでは：

    * 2水準の場合、係数 ``b`` がそのまま「+1側水準」の部分効用、
      ``-b`` が「基準水準」の部分効用。
    * 3水準以上の場合、各非基準水準の係数 ``b_k`` がその水準の部分効用、
      基準水準は ``-Σ b_k``。

    Parameters
    ----------
    result : ConjointResult
    ax : matplotlib.axes.Axes, optional
    title : str
    show_zero_line : bool, default True
        ゼロを示す垂直線（基準）を描画するか。

    Returns
    -------
    matplotlib.axes.Axes

    Notes
    -----
    水準名は符号化列名（例：``price_6``, ``os_apple``）から復元する。
    基準水準は ``{属性名}（基準）`` のように表示される。
    """
    _ensure_japanese_font()

    # 各属性ごとに部分効用を計算
    rows = []
    groups = _group_columns(result)
    for attr, cols in groups.items():
        bs = np.array([float(result.params[c]) for c in cols])
        # 各非基準水準
        for c, b in zip(cols, bs):
            level = c[len(attr) + 1:] if c.startswith(f"{attr}_") else c
            rows.append({"attribute": attr, "level": level, "partworth": float(b)})
        # 基準水準
        ref_value = -float(bs.sum())
        ref_label = _reference_level_label(result, attr)
        rows.append(
            {"attribute": attr, "level": ref_label, "partworth": ref_value}
        )

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

    labels = [f"{a} = {l}" for a, l in zip(df_pw["attribute"], df_pw["level"])]
    ax.barh(labels, df_pw["partworth"], color=colors)
    if show_zero_line:
        ax.axvline(0, color="gray", linewidth=0.8)
    ax.set_xlabel("部分効用（評点ポイント）")
    ax.set_title(title)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return ax


def plot_wtp(
    result: "ConjointResult",
    *,
    ax=None,
    title: Optional[str] = None,
    color: str = "#E45756",
    show_values: bool = True,
    sort: bool = True,
    price_unit: Optional[str] = None,
):
    """
    各非価格属性の **WTP（支払意思額）** を棒グラフで描画する。

    Parameters
    ----------
    result : ConjointResult
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
    wtp = result.wtp()
    if sort:
        wtp = wtp.sort_values("wtp")

    if ax is None:
        fig, ax = plt.subplots(figsize=(7, max(2.5, 0.5 * len(wtp) + 1.5)))

    ax.barh(wtp.index, wtp["wtp"], color=color)
    ax.axvline(0, color="gray", linewidth=0.8)

    label_unit = f"（{price_unit}）" if price_unit else ""
    ax.set_xlabel(f"WTP（支払意思額）{label_unit}")
    ax.set_title(title or "属性のWTP（支払意思額）")

    if show_values:
        max_abs = max(abs(wtp["wtp"]).max(), 1e-9)
        for y, v in enumerate(wtp["wtp"]):
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

def _group_columns(result: "ConjointResult") -> dict:
    """
    結果オブジェクトから属性 → 符号化列のマッピングを取得する。
    """
    from .analysis import _group_columns_by_attribute
    return _group_columns_by_attribute(
        result.encoded_columns, list(result.reference_levels.keys())
    )


def _reference_level_label(result: "ConjointResult", attr: str) -> str:
    """
    属性 ``attr`` の基準水準を「{値}（基準）」のラベルにして返す。
    reference_levels に登録されていれば値を、なければ「基準」とだけ表示。
    """
    if attr in result.reference_levels:
        return f"{result.reference_levels[attr]}（基準）"
    return "（基準）"
