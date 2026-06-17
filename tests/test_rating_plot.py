"""rating/plot.py（可視化）のスモークテスト。

特に今回の表示改善を検証する：
- plot_partworth：基準水準が「（基準）」ラベル＋ひし形マーカーで示され、
  効果コーディングでは基準も実値（−Σb）の棒として描かれる（歯抜けにならない）。
- plot_wtp：数値ラベルが棒の外側に置かれ、x 軸範囲に余白が確保される
  （正負どちらの棒でもラベルが軸・枠と重ならない）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import matplotlib

matplotlib.use("Agg")  # 描画ウィンドウを開かないバックエンド

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pytest

import py4conjoint.rating as pcr


@pytest.fixture(autouse=True)
def _close_figures():
    yield
    plt.close("all")


@pytest.fixture(scope="module")
def result():
    """価格2水準 + ブランド3水準 + OS2水準の評点データを推定した結果。

    価格を2水準にして WTP を単一の横棒グラフ（区間別でない）にする。
    ブランド C は基準 A より不人気にして **負の WTP** を作り、
    左向きの棒でもラベルが軸と重ならないことを確かめられるようにする。
    """
    rng = np.random.default_rng(0)
    rows = []
    for r in range(1, 31):
        for price in [6, 10]:
            for brand in ["A", "B", "C"]:
                for os in ["android", "apple"]:
                    u = 5.0 - 0.4 * price
                    u += {"A": 0.0, "B": 0.6, "C": -0.5}[brand]
                    u += 0.7 if os == "apple" else -0.7
                    u += rng.normal(0, 0.3)
                    rows.append({"respondent_id": r,
                                 "rating": float(np.clip(u, 1, 10)),
                                 "price": price, "brand": brand, "os": os})
    df = pd.DataFrame(rows)
    dc = pcr.encode(df, reference_levels={"price": 10, "brand": "A", "os": "android"})
    return pcr.fit(dc, price_col="price")


# ---------------------------------------------------------------------------
# plot_partworth
# ---------------------------------------------------------------------------

def test_plot_partworth_smoke(result):
    ax = result.plot_partworth()
    assert ax is not None
    assert ax.get_title() == "部分効用（パートワース）"
    labels = [t.get_text() for t in ax.get_yticklabels()]
    assert any("（基準）" in l for l in labels)
    # 効果コーディングでは全水準が棒になる：price 2 + brand 3 + os 2 = 7 本
    assert len(ax.patches) == 7


def test_plot_partworth_reference_marker(result):
    """基準水準はひし形マーカーで示され、効果コーディングでは −Σb の実値に出る。"""
    ax = result.plot_partworth()
    assert len(ax.collections) >= 1
    legend = ax.get_legend()
    assert legend is not None
    assert any("基準水準" in t.get_text() for t in legend.get_texts())
    assert any(line.get_linestyle() in (":", "dotted") for line in ax.lines)
    # 基準水準は price・brand・os の3つ。効果コーディングなので x は 0 でない。
    offsets = ax.collections[0].get_offsets()
    assert offsets.shape[0] == 3
    assert all(abs(float(x)) > 1e-9 for x, _ in offsets)


# ---------------------------------------------------------------------------
# plot_wtp：ラベル外側配置・xlim 余白
# ---------------------------------------------------------------------------

def test_plot_wtp_xlim_has_padding(result):
    """棒の外側ラベルが収まるよう、x 軸範囲が棒の最大値より広い。"""
    ax = result.plot_wtp(price_unit="万円")
    wtp = result.wtp()
    vmax = float(wtp["限界支払意思額"].max())
    vmin = float(wtp["限界支払意思額"].min())
    lo, hi = ax.get_xlim()
    # 正の棒があれば右側に、負の棒があれば左側に余白がある
    if vmax > 0:
        assert hi > vmax
    if vmin < 0:
        assert lo < vmin


def test_plot_functions_accept_ax(result):
    fig, axes = plt.subplots(1, 2, figsize=(12, 4))
    ax1 = pcr.plot_partworth(result, ax=axes[0])
    ax2 = pcr.plot_wtp(result, ax=axes[1])
    assert ax1 is axes[0]
    assert ax2 is axes[1]
