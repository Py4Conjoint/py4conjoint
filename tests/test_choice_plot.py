"""choice/plot.py（可視化）のスモークテスト。

例外なく描画でき、Axes が返り、日本語ラベルが設定されることを確認する。
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

import py4conjoint.choice as pcc

# ---------------------------------------------------------------------------
# フィクスチャ：小規模な人工データの推定結果
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def result():
    """価格 + 3水準ブランドの小さな選択データを推定した結果を返す。"""
    rng = np.random.default_rng(0)
    n_sets, n_alts = 500, 3
    price = rng.choice([2.0, 3.0, 4.0], size=(n_sets, n_alts))
    brand = rng.choice(["A", "B", "C"], size=(n_sets, n_alts))
    v = -0.8 * price + 0.5 * (brand == "B") - 0.4 * (brand == "C")
    u = v + rng.gumbel(size=(n_sets, n_alts))
    chosen = u.argmax(axis=1)

    df = pd.DataFrame({
        "選択セットID": np.repeat(np.arange(n_sets), n_alts),
        "choice": (np.tile(np.arange(n_alts), n_sets)
                   == np.repeat(chosen, n_alts)).astype(int),
        "price": price.ravel(),
        "brand": brand.ravel(),
    })
    df_coded = pcc.encode(df, reference_levels={"brand": "A"})
    return pcc.fit(df_coded, choice="choice", choice_set_id_col="選択セットID")


@pytest.fixture(autouse=True)
def _close_figures():
    yield
    plt.close("all")


# ---------------------------------------------------------------------------
# スモークテスト
# ---------------------------------------------------------------------------

def test_plot_importance_smoke(result):
    ax = result.plot_importance()
    assert ax is not None
    assert ax.get_xlabel() == "重要度（%）"
    assert ax.get_title() == "属性の重要度"
    # 属性数（brand, price）と同じ数の棒がある
    assert len(ax.patches) == 2


def test_plot_partworth_smoke(result):
    ax = result.plot_partworth()
    assert ax is not None
    assert ax.get_title() == "部分効用（パートワース）"
    labels = [t.get_text() for t in ax.get_yticklabels()]
    # 基準水準（係数0）が明示的に表示される
    assert any("（基準）" in l for l in labels)
    # 数値変数 price も表示される
    assert any("price" in l for l in labels)
    # 棒の数 = brand 3水準（基準含む） + price 1本
    assert len(ax.patches) == 4


def test_plot_wtp_smoke(result):
    ax = result.plot_wtp(price_unit="ドル")
    assert ax is not None
    assert ax.get_xlabel() == "限界支払意思額（ドル）"
    assert ax.get_title() == "属性の限界支払意思額"
    # 非価格変数（brand_B, brand_C）の2本
    assert len(ax.patches) == 2


def test_plot_functions_accept_ax(result):
    """既存の Axes を渡しても描画できる（rating 版と同じ使い勝手）。"""
    fig, axes = plt.subplots(1, 3, figsize=(15, 4))
    ax1 = pcc.plot_importance(result, ax=axes[0])
    ax2 = pcc.plot_partworth(result, ax=axes[1])
    ax3 = pcc.plot_wtp(result, ax=axes[2])
    assert ax1 is axes[0]
    assert ax2 is axes[1]
    assert ax3 is axes[2]
