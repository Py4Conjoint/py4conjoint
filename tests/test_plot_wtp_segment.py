"""plot_wtp が wtp() の表と常に一致することのテスト（rating / choice 共通）。

- 価格2水準：属性ごとに1本の棒グラフ（単一区間）。
- 価格3水準以上（method="segment"）：価格区間ごとに色分けしたグループ化棒グラフ。
- 棒の高さ（WTP）が result.wtp(method=...) の表と一致する。
- method="linear" のときはタイトルに「線形近似」と明示する。
- price_segment で特定区間だけ描ける。
- rating / choice で挙動が完全に揃う。
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
import py4conjoint.rating as pc


@pytest.fixture(autouse=True)
def _close_figures():
    yield
    plt.close("all")


# ---------------------------------------------------------------------------
# 人工データ
# ---------------------------------------------------------------------------

def _simulate_cbc(price_util, brand_util, *, n_sets=30_000, n_alts=3, seed=0):
    rng = np.random.default_rng(seed)
    prices = np.array(sorted(price_util))
    brands = np.array(sorted(brand_util))
    rows = []
    for t in range(n_sets):
        pr = rng.choice(prices, n_alts)
        br = rng.choice(brands, n_alts)
        v = np.array([price_util[p] + brand_util[b] for p, b in zip(pr, br)])
        p = np.exp(v - v.max())
        p /= p.sum()
        ch = (rng.random() < p.cumsum()).argmax()
        for j in range(n_alts):
            rows.append({"choice_set_id": t, "choice": int(j == ch),
                         "price": int(pr[j]), "brand": br[j]})
    return pd.DataFrame(rows)


def _make_rating_3level(seed=42, n_resp=30):
    rng = np.random.default_rng(seed)
    rows = []
    for r in range(1, n_resp + 1):
        for price in [6, 8, 10]:
            for os in ["android", "apple"]:
                u = 4.0 + {6: 1.5, 8: 1.0, 10: -1.5}[price]
                u += 0.7 if os == "apple" else -0.7
                u += rng.normal(0, 0.3)
                rows.append({"回答者ID": r, "rating": int(np.clip(round(u), 1, 7)),
                             "price": price, "os": os})
    return pd.DataFrame(rows)


@pytest.fixture(scope="module")
def choice_3level():
    df = _simulate_cbc({6: 0.0, 8: -0.3, 10: -1.6}, {"A": 0.0, "B": 0.6},
                       n_sets=40_000, seed=5)
    dc = pcc.encode(df, reference_levels={"price": 6, "brand": "A"})
    return pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                   encoded_columns=["price_8", "price_10", "brand_B"],
                   price_col="price")


@pytest.fixture(scope="module")
def choice_2level():
    df = _simulate_cbc({6: 0.0, 10: -1.2}, {"A": 0.0, "B": 0.5},
                       n_sets=30_000, seed=3)
    dc = pcc.encode(df, reference_levels={"price": 10, "brand": "A"})
    return pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                   encoded_columns=["price_6", "brand_B"], price_col="price")


@pytest.fixture(scope="module")
def rating_3level():
    df = _make_rating_3level(seed=42, n_resp=30)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"},
                   suffix_map={"price": ["low", "mid"]})
    return pc.fit(dc, price_col="price")


# ---------------------------------------------------------------------------
# 2水準：属性ごとに1本（単一区間）
# ---------------------------------------------------------------------------

def test_choice_2level_single_bars(choice_2level):
    ax = choice_2level.plot_wtp(price_unit="ドル")
    assert ax is not None
    # 区間がないので属性ごとに1本（brand_B の1本）
    assert len(ax.patches) == 1
    assert ax.get_xlabel() == "限界支払意思額（ドル）"
    assert ax.get_title() == "属性の限界支払意思額"
    # 凡例（価格区間）は付かない
    assert ax.get_legend() is None


# ---------------------------------------------------------------------------
# 3水準：価格区間ごとのグループ化棒グラフ
# ---------------------------------------------------------------------------

def test_choice_3level_grouped_bars(choice_3level):
    ax = choice_3level.plot_wtp()
    assert ax is not None
    # 属性1（brand_B）× 区間2（6〜8, 8〜10）= 2本
    assert len(ax.patches) == 2
    # 凡例に価格区間が表示される
    legend = ax.get_legend()
    assert legend is not None
    labels = {t.get_text() for t in legend.get_texts()}
    assert labels == {"6〜8", "8〜10"}
    assert ax.get_title() == "属性の限界支払意思額（価格区間別）"


def test_rating_3level_grouped_bars(rating_3level):
    ax = rating_3level.plot_wtp()
    assert ax is not None
    # os_0 × 区間2 = 2本
    assert len(ax.patches) == 2
    legend = ax.get_legend()
    assert legend is not None
    labels = {t.get_text() for t in legend.get_texts()}
    assert labels == {"6〜8", "8〜10"}


# ---------------------------------------------------------------------------
# 棒の高さが wtp(method="segment") の表と一致する
# ---------------------------------------------------------------------------

def test_choice_grouped_heights_match_table(choice_3level):
    w = choice_3level.wtp(method="segment")
    ax = choice_3level.plot_wtp()
    bar_heights = sorted(round(p.get_height(), 8) for p in ax.patches)
    table_vals = sorted(round(float(v), 8) for v in w["限界支払意思額"])
    assert bar_heights == table_vals


def test_rating_grouped_heights_match_table(rating_3level):
    w = rating_3level.wtp(method="segment")
    ax = rating_3level.plot_wtp()
    bar_heights = sorted(round(p.get_height(), 8) for p in ax.patches)
    table_vals = sorted(round(float(v), 8) for v in w["限界支払意思額"])
    assert bar_heights == table_vals


def test_choice_single_widths_match_table(choice_2level):
    """2水準（横棒）でも棒の長さが表と一致する。"""
    w = choice_2level.wtp(method="segment")
    ax = choice_2level.plot_wtp()
    bar_widths = sorted(round(p.get_width(), 8) for p in ax.patches)
    table_vals = sorted(round(float(v), 8) for v in w["限界支払意思額"])
    assert bar_widths == table_vals


# ---------------------------------------------------------------------------
# method="linear"：単一棒・タイトルに「線形近似」
# ---------------------------------------------------------------------------

def test_choice_linear_single_bar_and_title(choice_3level):
    ax = choice_3level.plot_wtp(method="linear")
    # 線形近似は属性ごとに1本（brand_B のみ）
    assert len(ax.patches) == 1
    assert "線形近似" in ax.get_title()
    assert ax.get_legend() is None


def test_rating_linear_single_bar_and_title(rating_3level):
    ax = rating_3level.plot_wtp(method="linear")
    assert len(ax.patches) == 1
    assert "線形近似" in ax.get_title()


# ---------------------------------------------------------------------------
# price_segment：特定区間だけ描く
# ---------------------------------------------------------------------------

def test_choice_price_segment_filter_plot(choice_3level):
    ax = choice_3level.plot_wtp(price_segment="8〜10")
    # 1属性 × 1区間 = 1本
    assert len(ax.patches) == 1
    legend = ax.get_legend()
    labels = {t.get_text() for t in legend.get_texts()}
    assert labels == {"8〜10"}


# ---------------------------------------------------------------------------
# 既存の Axes を渡せる
# ---------------------------------------------------------------------------

def test_grouped_accepts_ax(choice_3level):
    fig, my_ax = plt.subplots()
    ax = choice_3level.plot_wtp(ax=my_ax)
    assert ax is my_ax
