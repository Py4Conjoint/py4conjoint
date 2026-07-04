"""価格列の指定統一・区間別 WTP（rating / choice 共通仕様）のテスト。

仕様:
- price_col は「数値が入った数値列のラベル」。ダミー列は encoded_columns に入れる。
- 価格3水準以上では区間（隣接水準）ごとに別々の傾きで WTP を計算する。
- 価格2水準では区間が1つなので、ダミー経由でも数値直接でも従来と一致する。
- price_col 未指定（None）で wtp() を呼ぶと日本語で案内する。
- price_range_high のような紛らわしい列名でも価格列を誤検出しない（構成的照合）。
- encode() は価格をダミー化しても元の数値列を残す。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np
import pandas as pd
import pytest

import py4conjoint.choice as pcc
import py4conjoint.rating as pc

# ---------------------------------------------------------------------------
# 人工データ生成
# ---------------------------------------------------------------------------

def _simulate_cbc(price_util, brand_util, *, n_sets=40_000, n_alts=3, seed=0):
    """価格水準ごとの効用 price_util（dict）から条件付きロジット選択を生成。"""
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


def _make_rating_3level(seed=42, n_resp=25):
    """価格3水準（6/8/10、非線形）・os 2水準の評点データ。"""
    rng = np.random.default_rng(seed)
    rows = []
    for r in range(1, n_resp + 1):
        for price in [6, 8, 10]:
            for os in ["android", "apple"]:
                u = 4.0 + {6: 1.5, 8: 1.0, 10: -1.5}[price]  # 非線形（8→10で急落）
                u += 0.7 if os == "apple" else -0.7
                u += rng.normal(0, 0.3)
                rows.append({"respondent_id": r, "rating": int(np.clip(round(u), 1, 7)),
                             "price": price, "os": os})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# choice：到達点コードがそのまま動く
# ---------------------------------------------------------------------------

def test_choice_target_code_runs():
    """到達点：price をダミー化し price_col='price' で wtp() がエラーなく返る。"""
    rng = np.random.default_rng(1)
    rows = []
    for t in range(3000):
        pr = rng.choice([6, 10], 2)
        osv = rng.choice(["android", "apple"], 2)
        cam = rng.choice(["standard", "high"], 2)
        u = np.array([
            (-0.4 * (p == 10)) + 0.5 * (o == "apple") + 0.6 * (c == "high")
            for p, o, c in zip(pr, osv, cam)
        ])
        p = np.exp(u - u.max())
        p /= p.sum()
        ch = (rng.random() < p.cumsum()).argmax()
        for j in range(2):
            rows.append({"choice_set_id": t, "respondent_id": t % 30,
                         "choice": int(j == ch), "price": int(pr[j]),
                         "os": osv[j], "camera": cam[j]})
    df = pd.DataFrame(rows)
    df_coded = pcc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "standard"},
    )
    result = pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="choice_set_id",
        encoded_columns=["price_6", "os_apple", "camera_high"],
        respondent_id_col="respondent_id",
        price_col="price",
    )
    w = result.wtp()
    assert isinstance(w, pd.DataFrame)
    # price 2水準なので単一値（価格区間列なし）、価格ダミーは出力に含まれない
    assert "価格区間" not in w.columns
    assert set(w.index) == {"os_apple", "camera_high"}
    print("OK test_choice_target_code_runs")


# ---------------------------------------------------------------------------
# choice：2水準ならダミー経由と数値直接で WTP が一致
# ---------------------------------------------------------------------------

def test_choice_2level_dummy_matches_numeric():
    df = _simulate_cbc({6: 0.0, 10: -1.2}, {"A": 0.0, "B": 0.5},
                       n_sets=40_000, seed=3)
    # (A) ダミー価格
    dc = pcc.encode(df, reference_levels={"price": 10, "brand": "A"})
    r_dummy = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                      encoded_columns=["price_6", "brand_B"], price_col="price")
    w_dummy = float(r_dummy.wtp().loc["brand_B", "限界支払意思額"])
    # (B) 数値価格
    dc2 = pcc.encode(df, reference_levels={"brand": "A"})
    r_num = pcc.fit(dc2, choice="choice", choice_set_id_col="choice_set_id",
                    encoded_columns=["price", "brand_B"], price_col="price")
    w_num = float(r_num.wtp().loc["brand_B", "限界支払意思額"])
    assert np.isclose(w_dummy, w_num, rtol=1e-4), \
        f"2水準 WTP 不一致: dummy={w_dummy}, numeric={w_num}"
    print("OK test_choice_2level_dummy_matches_numeric")


# ---------------------------------------------------------------------------
# choice：3水準で区間ごとに異なる傾き・基準水準の特定
# ---------------------------------------------------------------------------

def test_choice_3level_segment_slopes():
    # 非線形な価格効用：6→8 はゆるやか、8→10 は急。基準水準 = 6。
    df = _simulate_cbc({6: 0.0, 8: -0.3, 10: -1.6}, {"A": 0.0, "B": 0.6},
                       n_sets=60_000, seed=5)
    dc = pcc.encode(df, reference_levels={"price": 6, "brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                     encoded_columns=["price_8", "price_10", "brand_B"],
                     price_col="price")
    w = result.wtp()  # 区間別（デフォルト）
    assert "価格区間" in w.columns
    assert set(w["価格区間"]) == {"6〜8", "8〜10"}

    # 基準水準 6 の効用は 0、各水準の効用は係数そのもの（ダミー）
    u6, u8, u10 = 0.0, float(result.params["price_8"]), float(result.params["price_10"])
    bB = float(result.params["brand_B"])
    slope1 = (u8 - u6) / (8 - 6)
    slope2 = (u10 - u8) / (10 - 8)
    exp1 = -bB / slope1
    exp2 = -bB / slope2
    got1 = float(w[w["価格区間"] == "6〜8"].loc["brand_B", "限界支払意思額"])
    got2 = float(w[w["価格区間"] == "8〜10"].loc["brand_B", "限界支払意思額"])
    assert np.isclose(got1, exp1), f"区間6〜8: {got1} != {exp1}"
    assert np.isclose(got2, exp2), f"区間8〜10: {got2} != {exp2}"
    # 価格感応度が高い 8〜10 区間の WTP は小さくなるはず
    assert got2 < got1
    print("OK test_choice_3level_segment_slopes")


def test_choice_price_segment_filter():
    df = _simulate_cbc({6: 0.0, 8: -0.3, 10: -1.6}, {"A": 0.0, "B": 0.6},
                       n_sets=20_000, seed=6)
    dc = pcc.encode(df, reference_levels={"price": 6, "brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                     encoded_columns=["price_8", "price_10", "brand_B"],
                     price_col="price")
    w_all = result.wtp()
    w_one = result.wtp(price_segment="8〜10")
    assert set(w_one["価格区間"]) == {"8〜10"}
    # タプル指定でも同じ
    w_tuple = result.wtp(price_segment=(8, 10))
    assert set(w_tuple["価格区間"]) == {"8〜10"}
    # 存在しない区間はエラー
    with pytest.raises(ValueError, match="価格区間"):
        result.wtp(price_segment="6〜10")
    assert len(w_all) == 2  # 全区間
    print("OK test_choice_price_segment_filter")


# ---------------------------------------------------------------------------
# rating：2水準一致・3水準区間別
# ---------------------------------------------------------------------------

def test_rating_2level_matches_legacy_formula():
    """rating 2水準価格は区間別でも従来式 factor*b と一致し、単一値を返す。"""
    rng = np.random.default_rng(7)
    rows = []
    for r in range(1, 21):
        for price in [6, 10]:
            for os in ["android", "apple"]:
                u = 4.0 + (1.0 if price == 6 else -1.0)
                u += 0.7 if os == "apple" else -0.7
                u += rng.normal(0, 0.3)
                rows.append({"respondent_id": r, "rating": int(np.clip(round(u), 1, 7)),
                             "price": price, "os": os})
    df = pd.DataFrame(rows)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(dc, price_col="price")
    w = result.wtp()
    # 2水準は区間が1つ → 価格区間列なし・単一値
    assert "価格区間" not in w.columns
    factor = w.attrs["wtp_price_factor"]
    b_os = float(result.params["os_0"])
    assert np.isclose(w.loc["os_0", "限界支払意思額"], factor * b_os)
    # segment と linear が一致する
    w_lin = result.wtp(method="linear")
    assert np.isclose(w.loc["os_0", "限界支払意思額"],
                      w_lin.loc["os_0", "限界支払意思額"])
    print("OK test_rating_2level_matches_legacy_formula")


def test_rating_3level_segment_slopes():
    df = _make_rating_3level(seed=42, n_resp=30)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"},
                   suffix_map={"price": ["low", "mid"]})
    result = pc.fit(dc, price_col="price")
    w = result.wtp()
    assert "価格区間" in w.columns
    assert set(w["価格区間"]) == {"6〜8", "8〜10"}

    # 価格水準の効用（効果コーディング）：基準 10 の効用 = -(b_low + b_mid)
    b_low = float(result.params["price_low"])   # 水準 6
    b_mid = float(result.params["price_mid"])   # 水準 8
    u6, u8, u10 = b_low, b_mid, -(b_low + b_mid)
    slope1 = (u8 - u6) / (8 - 6)
    slope2 = (u10 - u8) / (10 - 8)
    b_os = float(result.params["os_0"])
    # 効果コーディング：基準→水準 の効用差 = 2*b_os（2水準属性）
    exp1 = 2 * b_os * (-1.0 / slope1)
    exp2 = 2 * b_os * (-1.0 / slope2)
    got1 = float(w[w["価格区間"] == "6〜8"].loc["os_0", "限界支払意思額"])
    got2 = float(w[w["価格区間"] == "8〜10"].loc["os_0", "限界支払意思額"])
    assert np.isclose(got1, exp1), f"区間6〜8: {got1} != {exp1}"
    assert np.isclose(got2, exp2), f"区間8〜10: {got2} != {exp2}"
    print("OK test_rating_3level_segment_slopes")


# ---------------------------------------------------------------------------
# price_col 未指定（None）の案内
# ---------------------------------------------------------------------------

def test_price_col_none_message_rating():
    df = _make_rating_3level(seed=1, n_resp=10)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"},
                   suffix_map={"price": ["low", "mid"]})
    result = pc.fit(dc, price_col="price")
    # price_col 未指定（None）の状態で wtp() を呼ぶと日本語で案内する
    result.price_col = None
    with pytest.raises(ValueError, match="価格列が指定されていません"):
        result.wtp()
    print("OK test_price_col_none_message_rating")


def test_price_col_none_message_choice():
    df = _simulate_cbc({6: 0.0, 10: -1.0}, {"A": 0.0, "B": 0.4},
                       n_sets=2000, seed=2)
    dc = pcc.encode(df, reference_levels={"price": 10, "brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                     encoded_columns=["price_6", "brand_B"], price_col="price")
    result.price_col = None
    with pytest.raises(ValueError, match="価格列が指定されていません"):
        result.wtp()
    print("OK test_price_col_none_message_choice")


# ---------------------------------------------------------------------------
# 構成的照合：price_range_high のような紛らわしい列名を誤検出しない
# ---------------------------------------------------------------------------

def test_constructive_matching_choice():
    """choice：price_range というダミー属性があっても price 列を誤検出しない。"""
    df = _simulate_cbc({6: 0.0, 10: -1.0}, {"A": 0.0, "B": 0.4},
                       n_sets=8000, seed=8)
    # price_range（別属性）を追加：価格が安いと "low" 帯
    df["price_range"] = np.where(df["price"] <= 6, "low", "high")
    dc = pcc.encode(df, reference_levels={"price": 10, "brand": "A",
                                          "price_range": "high"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                     encoded_columns=["price_6", "brand_B", "price_range_low"],
                     price_col="price")
    w = result.wtp()
    # 価格列として price_6 のみを使い、price_range_low は非価格属性として残る
    assert "price_range_low" in w.index
    assert "price_6" not in w.index
    print("OK test_constructive_matching_choice")


def test_constructive_matching_rating():
    """rating：price_grade という紛らわしい属性があっても price を誤検出しない。"""
    rng = np.random.default_rng(9)
    rows = []
    for r in range(1, 16):
        for price in [6, 10]:
            for grade in ["A", "B"]:
                u = 4.0 + (1.0 if price == 6 else -1.0)
                u += 0.5 if grade == "A" else -0.5
                u += rng.normal(0, 0.3)
                rows.append({"respondent_id": r, "rating": int(np.clip(round(u), 1, 7)),
                             "price": price, "price_grade": grade})
    df = pd.DataFrame(rows)
    dc = pc.encode(df, reference_levels={"price": 10, "price_grade": "B"})
    result = pc.fit(dc, price_col="price")
    # price の符号化列は price_0 のみ（price_grade_0 を巻き込まない）
    assert result._find_encoded_for("price") == ["price_0"]
    w = result.wtp()
    assert "price_grade_0" in w.index
    assert "price_0" not in w.index
    print("OK test_constructive_matching_rating")


# ---------------------------------------------------------------------------
# price_segment は区間別 WTP のときだけ指定できる（黙って無視しない）
# ---------------------------------------------------------------------------

def test_price_segment_rejected_for_2level_rating():
    """rating：価格2水準（区間1つ）で price_segment を指定するとエラー。"""
    rng = np.random.default_rng(11)
    rows = []
    for r in range(1, 16):
        for price in [6, 10]:
            for os in ["android", "apple"]:
                u = 4.0 + (1.0 if price == 6 else -1.0)
                u += 0.7 if os == "apple" else -0.7
                u += rng.normal(0, 0.3)
                rows.append({"respondent_id": r, "rating": int(np.clip(round(u), 1, 7)),
                             "price": price, "os": os})
    df = pd.DataFrame(rows)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(dc, price_col="price")
    with pytest.raises(ValueError, match="price_segment"):
        result.wtp(price_segment="6〜10")


def test_price_segment_rejected_with_linear_rating():
    """rating：method='linear'（区間別でない）で price_segment を指定するとエラー。"""
    df = _make_rating_3level(seed=12, n_resp=15)
    dc = pc.encode(df, reference_levels={"price": 10, "os": "android"},
                   suffix_map={"price": ["low", "mid"]})
    result = pc.fit(dc, price_col="price")
    with pytest.raises(ValueError, match="price_segment"):
        result.wtp(method="linear", price_segment="6〜8")


def test_price_segment_rejected_for_single_segment_choice():
    """choice：数値（線形）価格＝区間1つで price_segment を指定するとエラー。"""
    df = _simulate_cbc({6: 0.0, 10: -1.0}, {"A": 0.0, "B": 0.4},
                       n_sets=2000, seed=13)
    dc = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id",
                     encoded_columns=["price", "brand_B"], price_col="price")
    with pytest.raises(ValueError, match="price_segment"):
        result.wtp(price_segment=(6, 10))


# ---------------------------------------------------------------------------
# encode() の数値列保持
# ---------------------------------------------------------------------------

def test_encode_keeps_numeric_price_rating():
    df = pd.DataFrame({"rating": [5, 3, 6, 4], "price": [6, 10, 6, 10],
                       "os": ["a", "b", "b", "a"]})
    out = pc.encode(df, reference_levels={"price": 10, "os": "a"})
    assert "price" in out.columns
    assert pd.api.types.is_numeric_dtype(out["price"])
    assert out["price"].tolist() == [6, 10, 6, 10]


def test_encode_keeps_numeric_price_choice():
    df = pd.DataFrame({"choice_set_id": [1, 1, 2, 2], "choice": [1, 0, 0, 1],
                       "price": [6, 10, 6, 10], "brand": ["A", "B", "A", "B"]})
    out = pcc.encode(df, reference_levels={"price": 10, "brand": "A"})
    assert "price" in out.columns
    assert pd.api.types.is_numeric_dtype(out["price"])
    assert out["price"].tolist() == [6, 10, 6, 10]
