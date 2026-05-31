"""
ノートブック例の再現テスト
==========================
``smartphone_apple_google-3attributes_demo_2.ipynb`` の処理を本パッケージで
書き換え、結果が一致するかを確認する。
"""
import sys
from pathlib import Path

# パッケージのパスを通す（src レイアウト対応）
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import numpy as np
import pandas as pd
import py4conjoint as pc


def make_synthetic_data(seed: int = 0, n_resp: int = 30) -> pd.DataFrame:
    """
    ノートブックと同じ設計（4プロファイル, 3属性）の合成データを作る。
    """
    rng = np.random.default_rng(seed)
    cards = pd.DataFrame(
        {
            "price":  [6, 10, 6, 10],
            "os":     ["android", "apple", "apple", "android"],
            "camera": ["標準", "標準", "高性能", "高性能"],
        },
        index=["P1", "P2", "P3", "P4"],
    )
    rows = []
    for r in range(1, n_resp + 1):
        for cid in cards.index:
            row = cards.loc[cid].to_dict()
            # 真の効用：価格安い+1, apple+0.7, 高性能+1.2 + ノイズ
            u = 4.0
            u += 1.0 if row["price"] == 6 else -1.0
            u += 0.7 if row["os"] == "apple" else -0.7
            u += 1.2 if row["camera"] == "高性能" else -1.2
            u += rng.normal(0, 0.5)
            rating = int(np.clip(round(u), 1, 7))
            rows.append({"回答者ID": r, "プロファイルID": cid, "rating": rating, **row})
    return pd.DataFrame(rows)


def test_encode_binary_three_attrs():
    """ノートブックの符号化が encode() で再現できるか"""
    df = make_synthetic_data(seed=1, n_resp=10)
    df_coded = pc.encode(
        df,
        reference_levels={
            "price":  10,
            "os":     "android",
            "camera": "標準",
        },
    )
    # 列名はノートブックと違ってよい（_6/_apple/_高性能）が、
    # 値が ±1 のみで、かつ意味が同じであることを確認
    assert "price_0" in df_coded.columns
    assert "os_0" in df_coded.columns
    assert "camera_0" in df_coded.columns
    assert set(df_coded["price_0"].unique()) == {-1, 1}
    assert set(df_coded["os_0"].unique()) == {-1, 1}
    assert set(df_coded["camera_0"].unique()) == {-1, 1}
    # 意味の対応
    assert (df_coded.loc[df_coded["price"] == 6, "price_0"] == 1).all()
    assert (df_coded.loc[df_coded["price"] == 10, "price_0"] == -1).all()
    print("OK test_encode_binary_three_attrs")


def test_encode_three_levels():
    """3水準の場合、K-1個のダミー列が生成される"""
    df = pd.DataFrame({
        "rating": [3, 5, 7, 4, 5, 6],
        "color":  ["赤", "青", "緑", "赤", "青", "緑"],
    })
    out = pc.encode(df, reference_levels={"color": "赤"})
    assert "color_0" in out.columns
    assert "color_1" in out.columns
    # 赤の行は両方 -1
    red = out[out["color"] == "赤"]
    assert (red["color_0"] == -1).all()
    assert (red["color_1"] == -1).all()
    blue = out[out["color"] == "青"]
    assert (blue["color_0"] == 1).all()
    assert (blue["color_1"] == 0).all()
    green = out[out["color"] == "緑"]
    assert (green["color_0"] == 0).all()
    assert (green["color_1"] == 1).all()
    print("OK test_encode_three_levels")


def test_fit_and_summary():
    df = make_synthetic_data(seed=2, n_resp=30)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded)
    s = result.summary()
    assert "コンジョイント分析の結果" in s
    assert "決定係数" in s
    # 真の係数の符号方向が出ているか
    assert result.params["price_0"] > 0
    assert result.params["os_0"] > 0
    assert result.params["camera_0"] > 0
    # R² がそこそこ
    assert result.rsquared > 0.3
    print("OK test_fit_and_summary")
    print(s)


def test_importance_sums_to_100():
    df = make_synthetic_data(seed=3, n_resp=30)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded)
    imp = result.importance(as_percent=True)
    assert abs(imp["importance"].sum() - 100.0) < 1e-6
    # 真の係数で「カメラ > 価格 > OS」の順のはず
    sorted_imp = imp.sort_values("importance", ascending=False)
    print(sorted_imp)
    print("OK test_importance_sums_to_100")


def test_wtp():
    df = make_synthetic_data(seed=4, n_resp=30)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded)
    wtp = result.wtp()
    print(wtp)
    # appleの方がandroidより魅力 → 正のWTP
    assert wtp.loc["os_0", "支払意思額"] > 0
    # カメラ高性能 → 正のWTP
    assert wtp.loc["camera_0", "支払意思額"] > 0
    # 新式: (price_max - price_min) / abs(b_price * 2)
    b_price = result.params["price_0"]
    expected_unit = (10 - 6) / abs(b_price * 2)
    assert abs(result.unit_rating_money() - expected_unit) < 1e-9
    print("OK test_wtp")


def test_market_share():
    df = make_synthetic_data(seed=5, n_resp=30)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded)
    products = pd.DataFrame(
        {
            "price_0":      [1, -1],   # 製品A: 6万円, 製品B: 10万円
            "os_0":     [-1, 1],   # 製品A: android, 製品B: apple
            "camera_0": [1, -1],   # 製品A: 高性能, 製品B: 標準
        },
        index=["製品A", "製品B"],
    )
    share = result.market_share(products)
    assert abs(share.sum() - 1.0) < 1e-9
    print(share)
    print("OK test_market_share")


def test_diagnostics_warning_low_r2():
    """R² < 0.20 の場合に重大度「大」の r2_low 警告が出る"""
    rng = np.random.default_rng(0)
    df = pd.DataFrame(
        {
            "rating": rng.integers(1, 8, 50),  # 完全ランダム
            "price":  rng.choice([6, 10], 50),
            "os":     rng.choice(["android", "apple"], 50),
        }
    )
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(df_coded)

    # 新しい Diagnostic API で確認
    w_df = result.warnings()
    print(f"R² = {result.rsquared:.3f}")
    print(w_df)
    assert "r2_low" in w_df["category"].values, "r2_low 警告が出ていない"
    r2_row = w_df[w_df["category"] == "r2_low"].iloc[0]
    assert r2_row["severity"] == "大"

    # severity フィルタが動く
    major_only = result.warnings(severity="大")
    assert len(major_only) >= 1

    # category フィルタが動く
    r2_only = result.warnings(category="r2_low")
    assert len(r2_only) == 1

    # as_dataframe=False でも取得できる
    diags = result.warnings(as_dataframe=False)
    assert all(hasattr(d, "severity") for d in diags)

    print("OK test_diagnostics_warning_low_r2")


def test_diagnostics_price_insignificant():
    """価格に真の効果がない場合に price_insignificant 中警告が出る（確定的テスト）。

    直交バランスデザイン（price_0 ⊥ os_0）で os のみ評点に影響するデータを生成する。
    このとき b_hat_price ≈ 0 となり p 値が大きくなることが設計上保証される。
    """
    cards_order = ["P1", "P2", "P3", "P4"]
    price_vals  = [6, 10, 6, 10]
    os_vals     = ["android", "apple", "apple", "android"]
    n_resp = 20
    rows = []
    for r in range(1, n_resp + 1):
        resp_offset = float(r % 2)      # 回答者ごとの定数（price_0 ⊥ offset）
        for i, cid in enumerate(cards_order):
            os_effect     = 1.5 if os_vals[i] == "apple" else -1.5
            tiny_price_ef = 0.001 * (1 if price_vals[i] == 6 else -1)  # 極小、事実上ゼロ
            rating = float(np.clip(4.0 + os_effect + resp_offset + tiny_price_ef, 1, 7))
            rows.append({
                "回答者ID": r, "プロファイルID": cid, "rating": rating,
                "price": price_vals[i], "os": os_vals[i],
            })
    df = pd.DataFrame(rows)
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(df_coded)
    _ = result.wtp()

    p_price = float(result.ols.pvalues["price_0"])
    print(f"  price p値 = {p_price:.4f}")
    assert p_price >= 0.10, \
        f"price 係数が有意になった（p={p_price:.4f}）。データ構築を確認してください"

    w_df = result.warnings()
    assert "price_insignificant" in w_df["category"].values, \
        f"price_insignificant が出ていない（p={p_price:.4f}）"
    row = w_df[w_df["category"] == "price_insignificant"].iloc[0]
    assert row["severity"] == "中"
    print("OK test_diagnostics_price_insignificant")


def test_diagnostics_wtp_extrapolation():
    """|WTP| > 価格レンジ×2 の場合に wtp_extrapolation 警告が出る"""
    # 価格係数を小さく、OS係数を大きくして外挿を意図的に起こす
    rng = np.random.default_rng(7)
    n = 40
    os_vals = rng.choice(["android", "apple"], n)
    df = pd.DataFrame({
        "price":  rng.choice([6, 10], n),
        "os":     os_vals,
        # OS の効果を非常に大きく設定（係数 ≒ 4 → WTP ≒ 4/b_price × 4 ≫ レンジ4）
        "rating": np.where(os_vals == "apple", 6, 2).astype(float)
              + rng.normal(0, 0.3, n),
    })
    df["rating"] = df["rating"].clip(1, 7)
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(df_coded)
    wtp = result.wtp()
    w_df = result.warnings()
    print(wtp)
    print(w_df)

    # 外挿警告が出ているはず（price_range=4, threshold=8, WTP >> 8 のはず）
    assert "wtp_extrapolation" in w_df["category"].values, \
        f"wtp_extrapolation 警告が出ていない。WTP:\n{wtp}"
    ext_rows = w_df[w_df["category"] == "wtp_extrapolation"]
    # wtp_extrapolation は常に「中」
    assert ext_rows.iloc[0]["severity"] == "中", \
        f"期待: 中, 実際: {ext_rows.iloc[0]['severity']}"
    print(f"  wtp_extrapolation 警告を確認（重大度={ext_rows.iloc[0]['severity']}）")
    print("OK test_diagnostics_wtp_extrapolation")


def test_warnings_not_duplicated():
    """wtp()を複数回呼んでも警告が重複登録されない"""
    df = make_synthetic_data(seed=4, n_resp=30)
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android", "camera": "標準"})
    result = pc.fit(df_coded)

    _ = result.wtp()
    _ = result.wtp()   # 2回目
    _ = result.wtp()   # 3回目

    w_df = result.warnings(category="price_insignificant")
    # 高々1件しかないはず
    assert len(w_df) <= 1, f"price_insignificant が重複: {len(w_df)} 件"

    all_cats = result.warnings()["category"].tolist() if len(result.warnings()) > 0 else []
    from collections import Counter
    dupes = [k for k, v in Counter(all_cats).items() if v > 1
             and not k.startswith("wtp_extrapolation_")]  # 属性ごとなので別カテゴリ扱い
    assert not dupes, f"重複した警告カテゴリ: {dupes}"
    print("OK test_warnings_not_duplicated")


def test_price_col_override():
    """price以外の価格列名でも動く"""
    df = make_synthetic_data(seed=6, n_resp=20).rename(columns={"price": "値段"})
    df_coded = pc.encode(
        df,
        reference_levels={"値段": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded, price_col="値段")
    wtp = result.wtp()
    assert "os_0" in wtp.index
    assert "camera_0" in wtp.index
    print(wtp)
    print("OK test_price_col_override")


def test_e2e_real_data():
    """
    実データ（test.xlsx）を使ったエンドツーエンドのテスト。
    ノートブック方式と数値完全一致 + 新警告が実データで正しく出ることを確認する。
    """
    import os, statsmodels.formula.api as smf
    xlsx = "/mnt/user-data/uploads/test.xlsx"
    if not os.path.exists(xlsx):
        print("  test.xlsx が見つからないためスキップ")
        return

    profiles = pd.DataFrame(
        {
            "price":  [6, 10, 6, 10],
            "os":     ["android", "apple", "apple", "android"],
            "camera": ["標準", "標準", "高性能", "高性能"],
        },
        index=["P1", "P2", "P3", "P4"],
    )

    # フルパイプライン
    df = pc.forms_to_conjoint_data(
        responses_file=xlsx,
        attributes=profiles,
    )
    df = pc.encode(df, reference_levels={"price": 10, "os": "android", "camera": "標準"})
    result = pc.fit(df)

    # ノートブック方式と数値一致
    df_nb = pc.forms_to_conjoint_data(responses_file=xlsx, attributes=profiles)
    df_nb["price_low"]   = df_nb["price"].map({10: -1, 6: 1})
    df_nb["os_apple"]    = df_nb["os"].map({"android": -1, "apple": 1})
    df_nb["camera_high"] = df_nb["camera"].map({"標準": -1, "高性能": 1})
    res_nb = smf.ols("rating ~ price_low + os_apple + camera_high", data=df_nb).fit()

    for nb_col, pkg_col in [
        ("price_low", "price_0"), ("os_apple", "os_0"), ("camera_high", "camera_0")
    ]:
        assert abs(res_nb.params[nb_col] - result.params[pkg_col]) < 1e-9, \
            f"{pkg_col} が一致しない"

    wtp = result.wtp()
    apple_wtp_nb = -(6-10) / res_nb.params["price_low"] * res_nb.params["os_apple"]
    assert abs(wtp.loc["os_0", "支払意思額"] - apple_wtp_nb) < 1e-9

    # 実データで出るはずの警告を確認
    w_df = result.warnings()
    print(f"\n--- 実データの警告一覧 ---")
    print(w_df.to_string(index=False))
    print()

    # 実データ: price p値=0.138 → price_insignificant (中) が出るはず
    assert "price_insignificant" in w_df["category"].values, \
        "price_insignificant が出ていない"
    # 実データ: os_0 (apple) WTP = 15.83 >> 4×2=8 → wtp_extrapolation が出るはず
    assert "wtp_extrapolation" in w_df["category"].values, \
        "wtp_extrapolation が出ていない"

    # summary に「大」だけ出ているか
    s = result.summary()
    print(s)
    major = result.warnings(severity="大")
    if len(major) > 0:
        for _, row in major.iterrows():
            assert row["message"][:10] in s, \
                f"重大度「大」の警告がsummaryに出ていない: {row['message'][:20]}"

    print("OK test_e2e_real_data")


def test_unit_rating_money_returns_float():
    """unit_rating_money() が float を返し、計算式が正しいことを確認"""
    df = make_synthetic_data(seed=4, n_resp=30)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
    )
    result = pc.fit(df_coded)

    unit = result.unit_rating_money()

    # float を返すこと
    assert isinstance(unit, float), f"float が返されるべきだが {type(unit)} が返された"

    # 計算式: (price_max - price_min) / abs(b_price * 2)
    b_price = result.params["price_0"]
    expected = (10 - 6) / abs(b_price * 2)
    assert abs(unit - expected) < 1e-9, f"期待値 {expected} に対し {unit} が返された"

    # wtp() を一度も呼ばなくても動くこと（副作用なし）
    result2 = pc.fit(df_coded)
    unit2 = result2.unit_rating_money()
    assert isinstance(unit2, float)
    assert abs(unit2 - expected) < 1e-9

    print(f"評点1点 = {unit:.4f} 万円")
    print("OK test_unit_rating_money_returns_float")


def test_few_respondents_major():
    """回答者が1人のとき few_respondents 大警告が出る"""
    df = make_synthetic_data(seed=0, n_resp=1)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    w_df = result.warnings(category="few_respondents")
    assert len(w_df) == 1, f"few_respondents 警告が {len(w_df)} 件"
    assert w_df.iloc[0]["severity"] == "大"
    print("OK test_few_respondents_major")


def test_few_respondents_minor():
    """回答者が3人のとき few_respondents 中警告が出る"""
    df = make_synthetic_data(seed=0, n_resp=3)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    w_df = result.warnings(category="few_respondents")
    assert len(w_df) == 1, f"few_respondents 警告が {len(w_df)} 件"
    assert w_df.iloc[0]["severity"] == "中"
    print("OK test_few_respondents_minor")


def test_few_respondents_no_warning():
    """回答者が5人（境界値）のとき few_respondents 警告が出ない"""
    df = make_synthetic_data(seed=0, n_resp=5)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    w_df = result.warnings(category="few_respondents")
    assert len(w_df) == 0, f"n=5 なのに few_respondents 警告が出た:\n{w_df}"
    print("OK test_few_respondents_no_warning")


def test_price_sign_negative():
    """高価格ほど評点が高い（高級品）データで price_sign_negative 中警告が出る"""
    rng = np.random.default_rng(99)
    n_resp = 20
    rows = []
    for r in range(1, n_resp + 1):
        for price, os in [(6, "android"), (10, "android"), (6, "apple"), (10, "apple")]:
            luxury_effect = 2.0 if price == 10 else -2.0
            rating = float(np.clip(4.0 + luxury_effect + rng.normal(0, 0.3), 1, 7))
            rows.append({"回答者ID": r, "price": price, "os": os, "rating": rating})
    df = pd.DataFrame(rows)
    # 高い方を基準 → price_0=+1 が price=6（安い方）→ luxury では b_price < 0
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(df_coded)
    b_price = float(result.params["price_0"])
    assert b_price < 0, f"b_price={b_price:.4f}: 高級品データなのに符号が正"
    w_df = result.warnings(category="price_sign_negative")
    assert len(w_df) == 1
    assert w_df.iloc[0]["severity"] == "中"
    print(f"  b_price = {b_price:.4f}")
    print("OK test_price_sign_negative")


def test_summary_slim_false():
    """summary(slim=False) は statsmodels の詳細表（英語）を返す"""
    df = make_synthetic_data(seed=0, n_resp=20)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    s_full = result.summary(slim=False)
    assert "OLS Regression Results" in s_full
    s_slim = result.summary(slim=True)
    assert s_full != s_slim
    assert "コンジョイント分析" in s_slim
    print("OK test_summary_slim_false")


def test_auto_reference_levels():
    """auto_reference_levels: 数値列→最大値、カテゴリ列→辞書順先頭"""
    import warnings as _warnings
    df = make_synthetic_data(seed=0, n_resp=5)
    with _warnings.catch_warnings(record=True) as w:
        _warnings.simplefilter("always")
        refs = pc.auto_reference_levels(df, ["price", "os", "camera"])
    # 数値列 price: 最大値 = 10
    assert refs["price"] == 10, f"price の基準値が {refs['price']} (期待: 10)"
    # カテゴリ列 os: 辞書順先頭（"android" < "apple"）
    assert refs["os"] == "android", f"os の基準値が {refs['os']} (期待: 'android')"
    # カテゴリ列 camera: 辞書順先頭（"標準" < "高性能" in Unicode order）
    assert refs["camera"] == sorted(["標準", "高性能"])[0]
    # UserWarning が発生していること
    assert any(issubclass(wi.category, UserWarning) for wi in w)
    print(f"  refs = {refs}")
    print("OK test_auto_reference_levels")


def test_market_share_max():
    """method='max' は最大効用の製品にシェア1.0、他は0.0を割り当てる"""
    df = make_synthetic_data(seed=0, n_resp=20)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    products = pd.DataFrame(
        {
            "price_0":  [ 1, -1],
            "os_0":     [ 1, -1],
            "camera_0": [ 1, -1],
        },
        index=["製品A", "製品B"],
    )
    share = result.market_share(products, method="max")
    # 製品Aがすべてプラス属性 → 最大効用 → シェア1
    assert share["製品A"] == 1.0
    assert share["製品B"] == 0.0
    assert abs(share.sum() - 1.0) < 1e-9
    print("OK test_market_share_max")


def test_encode_drop_original():
    """drop_original=True で元の属性列が削除される"""
    df = make_synthetic_data(seed=0, n_resp=5)
    out = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        drop_original=True,
    )
    for col in ("price", "os", "camera"):
        assert col not in out.columns, f"'{col}' が残っている"
    for col in ("price_0", "os_0", "camera_0"):
        assert col in out.columns, f"'{col}' がない"
    print("OK test_encode_drop_original")


def test_encode_inplace():
    """inplace=True は入力 DataFrame を直接書き換えて同じオブジェクトを返す"""
    df = make_synthetic_data(seed=0, n_resp=5)
    original_id = id(df)
    out = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        inplace=True,
    )
    assert id(out) == original_id, "inplace=True なのに別オブジェクトが返された"
    assert "price_0" in out.columns
    assert "os_0" in out.columns
    print("OK test_encode_inplace")


def test_encode_binary_suffix_map():
    """binary_suffix_map でカスタム列名サフィックスが使われる"""
    df = make_synthetic_data(seed=0, n_resp=5)
    out = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        binary_suffix_map={"price": "low", "os": "apple"},
    )
    assert "price_low" in out.columns, "price_low がない"
    assert "os_apple" in out.columns, "os_apple がない"
    assert "price_0" not in out.columns, "price_0 が残っている（カスタム名に上書きされるはず）"
    assert "os_0" not in out.columns, "os_0 が残っている（カスタム名に上書きされるはず）"
    # binary_suffix_map 未指定の属性はデフォルト命名
    assert "camera_0" in out.columns, "camera_0 がない"
    print("OK test_encode_binary_suffix_map")


def test_importance_ratio():
    """importance(as_percent=False) は合計が1.0の比率を返す"""
    df = make_synthetic_data(seed=0, n_resp=20)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    imp = result.importance(as_percent=False)
    assert abs(imp["importance"].sum() - 1.0) < 1e-6, \
        f"合計が1.0にならない: {imp['importance'].sum()}"
    assert (imp["importance"] > 0).all(), "重要度が0以下の属性がある"
    print("OK test_importance_ratio")


def test_wtp_attrs():
    """wtp() の戻り値 DataFrame.attrs に必要なキーと正しい値が入っている"""
    df = make_synthetic_data(seed=0, n_resp=20)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)
    wtp = result.wtp()
    for key in ("price_range", "wtp_price_factor", "p_price", "price_low", "price_high"):
        assert key in wtp.attrs, f"attrs に '{key}' がない"
    assert wtp.attrs["price_low"] == 6
    assert wtp.attrs["price_high"] == 10
    assert abs(wtp.attrs["price_range"] - 4.0) < 1e-9
    b_price = float(result.params["price_0"])
    expected_factor = 4.0 / b_price   # -(low-high)/b = -(6-10)/b = 4/b
    assert abs(wtp.attrs["wtp_price_factor"] - expected_factor) < 1e-9
    print("OK test_wtp_attrs")


if __name__ == "__main__":
    test_encode_binary_three_attrs()
    test_encode_three_levels()
    test_fit_and_summary()
    test_importance_sums_to_100()
    test_wtp()
    test_market_share()
    test_diagnostics_warning_low_r2()
    test_diagnostics_price_insignificant()
    test_diagnostics_wtp_extrapolation()
    test_warnings_not_duplicated()
    test_price_col_override()
    test_e2e_real_data()
    test_unit_rating_money_returns_float()
    test_few_respondents_major()
    test_few_respondents_minor()
    test_few_respondents_no_warning()
    test_price_sign_negative()
    test_summary_slim_false()
    test_auto_reference_levels()
    test_market_share_max()
    test_encode_drop_original()
    test_encode_inplace()
    test_encode_binary_suffix_map()
    test_importance_ratio()
    test_wtp_attrs()
    print("\nすべてのテストがパスしました。")
