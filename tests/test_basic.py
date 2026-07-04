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
import py4conjoint.rating as pc


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
            rows.append({"respondent_id": r, "profile_id": cid, "rating": rating, **row})
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


def test_encode_three_levels_keeps_nan():
    """3水準以上でも欠損は欠損のまま残る（0 に化けない）。

    回帰テスト：以前は Series.map が NaN にも関数を適用したため、
    欠損が全ダミー列 0（＝全水準平均の効用）として静かに回帰に混入していた。
    2水準（map(dict)）・choice 版（pd.NA 保持）と同じ挙動に揃える。
    """
    df = pd.DataFrame({
        "rating": [3, 5, 7, 4],
        "color":  ["赤", "青", "緑", np.nan],
    })
    out = pc.encode(df, reference_levels={"color": "赤"})
    nan_rows = out[out["color"].isna()]
    assert nan_rows["color_0"].isna().all(), "欠損行の color_0 が NaN でない"
    assert nan_rows["color_1"].isna().all(), "欠損行の color_1 が NaN でない"
    # 欠損でない行の符号化は従来どおり
    assert (out.loc[out["color"] == "赤", ["color_0", "color_1"]] == -1).all().all()
    print("OK test_encode_three_levels_keeps_nan")


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
    assert abs(imp["重要度"].sum() - 100.0) < 1e-6
    # 真の係数で「カメラ > 価格 > OS」の順のはず
    sorted_imp = imp.sort_values("重要度", ascending=False)
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
    assert wtp.loc["os_0", "限界支払意思額"] > 0
    # カメラ高性能 → 正のWTP
    assert wtp.loc["camera_0", "限界支払意思額"] > 0
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
                "respondent_id": r, "profile_id": cid, "rating": rating,
                "price": price_vals[i], "os": os_vals[i],
            })
    df = pd.DataFrame(rows)
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    # このテストは決定的データ（回答者内の残差がゼロ）を使うため、
    # クラスタSEでは価格係数のSEが0に近づきp値が0になる。
    # 警告ロジックの検証が目的なので通常SEを明示する。
    result = pc.fit(df_coded, cluster_se=False)
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
    実データ（examples/responses_os.csv）を使ったエンドツーエンドのテスト。
    ノートブック方式（手動符号化 + 通常OLS）と係数が完全一致し、
    クラスタロバスト標準誤差がデフォルトで適用されることを確認する。
    """
    import statsmodels.formula.api as smf
    csv = Path(__file__).parent.parent / "examples" / "responses_os.csv"
    if not csv.exists():
        print("  responses_os.csv が見つからないためスキップ")
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
    df = pc.forms_to_data(
        responses_file=str(csv),
        profiles=profiles,
        forms="google",
    )
    df = pc.encode(df, reference_levels={"price": 10, "os": "android", "camera": "標準"})
    result = pc.fit(df)

    # 回答者ID列があるのでクラスタロバストSEが使われる
    assert result.se_type == "cluster"
    assert "クラスタロバスト" in result.summary()

    # ノートブック方式（手動符号化 + 通常OLS）と係数が一致
    # （クラスタSEは標準誤差のみ変え、係数の推定値は変えない）
    df_nb = pc.forms_to_data(
        responses_file=str(csv), profiles=profiles, forms="google"
    )
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
    assert abs(wtp.loc["os_0", "限界支払意思額"] - apple_wtp_nb) < 1e-9

    # 警告の整合性チェック（データの中身に依存しない条件付き検証）
    w_df = result.warnings()
    print("\n--- 実データの警告一覧 ---")
    print(w_df.to_string(index=False))
    cats = w_df["category"].values
    p_price = wtp.attrs["p_price"]
    if p_price >= 0.10:
        assert "price_insignificant" in cats, \
            f"p_price={p_price:.3f} ≥ 0.10 なのに price_insignificant が出ていない"
    else:
        assert "price_insignificant" not in cats, \
            f"p_price={p_price:.3f} < 0.10 なのに price_insignificant が出た"
    threshold = wtp.attrs["price_range"] * 2
    if (wtp["限界支払意思額"].abs() > threshold).any():
        assert "wtp_extrapolation" in cats, "wtp_extrapolation が出ていない"
    # 回答者ID列があるので independence_assumed は出ない
    assert "independence_assumed" not in cats

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
            rows.append({"respondent_id": r, "price": price, "os": os, "rating": rating})
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
    assert abs(imp["重要度"].sum() - 1.0) < 1e-6, \
        f"合計が1.0にならない: {imp['重要度'].sum()}"
    assert (imp["重要度"] > 0).all(), "重要度が0以下の属性がある"
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


def test_encode_multi_with_suffix_map():
    """suffix_map で3水準のサフィックスを指定できるか"""
    # 存在しない基準水準を指定すると ValueError
    df_bad = pd.DataFrame({
        "rating": [5, 3, 7, 4, 6, 2],
        "color": ["赤", "青", "緑", "赤", "青", "緑"],
    })
    try:
        pc.encode(df_bad, reference_levels={"color": "白"})
        assert False, "ValueError が出るはず"
    except ValueError:
        pass

    # suffix_map で3水準のサフィックスを指定
    df2 = pd.DataFrame({
        "rating": [5, 3, 7, 4, 6, 2],
        "color": ["赤", "青", "緑", "赤", "青", "緑"],
    })
    df2_coded = pc.encode(
        df2,
        reference_levels={"color": "赤"},
        suffix_map={"color": ["blue", "green"]},
    )
    assert "color_blue" in df2_coded.columns, "color_blue がない"
    assert "color_green" in df2_coded.columns, "color_green がない"
    assert "color_0" not in df2_coded.columns, "color_0 が残っている"
    # 値の正当性
    assert (df2_coded.loc[df2_coded["color"] == "青", "color_blue"] == 1).all()
    assert (df2_coded.loc[df2_coded["color"] == "赤", "color_blue"] == -1).all()
    assert (df2_coded.loc[df2_coded["color"] == "緑", "color_blue"] == 0).all()
    print("OK test_encode_multi_with_suffix_map")


def test_encode_multi_suffix_length_mismatch():
    """suffix_map のリスト長が水準数と一致しない場合に ValueError"""
    df = pd.DataFrame({"rating": [1, 2, 3], "color": ["赤", "青", "緑"]})
    try:
        pc.encode(df, reference_levels={"color": "赤"}, suffix_map={"color": ["only_one"]})
        assert False, "ValueError が出るはず"
    except ValueError as e:
        assert "suffix_map" in str(e), f"エラーメッセージに suffix_map が含まれない: {e}"
    print("OK test_encode_multi_suffix_length_mismatch")


def test_binary_suffix_map_deprecation():
    """binary_suffix_map を渡すと DeprecationWarning が出る"""
    import warnings as _warnings
    df = pd.DataFrame({"rating": [1, 2], "price": [6, 10]})
    with _warnings.catch_warnings(record=True) as w:
        _warnings.simplefilter("always")
        pc.encode(df, reference_levels={"price": 10}, binary_suffix_map={"price": "low"})
    assert any(issubclass(wi.category, DeprecationWarning) for wi in w), \
        "DeprecationWarning が出ていない"
    dep_warns = [wi for wi in w if issubclass(wi.category, DeprecationWarning)]
    assert "binary_suffix_map" in str(dep_warns[0].message)
    print("OK test_binary_suffix_map_deprecation")


def test_check_design_basic():
    """check_design() の基本動作：完全直交デザインはバランス・相関・χ²に問題なし"""
    profiles = pd.DataFrame({
        "price":  [6, 10, 6, 10],
        "os":     ["android", "apple", "apple", "android"],
        "camera": ["標準", "標準", "高性能", "高性能"],
    })
    result = pc.check_design(profiles)
    assert isinstance(result.balance, pd.DataFrame)
    assert isinstance(result.correlation, pd.DataFrame)
    assert isinstance(result.chi2, pd.DataFrame)
    assert isinstance(result.diagnostics, list)

    # バランス・相関・χ² に関するアクティブな問題はないはず
    cats = [d.category for d in result.diagnostics]
    assert not any("balance" in c for c in cats), f"balance 警告が出た: {cats}"
    assert not any("correlation" in c for c in cats), f"correlation 警告が出た: {cats}"
    assert not any("chi2" in c for c in cats), f"chi2 警告が出た: {cats}"

    # summary() が文字列を返す
    s = result.summary()
    assert "デザイン直交性チェック" in s
    # warnings() が DataFrame を返す
    w_df = result.warnings()
    assert isinstance(w_df, pd.DataFrame)
    print("OK test_check_design_basic")


def test_check_design_imbalanced():
    """check_design() でバランスが崩れている場合に balance_* 警告が出る"""
    profiles = pd.DataFrame({
        "price": [6, 6, 6, 10],   # 6が3回、10が1回（偏り大）
        "os":    ["android", "apple", "android", "apple"],
    })
    result = pc.check_design(profiles)
    cats = [d.category for d in result.diagnostics]
    assert any("balance" in c for c in cats), \
        f"balance 警告が出ていない。警告: {cats}"
    print("OK test_check_design_imbalanced")


def test_check_design_ignores_pandas_index_column(tmp_path):
    """index=False を付け忘れた profiles CSV でも、行番号列を属性として診断しない。

    回帰テスト：以前は Unnamed: 0 列（P1, P2, … のラベル）が「プロファイル数と
    同数の水準を持つ属性」として扱われ、パラメータ数が架空に膨らんで
    insufficient_profiles などの誤警告が大量に出ていた。
    """
    profiles = pd.DataFrame({
        "price":  [6, 10, 6, 10],
        "os":     ["android", "apple", "apple", "android"],
        "camera": ["標準", "標準", "高性能", "高性能"],
    }, index=["P1", "P2", "P3", "P4"])
    csv = tmp_path / "profiles.csv"
    profiles.to_csv(csv)                      # index=False を付け忘れたケース
    loaded = pd.read_csv(csv)
    assert "Unnamed: 0" in loaded.columns

    result = pc.check_design(loaded)
    # 行番号列は診断対象にならず、本来の属性だけが並ぶ
    assert sorted(result.balance.index) == ["camera", "os", "price"]
    # 完全直交デザインなので誤警告（insufficient_profiles 等）は出ない
    cats = [d.category for d in result.diagnostics]
    assert "insufficient_profiles" not in cats, f"誤警告が出た: {cats}"
    assert not any("Unnamed" in c for c in cats), f"行番号列由来の警告: {cats}"
    print("OK test_check_design_ignores_pandas_index_column")


def test_encode_binary_suffix_map_as_list():
    """2水準属性に suffix_map で1要素リストを渡しても正しく動く（str と等価）"""
    df = pd.DataFrame({"rating": [1, 2], "price": [6, 10]})
    out = pc.encode(df, reference_levels={"price": 10}, suffix_map={"price": ["low"]})
    assert "price_low" in out.columns, "price_low がない"
    assert "price_0" not in out.columns, "price_0 が残っている"

    # 2水準属性に2要素以上のリストを渡すと ValueError
    try:
        pc.encode(df, reference_levels={"price": 10}, suffix_map={"price": ["low", "high"]})
        assert False, "ValueError が出るはず"
    except ValueError:
        pass
    print("OK test_encode_binary_suffix_map_as_list")


def test_wtp_three_level_price_no_extra_columns():
    """3水準価格の場合、WTP 出力に価格の符号化列（price_low, price_mid）が混入しない"""
    rng = np.random.default_rng(42)
    n_resp = 25
    rows = []
    for r in range(1, n_resp + 1):
        for price, os in [(6, "android"), (8, "android"), (10, "android"),
                          (6, "apple"), (8, "apple"), (10, "apple")]:
            u = 4.0 + {6: 1.5, 8: 0.0, 10: -1.5}[price]
            u += 0.7 if os == "apple" else -0.7
            u += rng.normal(0, 0.3)
            rating = int(np.clip(round(u), 1, 7))
            rows.append({"respondent_id": r, "rating": rating, "price": price, "os": os})
    df = pd.DataFrame(rows)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android"},
        suffix_map={"price": ["low", "mid"]},
    )
    result = pc.fit(df_coded, price_col="price")
    wtp = result.wtp()  # 区間別（デフォルト）
    # 価格列（price_low, price_mid）が WTP 出力に含まれないこと
    assert "price_low" not in wtp.index, "price_low が WTP 出力に混入している"
    assert "price_mid" not in wtp.index, "price_mid が WTP 出力に混入している"
    # 非価格属性は os_0 のみ（区間ごとに1行ずつ出力される）
    assert set(wtp.index) == {"os_0"}, f"os_0 以外が混入: {set(wtp.index)}"
    # 区間別なので 価格区間 列が付き、2区間（6〜8, 8〜10）になる
    assert "価格区間" in wtp.columns
    assert len(wtp) == 2, f"WTP 出力の行数が2以外: {len(wtp)}"
    assert set(wtp["価格区間"]) == {"6〜8", "8〜10"}
    print("OK test_wtp_three_level_price_no_extra_columns")


def test_wtp_three_level_price():
    """3水準の価格でもwtp()がNotImplementedErrorを出さず、正のWTPを返す"""
    rng = np.random.default_rng(42)
    n_resp = 25
    rows = []
    for r in range(1, n_resp + 1):
        for price, os in [(6, "android"), (8, "android"), (10, "android"),
                          (6, "apple"), (8, "apple"), (10, "apple")]:
            u = 4.0
            u += {6: 1.5, 8: 0.0, 10: -1.5}[price]
            u += 0.7 if os == "apple" else -0.7
            u += rng.normal(0, 0.3)
            rating = int(np.clip(round(u), 1, 7))
            rows.append({"respondent_id": r, "rating": rating, "price": price, "os": os})
    df = pd.DataFrame(rows)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android"},
        suffix_map={"price": ["low", "mid"]},
    )
    result = pc.fit(df_coded, price_col="price")

    # エラーを出さず、区間別の DataFrame を返すこと
    wtp = result.wtp()
    assert isinstance(wtp, pd.DataFrame), "wtp() が DataFrame を返さない"
    assert "価格区間" in wtp.columns, "区間別 WTP に 価格区間 列がない"
    # apple の WTP は各区間で正のはず
    assert (wtp["限界支払意思額"] > 0).all(), \
        f"apple の区間別 WTP に負の値: {wtp['限界支払意思額'].tolist()}"

    # デフォルト（method='segment'）では線形近似警告は出ない
    cats = [d.category for d in result._diagnostics]
    assert "wtp_price_linear_approx" not in cats, \
        f"segment なのに線形近似警告が出ている: {cats}"

    # method='linear' のときだけ線形近似1本の単一値＋警告になる
    wtp_lin = result.wtp(method="linear")
    assert "価格区間" not in wtp_lin.columns
    assert wtp_lin.loc["os_0", "限界支払意思額"] > 0
    cats2 = [d.category for d in result._diagnostics]
    assert "wtp_price_linear_approx" in cats2, \
        f"method='linear' で線形近似警告が出ていない: {cats2}"

    print("OK test_wtp_three_level_price")


ATTR_LEVELS = {
    "price":  [6, 8, 10],
    "os":     ["android", "apple"],
    "camera": ["標準", "高性能", "超高性能"],
}
REF_LEVELS = {"price": 10, "os": "android", "camera": "標準"}


def test_design_profiles_basic():
    """design_profiles() が正しい件数・列・水準のプロファイルを返す"""
    profiles = pc.design_profiles(ATTR_LEVELS, n_profiles=12, seed=42)

    # 件数・列名
    assert len(profiles) == 12, f"件数が12でない: {len(profiles)}"
    assert list(profiles.columns) == list(ATTR_LEVELS.keys())

    # インデックスが P1〜P12
    assert profiles.index.tolist() == [f"P{i}" for i in range(1, 13)]

    # すべての行が完全交差の中に存在する
    full = pd.DataFrame(
        [dict(zip(ATTR_LEVELS.keys(), c))
         for c in __import__("itertools").product(*ATTR_LEVELS.values())]
    )
    for _, row in profiles.iterrows():
        match = (full == row.values).all(axis=1)
        assert match.any(), f"プロファイル {row.to_dict()} が完全交差に存在しない"

    # d_efficiency が attrs に保存されている
    assert "d_efficiency" in profiles.attrs
    assert "n_candidates" in profiles.attrs
    assert "det_xpx" in profiles.attrs
    assert 0 < profiles.attrs["d_efficiency"] <= 1.0
    assert profiles.attrs["n_candidates"] == 18
    print(f"  d_efficiency = {profiles.attrs['d_efficiency']:.4f}")
    print("OK test_design_profiles_basic")


def test_design_profiles_check_design():
    """design_profiles() の出力を check_design() に渡してもエラーが出ない"""
    profiles = pc.design_profiles(
        ATTR_LEVELS, n_profiles=12,
        reference_levels=REF_LEVELS, seed=42,
    )
    result = pc.check_design(profiles)
    assert isinstance(result.balance, pd.DataFrame)
    assert isinstance(result.diagnostics, list)
    print(result.summary())
    print("OK test_design_profiles_check_design")


def test_design_profiles_reproducible():
    """同じ seed を渡すと同じプロファイルが返る"""
    p1 = pc.design_profiles(ATTR_LEVELS, n_profiles=9, seed=0)
    p2 = pc.design_profiles(ATTR_LEVELS, n_profiles=9, seed=0)
    assert (p1.values == p2.values).all(), "同じ seed で結果が異なる"
    print("OK test_design_profiles_reproducible")


def test_design_profiles_full_factorial():
    """n_profiles == N のとき完全交差をすべて返す"""
    N = 3 * 2 * 3  # = 18
    profiles = pc.design_profiles(ATTR_LEVELS, n_profiles=N, seed=0)
    assert len(profiles) == N
    assert profiles.attrs["d_efficiency"] == 1.0
    print("OK test_design_profiles_full_factorial")


def test_design_profiles_errors():
    """不正な n_profiles で ValueError が出る"""
    # N を超える
    try:
        pc.design_profiles(ATTR_LEVELS, n_profiles=19)
        assert False, "ValueError が出るはず"
    except ValueError as e:
        assert "候補数" in str(e)

    # パラメータ数より少ない（p = 1+2+1+2 = 6）
    try:
        pc.design_profiles(ATTR_LEVELS, n_profiles=5)
        assert False, "ValueError が出るはず"
    except ValueError as e:
        assert "パラメータ数" in str(e)

    print("OK test_design_profiles_errors")


def test_design_duplicate_levels_rejected():
    """水準リストに重複があると design_profiles / suggest_n_profiles が
    ValueError を出す（重複は候補数 N・パラメータ数 p を架空に膨らませる）"""
    dup_levels = {"price": [6, 10, 6], "os": ["android", "apple"]}
    try:
        pc.design_profiles(dup_levels, n_profiles=4, seed=0)
        assert False, "ValueError が出るはず"
    except ValueError as e:
        assert "重複" in str(e) and "price" in str(e)

    try:
        pc.suggest_n_profiles(dup_levels)
        assert False, "ValueError が出るはず"
    except ValueError as e:
        assert "重複" in str(e) and "price" in str(e)

    print("OK test_design_duplicate_levels_rejected")


def test_design_profiles_d_efficiency_better_than_random():
    """D 最適設計の det(X'X) がランダム選択より大きい（n_starts=10）"""
    import itertools

    n_profiles = 9
    profiles_opt = pc.design_profiles(
        ATTR_LEVELS, n_profiles=n_profiles,
        reference_levels=REF_LEVELS, n_starts=10, seed=42,
    )
    det_opt = profiles_opt.attrs["det_xpx"]

    # ランダム選択を 30 回試行し、中央値と比較
    rng = np.random.default_rng(99)
    from py4conjoint.rating.design import _build_effect_matrix
    full = pd.DataFrame(
        [dict(zip(ATTR_LEVELS.keys(), c))
         for c in itertools.product(*ATTR_LEVELS.values())]
    )
    ref = {k: v[0] for k, v in ATTR_LEVELS.items()}
    X_full = _build_effect_matrix(full, ATTR_LEVELS, ref)
    N = len(full)
    rand_dets = []
    for _ in range(30):
        idx = rng.choice(N, n_profiles, replace=False)
        X_r = X_full[idx]
        rand_dets.append(np.linalg.det(X_r.T @ X_r))

    median_rand = float(np.median(rand_dets))
    assert det_opt > median_rand, (
        f"D最適の det ({det_opt:.2f}) がランダム中央値 ({median_rand:.2f}) 以下"
    )
    print(f"  det(D最適) = {det_opt:.2f}, ランダム中央値 = {median_rand:.2f}")
    print("OK test_design_profiles_d_efficiency_better_than_random")


def test_design_profiles_mixed_levels():
    """属性間で水準数が異なる場合も正しく動作する（2×3×4）"""
    attr_levels = {
        "price":   [6, 10],                         # 2水準
        "os":      ["android", "apple", "other"],   # 3水準
        "quality": ["低", "中", "高", "超高"],       # 4水準
    }
    # p = 1+1+2+3 = 7, N = 2×3×4 = 24
    profiles = pc.design_profiles(attr_levels, n_profiles=12, seed=0)
    assert len(profiles) == 12
    assert profiles.attrs["n_candidates"] == 24
    assert 0 < profiles.attrs["d_efficiency"] <= 1.0
    print("OK test_design_profiles_mixed_levels")


def test_suggest_n_profiles_basic():
    """suggest_n_profiles() が DataFrame を返し attrs に正しい値が入る"""
    result = pc.suggest_n_profiles(ATTR_LEVELS)
    assert isinstance(result, pd.DataFrame)
    assert "推奨 n_profiles" in result.columns
    assert "回答者数" in result.columns
    # attrs に基本統計が入っていること
    assert result.attrs["n_params"] == 6       # 切片1 + price2 + os1 + camera2
    assert result.attrs["n_encoded"] == 5
    assert result.attrs["n_candidates"] == 18  # 3×2×3
    assert result.attrs["m_min"] == 6          # = p
    assert result.attrs["m_orme"] == 10        # = 2×5
    # n_respondents 省略時は複数行
    assert len(result) > 1
    print("OK test_suggest_n_profiles_basic")


def test_suggest_n_profiles_with_respondents():
    """n_respondents 指定時は1行のみ返る"""
    result = pc.suggest_n_profiles(ATTR_LEVELS, n_respondents=30)
    assert len(result) == 1
    assert result.iloc[0]["回答者数"] == 30
    # 推奨値が m_min 以上
    rec = result.iloc[0]["推奨 n_profiles"]
    assert rec >= result.attrs["m_min"]
    # 推奨値が max_burden(20) 以下
    assert rec <= 20
    # obs/pred が obs_per_predictor(10) 以上
    assert result.iloc[0]["obs/pred（達成）"] >= 10.0
    print("OK test_suggest_n_profiles_with_respondents")


def test_suggest_n_profiles_small_respondents():
    """回答者が少ない場合は obs 条件が効いて推奨 M が大きくなる"""
    r5  = pc.suggest_n_profiles(ATTR_LEVELS, n_respondents=5)
    r50 = pc.suggest_n_profiles(ATTR_LEVELS, n_respondents=50)
    rec5  = r5.iloc[0]["推奨 n_profiles"]
    rec50 = r50.iloc[0]["推奨 n_profiles"]
    assert rec5 >= rec50, "回答者が少ないほど推奨Mが大きくなるか等しいはず"
    print("OK test_suggest_n_profiles_small_respondents")


def test_design_profiles_prefix():
    """profile_id_prefix を変えるとインデックスが変わる"""
    profiles = pc.design_profiles(ATTR_LEVELS, n_profiles=6, seed=0, profile_id_prefix="Card")
    assert profiles.index.tolist() == [f"Card{i}" for i in range(1, 7)]
    print("OK test_design_profiles_prefix")


def _make_mixed_level_data(seed: int = 10, n_resp: int = 20) -> pd.DataFrame:
    """price=2水準・os=2水準・camera=3水準の混在データを作る。"""
    rng = np.random.default_rng(seed)
    rows = []
    for r in range(1, n_resp + 1):
        for price in [6, 10]:
            for os in ["android", "apple"]:
                for cam in ["標準", "高性能", "超高性能"]:
                    u = 4.0
                    u += 1.0 if price == 6 else -1.0
                    u += 0.7 if os == "apple" else -0.7
                    u += {"標準": -1.2, "高性能": 0.5, "超高性能": 1.2}[cam]
                    u += rng.normal(0, 0.4)
                    rating = int(np.clip(round(u), 1, 7))
                    rows.append({
                        "respondent_id": r, "rating": rating,
                        "price": price, "os": os, "camera": cam,
                    })
    return pd.DataFrame(rows)


def test_importance_multi_level_non_price():
    """非価格属性に3水準以上がある場合も importance() の合計が100になる"""
    df = _make_mixed_level_data(seed=10, n_resp=20)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        suffix_map={"camera": ["high", "ultra"]},
    )
    result = pc.fit(df_coded)

    imp = result.importance(as_percent=True)
    # 合計が100%
    assert abs(imp["重要度"].sum() - 100.0) < 1e-6, \
        f"合計が100にならない: {imp['重要度'].sum()}"
    # 3属性が揃っていること
    assert set(imp.index) == {"price", "os", "camera"}, \
        f"属性がそろっていない: {imp.index.tolist()}"
    # 各属性の重要度が正の値
    assert (imp["重要度"] > 0).all()
    print("OK test_importance_multi_level_non_price")
    print(imp)


def test_wtp_three_level_non_price():
    """非価格属性が3水準以上の場合、wtp() に K-1 行（各非基準水準ごと1行）が出力される"""
    df = _make_mixed_level_data(seed=11, n_resp=20)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        suffix_map={"camera": ["high", "ultra"]},
    )
    result = pc.fit(df_coded)
    wtp = result.wtp()

    # 価格列はWTP出力に含まれないこと
    assert "price_0" not in wtp.index, "price_0 が WTP 出力に混入"
    # os: 2水準 → 1行
    assert "os_0" in wtp.index, "os_0 が WTP 出力にない"
    # camera: 3水準 → 2行（K-1 = 2）
    assert "camera_high" in wtp.index, "camera_high が WTP 出力にない"
    assert "camera_ultra" in wtp.index, "camera_ultra が WTP 出力にない"
    # 合計 3行（os_0, camera_high, camera_ultra）
    assert len(wtp) == 3, f"WTP 出力の行数が3以外: {len(wtp)}"
    print("OK test_wtp_three_level_non_price")
    print(wtp)


def test_cluster_robust_se_default():
    """回答者ID列があるとき、デフォルトでクラスタロバスト標準誤差が使われる"""
    df = make_synthetic_data(seed=2, n_resp=30)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    res_cluster = pc.fit(df_coded)
    res_plain = pc.fit(df_coded, cluster_se=False)
    assert res_cluster.se_type == "cluster"
    assert res_plain.se_type == "nonrobust"
    # 係数の推定値は同一、標準誤差は異なる
    assert np.allclose(res_cluster.params.values, res_plain.params.values), \
        "クラスタSEで係数が変わった（標準誤差のみ変わるはず）"
    assert not np.allclose(res_cluster.ols.bse.values, res_plain.ols.bse.values), \
        "クラスタSEと通常SEの標準誤差が同一"
    # summary に標準誤差の種別が表示される
    assert "クラスタロバスト" in res_cluster.summary()
    assert "独立性を仮定" in res_plain.summary()
    print("OK test_cluster_robust_se_default")


def test_independence_assumed_warning():
    """回答者ID列がない場合、通常SEになり independence_assumed 中警告が出る"""
    rng = np.random.default_rng(3)
    n = 40
    os_vals = rng.choice(["android", "apple"], n)
    df = pd.DataFrame({
        "rating": np.where(os_vals == "apple", 5.5, 2.5) + rng.normal(0, 0.5, n),
        "price":  rng.choice([6, 10], n),
        "os":     os_vals,
    })
    df_coded = pc.encode(df, reference_levels={"price": 10, "os": "android"})
    result = pc.fit(df_coded)
    assert result.se_type == "nonrobust"
    w = result.warnings(category="independence_assumed")
    assert len(w) == 1, f"independence_assumed が {len(w)} 件"
    assert w.iloc[0]["severity"] == "中"
    print("OK test_independence_assumed_warning")


def test_cluster_se_single_respondent_fallback():
    """回答者が1人の場合はクラスタリングせず通常SEで動く"""
    df = make_synthetic_data(seed=0, n_resp=1)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded)  # エラーにならないこと
    assert result.se_type == "nonrobust"
    # few_respondents（大）は出るが independence_assumed は出ない（列はあるため）
    cats = result.warnings()["category"].values
    assert "few_respondents" in cats
    assert "independence_assumed" not in cats
    print("OK test_cluster_se_single_respondent_fallback")


def test_encode_attrs_no_side_effect():
    """encode() が入力 DataFrame の attrs（ネスト辞書）を書き換えない"""
    df = make_synthetic_data(seed=0, n_resp=5)
    out1 = pc.encode(df, reference_levels={"price": 10})
    refs1 = dict(out1.attrs["py4conjoint"]["reference_levels"])
    out2 = pc.encode(out1, reference_levels={"os": "android"})
    # out1 の attrs は変わっていないこと
    assert out1.attrs["py4conjoint"]["reference_levels"] == refs1, \
        "encode() が入力側の attrs を書き換えた"
    assert "os" not in out1.attrs["py4conjoint"]["reference_levels"]
    # out2 には両方の基準水準が引き継がれていること
    assert set(out2.attrs["py4conjoint"]["reference_levels"]) == {"price", "os"}
    print("OK test_encode_attrs_no_side_effect")


def test_detect_encoded_columns_prefix_collision():
    """属性名が別の属性名の接頭辞でも符号化列の検出が重複・誤検出しない"""
    df = pd.DataFrame({
        "rating": [5, 3, 6, 4, 5, 3, 6, 4],
        "os": ["android", "apple"] * 4,
        "os_version": ["v1", "v2", "v2", "v1"] * 2,
    })
    df_coded = pc.encode(
        df, reference_levels={"os": "android", "os_version": "v1"}
    )
    result = pc.fit(df_coded)
    # "os_" が "os_version"（元の文字列列）や "os_version_0"（別属性の符号化列）に
    # 前方一致しても、重複なく2列だけ検出されること
    assert sorted(result.encoded_columns) == ["os_0", "os_version_0"], \
        f"検出された列: {result.encoded_columns}"
    # importance() も属性ごとに正しくグルーピングされること
    imp = result.importance()
    assert set(imp.index) == {"os", "os_version"}
    print("OK test_detect_encoded_columns_prefix_collision")


def test_detect_encoded_columns_fallback_excludes_01():
    """reference_levels なしの自動検出は 0/1 列（respondent_encode 出力）を含めない"""
    df = pd.DataFrame({
        "rating": [5, 3, 6, 4, 5, 4, 6, 3],
        "price_0": [1, -1, 1, -1, 1, -1, 1, -1],
        "gender_0": [0, 1, 0, 1, 1, 0, 1, 0],
    })
    result = pc.fit(df)  # attrs なし → フォールバック検出
    assert result.encoded_columns == ["price_0"], \
        f"0/1 列が混入した: {result.encoded_columns}"
    print("OK test_detect_encoded_columns_fallback_excludes_01")


def test_fit_with_formula_consistent():
    """formula 指定時は説明変数が formula から取得され importance()/wtp() が動く"""
    df = make_synthetic_data(seed=8, n_resp=20)
    df_coded = pc.encode(
        df, reference_levels={"price": 10, "os": "android", "camera": "標準"}
    )
    result = pc.fit(df_coded, formula="rating ~ price_0 + os_0")
    # encoded_columns は formula の右辺と一致（camera_0 は含まれない）
    assert result.encoded_columns == ["price_0", "os_0"], \
        f"encoded_columns: {result.encoded_columns}"
    # importance()/wtp() が KeyError にならない
    imp = result.importance()
    assert set(imp.index) == {"price", "os"}
    wtp = result.wtp()
    assert list(wtp.index) == ["os_0"]
    print("OK test_fit_with_formula_consistent")


def test_wtp_p_price_three_levels_joint_f_test():
    """3水準価格の p_price は全価格係数の同時F検定のp値"""
    rng = np.random.default_rng(42)
    rows = []
    for r in range(1, 26):
        for price, os in [(6, "android"), (8, "android"), (10, "android"),
                          (6, "apple"),   (8, "apple"),   (10, "apple")]:
            u = 4.0 + {6: 1.5, 8: 0.0, 10: -1.5}[price]
            u += 0.7 if os == "apple" else -0.7
            u += rng.normal(0, 0.3)
            rows.append({"respondent_id": r, "rating": int(np.clip(round(u), 1, 7)),
                         "price": price, "os": os})
    df = pd.DataFrame(rows)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android"},
        suffix_map={"price": ["low", "mid"]},
    )
    result = pc.fit(df_coded, price_col="price")
    wtp = result.wtp()
    # 同時F検定のp値と一致すること
    exog = list(result.ols.model.exog_names)
    R = np.zeros((2, len(exog)))
    R[0, exog.index("price_low")] = 1.0
    R[1, exog.index("price_mid")] = 1.0
    expected = float(result.ols.f_test(R).pvalue)
    assert abs(wtp.attrs["p_price"] - expected) < 1e-12, \
        f"p_price={wtp.attrs['p_price']} != F検定 {expected}"
    # 価格効果が強いデータなので有意のはず
    assert wtp.attrs["p_price"] < 0.05
    print("OK test_wtp_p_price_three_levels_joint_f_test")


def test_wtp_multi_level_non_price_matches_definition():
    """3水準の非価格属性のWTPが「基準水準からの効用差 × 評点1点の金額」と一致する"""
    df = _make_mixed_level_data(seed=12, n_resp=25)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android", "camera": "標準"},
        suffix_map={"camera": ["high", "ultra"]},
    )
    result = pc.fit(df_coded)
    wtp = result.wtp()
    factor = wtp.attrs["wtp_price_factor"]
    b_high = float(result.params["camera_high"])
    b_ultra = float(result.params["camera_ultra"])
    # 効果コーディングでは基準水準（標準）の効用 = -(b_high + b_ultra) なので
    # 標準→高性能 の効用差 = b_high - (-(b_high+b_ultra)) = 2*b_high + b_ultra
    expected_high = (2 * b_high + b_ultra) * factor / 2
    expected_ultra = (b_high + 2 * b_ultra) * factor / 2
    assert abs(wtp.loc["camera_high", "限界支払意思額"] - expected_high) < 1e-9, \
        f"camera_high: {wtp.loc['camera_high', '限界支払意思額']} != {expected_high}"
    assert abs(wtp.loc["camera_ultra", "限界支払意思額"] - expected_ultra) < 1e-9, \
        f"camera_ultra: {wtp.loc['camera_ultra', '限界支払意思額']} != {expected_ultra}"
    # 2水準属性は従来式 factor * b と同値のまま
    b_os = float(result.params["os_0"])
    assert abs(wtp.loc["os_0", "限界支払意思額"] - factor * b_os) < 1e-9
    # データ生成の真値（標準→高性能 +1.7 効用、価格感度 0.5/万円）に近いこと
    # WTP ≈ 1.7 / 0.5 = 3.4 万円（ノイズがあるため緩めの範囲で確認）
    assert 2.5 < wtp.loc["camera_high", "限界支払意思額"] < 4.5, \
        f"WTPが真値3.4から大きく外れている: {wtp.loc['camera_high', '限界支払意思額']}"
    print("OK test_wtp_multi_level_non_price_matches_definition")


def test_suggest_n_profiles_max_burden_below_p():
    """max_burden < p のとき推奨値が p に引き上げられ UserWarning が出る"""
    import warnings as _warnings
    big = {f"a{i}": [1, 2, 3, 4] for i in range(8)}  # p = 1 + 3×8 = 25 > 20
    with _warnings.catch_warnings(record=True) as w:
        _warnings.simplefilter("always")
        result = pc.suggest_n_profiles(big, n_respondents=30)
    p = result.attrs["n_params"]
    assert p == 25
    rec = result.iloc[0]["推奨 n_profiles"]
    assert rec >= p, f"推奨 n_profiles ({rec}) がパラメータ数 p ({p}) 未満（回帰不能）"
    assert any(issubclass(wi.category, UserWarning) for wi in w), \
        "max_burden < p なのに UserWarning が出ていない"
    # 推奨値が design_profiles() でそのまま使えること
    profiles = pc.design_profiles(big, n_profiles=int(rec), n_starts=1, seed=0)
    assert len(profiles) == rec
    print("OK test_suggest_n_profiles_max_burden_below_p")


def test_unit_rating_money_three_level_price():
    """unit_rating_money() が3水準以上の価格でも正の float を返す"""
    rng = np.random.default_rng(42)
    n_resp = 25
    rows = []
    for r in range(1, n_resp + 1):
        for price, os in [(6, "android"), (8, "android"), (10, "android"),
                          (6, "apple"),   (8, "apple"),   (10, "apple")]:
            u = 4.0 + {6: 1.5, 8: 0.0, 10: -1.5}[price]
            u += 0.7 if os == "apple" else -0.7
            u += rng.normal(0, 0.3)
            rating = int(np.clip(round(u), 1, 7))
            rows.append({"respondent_id": r, "rating": rating, "price": price, "os": os})
    df = pd.DataFrame(rows)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android"},
        suffix_map={"price": ["low", "mid"]},
    )
    result = pc.fit(df_coded, price_col="price")

    unit = result.unit_rating_money()
    assert isinstance(unit, float), f"float が返されるべきだが {type(unit)} が返された"
    assert unit > 0, f"unit_rating_money が正でない: {unit}"
    # wtp_price_factor / 2 と等価であることを確認
    wtp = result.wtp()
    expected = wtp.attrs["wtp_price_factor"] / 2.0
    assert abs(unit - expected) < 1e-9, f"wtp_price_factor/2 との不一致: {unit} vs {expected}"
    print(f"  3水準価格 unit_rating_money = {unit:.4f}")
    print("OK test_unit_rating_money_three_level_price")


def test_wtp_three_level_price_warnings_not_duplicated():
    """3水準価格で wtp(method='linear') を複数回呼んでも警告が重複しない"""
    rng = np.random.default_rng(42)
    n_resp = 25
    rows = []
    for r in range(1, n_resp + 1):
        for price, os in [(6, "android"), (8, "android"), (10, "android"),
                          (6, "apple"),   (8, "apple"),   (10, "apple")]:
            u = 4.0 + {6: 1.5, 8: 0.0, 10: -1.5}[price]
            u += 0.7 if os == "apple" else -0.7
            u += rng.normal(0, 0.3)
            rating = int(np.clip(round(u), 1, 7))
            rows.append({"respondent_id": r, "rating": rating, "price": price, "os": os})
    df = pd.DataFrame(rows)
    df_coded = pc.encode(
        df,
        reference_levels={"price": 10, "os": "android"},
        suffix_map={"price": ["low", "mid"]},
    )
    result = pc.fit(df_coded, price_col="price")

    _ = result.wtp(method="linear")
    _ = result.wtp(method="linear")
    _ = result.wtp(method="linear")

    from collections import Counter
    all_cats = [d.category for d in result._diagnostics]
    dupes = [k for k, v in Counter(all_cats).items()
             if v > 1 and not k.startswith("wtp_extrapolation_")]
    assert not dupes, f"重複した警告カテゴリ: {dupes}"
    # wtp_price_linear_approx は1件だけ
    approx_count = sum(1 for c in all_cats if c == "wtp_price_linear_approx")
    assert approx_count == 1, f"wtp_price_linear_approx が {approx_count} 件（期待: 1）"
    print("OK test_wtp_three_level_price_warnings_not_duplicated")


def test_check_design_three_level_attribute():
    """3水準以上の属性を含むデザインでcheck_design()が正しく機能し、誤った相関警告が出ない"""
    # price=3水準 × os=2水準 の完全直交デザイン
    profiles = pd.DataFrame({
        "price": [6, 8, 10, 6, 8, 10],
        "os":    ["android", "android", "android", "apple", "apple", "apple"],
    })
    result = pc.check_design(profiles)

    # balance: すべてCV=0 → 均等
    assert (result.balance["CV"] == 0.0).all(), \
        f"バランスが崩れている: {result.balance}"

    # 相関: 同一属性内の price__0 × price__1 ペアが誤検知されないこと
    cats = [d.category for d in result.diagnostics]
    assert not any("correlation" in c for c in cats), \
        f"誤った相関警告が出た（同一属性内ペアのスキップ漏れ）: {cats}"

    # chi2: 完全直交なので警告なし
    assert not any("chi2" in c for c in cats), \
        f"chi2 警告が出た: {cats}"

    # balance の水準数が正しい
    assert result.balance.loc["price", "水準数"] == 3
    assert result.balance.loc["os", "水準数"] == 2
    print("OK test_check_design_three_level_attribute")


def test_check_design_correlated():
    """相関が高い（完全共線）デザインで相関警告とχ²警告が出る"""
    # price と os が常に同じ組み合わせで出現（price=低 ↔ os=android、price=高 ↔ os=apple）
    profiles = pd.DataFrame({
        "price": [6, 6, 10, 10],
        "os":    ["android", "android", "apple", "apple"],
    })
    result = pc.check_design(profiles)
    cats = [d.category for d in result.diagnostics]

    # 完全相関 → 相関警告（大）が出るはず
    assert any("correlation" in c for c in cats), \
        f"相関警告が出ていない: {cats}"
    corr_diags = [d for d in result.diagnostics if "correlation" in d.category]
    assert corr_diags[0].severity == "大", \
        f"相関警告の重大度が「大」でない: {corr_diags[0].severity}"

    # χ²も高いはず → chi2警告が出るはず
    assert any("chi2" in c for c in cats), \
        f"chi2 警告が出ていない: {cats}"
    print("OK test_check_design_correlated")


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
    test_encode_multi_with_suffix_map()
    test_encode_multi_suffix_length_mismatch()
    test_binary_suffix_map_deprecation()
    test_check_design_basic()
    test_check_design_imbalanced()
    test_encode_binary_suffix_map_as_list()
    test_wtp_three_level_price_no_extra_columns()
    test_wtp_three_level_price()
    test_importance_multi_level_non_price()
    test_wtp_three_level_non_price()
    test_wtp_multi_level_non_price_matches_definition()
    test_suggest_n_profiles_max_burden_below_p()
    test_cluster_robust_se_default()
    test_independence_assumed_warning()
    test_cluster_se_single_respondent_fallback()
    test_encode_attrs_no_side_effect()
    test_detect_encoded_columns_prefix_collision()
    test_detect_encoded_columns_fallback_excludes_01()
    test_fit_with_formula_consistent()
    test_wtp_p_price_three_levels_joint_f_test()
    test_unit_rating_money_three_level_price()
    test_wtp_three_level_price_warnings_not_duplicated()
    test_check_design_three_level_attribute()
    test_check_design_correlated()
    test_suggest_n_profiles_basic()
    test_suggest_n_profiles_with_respondents()
    test_suggest_n_profiles_small_respondents()
    test_design_profiles_basic()
    test_design_profiles_check_design()
    test_design_profiles_reproducible()
    test_design_profiles_full_factorial()
    test_design_profiles_errors()
    test_design_profiles_d_efficiency_better_than_random()
    test_design_profiles_mixed_levels()
    test_design_profiles_prefix()
    print("\nすべてのテストがパスしました。")
