"""choice サブパッケージ（選択型コンジョイント分析）のテスト。

(a) 人工データでの真値回復
(b) 完全分離データでの収束警告
(c) 外部検証: R の logitr による yogurt データの推定結果との一致
    （参照値・許容誤差は tests/data/logitr_yogurt_reference.py で定義）
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent / "data"))

import logitr_yogurt_reference as ref
import numpy as np
import pandas as pd
import pytest

import py4conjoint.choice as pcc

DATA_DIR = Path(__file__).resolve().parent / "data"


# ---------------------------------------------------------------------------
# ヘルパー：人工データの生成
# ---------------------------------------------------------------------------

TRUE_BETA = {"price": -0.8, "brand_B": 0.5, "brand_C": -0.4}


def _simulate_choice_data(n_sets=100_000, n_alts=3, seed=42) -> pd.DataFrame:
    """真の係数 TRUE_BETA から条件付きロジットに従う選択データを生成する。"""
    rng = np.random.default_rng(seed)
    price = rng.choice([2.0, 3.0, 4.0], size=(n_sets, n_alts))
    brand = rng.choice(["A", "B", "C"], size=(n_sets, n_alts))
    v = (
        TRUE_BETA["price"] * price
        + (brand == "B") * TRUE_BETA["brand_B"]
        + (brand == "C") * TRUE_BETA["brand_C"]
    )
    p = np.exp(v - v.max(axis=1, keepdims=True))
    p /= p.sum(axis=1, keepdims=True)
    chosen = (rng.random((n_sets, 1)) < p.cumsum(axis=1)).argmax(axis=1)
    return pd.DataFrame(
        {
            "選択セットID": np.repeat(np.arange(n_sets), n_alts),
            "choice": (
                np.tile(np.arange(n_alts), n_sets) == np.repeat(chosen, n_alts)
            ).astype(int),
            "price": price.ravel(),
            "brand": brand.ravel(),
        }
    )


# ---------------------------------------------------------------------------
# encode()（ダミーコーディング）
# ---------------------------------------------------------------------------


def test_encode_dummy_coding():
    """encode() は 0/1 のダミーコーディングを行い、基準水準は全列 0 になる"""
    df = pd.DataFrame(
        {
            "選択セットID": [1, 1, 1, 2, 2, 2],
            "choice": [1, 0, 0, 0, 1, 0],
            "brand": ["A", "B", "C", "B", "A", "C"],
        }
    )
    out = pcc.encode(df, reference_levels={"brand": "A"})
    assert "brand_B" in out.columns and "brand_C" in out.columns
    # 基準水準 A の行はすべて 0
    a_rows = out[out["brand"] == "A"]
    assert (a_rows["brand_B"] == 0).all() and (a_rows["brand_C"] == 0).all()
    # 各ダミー列は 0/1 のみ
    assert set(out["brand_B"].unique()) == {0, 1}
    assert set(out["brand_C"].unique()) == {0, 1}
    # B の行は brand_B=1, brand_C=0
    b_rows = out[out["brand"] == "B"]
    assert (b_rows["brand_B"] == 1).all() and (b_rows["brand_C"] == 0).all()
    # メタ情報が attrs に保存される
    meta = out.attrs["py4conjoint"]
    assert meta["encoding"] == "dummy"
    assert meta["reference_levels"] == {"brand": "A"}
    assert meta["encoded_columns"] == {"brand": ["brand_B", "brand_C"]}
    print("OK test_encode_dummy_coding")


def test_encode_errors():
    """encode() の入力チェック（日本語エラー）"""
    df = pd.DataFrame({"brand": ["A", "B"]})
    with pytest.raises(ValueError, match="基準水準"):
        pcc.encode(df, reference_levels={"brand": "Z"})
    with pytest.raises(ValueError, match="DataFrame にありません"):
        pcc.encode(df, reference_levels={"no_such": "A"})
    with pytest.raises(ValueError, match="水準が1つ"):
        pcc.encode(pd.DataFrame({"brand": ["A", "A"]}), reference_levels={"brand": "A"})
    with pytest.raises(ValueError, match="reference_levels"):
        pcc.encode(df, reference_levels={})
    print("OK test_encode_errors")


# ---------------------------------------------------------------------------
# (a) 人工データでの真値回復
# ---------------------------------------------------------------------------


def test_synthetic_recovery():
    """人工データから真の係数を許容誤差 1e-2 程度で回復できる"""
    df = _simulate_choice_data(n_sets=100_000, seed=42)
    df_coded = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(df_coded, choice="choice", choice_set_id_col="選択セットID")

    assert result.converged
    # encoded_columns は encode() のメタ情報 + price から自動検出される
    assert result.encoded_columns == ["price", "brand_B", "brand_C"]
    for name, true_val in TRUE_BETA.items():
        assert abs(result.params[name] - true_val) < 1.5e-2, (
            f"{name}: 推定値 {result.params[name]:.4f} が真値 {true_val} から乖離"
        )
    print("OK test_synthetic_recovery")


def test_synthetic_api_consistency():
    """summary/importance/wtp/market_share が rating 版と同じ形式で動く"""
    df = _simulate_choice_data(n_sets=5_000, seed=7)
    df_coded = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(df_coded, choice="choice", choice_set_id_col="選択セットID")

    # summary は和文の文字列
    s = result.summary()
    assert isinstance(s, str)
    assert "選択型コンジョイント分析の結果" in s
    assert "対数尤度" in s and "McFadden" in s
    s_full = result.summary(slim=False)
    assert "標準誤差" in s_full and "z値" in s_full

    # importance：列名は rating 版と同じ（効用範囲・重要度）、合計100
    imp = result.importance(as_percent=True)
    assert list(imp.columns) == ["効用範囲", "重要度"]
    assert imp.index.name == "属性"
    assert set(imp.index) == {"price", "brand"}
    assert abs(imp["重要度"].sum() - 100.0) < 1e-9
    imp_ratio = result.importance(as_percent=False)
    assert abs(imp_ratio["重要度"].sum() - 1.0) < 1e-12

    # wtp：列名は rating 版と同じ（係数・限界支払意思額）
    w = result.wtp()
    assert list(w.columns) == ["係数", "限界支払意思額"]
    assert set(w.index) == {"brand_B", "brand_C"}
    # MWTP = -b_attr / b_price
    expected = -result.params["brand_B"] / result.params["price"]
    assert np.isclose(w.loc["brand_B", "限界支払意思額"], expected)

    # market_share：合計1の Series
    products = pd.DataFrame(
        {
            "price": [2.0, 4.0],
            "brand_B": [1, 0],
            "brand_C": [0, 1],
        },
        index=["製品X", "製品Y"],
    )
    share = result.market_share(products)
    assert isinstance(share, pd.Series)
    assert abs(share.sum() - 1.0) < 1e-12
    assert share["製品X"] > share["製品Y"]  # 安い & 人気ブランド
    share_max = result.market_share(products, method="max")
    assert share_max["製品X"] == 1.0 and share_max["製品Y"] == 0.0

    # warnings：rating 版と同じ列構成の DataFrame
    w_df = result.warnings()
    assert list(w_df.columns) == ["severity", "category", "message", "recommendation"]
    print("OK test_synthetic_api_consistency")


# ---------------------------------------------------------------------------
# (b) 完全分離データでの警告
# ---------------------------------------------------------------------------


def test_separation_warning():
    """完全分離データでは separation 警告（重大度：大）が出る"""
    n_sets = 40
    # 選ばれた代替案は常に x=1、選ばれなかった代替案は常に x=0 → 完全分離
    df = pd.DataFrame(
        {
            "選択セットID": np.repeat(np.arange(n_sets), 2),
            "choice": np.tile([1, 0], n_sets),
            "x": np.tile([1.0, 0.0], n_sets),
        }
    )
    result = pcc.fit(
        df,
        choice="choice",
        choice_set_id_col="選択セットID",
        encoded_columns=["x"],
    )
    sep = result.warnings(category="separation")
    assert len(sep) == 1
    assert sep.iloc[0]["severity"] == "大"
    assert "完全分離" in sep.iloc[0]["message"]
    # summary（重大度「大」）にも表示される
    assert "完全分離" in result.summary()
    print("OK test_separation_warning")


# ---------------------------------------------------------------------------
# (c) 外部検証：R の logitr による yogurt データの推定結果と一致
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def yogurt_result():
    df = pd.read_csv(DATA_DIR / "yogurt.csv")
    assert len(df) == ref.N_OBS * ref.N_ALTS  # 2412 × 4 = 9648 行
    # yogurt.csv の選択セット列名は logitr 由来の慣習名 "obsID"。
    # 公開 API の列名（choice_set_id）にリネームしてから fit に渡す
    # （外部データの CSV 自体は obsID のまま）。
    df = df.rename(columns={"obsID": "choice_set_id"})
    df_coded = pcc.encode(df, reference_levels={"brand": ref.REFERENCE_LEVEL})
    return pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="choice_set_id",
        encoded_columns=[
            "price",
            "feat",
            "brand_hiland",
            "brand_weight",
            "brand_yoplait",
        ],
    )


# logitr の係数名（brandhiland）→ encode() の列名（brand_hiland）の対応
_NAME_MAP = {
    "price": "price",
    "feat": "feat",
    "brandhiland": "brand_hiland",
    "brandweight": "brand_weight",
    "brandyoplait": "brand_yoplait",
}


def test_yogurt_coefficients_match_logitr(yogurt_result):
    """係数が logitr の参照値と一致する（rtol=RTOL_COEF）"""
    assert yogurt_result.converged
    assert yogurt_result.n_sets == ref.N_OBS
    for logitr_name, our_name in _NAME_MAP.items():
        est = yogurt_result.params[our_name]
        expected = ref.COEF[logitr_name]
        assert est == pytest.approx(expected, rel=ref.RTOL_COEF), (
            f"{our_name}: {est:.6f} != logitr {expected:.6f}"
        )
    print("OK test_yogurt_coefficients_match_logitr")


def test_yogurt_std_errors_match_logitr(yogurt_result):
    """標準誤差（ヘッセ行列の逆行列）が logitr の参照値と一致する（rtol=RTOL_SE）"""
    assert yogurt_result.se_type == "nonrobust"
    for logitr_name, our_name in _NAME_MAP.items():
        est = yogurt_result.bse[our_name]
        expected = ref.STD_ERR[logitr_name]
        assert est == pytest.approx(expected, rel=ref.RTOL_SE), (
            f"{our_name}: SE {est:.6f} != logitr {expected:.6f}"
        )
    print("OK test_yogurt_std_errors_match_logitr")


def test_yogurt_loglik_match_logitr(yogurt_result):
    """対数尤度・帰無対数尤度が logitr の参照値と一致する（atol=ATOL_LL）"""
    assert yogurt_result.loglik == pytest.approx(ref.LOG_LIKELIHOOD, abs=ref.ATOL_LL)
    assert yogurt_result.null_loglik == pytest.approx(
        ref.NULL_LOG_LIKELIHOOD, abs=ref.ATOL_LL
    )
    print("OK test_yogurt_loglik_match_logitr")


def test_yogurt_cluster_se():
    """回答者ID列を指定するとクラスタロバスト標準誤差になる"""
    df = pd.read_csv(DATA_DIR / "yogurt.csv")
    # CSV の慣習列名 obsID を公開 API の choice_set_id にリネーム
    df = df.rename(columns={"obsID": "choice_set_id"})
    df_coded = pcc.encode(df, reference_levels={"brand": ref.REFERENCE_LEVEL})
    cols = ["price", "feat", "brand_hiland", "brand_weight", "brand_yoplait"]
    result = pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="choice_set_id",
        encoded_columns=cols,
        respondent_id_col="id",
    )
    assert result.se_type == "cluster"
    # 係数はクラスタリングの有無で変わらない
    nonrobust = pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="choice_set_id",
        encoded_columns=cols,
        cluster_se=False,
        respondent_id_col="id",
    )
    assert np.allclose(result.params.to_numpy(), nonrobust.params.to_numpy())
    # 標準誤差は変わる（同一回答者の繰り返し選択があるため通常は大きくなる）
    assert not np.allclose(result.bse.to_numpy(), nonrobust.bse.to_numpy())
    assert (result.bse.to_numpy() > nonrobust.bse.to_numpy()).all()
    print("OK test_yogurt_cluster_se")


# ---------------------------------------------------------------------------
# CBC固有の警告カテゴリ
# ---------------------------------------------------------------------------


def test_few_choice_sets_warning():
    """選択セット数が説明変数数の5倍未満なら few_choice_sets（大）が出る"""
    df = _simulate_choice_data(n_sets=8, seed=0)
    df_coded = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(df_coded, choice="choice", choice_set_id_col="選択セットID")
    # 8 セット / 3 変数 = 2.7 倍 < 5
    w = result.warnings(category="few_choice_sets")
    assert len(w) == 1
    assert w.iloc[0]["severity"] == "大"
    print("OK test_few_choice_sets_warning")


def test_unbalanced_choices_warning():
    """選択が特定の代替案位置に偏ると unbalanced_choices が出る"""
    n_sets = 50
    rng = np.random.default_rng(3)
    # 9割の選択セットで1番目の代替案が選ばれる（x はほぼ無関係）
    first = rng.random(n_sets) < 0.9
    rows = []
    for t in range(n_sets):
        chosen_pos = 0 if first[t] else 1
        for j in range(2):
            rows.append(
                {
                    "選択セットID": t,
                    "choice": int(j == chosen_pos),
                    "x": float(rng.random()),
                }
            )
    df = pd.DataFrame(rows)
    result = pcc.fit(
        df, choice="choice", choice_set_id_col="選択セットID", encoded_columns=["x"]
    )
    w = result.warnings(category="unbalanced_choices")
    assert len(w) == 1
    assert "代替案" in w.iloc[0]["message"]
    print("OK test_unbalanced_choices_warning")


# ---------------------------------------------------------------------------
# 入力検証（日本語エラー）
# ---------------------------------------------------------------------------


def test_fit_validation_errors():
    """fit() の入力チェック（選択セット構造の検証を含む）"""
    base = pd.DataFrame(
        {
            "選択セットID": [1, 1, 2, 2],
            "choice": [1, 0, 0, 1],
            "x": [1.0, 0.0, 0.5, 0.2],
        }
    )

    # choice 列がない
    with pytest.raises(ValueError, match="choice"):
        pcc.fit(
            base.drop(columns=["choice"]),
            choice="choice",
            choice_set_id_col="選択セットID",
            encoded_columns=["x"],
        )

    # choice 列が 0/1 でない
    bad = base.copy()
    bad["choice"] = [1, 2, 0, 1]
    with pytest.raises(ValueError, match="0/1"):
        pcc.fit(
            bad,
            choice="choice",
            choice_set_id_col="選択セットID",
            encoded_columns=["x"],
        )

    # 選択がちょうど1つでない選択セット
    bad2 = base.copy()
    bad2["choice"] = [1, 1, 0, 1]
    with pytest.raises(ValueError, match="ちょうど1つ"):
        pcc.fit(
            bad2,
            choice="choice",
            choice_set_id_col="選択セットID",
            encoded_columns=["x"],
        )

    # 代替案が1つしかない選択セット
    bad3 = pd.DataFrame(
        {
            "選択セットID": [1, 1, 2],
            "choice": [1, 0, 1],
            "x": [1.0, 0.0, 0.5],
        }
    )
    with pytest.raises(ValueError, match="代替案が1つ"):
        pcc.fit(
            bad3,
            choice="choice",
            choice_set_id_col="選択セットID",
            encoded_columns=["x"],
        )

    # 説明変数が見つからない（encode() 未実施・encoded_columns 未指定）
    with pytest.raises(ValueError, match="説明変数が見つかりません"):
        pcc.fit(
            base[["選択セットID", "choice"]],
            choice="choice",
            choice_set_id_col="選択セットID",
        )

    print("OK test_fit_validation_errors")


def test_wtp_requires_linear_price():
    """価格が説明変数に含まれない場合、wtp() は日本語エラーを出す"""
    df = _simulate_choice_data(n_sets=200, seed=5)
    df_coded = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="選択セットID",
        encoded_columns=["brand_B", "brand_C"],
    )
    with pytest.raises(ValueError, match="価格列"):
        result.wtp()
    print("OK test_wtp_requires_linear_price")
