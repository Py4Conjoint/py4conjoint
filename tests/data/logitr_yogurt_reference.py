"""logitr (R) による yogurt データの MNL 推定結果（外部検証用の参照値）。

生成方法（R で一度だけ実行）:

    library(logitr)
    mnl <- logitr(
      data    = yogurt,
      outcome = "choice",
      obsID   = "obsID",
      pars    = c("price", "feat", "brand")
    )
    summary(mnl)

- logitr バージョン: ユーザー実行環境（2026-06-12 取得）
- アルゴリズム: NLOPT_LD_LBFGS
- 基準水準: brand = "dannon"（R の factor が自動的に脱落させる水準）
- データ: tests/data/yogurt.csv（logitr 付属データを write.csv でエクスポート）
  2412 選択セット × 4 代替案 = 9648 行

Python 実装（scipy.optimize, BFGS, 解析的勾配）による事前検証で、
係数・SE は相対誤差 1e-4 未満、対数尤度は小数第6位まで一致することを
確認済み（2026-06-12）。テスト許容誤差は rtol=1e-3 を推奨。
"""

# 変数の並び順は encode() の出力と揃えること
COEF_NAMES = ["price", "feat", "brandhiland", "brandweight", "brandyoplait"]

# fmt: off
# R の logitr の出力と目視で照合するための表。小数点の位置を揃えてある。
COEF = {
    "price":        -0.366555,
    "feat":          0.491439,
    "brandhiland":  -3.715477,
    "brandweight":  -0.641138,
    "brandyoplait":  0.734519,
}

STD_ERR = {
    "price":         0.024365,
    "feat":          0.120062,
    "brandhiland":   0.145417,
    "brandweight":   0.054498,
    "brandyoplait":  0.080642,
}
# fmt: on

LOG_LIKELIHOOD = -2656.8878790
NULL_LOG_LIKELIHOOD = -3343.7419990
N_OBS = 2412  # 選択セット数
N_ALTS = 4  # 1セットあたりの代替案数
REFERENCE_LEVEL = "dannon"  # brand の基準水準

# 検証テストの推奨許容誤差
RTOL_COEF = 1e-3
RTOL_SE = 1e-3
ATOL_LL = 1e-3
