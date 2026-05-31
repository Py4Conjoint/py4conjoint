# py4conjoint

評点型コンジョイント分析を **Python初心者でも直感的に** 行えるパッケージです。

Microsoft Forms / Google Forms のアンケート回答ファイルを読み込み、
符号化・回帰分析・結果の解釈・可視化までを一貫して行えます。

## インストール

```bash
pip install py4conjoint
```

Google Colab では：

```python
!pip install py4conjoint
```

## クイックスタート

### 1. アンケートデータを読み込む

```python
import pandas as pd
import py4conjoint as pc

# プロファイル設計を作成
profiles = { # P1         P2       P3        P4
    "price":  [6,         10,      6,        10],
    "os":     ["android", "apple", "apple",  "android"],
    "camera": ["標準",    "標準",  "高性能", "高性能"],
}

# Microsoft Forms の回答ファイルを読み込む（デフォルト）
df = pc.forms_to_conjoint_data(
    responses_file  = "responses.xlsx",
    attributes      = profiles,
    respondent_cols = {"性別": "gender"},
)

# Google Forms の場合
df = pc.forms_to_conjoint_data(
    responses_file  = "responses.csv",
    attributes      = profiles,
    respondent_cols = {"性別": "gender"},
    forms           = "google",
)
```

### 2. 符号化する

```python
df_coded = pc.encode(
    df,
    reference_levels = {
        "price":  10,
        "os":     "android",
        "camera": "標準"
    },
    binary_suffix_map = {
        "price":  "low",
        "os":     "apple",
        "camera": "high"
    },
)
```

### 3. 回帰分析を実行する

```python
result = pc.fit(df_coded)
print(result.summary())
```

```
============================================================
コンジョイント分析の結果（和文サマリー）
============================================================
観測数         : 120
説明変数の数   : 3
決定係数 R²    : 0.9582
自由度修正 R²  : 0.9571

【推定された係数（部分効用 part-worth）】
  変数                            係数        p値  有意性
  ------------------------- ---------- ----------  ------
  Intercept                     4.2417     0.0000    ***
  price_low                     1.2583     0.0000    ***
  os_apple                      0.9083     0.0000    ***
  camera_high                   0.5917     0.0000    ***

  有意水準: *** p<0.001  ** p<0.01  * p<0.05  . p<0.1
============================================================
```

### 4. 結果を解釈する

```python
# 相対重要度（合計100%）
result.importance()
#              range  importance
# 属性
# price       2.5167      45.62
# os          1.8167      32.94
# camera      1.1833      21.44

# WTP（支払意思額）
result.wtp()
#                      係数  支払意思額
# 属性（符号化列名）
# os_apple           0.9083     2.8874
# camera_high        0.5917     1.8805

# 評点1点の金額換算
result.unit_rating_money()
# 1.5894

# 市場シェア予測
products = pd.DataFrame({
    "price_low":   [1, -1],
    "os_apple":    [1,  1],
    "camera_high": [1, -1],
}, index=["製品A", "製品B"])

result.market_share(products)
# 製品A    0.926
# 製品B    0.074
```

### 5. 可視化する

```python
result.plot_importance()   # 相対重要度の棒グラフ
result.plot_partworth()    # 部分効用の棒グラフ
result.plot_wtp(price_unit="万円")  # WTPの棒グラフ
```

## 主な機能

| 関数 / メソッド | 説明 |
|----------------|------|
| `forms_to_conjoint_data()` | Microsoft/Google Forms の回答ファイルを long 形式 DataFrame に変換 |
| `encode()` | 属性列を効果コーディング（-1/+1）に自動変換 |
| `auto_reference_levels()` | 基準水準を自動推測（補助関数） |
| `fit()` | OLS 回帰を実行し `ConjointResult` を返す |
| `result.summary()` | 係数表・R²・落とし穴チェックの和文サマリー |
| `result.warnings()` | 落とし穴の一覧（severity / category でフィルタ可） |
| `result.importance()` | 各属性の相対重要度（合計100%） |
| `result.wtp()` | 各属性の WTP（支払意思額） |
| `result.unit_rating_money()` | 評点1点の金額換算（float） |
| `result.market_share()` | 市場シェア予測（logit / max） |
| `result.plot_importance()` | 相対重要度の棒グラフ |
| `result.plot_partworth()` | 部分効用（パートワース）の棒グラフ |
| `result.plot_wtp()` | WTP の棒グラフ |

### 落とし穴の自動検出

`fit()` と `wtp()` は、以下の問題を自動的に検出して警告します：

| カテゴリ | 重大度 | 内容 |
|----------|:------:|------|
| `r2_low` | 大 | R² < 0.20（説明力が低い） |
| `obs_per_predictor` | 大/中 | 観測数／説明変数数の比率が低い（< 5 → 大、< 10 → 中） |
| `few_respondents` | 大/中 | 回答者数が少ない（1人→大、2〜4人→中） |
| `price_sign_negative` | 中 | 価格係数の符号が逆かつ有意（符号化ミスの疑い） |
| `price_insignificant` | 中 | 価格係数の p 値 ≥ 0.10（WTP の信頼性低下） |
| `wtp_extrapolation` | 中 | \|WTP\| > 価格レンジ × 2（外挿値） |

```python
result.warnings()                        # すべての警告
result.warnings(severity="大")           # 重大度「大」のみ
result.warnings(category="r2_low")      # カテゴリでフィルタ
```

## 依存パッケージ

| パッケージ | バージョン |
|-----------|-----------|
| pandas | ≥ 1.5 |
| numpy | ≥ 1.21 |
| statsmodels | ≥ 0.13 |
| matplotlib | ≥ 3.4 |
| openpyxl | ≥ 3.0 |

## ライセンス

MIT
