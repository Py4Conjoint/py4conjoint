# py4conjoint

コンジョイント分析を **Python初心者でも直感的に** 行えるパッケージです。

Microsoft Forms / Google Forms のアンケート回答ファイルを読み込み、
符号化・推定・結果の解釈・可視化までを一貫して行えます。

2つの分析方式をサブパッケージとして提供します：

| サブパッケージ | 方式 | 回答者のタスク | 統計モデル |
|---------------|------|---------------|-----------|
| `py4conjoint.rating` | 評点型 | 製品案を1つずつ採点（例：7点満点） | OLS 回帰 |
| `py4conjoint.choice` | 選択型 | 複数の製品案から1つを選ぶ | 条件付きロジット |

どちらの方式も **同じ関数名**（`forms_to_data()`・`encode()`・`fit()` など）を持ち、
名前空間（`pcr` / `pcc`）で使い分けます。たとえば回答ファイルの読み込みは、評点型なら
`pcr.forms_to_data(...)`、選択型なら `pcc.forms_to_data(...)` です。

## インストール

```bash
pip install py4conjoint
```

Google Colab もしくは JupyterLite では：

```python
%pip install py4conjoint
```

### Excel ファイル（.xlsx）を読む場合

Microsoft Forms からダウンロードした `.xlsx` を読むには、追加のパッケージが必要です。

```bash
pip install py4conjoint[excel]
```

**`.csv` だけを使う場合、追加インストールは不要です。**
Microsoft Forms のファイルも `.csv` に保存し直せばそのまま読めます
（→ [Microsoft Forms のデータを使う](#microsoft-forms-のデータを使う)）。

高速な代替エンジンを使いたい場合は、次でも `.xlsx` を読めます（任意）。

```bash
pip install py4conjoint[excel-fast]
```

## 評点型コンジョイント分析（rating）

### クイックスタート

#### 1. アンケートデータを読み込む

```python
import pandas as pd
import py4conjoint.rating as pcr

# プロファイル設計を作成（この4プロファイルは レポジトリのexamples/ ある design_rating_2price.csv と同じ内容。
# このように手で書いてもよいし、pd.read_csv("design_rating_2price.csv") で
# 生成済みの CSV を読み込んで渡してもよい）
profiles = { # P1         P2        P3         P4
    "price":  [6,         6,        10,        10],
    "os":     ["android", "apple",  "android", "apple"],
    "camera": ["標準",    "高性能", "高性能",  "標準"],
}

# Microsoft Forms の回答ファイルを読み込む（デフォルト）
# .csv も渡せます（推奨。→「Microsoft Forms のデータを使う」節）
df = pcr.forms_to_data(
    responses_file  = "responses_rating_2price.csv",
    profiles        = profiles,
    respondent_cols = {"あなたの性別を教えてください。": "gender"},
)

# Google Forms の場合（ファイル名は例。forms="google" の書き方を示すもの）
df = pcr.forms_to_data(
    responses_file  = "responses.csv",
    profiles        = profiles,
    respondent_cols = {"性別": "gender"},
    forms           = "google",
)
```

> **データの入手先**：`responses_rating_2price.csv` などのコード例で使うデータは、
> リポジトリの `examples/` にあります（`pip install` には含まれません）。
> `python examples/make_demo_data.py` で再生成できます。

#### 2. 符号化する

```python
df_coded = pcr.encode(
    df,
    reference_levels = {
        "price":  10,
        "os":     "android",
        "camera": "標準"
    },
    suffix_map = {
        "price":  "low",
        "os":     "apple",
        "camera": "high"
    }
)

# 回答者属性も 0/1 にしたい場合は respondent_encode を使います
# （性別のような2水準の属性は respondent_encode={"gender": ["女性", "male"]} とかくことができる）
```

#### 3. 回帰分析を実行する

```python
result = pcr.fit(df_coded)
print(result)
```

> ※ 以下は examples/ のデモデータで実際に実行した出力です。
>  データや実装が変わると数字は変わります。

```
============================================================
コンジョイント分析の結果（和文サマリー）
============================================================
観測数        : 400
説明変数の数  : 3
決定係数 R²   : 0.6757
自由度修正 R² : 0.6732
標準誤差      : クラスタロバスト（respondent_id）

【推定された係数（部分効用 part-worth）】
  変数                            係数        p値  有意性
  ------------------------- ---------- ----------  ------
  Intercept                     5.3350     0.0000  ***
  price_low                     1.4450     0.0000  ***
  os_apple                      0.7250     0.0000  ***
  camera_high                   0.6150     0.0000  ***

  有意水準: *** p<0.001  ** p<0.01  * p<0.05  . p<0.1
============================================================
```


#### 4. 結果を解釈する

```python
# 重要度（合計100%）
result.importance()
# → 属性を行、「効用範囲」と「重要度」を列とする DataFrame
#   （重要度は合計が 100 になるように正規化した値）

# WTP（限界支払意思額）
result.wtp()
# → 符号化列名を行、「係数」と「限界支払意思額」を列とする DataFrame
#   （価格以外の属性について、価格何単位分に相当するか）

# 評点1点の金額換算
result.unit_rating_money()
# → 評点が1点上がることに相当する金額（float）

# 市場シェア予測
products = pd.DataFrame({
    "price_low":   [1, -1],
    "os_apple":    [1,  1],
    "camera_high": [1, -1],
}, index=["製品A", "製品B"])

result.market_share(products)
# → 製品ごとの予測シェア（合計が 1 になる Series）
```

#### 5. 可視化する

```python
result.plot_importance()            # 重要度の棒グラフ
result.plot_partworth()             # 部分効用の棒グラフ
result.plot_wtp(price_unit="万円")  # WTPの棒グラフ
```

### 主な機能

| 関数 / メソッド | 説明 |
|----------------|------|
| `forms_to_data()` | Microsoft/Google Forms の回答ファイルを long 形式 DataFrame に変換（`rating_range` で評点の値域を指定できる） |
| `design_profiles()` | D 最適計画法によるプロファイル設計（`auto_balance=True` で水準バランスを優先できる） |
| `check_design()` | プロファイル設計の直交性チェック |
| `suggest_n_profiles()` | 推奨プロファイル数の目安 |
| `encode()` | 属性列を効果コーディング（-1/+1）に自動変換 |
| `auto_reference_levels()` | 基準水準を自動推測（補助関数） |
| `fit()` | OLS 回帰を実行し `ConjointResult` を返す |
| `result.summary()` | 係数表・R²・落とし穴チェックの和文サマリー |
| `result.warnings()` | 落とし穴の一覧（severity / category でフィルタ可） |
| `result.importance()` | 各属性の重要度（合計100%） |
| `result.wtp()` | 各属性の WTP（限界支払意思額） |
| `result.unit_rating_money()` | 評点1点の金額換算（float） |
| `result.market_share()` | 市場シェア予測（logit / max） |
| `result.plot_importance()` | 重要度の棒グラフ |
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
result.warnings()                    # すべての警告
result.warnings(severity="大")       # 重大度「大」のみ
result.warnings(category="r2_low")   # カテゴリでフィルタ
```

## 選択型コンジョイント分析（choice / CBC）

### クイックスタート

#### 1. 選択セットを設計する

```python
import pandas as pd
import py4conjoint.choice as pcc


# プロファイル設計を作成
attributes = {
    "price":  [6, 10],                  # 万円
    "os":     ["android", "apple"],
    "camera": ["標準", "高性能"],
}

# 選択セット（設問）を生成：6設問 × 3代替案
design = pcc.design_choice_sets(
    attributes,
    n_sets=6,
    n_alts=3,
    seed=180,
)

# 設計の品質を診断（バランス・独立性・オーバーラップ）
print(pcc.check_design(design).summary())

# 必要回答者数の目安（Johnson-Orme の経験則 n ≥ 500c/(t×a)）
pcc.suggest_n_respondents(attributes, n_sets=6, n_alts=3)

# 作成した設計はすぐ保存し、分析時は作り直さず読み込んで使う（下の警告を参照）
design.to_csv("design_choice_2price.csv", index=False)
```

> ⚠️ **アンケート作成に使った design と、`forms_to_data()` に渡す design は
> 完全に同一にすること。** 属性名・水準・**水準の順序**・`seed`・`n_sets`・
> `n_alts` が1つでも違うと、同じ `seed` でも別の設計になり、回答と代替案の
> 対応が **エラーなく食い違って結果が誤ります**（例：本来 正 の係数が 負 に出る）。
> このような問題を防ぐため、design は作成後すぐ `design.to_csv(...)` で保存し、分析時は
> `pd.read_csv(...)` で**同じファイルを読み込んで**渡してください。2つの design が
> 同一かは `pcc.design_signature(design)` の署名で確認できます。

#### 2. アンケートデータを読み込む

```python
# 1設問 = 1選択セット。回答選択肢（例：「製品A」「製品B」「製品C」）が代替案
# .csv も渡せます（推奨。→「Microsoft Forms のデータを使う」節）
# 設計は作り直さず、step 1 で保存したファイルを読み込んで使う
design = pd.read_csv("design_choice_2price.csv")

df = pcc.forms_to_data(
    responses_file = "responses_choice_2price.csv",
    design         = design,
    choice_labels  = ["製品A", "製品B", "製品C"],   # 回答文字列に含まれるラベル（alt_id の順）
)

# Google Forms の場合は forms="google"、.csv ファイルを渡す
```

> `responses_choice_2price.csv`は rating の例と同じくリポジトリの `examples/` にあり、
> `python examples/make_demo_data.py` で再生成できます。

#### 3. 符号化する

```python
# ダミーコーディング（0/1）。基準水準の効用が 0 に固定される
df_coded = pcc.encode(df, reference_levels={"os": "android", "camera": "標準"})
# → os_apple, camera_高性能 列が追加される（価格などの数値変数はそのまま使う）
```

#### 4. 条件付きロジットを推定する

```python
result = pcc.fit(
    df_coded,
    choice_set_id_col = "choice_set_id",  # 選択セットID（回答者×設問）
    respondent_id_col = "respondent_id",  # クラスタロバスト標準誤差に使用
)
print(result)
```

> ※ 以下は examples/ のデモデータで実際に実行した出力です。
>  データや実装が変わると数字は変わります。

```
============================================================
選択型コンジョイント分析の結果（和文サマリー）
============================================================
観測数（行数）                       : 1800
選択セット数（回答者数 × 設問数/人） : 600（100 × 6/人）
説明変数の数                         : 3
対数尤度                             : -557.7745
擬似決定係数 R²（McFadden）          : 0.1538
標準誤差                             : クラスタロバスト（respondent_id）

【推定された係数（部分効用 part-worth）】
  変数                            係数        p値  有意性
  ------------------------- ---------- ----------  ------
  price                        -0.3261     0.0000  ***
  os_apple                      0.6533     0.0000  ***
  camera_高性能                 0.4605     0.0002  ***

  有意水準: *** p<0.001  ** p<0.01  * p<0.05  . p<0.1
============================================================
```

> 各回答者が同じ設問数に答える設計（バランス済み）では、「選択セット数」行に
> 内訳が付き `選択セット数（回答者数 × 設問数/人）: 600（100 × 6/人）` のように
> 表示されます。

#### 5. 結果を解釈する

```python
result.importance()      # 重要度（合計100%）
result.wtp()             # WTP（限界支払意思額 = -係数/価格係数）
result.warnings()        # 落とし穴チェック

# 市場シェア予測（ダミー列は 0/1、数値列は実際の値）
products = pd.DataFrame({
    "price":        [6, 10],
    "os_apple":     [1,  0],
    "camera_高性能": [1,  1],
}, index=["製品X", "製品Y"])

result.market_share(products)
```

#### 6. 可視化する

```python
result.plot_importance()   # 重要度の棒グラフ
result.plot_partworth()    # 部分効用の棒グラフ（基準水準＝0 も明示）
result.plot_wtp(price_unit="円")  # WTPの棒グラフ
```

### 主な機能

| 関数 / メソッド | 説明 |
|----------------|------|
| `design_choice_sets()` | CBC 用の選択セット生成（完全交差からのランダム割り当て） |
| `design_signature()` | 設計の署名（作成時と分析時の design が同一か確認する） |
| `check_design()` | 選択セット設計の診断（バランス・独立性・オーバーラップ） |
| `suggest_n_respondents()` | Johnson-Orme の経験則による必要回答者数の目安 |
| `forms_to_data()` | Microsoft/Google Forms の回答ファイルを long 形式 DataFrame に変換 |
| `encode()` | 属性列をダミーコーディング（0/1）に自動変換 |
| `fit()` | 条件付きロジットを推定し `ChoiceConjointResult` を返す |
| `result.summary()` | 係数表・対数尤度・擬似R²・落とし穴チェックの和文サマリー |
| `result.warnings()` | 落とし穴の一覧（severity / category でフィルタ可） |
| `result.importance()` | 各属性の重要度（合計100%） |
| `result.wtp()` | 各属性の WTP（限界支払意思額） |
| `result.market_share()` | 市場シェア予測（logit / max） |
| `result.plot_importance()` | 重要度の棒グラフ |
| `result.plot_partworth()` | 部分効用（パートワース）の棒グラフ |
| `result.plot_wtp()` | WTP の棒グラフ |

### 落とし穴の自動検出

`fit()` と `wtp()` は、以下の問題を自動的に検出して警告します：

| カテゴリ | 重大度 | 内容 |
|----------|:------:|------|
| `separation` | 大 | 完全分離の疑い（収束失敗・係数が異常に大きい） |
| `few_choice_sets` | 大/中 | 選択セット数／説明変数数の比率が低い（< 5 → 大、< 10 → 中） |
| `unbalanced_choices` | 中 | 特定の位置の代替案ばかり選ばれている（≥ 80%） |
| `few_respondents` | 大/中 | 回答者数が少ない（1人→大、2〜4人→中） |
| `independence_assumed` | 中 | 回答者ID列が無く、観測の独立性を仮定 |
| `price_sign_positive` | 中 | 価格係数が正かつ有意（データ品質の疑い） |
| `price_insignificant` | 中 | 価格係数の p 値 ≥ 0.10（WTP の信頼性低下） |
| `wtp_extrapolation` | 中 | \|WTP\| > 価格レンジ × 2（外挿値） |

## Microsoft Forms のデータを使う

Microsoft Forms の回答ファイルは `.xlsx` / `.csv` のどちらでも読めますが、
**`.csv` を推奨します**。理由は2つあります。

- Excel を読むための追加パッケージ（`py4conjoint[excel]`）が不要
- ブラウザ上で動く Jupyter（JupyterLite / Pyodide）でファイルが壊れることがない

手順は次のとおりです。

1. Microsoft Forms から回答ファイル（`.xlsx`）をダウンロードする
2. そのファイルを Excel で開く
3. 「名前を付けて保存」で **「CSV UTF-8（コンマ区切り）(*.csv)」** を選んで保存する
4. 保存した `.csv` を `forms_to_data()` に渡す

```python
# forms="microsoft" のまま .csv を渡せます（警告は出ません）
df = pcr.forms_to_data(
    responses_file = "responses.csv",
    profiles       = profiles,
)
```

`forms="google"` の引数を使う必要はありません。`forms` は「どの Forms で
集めたか」を指定する引数で、ファイル形式を指定するものではないためです。

## ブラウザ版 Jupyter（JupyterLite / Pyodide）で使う場合

py4conjoint は `.xlsx` を読むとき、python-calamine → openpyxl の順に
エンジンを試します。

- **openpyxl は Pyodide に同梱されていません。**
- 新しめの Pyodide には python-calamine が同梱されているため、
  追加インストールなしで `.xlsx` を読める場合があります。
- ただし `engine="calamine"` は pandas 2.2 で追加された機能のため、
  **pandas 2.2 より古い Pyodide では python-calamine を指定できません。
  この環境には openpyxl も無いため、`.xlsx` はどちらのエンジンでも
  読めません**（この場合は追加インストールを案内する日本語のエラーが
  出ます）。なお、ローカル環境で pandas 2.2 より古い場合は、openpyxl が
  入っていればそちらにフォールバックして読み込めます。
- また `.xlsx` は、ブラウザ環境でのファイル転送時に壊れることが
  あります。その場合も日本語のエラーで `.csv` への変換手順を案内します。

**`.csv` であれば確実です。** 追加パッケージも読み込みエンジンも不要で、
どの環境でも読み込めます。上の「Microsoft Forms のデータを使う」の手順で
`.csv` に変換しておくことをおすすめします。

## 依存パッケージ

### 必須

| パッケージ | バージョン |
|-----------|-----------|
| pandas | ≥ 1.5 |
| numpy | ≥ 1.21 |
| scipy | ≥ 1.8 |
| statsmodels | ≥ 0.13 |
| matplotlib | ≥ 3.4 |

### オプション（Excel ファイルを読む場合のみ）

`.csv` だけを使う場合は不要です。

| パッケージ | バージョン | extra 名 | 用途 |
|-----------|-----------|---------|------|
| openpyxl | ≥ 3.1.5 | `excel` | `.xlsx`（Excel ファイル）を読むために必要 |
| python-calamine | ≥ 0.6 | `excel-fast` | `.xlsx` を読むための代替エンジン |

`.xlsx` を読むときは python-calamine → openpyxl の順に試します。
**openpyxl が入っていれば確実に読めます。**
（python-calamine は Python 3.10 以上、および pandas 2.2 以上が
必要です。openpyxl はどの pandas でも使えます。）

## ライセンス

MIT
