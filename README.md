# py4conjoint

Google Forms の回答CSVを評点型コンジョイント分析用のlong形式DataFrameに変換するPythonパッケージです。

## インストール

```bash
pip install py4conjoint
```

Google Colab では：

```python
!pip install py4conjoint
```

## 使い方

```python
import pandas as pd
import py4conjoint as pc

# カード設計（プロファイル）を作成
cards = pd.DataFrame({    # P1         P2       P3        P4
                "price":   [6,         10,      6,        10],
                "os":      ["android", "apple", "apple",  "android"],
                "camera":  ["標準",    "標準",  "高性能", "高性能"]
}, index=["P1", "P2", "P3", "P4"])


# Microsoft Forms の回答xlsxをlong形式に変換（デフォルト）
df = pc.forms_to_conjoint_data(
    responses_csv = "responses.xlsx",
    n_cards       = 4,
    attributes    = cards,
    respondent_cols= {"性別": "gender"},
)

# Google Forms の回答CSVをlong形式に変換
df = pc.forms_to_conjoint_data(
    responses_csv = "responses.csv",
    n_cards       = 4,
    attributes    = cards,
    respondent_cols= {"性別": "gender"},
    forms         = "google",
)
```

## 出力形式

```
   回答者ID プロファイルID  rating gender  price       os    camera
0        1            P1       4     女性      6  android      標準
1        1            P2       3     女性     10    apple      標準 
...
```

## ライセンス

MIT
