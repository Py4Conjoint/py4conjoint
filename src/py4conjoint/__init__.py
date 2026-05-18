"""
py4conjoint
===========
評点型コンジョイント分析を **Python初心者でも直感的に** 行えるパッケージ。

主な機能
--------
1. **データ作成**: Microsoft Forms / Google Forms の回答ファイルを
   long形式DataFrameに変換（:func:`forms_to_conjoint_data`）
2. **符号化**: ``-1/1`` 効果コーディングを自動化（:func:`encode`）
3. **回帰分析**: 回帰モデルを推定し、解釈オブジェクトを返す（:func:`fit`）
4. **解釈**: 相対重要度・WTP・市場シェアの計算（:class:`ConjointResult`）
5. **可視化**: 棒グラフによる結果の可視化
6. **落とし穴の自動検出**: データ品質や仮定の問題を警告

クイックスタート
----------------
.. code-block:: python

    import pandas as pd
    import py4conjoint as pc

    # ---- 1. アンケートのカード設計 ----
    cards = pd.DataFrame({
        "price":  [6, 10, 6, 10],
        "os":     ["android", "apple", "apple", "android"],
        "camera": ["標準", "標準", "高性能", "高性能"],
    }, index=["P1", "P2", "P3", "P4"])

    # ---- 2. Forms 回答ファイルを分析用データに変換 ----
    df = pc.forms_to_conjoint_data(
        responses_file="responses.xlsx",
        n_cards=4,
        attributes=cards,
    )

    # ---- 3. 符号化（基準水準を指定するだけ）----
    df_coded = pc.encode(
        df,
        reference_levels={
            "price":  10,         # 高い方を基準
            "os":     "android",
            "camera": "標準",
        },
    )

    # ---- 4. 回帰分析 ----
    result = pc.fit(df_coded)
    print(result.summary())

    # ---- 5. 解釈 ----
    result.importance()                         # 相対重要度（合計100%）
    result.wtp()                                # WTP
    result.market_share(products_df)            # 市場シェア予測
    result.plot_importance()                    # 重要度の棒グラフ
    result.plot_partworth()                     # 部分効用の棒グラフ
    result.plot_wtp(price_unit="万円")          # WTPの棒グラフ
"""
from __future__ import annotations

# 既存のデータ作成関数（後方互換）
from ._forms import forms_to_conjoint_data

# 新規追加：符号化
from .encoding import encode, auto_reference_levels

# 新規追加：回帰分析
from .analysis import fit, ConjointResult

# 新規追加：可視化
from .plot import plot_importance, plot_partworth, plot_wtp


__version__ = "0.2.0"

__all__ = [
    # データ作成
    "forms_to_conjoint_data",
    # 符号化
    "encode",
    "auto_reference_levels",
    # 分析
    "fit",
    "ConjointResult",
    # 可視化
    "plot_importance",
    "plot_partworth",
    "plot_wtp",
]
