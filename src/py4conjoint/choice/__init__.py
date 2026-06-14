"""
py4conjoint.choice
==================
選択型コンジョイント分析（CBC：Choice-Based Conjoint）のサブパッケージ。

回答者に「複数の製品案からどれか1つを選んでもらう」形式のデータを、
条件付きロジット（conditional logit）モデルで分析する。
rating（評点型）と対称的なAPIを提供する。

クイックスタート
----------------
.. code-block:: python

    import pandas as pd
    import py4conjoint.choice as pcc

    # df は long形式：1行 = 1つの選択セット内の1つの代替案
    # 列の例：選択セットID, choice(0/1), price, brand

    # ---- 1. ダミーコーディング（基準水準を指定するだけ）----
    df_coded = pcc.encode(df, reference_levels={"brand": "dannon"})

    # ---- 2. 条件付きロジットの推定 ----
    result = pcc.fit(
        df_coded,
        choice="choice",
        choice_set_id_col="選択セットID",
        encoded_columns=["price", "brand_hiland", "brand_yoplait"],
    )
    print(result.summary())

    # ---- 3. 解釈 ----
    result.importance()                  # 重要度（合計100%）
    result.wtp()                         # WTP（限界支払意思額）
    result.market_share(products_df)     # 市場シェア予測
    result.warnings()                    # 落とし穴チェック
"""
from __future__ import annotations

# バージョン（親パッケージと共通）
from .. import __version__

# Forms 回答ファイルの読み込み
from ._forms import cbc_forms_to_data

# 条件付きロジットの推定
from .analysis import ChoiceConjointResult, fit

# 選択セット設計
from .design import (
    ChoiceDesignCheckResult,
    check_design,
    design_choice_sets,
    suggest_n_respondents,
)

# ダミーコーディング（0/1）
from .encoding import encode

# 可視化
from .plot import plot_importance, plot_partworth, plot_wtp

__all__ = [
    # Forms 読み込み
    "cbc_forms_to_data",
    # 設計
    "design_choice_sets",
    "check_design",
    "ChoiceDesignCheckResult",
    "suggest_n_respondents",
    # 符号化
    "encode",
    # 分析
    "fit",
    "ChoiceConjointResult",
    # 可視化
    "plot_importance",
    "plot_partworth",
    "plot_wtp",
]
