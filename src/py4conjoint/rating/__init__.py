"""
py4conjoint.rating
==================
評点型コンジョイント分析のサブパッケージ。
"""
from __future__ import annotations

# バージョン（親パッケージと共通）
from .. import __version__

# データ作成
from ._forms import forms_to_data

# プロファイル設計
from .design import design_profiles, suggest_n_profiles

# 符号化
from .encoding import encode, auto_reference_levels

# 回帰分析
from .analysis import fit, ConjointResult, check_design, DesignCheckResult

# 可視化
from .plot import plot_importance, plot_partworth, plot_wtp

__all__ = [
    # データ作成
    "forms_to_data",
    # プロファイル設計
    "design_profiles",
    "suggest_n_profiles",
    # 符号化
    "encode",
    "auto_reference_levels",
    # 分析
    "fit",
    "ConjointResult",
    "check_design",
    "DesignCheckResult",
    # 可視化
    "plot_importance",
    "plot_partworth",
    "plot_wtp",
]
