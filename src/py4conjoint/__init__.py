"""
py4conjoint
===========
コンジョイント分析を **Python初心者でも直感的に** 行えるパッケージ。

v0.4.0 から、機能はサブパッケージとして提供されます：

- :mod:`py4conjoint.rating` — 評点型コンジョイント分析
- :mod:`py4conjoint.choice` — 選択型コンジョイント分析（CBC）

クイックスタート
----------------
.. code-block:: python

    import py4conjoint.rating as pcr

    df_coded = pcr.encode(df, reference_levels={...})
    result = pcr.fit(df_coded)
    print(result.summary())

旧バージョン（v0.3.x 以前）のトップレベルAPI（``pc.fit`` など）は
v0.4.0 で廃止されました。``import py4conjoint.rating as pcr`` を
使ってください。
"""
from __future__ import annotations

import importlib

__version__ = "0.4.0"

# v0.3.x までトップレベルに存在した旧API名
_REMOVED_API = frozenset({
    # データ作成
    "forms_to_conjoint_data",
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
})

_SUBPACKAGES = ("rating", "choice")


def __getattr__(name: str):
    if name in _SUBPACKAGES:
        return importlib.import_module(f".{name}", __name__)
    if name in _REMOVED_API:
        raise AttributeError(
            f"py4conjoint.{name} は v0.4.0 で廃止されました。"
            "`import py4conjoint.rating as pcr` を使ってください。"
        )
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
