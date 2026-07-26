"""
encoding.py（choice 版）
========================
選択型コンジョイント分析（CBC）用の **ダミーコーディング（0/1）** を
自動化するモジュール。

rating 版の :func:`py4conjoint.rating.encode`（効果コーディング -1/+1）とは
**別物** である点に注意。

* 条件付きロジット（conditional logit）では、属性ごとに **基準水準** を
  1つ決め、残りの水準に 0/1 のダミー列を作るのが標準的。
* 基準水準の効用は 0 に固定され、各ダミー係数は
  「基準水準と比べてどれだけ選ばれやすいか」を表す。
* 価格のような **数値（連続）属性** はダミー化せず、そのまま
  ``fit()`` の説明変数に入れる（係数は「1単位あたりの効用変化」）。

API の形式（``reference_levels`` 引数で基準水準を指定する）は
rating 版と揃えてある。
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

import pandas as pd

# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------


def encode(
    df: pd.DataFrame,
    reference_levels: Dict[str, object],
    *,
    suffix_map: Optional[Dict[str, Union[str, List[str]]]] = None,
    drop_original: bool = False,
    inplace: bool = False,
) -> pd.DataFrame:
    """
    カテゴリ属性列を **0/1 のダミーコーディング** に自動変換する。

    各属性について **基準水準** を指定すると、基準水準 **以外** の水準
    それぞれに ``0/1`` のダミー列が1つずつ作られる（K水準なら K-1 列）。
    基準水準の行はすべてのダミー列で 0 になる。

    Parameters
    ----------
    df : pd.DataFrame
        long形式の選択データ。1行が「1つの選択セット内の1つの代替案」を
        表すことを前提とする。

    reference_levels : dict
        ``{"属性名": 基準水準}`` の辞書。
        基準水準の効用が 0 に固定され、各ダミー係数は基準水準との差を表す。

        例::

            reference_levels = {"brand": "dannon"}

    suffix_map : dict, optional
        生成する列名のサフィックスを手動指定する辞書。
        非基準水準の数と同じ長さの ``List[str]``（非基準水準が1つなら
        ``str`` でも可）を渡す。順序はデータ上の水準の出現順に合わせること。

        省略時は ``{属性名}_{水準名}`` の形式になる
        （例：``brand_hiland``, ``brand_yoplait``）。

    drop_original : bool, default False
        ``True`` にすると元の属性列（``brand`` など）を削除する。
        授業では元の列も確認したいことが多いため、デフォルトは残す。

    inplace : bool, default False
        ``True`` にすると入力 ``df`` を直接書き換える。
        デフォルトはコピーを返す。

    Returns
    -------
    pd.DataFrame
        ダミー列が追加されたDataFrame。
        例：``brand`` ∈ {dannon, hiland, yoplait}, 基準=dannon
        → ``brand_hiland``, ``brand_yoplait`` 列が追加される
        （dannon の行は両方とも 0）。

    Raises
    ------
    TypeError
        ``df`` が ``pd.DataFrame`` でない場合。
    ValueError
        ``reference_levels`` が空または辞書でない場合。
        指定した属性名が ``df`` にない場合。
        指定した基準水準が実際のデータに存在しない場合。
        属性が1水準しかない場合（分析不能）。
        ``suffix_map`` のリスト長が非基準水準数と一致しない場合。

    Notes
    -----
    rating 版の :func:`py4conjoint.rating.encode` は ``-1/+1`` の
    効果コーディングを使うが、choice 版は ``0/1`` のダミーコーディングを
    使う。条件付きロジットでは選択セット内で共通の定数（切片）が
    消えるため、ダミーコーディングの方が係数の解釈が直感的になる
    （「基準水準と比べた選ばれやすさ」）。

    Examples
    --------
    >>> import pandas as pd
    >>> import py4conjoint.choice as pcc
    >>> df = pd.DataFrame({
    ...     "choice_set_id": [1, 1, 2, 2],
    ...     "choice": [1, 0, 0, 1],
    ...     "price": [100, 150, 150, 100],
    ...     "brand": ["A社", "B社", "A社", "B社"],
    ... })
    >>> df_coded = pcc.encode(df, reference_levels={"brand": "A社"})
    >>> df_coded["brand_B社"].tolist()
    [0, 1, 0, 1]
    """
    # ---------- 入力チェック ----------
    if not isinstance(df, pd.DataFrame):
        raise TypeError(
            "df は pandas.DataFrame である必要があります。\n"
            f"  受け取った型: {type(df).__name__}"
        )
    if not isinstance(reference_levels, dict) or len(reference_levels) == 0:
        raise ValueError(
            "reference_levels は空でない辞書を指定してください。\n"
            "  例: {'brand': 'dannon'}"
        )

    suffix_map = suffix_map or {}
    out = df if inplace else df.copy()

    encoded_map: Dict[str, List[str]] = {}
    for attr, ref_level in reference_levels.items():
        if attr not in out.columns:
            raise ValueError(
                f"列 '{attr}' が DataFrame にありません。\n"
                f"  存在する列: {list(out.columns)}"
            )

        levels = _unique_levels(out[attr])
        if ref_level not in levels:
            raise ValueError(
                f"属性 '{attr}' に基準水準 '{ref_level}' が見つかりません。\n"
                f"  存在する水準: {levels}\n"
                "  reference_levels の値を確認してください。"
            )
        if len(levels) < 2:
            raise ValueError(
                f"属性 '{attr}' は水準が1つしかありません（{levels}）。\n"
                "  分析するには最低でも2水準が必要です。"
            )

        others = [lv for lv in levels if lv != ref_level]

        # suffix の決定（省略時は水準名そのもの）
        raw_suffix = suffix_map.get(attr)
        if raw_suffix is None:
            suffix_list = [str(lv) for lv in others]
        elif isinstance(raw_suffix, str):
            suffix_list = [raw_suffix]
        else:
            suffix_list = [str(s) for s in raw_suffix]
        if len(suffix_list) != len(others):
            raise ValueError(
                f"属性 '{attr}' の非基準水準は {len(others)} 個ですが、\n"
                f"suffix_map に {len(suffix_list)} 個のサフィックスが指定されています。\n"
                f"  非基準水準（基準='{ref_level}' 以外）: {others}\n"
                f"  指定されたサフィックス: {suffix_list}\n"
                "  suffix_map の値を非基準水準と同じ順・同じ数のリストにしてください。"
            )

        new_cols: List[str] = []
        for target, suffix in zip(others, suffix_list):
            new_col = f"{attr}_{suffix}"
            out[new_col] = (out[attr] == target).astype(int)
            # 欠損は 0 にしない（後段の fit() で気づけるよう NaN を残す）
            if out[attr].isna().any():
                out.loc[out[attr].isna(), new_col] = pd.NA
            new_cols.append(new_col)
        encoded_map[attr] = new_cols

    if drop_original:
        out = out.drop(columns=list(reference_levels.keys()))

    # 後で fit() が再利用できるよう、メタ情報を attrs に保存
    # （rating 版と同じ "py4conjoint" キーを使い、encoding 種別で区別する）
    existing_meta = {
        k: (dict(v) if isinstance(v, dict) else v)
        for k, v in out.attrs.get("py4conjoint", {}).items()
    }
    existing_meta.setdefault("reference_levels", {}).update(reference_levels)
    existing_meta["encoding"] = "dummy"
    existing_meta.setdefault("encoded_columns", {}).update(encoded_map)
    out.attrs["py4conjoint"] = existing_meta

    return out


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _unique_levels(s: pd.Series) -> List[Any]:
    """
    Series から欠損を除いたユニーク水準のリストを返す。
    数値は数値のまま返し、出現順を保つ（rating 版と同じ規約）。
    """
    return list(pd.Series(s.dropna().unique()))
