"""
encoding.py
============
コンジョイント分析用の **符号化（エンコーディング）** を自動化するモジュール。

授業のノートブックで行っている次のような手作業：

>>> df["price_low"]   = df["price"].map({10: -1,  6: 1})
>>> df["os_apple"]    = df["os"].map({"android": -1, "apple": 1})
>>> df["camera_high"] = df["camera"].map({"標準": -1, "高性能": 1})

を、**基準水準（reference level）を指定するだけ**で自動化する。

設計思想
--------
* ``-1 / 1`` の効果コーディング（effect coding）を採用。
  授業の資料に合わせる。
* **基準水準＝ -1 になる水準** と明示的に定義する。
  「評点が低くなりそうな水準」を基準にすると係数の符号が直感的になる。
* 3水準以上の属性については ``-1, 0, 1`` 形式の効果コーディング
  （``K-1`` 個のダミー変数を生成）に自動拡張する。
* 生成される列名はデフォルトで ``{属性名}_{インデックス}`` の形式（例：``price_0``, ``camera_0``）。
  ``encode()`` の ``suffix_map`` 引数で任意の名前に変更できる。元の列はそのまま残す。
"""

from __future__ import annotations

import warnings
from typing import Any, Dict, List, Optional, Union

import pandas as pd


# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------


def encode(
    df: pd.DataFrame,
    reference_levels: Dict[str, object],
    *,
    respondent_encode: Optional[Dict[str, object]] = None,
    suffix_map: Optional[Dict[str, Union[str, List[str]]]] = None,
    binary_suffix_map: Optional[Dict[str, str]] = None,
    drop_original: bool = False,
    inplace: bool = False,
) -> pd.DataFrame:
    """
    属性列を ``-1/1``（または効果コーディング）に自動変換する。

    各属性について **基準水準** を指定すると、その水準が ``-1`` になり、
    残りの水準が ``+1`` になる（2水準の場合）。
    3水準以上の場合は ``K-1`` 個のダミー列を自動生成する。

    Parameters
    ----------
    df : pd.DataFrame
        ``forms_to_data`` の出力など、long形式のデータ。
        各プロファイルの属性が列として入っていることを前提とする。

    reference_levels : dict
        ``{"属性名": 基準水準}`` の辞書。
        基準水準は ``-1`` にコード化される（＝評点が低くなると思われる水準を選ぶ）。

        例：
        ::

            reference_levels = {
                "price":  10,         # 高い方を基準（評点が低くなりそう）
                "os":     "android",
                "camera": "標準",
            }

    respondent_encode : dict, optional
        回答者属性を ``0/1`` の2値にコード化したい場合に指定する辞書。

        値に **文字列** を渡すと ``{列名}_0`` という列名になる。
        値に **[0にしたい水準, 列名サフィックス]** のリストを渡すと
        ``{列名}_{サフィックス}`` という列名になる。

        * ``0`` になる水準を明示的に選べる（例：女性→0、男性→1）。
        * 効果コーディング（-1/1）ではなく、単純な 0/1 の2値変換。

        例：
        ::

            respondent_encode = {"gender": "女性"}
            # → gender_0 列が追加される（女性→0, 男性→1）

            respondent_encode = {"gender": ["女性", "female"]}
            # → gender_female 列が追加される（女性→0, 男性→1）

    suffix_map : dict, optional
        生成する列名のサフィックスを手動指定する辞書。
        2水準の属性には ``str``、3水準以上の属性には ``List[str]``（非基準水準の数と
        同じ長さ）を渡す。

        例::

            suffix_map={
                "price": "low",           # 2水準 → price_low
                "color": ["red", "blue"], # 3水準（基準=白、非基準=赤・青）→ color_red, color_blue
            }

        .. note::
            3水準以上の場合、``List[str]`` の順序はデータ上の水準の出現順
            （``_unique_levels()`` が返す順序）と一致させてください。
            ``pd.Series.unique()`` は最初に出現した順を返します。

        省略時は ``{属性名}_0``, ``{属性名}_1``, ... の形式になります。

    binary_suffix_map : dict, optional
        **非推奨**。``suffix_map`` を使ってください。
        後方互換のために残してあります。指定した場合は ``DeprecationWarning`` が出ます。

    drop_original : bool, default False
        ``True`` にすると元の属性列（``price`` など）を削除する。
        授業では元の列も確認したいことが多いため、デフォルトは残す。

    inplace : bool, default False
        ``True`` にすると入力 ``df`` を直接書き換える。
        デフォルトはコピーを返す。

    Returns
    -------
    pd.DataFrame
        符号化された列が追加されたDataFrame。

        2水準の場合（``suffix_map`` 未指定時）：``{属性名}_0`` という列が1つ増える。
            例：``price`` ∈ {6, 10}, 基準=10 → ``price_0`` 列（6→1, 10→-1）

        3水準以上の場合（``suffix_map`` 未指定時）：``{属性名}_{0,1,...}`` という列が ``K-1`` 個増える。
            例：``color`` ∈ {赤, 青, 緑}, 基準=赤 → ``color_0``（青）, ``color_1``（緑）
            （青なら[1,0]、緑なら[0,1]、赤なら[-1,-1]）

    Raises
    ------
    TypeError
        ``df`` が ``pd.DataFrame`` でない場合。
    ValueError
        ``reference_levels`` が空または辞書でない場合。
        ``reference_levels`` で指定した属性名が ``df`` にない場合。
        指定した基準水準が実際のデータに存在しない場合。
        属性が1水準しかない場合（分析不能）。
        ``suffix_map`` のリスト長が非基準水準数と一致しない場合。

    Notes
    -----
    効果コーディング（-1/1）はダミーコーディング（0/1）と異なり、
    切片 ``b_0`` が **全水準の平均効用** を表すという性質を持つ。
    部分効用（パートワース）の解釈がしやすくなるため、
    コンジョイント分析では効果コーディングが標準的に使われる。

    Examples
    --------
    >>> import pandas as pd
    >>> import py4conjoint as pc
    >>> df = pd.DataFrame({
    ...     "rating": [5, 3, 7, 4],
    ...     "price":  [6, 10, 6, 10],
    ...     "os":     ["android", "apple", "apple", "android"],
    ...     "camera": ["標準", "標準", "高性能", "高性能"],
    ... })
    >>> df_coded = pc.encode(
    ...     df,
    ...     reference_levels={"price": 10, "os": "android", "camera": "標準"},
    ... )
    >>> df_coded.columns.tolist()
    ['rating', 'price', 'os', 'camera', 'price_0', 'os_0', 'camera_0']

    3水準以上の属性にサフィックスを指定する例::

        df_coded = pc.encode(
            df,
            reference_levels={"price": 10, "color": "白"},
            suffix_map={
                "price": "low",          # 2水準 → price_low
                "color": ["red", "blue"] # 3水準（基準=白、非基準=赤・青）→ color_red, color_blue
            },
        )
        # → 生成される列: price_low, color_red, color_blue
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
            "  例: {'price': 10, 'os': 'android', 'camera': '標準'}"
        )

    # binary_suffix_map が渡されたら suffix_map にマージし、非推奨警告を出す
    if binary_suffix_map is not None:
        warnings.warn(
            "binary_suffix_map は非推奨です。suffix_map を使ってください。\n"
            "例: suffix_map={'price': 'low', 'os': 'apple'}",
            DeprecationWarning,
            stacklevel=2,
        )
        merged = dict(suffix_map) if suffix_map else {}
        merged.update(binary_suffix_map)
        suffix_map = merged

    suffix_map = suffix_map or {}

    out = df if inplace else df.copy()

    for attr, ref_level in reference_levels.items():
        # 列の存在確認
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

        # 2水準 と 3水準以上 で分岐（結果はすでに out に書き込み済み）
        raw_suffix = suffix_map.get(attr)
        if len(levels) == 2:
            # 2水準属性に List[str] が渡された場合は1要素のみ許容して str に変換
            if isinstance(raw_suffix, list):
                if len(raw_suffix) != 1:
                    raise ValueError(
                        f"2水準の属性 '{attr}' の suffix_map には str か1要素のリストを"
                        f"指定してください。{len(raw_suffix)} 要素のリストが渡されました。"
                    )
                raw_suffix = raw_suffix[0]
            _encode_binary(out, attr, ref_level, raw_suffix)
        else:
            _encode_multi(out, attr, ref_level, raw_suffix)

    # ---------- 回答者属性の 0/1 コーディング ----------
    if respondent_encode:
        for attr, spec in respondent_encode.items():
            if attr not in out.columns:
                raise ValueError(
                    f"列 '{attr}' が DataFrame にありません。\n"
                    f"  存在する列: {list(out.columns)}"
                )
            # spec は str（zero_levelのみ）または [zero_level, suffix] のリスト
            if isinstance(spec, list):
                if len(spec) != 2:
                    raise ValueError(
                        f"respondent_encode の '{attr}' の値をリストで指定する場合は\n"
                        f"  [0にしたい水準, 列名サフィックス] の2要素にしてください。\n"
                        f"  例: ['女性', 'female']"
                    )
                zero_level, suffix = spec[0], str(spec[1])
            else:
                zero_level, suffix = spec, "0"

            levels = _unique_levels(out[attr])
            if zero_level not in levels:
                raise ValueError(
                    f"属性 '{attr}' に水準 '{zero_level}' が見つかりません。\n"
                    f"  存在する水準: {levels}"
                )
            if len(levels) != 2:
                raise ValueError(
                    f"respondent_encode の属性 '{attr}' は2水準である必要があります。\n"
                    f"  現在の水準数: {len(levels)}（{levels}）"
                )
            one_level = [lv for lv in levels if lv != zero_level][0]
            new_col = f"{attr}_{suffix}"
            out[new_col] = out[attr].map({zero_level: 0, one_level: 1})

    if drop_original:
        drop_cols = list(reference_levels.keys())
        if respondent_encode:
            drop_cols += list(respondent_encode.keys())
        out = out.drop(columns=drop_cols)

    # 後で fit() が再利用できるよう、メタ情報を attrs に保存
    # （df.attrs は pandas のユーザー定義メタデータ機構）
    # ネストした辞書もコピーし、入力 df の attrs を書き換えないようにする
    existing_meta = {
        k: (dict(v) if isinstance(v, dict) else v)
        for k, v in out.attrs.get("py4conjoint", {}).items()
    }
    existing_meta.setdefault("reference_levels", {}).update(reference_levels)
    if suffix_map:
        existing_meta.setdefault("suffix_map", {}).update(suffix_map)
    out.attrs["py4conjoint"] = existing_meta

    return out


def auto_reference_levels(
    df: pd.DataFrame,
    attribute_columns: List[str],
    *,
    price_columns: Optional[List[str]] = None,
) -> Dict[str, object]:
    """
    各属性について基準水準を自動的に推測する補助関数。

    判定ルール（**推論**ベース。最終的にはユーザーが確認すべき）：

    * 数値列かつ ``price_columns`` に含まれる（価格相当列）：**最大値** を基準（高い方を ``-1``）
    * 数値列かつ ``price_columns`` に含まれない：**最小値** を基準（低い方を ``-1``）
    * カテゴリ列：水準を文字列でソートして **先頭** を基準

    自動推測の結果は警告として表示されるので、必ず内容を確認してから使うこと。

    Parameters
    ----------
    df : pd.DataFrame
    attribute_columns : list of str
        基準水準を推測したい属性名のリスト。
    price_columns : list of str, optional
        「価格相当」の列名。これらは数値の **大きい方** を基準にする
        （高価格＝評点が低くなりそう、という前提）。
        省略時は ``["price"]`` を使う。

    Returns
    -------
    dict
        ``{属性名: 推測された基準水準}`` の辞書。
        そのまま ``encode()`` の ``reference_levels`` に渡せる。

    Raises
    ------
    ValueError
        ``attribute_columns`` に含まれる列名が ``df`` に存在しない場合。

    Notes
    -----
    この関数はあくまで **推論** に基づくショートカットであり、
    属性の意味を理解した上での明示的な指定が望ましい。
    """
    price_cols = set(price_columns or ["price"])
    refs: Dict[str, object] = {}
    notes: List[str] = []

    for col in attribute_columns:
        if col not in df.columns:
            raise ValueError(f"列 '{col}' が DataFrame にありません。")

        levels = _unique_levels(df[col])
        if pd.api.types.is_numeric_dtype(df[col]):
            ref = max(levels) if col in price_cols else min(levels)
            reason = (
                "数値・価格列なので最大値"
                if col in price_cols
                else "数値列なので最小値"
            )
        else:
            # 文字列としてソートして先頭
            ref = sorted(levels, key=str)[0]
            reason = "カテゴリ列なので辞書順で先頭"
        refs[col] = ref
        notes.append(f"  {col}: {ref!r} （{reason}）")

    warnings.warn(
        "基準水準を自動で推測しました（推論ベース）。\n"
        "結果を確認し、必要なら明示的に reference_levels を指定してください。\n"
        + "\n".join(notes),
        UserWarning,
        stacklevel=2,
    )
    return refs


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _unique_levels(s: pd.Series) -> List[Any]:
    """
    Series から欠損を除いたユニーク水準のリストを返す。
    数値は数値のまま返し、出現順を保つ。
    """
    return list(pd.Series(s.dropna().unique()))


def _encode_binary(
    df: pd.DataFrame, attr: str, ref_level: object, manual_suffix: Optional[str]
) -> str:
    """
    2水準の属性を ``-1/1`` に変換する。
    新しい列を ``df`` に追加し、列名を返す。

    列名規則
    --------
    * ``manual_suffix`` が指定されていれば ``{attr}_{manual_suffix}``
    * そうでなければ ``{attr}_0``
    """
    levels = _unique_levels(df[attr])
    other = [lv for lv in levels if lv != ref_level][0]

    if manual_suffix is not None:
        new_col = f"{attr}_{manual_suffix}"
    else:
        new_col = f"{attr}_0"

    df[new_col] = df[attr].map({ref_level: -1, other: 1})

    # マッピング後にNaNがある = データに想定外の値が混入している
    if df[new_col].isna().any() and not df[attr].isna().any():
        bad = df.loc[df[new_col].isna(), attr].unique().tolist()
        raise ValueError(
            f"属性 '{attr}' に想定外の値が含まれています: {bad}\n"
            f"  想定されていた水準: {levels}"
        )

    return new_col


def _encode_multi(
    df: pd.DataFrame,
    attr: str,
    ref_level: object,
    suffixes: Optional[Union[str, List[str]]] = None,
) -> List[str]:
    """
    3水準以上の属性を効果コーディングで K-1 列に展開する。

    例：``color`` ∈ {赤, 青, 緑}, 基準=赤
        → ``color_0``（青）: 青→1, 赤→-1, 緑→0
        → ``color_1``（緑）: 緑→1, 赤→-1, 青→0

    ``others`` の順序は ``_unique_levels()`` が返す出現順（``pd.Series.unique()`` ベース）
    に依存する。``suffixes`` を指定する際はデータ上の水準の出現順を確認してから指定すること。
    """
    levels = _unique_levels(df[attr])
    others = [lv for lv in levels if lv != ref_level]

    # suffixes の型を List[str] に統一
    if suffixes is None:
        suffix_list = None
    elif isinstance(suffixes, str):
        suffix_list = [suffixes]
    else:
        suffix_list = list(suffixes)

    # 長さの検証
    if suffix_list is not None and len(suffix_list) != len(others):
        raise ValueError(
            f"属性 '{attr}' の非基準水準は {len(others)} 個ですが、\n"
            f"suffix_map に {len(suffix_list)} 個のサフィックスが指定されています。\n"
            f"  非基準水準（基準='{ref_level}' 以外）: {others}\n"
            f"  指定されたサフィックス: {suffix_list}\n"
            "  suffix_map の値を非基準水準と同じ順・同じ数のリストにしてください。"
        )

    new_cols: List[str] = []
    for i, target in enumerate(others):
        suffix = suffix_list[i] if suffix_list else str(i)
        new_col = f"{attr}_{suffix}"

        def _map(v, t=target, r=ref_level):
            if v == t:
                return 1
            if v == r:
                return -1
            return 0

        # na_action="ignore" で欠損を欠損のまま残す。
        # （指定しないと NaN に _map が適用されて 0 になり、欠損が
        #   「全水準の平均」を意味する値として回帰に静かに混入してしまう。
        #   2水準の _encode_binary は map(dict) なので元々 NaN が保持される。）
        df[new_col] = df[attr].map(_map, na_action="ignore")
        new_cols.append(new_col)

    return new_cols
