"""
py4conjoint
===========
Microsoft Forms / Google Forms の回答CSVを
評点型コンジョイント分析用のlong形式DataFrameに変換する。

インストール:
    pip install py4conjoint

使い方A（cardsのDataFrameをそのまま渡す・推奨）:
    import py4conjoint as pc

    cards = pd.DataFrame([
        {"price": 6,  "os": "android", "camera": "standard"},  # P1
        {"price": 10, "os": "apple",   "camera": "standard"},  # P2
        {"price": 6,  "os": "apple",   "camera": "high"},      # P3
        {"price": 10, "os": "android", "camera": "high"},      # P4
    ], index=["P1", "P2", "P3", "P4"])

    # Microsoft Forms（デフォルト）
    df = pc.forms_to_conjoint_data(
        responses_csv = "responses.xlsx",
        n_cards       = 4,
        attributes    = cards,
    )

    # Google Forms
    df = pc.forms_to_conjoint_data(
        responses_csv = "responses.csv",
        n_cards       = 4,
        attributes    = cards,
        forms         = "google",
    )

使い方B（辞書のリストで渡す・従来形式）:
    import py4conjoint as pc

    attributes = [
        {"price":   [6, 10, 6, 10]},
        {"os":      ["android", "apple", "apple", "android"]},
        {"camera":  ["standard", "standard", "high", "high"]},
    ]

    df = pc.forms_to_conjoint_data(
        responses_csv = "responses.xlsx",
        n_cards       = 4,
        attributes    = attributes,
    )
"""

from __future__ import annotations

import re
import warnings
from pathlib import Path
from typing import Dict, List, Literal, Optional, Sequence

import pandas as pd


# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------

def forms_to_conjoint_data(
    responses_csv: str,
    n_cards: int,
    attributes: "pd.DataFrame | Sequence[Dict[str, Sequence]]",
    *,
    forms: Literal["microsoft", "google"] = "microsoft",
    respondent_cols: Optional[Dict[str, str]] = None,
    card_id_prefix: str = "P",
    rating_colname: str = "rating",
    respondent_id_colname: str = "回答者ID",
    card_id_colname: str = "カードID",
    out_csv: Optional[str] = None,
) -> pd.DataFrame:
    """
    Microsoft Forms / Google Forms の回答ファイルをlong形式DataFrameに変換する。

    Parameters
    ----------
    responses_csv : str
        Forms からダウンロードした回答ファイルのパス。
        Microsoft Forms の場合は .xlsx、Google Forms の場合は .csv。

    n_cards : int
        アンケートで提示したカード（プロファイル）の枚数。
        例：4

    attributes : pd.DataFrame または list of dict
        カード設計を指定する。以下の2形式を受け付ける。

        【形式A：DataFrameをそのまま渡す（推奨）】
            授業で作成した cards をそのまま渡すことができる。
            行がカード、列が属性に対応する。
            インデックスは ["P1","P2",...] でも整数でも可。

            例：
            cards = pd.DataFrame([
                {"price": 6,  "os": "android", "camera": "standard"},
                {"price": 10, "os": "apple",   "camera": "standard"},
                {"price": 6,  "os": "apple",   "camera": "high"},
                {"price": 10, "os": "android", "camera": "high"},
            ], index=["P1", "P2", "P3", "P4"])

            df = pc.forms_to_conjoint_data(..., attributes=cards)

        【形式B：辞書のリスト（従来形式）】
            各属性を1辞書で指定する。辞書のキーが属性名、値がカード順の水準リスト。

            例：
            [
                {"price":   [6, 10, 6, 10]},
                {"os":      ["android", "apple", "apple", "android"]},
                {"camera":  ["standard", "standard", "high", "high"]},
            ]

        いずれの形式でも、行数（または水準リストの長さ）は n_cards と一致する必要がある。

    forms : {"microsoft", "google"}, default "microsoft"
        使用するFormsの種類を指定する。
        "microsoft" : Microsoft Forms（.xlsx形式）
        "google"    : Google Forms（.csv形式）

    respondent_cols : dict, optional
        回答者属性として残したい列の対応辞書。
        {"CSVの列名": "出力DataFrameの列名"} の形式。
        例：{"性別": "gender", "学年": "year"}
        省略した場合は回答者属性を付与しない。

    card_id_prefix : str, default "P"
        プロファイルIDの接頭辞。"P" なら P1, P2, P3, P4 となる。

    rating_colname : str, default "rating"
        出力DataFrameの評点列名。

    respondent_id_colname : str, default "回答者ID"
        出力DataFrameの回答者ID列名。

    card_id_colname : str, default "カードID"
        出力DataFrameの列名。

    out_csv : str, optional
        変換後のDataFrameをCSVとして保存するパス。
        省略した場合は保存しない。

    Returns
    -------
    pd.DataFrame
        long形式のDataFrame。
        列：回答者ID, カードID, rating, [回答者属性], [カード属性]

    Raises
    ------
    FileNotFoundError
        responses_csv が存在しない場合。
    ValueError
        forms が "microsoft" または "google" 以外の場合。
        attributes の行数（または水準リストの長さ）が n_cards と一致しない場合。
        評点列が n_cards 列分見つからない場合。
    """

    # ------------------------------------------------------------------
    # 0. 入力チェック
    # ------------------------------------------------------------------
    if forms not in ("microsoft", "google"):
        raise ValueError(
            f"forms='{forms}' は無効な値です。\n"
            "'microsoft' または 'google' を指定してください。"
        )

    attributes = _normalize_attributes(attributes, n_cards)
    _check_attributes(attributes, n_cards)

    csv_path = Path(responses_csv)
    if not csv_path.exists():
        raise FileNotFoundError(
            f"ファイルが見つかりません: {responses_csv}\n"
            "ファイル名とパスを確認してください。"
        )

    # forms="microsoft" なのに .xlsx/.xls 以外の拡張子の場合は警告を出す
    if forms == "microsoft" and csv_path.suffix.lower() not in (".xlsx", ".xls"):
        warnings.warn(
            f"forms='microsoft' が指定されていますが、\n"
            f"ファイルの拡張子が '{csv_path.suffix}' です。\n"
            "Microsoft Forms のダウンロードファイルは通常 .xlsx 形式です。\n"
            "Google Forms のファイルを使う場合は forms='google' を指定してください。",
            UserWarning,
            stacklevel=2,
        )

    # ------------------------------------------------------------------
    # 1. ファイル読み込み
    #    Microsoft Forms → .xlsx（openpyxl）
    #    Google Forms   → .csv（UTF-8 / BOM付きUTF-8）
    # ------------------------------------------------------------------
    if forms == "microsoft":
        raw = _read_microsoft_forms(csv_path)
    else:
        raw = _read_google_forms(csv_path)

    # ------------------------------------------------------------------
    # 2. 管理列を除外して評点列・回答者属性列を特定する
    # ------------------------------------------------------------------
    if forms == "microsoft":
        system_cols = _detect_microsoft_system_cols(raw)
    else:
        system_cols = _detect_google_system_cols(raw)

    respondent_rename: Dict[str, str] = respondent_cols or {}
    respondent_src_cols = list(respondent_rename.keys())

    non_rating_cols = set(system_cols) | set(respondent_src_cols)
    rating_candidate_cols = [c for c in raw.columns if c not in non_rating_cols]

    rating_cols = _pick_rating_cols(rating_candidate_cols, raw, n_cards, responses_csv)

    # ------------------------------------------------------------------
    # 3. 回答者IDを付与
    # ------------------------------------------------------------------
    raw[respondent_id_colname] = range(1, len(raw) + 1)

    # ------------------------------------------------------------------
    # 4. 回答者属性列を選択・リネーム
    # ------------------------------------------------------------------
    keep_cols = [respondent_id_colname] + respondent_src_cols + rating_cols
    df_wide = raw[keep_cols].copy()

    if respondent_rename:
        df_wide = df_wide.rename(columns=respondent_rename)
        respondent_dst_cols = list(respondent_rename.values())
    else:
        respondent_dst_cols = []

    # 評点列をプロファイルID（文字列）にリネームして wide→long 変換しやすくする
    card_ids = [f"{card_id_prefix}{i+1}" for i in range(n_cards)]
    rating_rename = dict(zip(rating_cols, card_ids))
    df_wide = df_wide.rename(columns=rating_rename)

    # ------------------------------------------------------------------
    # 5. wide → long 変換
    # ------------------------------------------------------------------
    id_vars = [respondent_id_colname] + respondent_dst_cols
    df_long = df_wide.melt(
        id_vars=id_vars,
        value_vars=card_ids,
        var_name=card_id_colname,
        value_name=rating_colname,
    )
    df_long = df_long.sort_values([respondent_id_colname, card_id_colname])
    df_long = df_long.reset_index(drop=True)

    # ------------------------------------------------------------------
    # 6. カード設計（属性・水準）をマージ
    # ------------------------------------------------------------------
    card_design = _build_card_design(card_ids, attributes, card_id_colname)
    df_long = df_long.merge(card_design, on=card_id_colname)

    # ------------------------------------------------------------------
    # 7. 列順を整理：回答者ID, カードID, rating, 回答者属性, カード属性
    # ------------------------------------------------------------------
    attr_names = [list(a.keys())[0] for a in attributes]
    col_order = (
        [respondent_id_colname, card_id_colname, rating_colname]
        + respondent_dst_cols
        + attr_names
    )
    df_long = df_long[col_order]

    # ------------------------------------------------------------------
    # 8. CSV保存（任意）
    # ------------------------------------------------------------------
    if out_csv is not None:
        df_long.to_csv(out_csv, index=False, encoding="utf-8-sig")
        print(f"保存しました: {out_csv}")

    return df_long


# ---------------------------------------------------------------------------
# 内部ヘルパー関数：ファイル読み込み
# ---------------------------------------------------------------------------

def _read_microsoft_forms(path: Path) -> pd.DataFrame:
    """
    Microsoft Forms の回答ファイルを読み込む。
    .xlsx を想定するが、.csv（BOM付きUTF-8）も受け付ける。
    """
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        try:
            return pd.read_excel(path, engine="openpyxl")
        except ImportError:
            raise ImportError(
                "Microsoft Forms の .xlsx ファイルを読み込むには openpyxl が必要です。\n"
                "以下のコマンドでインストールしてください：\n"
                "  pip install openpyxl"
            )
    # .csv の場合（BOM付きUTF-8）
    return pd.read_csv(path, encoding="utf-8-sig")


def _read_google_forms(path: Path) -> pd.DataFrame:
    """Google Forms の回答CSVを読み込む（UTF-8 / BOM付きUTF-8）。"""
    return pd.read_csv(path, encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# 内部ヘルパー関数：管理列の検出
# ---------------------------------------------------------------------------

# Microsoft Forms が自動生成する管理列のパターン
_MICROSOFT_SYSTEM_PATTERNS = [
    r"^id$",
    r"^start\s*time$",
    r"^completion\s*time$",
    r"^email$",
    r"^name$",
    r"^last\s*modified\s*time$",
    r"^開始時刻$",
    r"^完了時刻$",
    r"^最終変更時刻$",
    r"^メール(アドレス)?$",
    r"^名前$",
]

# Google Forms が自動生成する管理列のパターン
_GOOGLE_SYSTEM_PATTERNS = [
    r"^timestamp$",
    r"^タイムスタンプ$",
    r"^開始時刻$",
    r"^完了時刻$",
    r"^最終変更時刻$",
    r"^メール(アドレス)?$",
    r"^名前$",
    r"^email$",
    r"^email\s*address$",
    r"^start\s*time$",
    r"^completion\s*time$",
    r"^last\s*modified\s*time$",
]


def _detect_microsoft_system_cols(df: pd.DataFrame) -> List[str]:
    """Microsoft Forms の管理列を検出する。"""
    return _detect_system_cols(df, _MICROSOFT_SYSTEM_PATTERNS)


def _detect_google_system_cols(df: pd.DataFrame) -> List[str]:
    """Google Forms の管理列を検出する。"""
    return _detect_system_cols(df, _GOOGLE_SYSTEM_PATTERNS)


def _detect_system_cols(df: pd.DataFrame, patterns: List[str]) -> List[str]:
    """指定したパターンに一致する管理列を検出する共通処理。"""
    system = []
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in patterns:
            if re.match(pattern, col_lower, re.IGNORECASE):
                system.append(col)
                break
    return system


# ---------------------------------------------------------------------------
# 内部ヘルパー関数：評点列の選択・バリデーション
# ---------------------------------------------------------------------------

def _pick_rating_cols(
    candidates: List[str],
    df: pd.DataFrame,
    n_cards: int,
    csv_path: str,
) -> List[str]:
    """
    評点列を candidates から n_cards 列分選ぶ。

    優先順位：
    1. 候補列の中で数値型の列が n_cards 個ある → それを採用
    2. 候補列の右端 n_cards 列を採用（数値変換できるか確認）
    3. 上記でも取得できなければ ValueError
    """
    numeric_candidates = [
        c for c in candidates
        if pd.api.types.is_numeric_dtype(df[c])
        or _is_coercible_to_numeric(df[c])
    ]

    if len(numeric_candidates) >= n_cards:
        return numeric_candidates[-n_cards:]

    if len(candidates) >= n_cards:
        selected = candidates[-n_cards:]
        for col in selected:
            if not _is_coercible_to_numeric(df[col]):
                raise ValueError(
                    f"評点列の自動検出に失敗しました。\n"
                    f"列 '{col}' を数値に変換できません。\n"
                    f"ファイルの列構造を確認してください: {csv_path}"
                )
        return selected

    raise ValueError(
        f"評点列が {n_cards} 列分見つかりませんでした。\n"
        f"評点列の候補: {candidates}\n"
        f"n_cards={n_cards} に対して候補が {len(candidates)} 列しかありません。\n"
        f"ファイルの列構造を確認してください: {csv_path}"
    )


def _is_coercible_to_numeric(series: pd.Series) -> bool:
    """pd.to_numeric で変換できるか（NaN以外の値が1つ以上あるか）を確認する。"""
    return pd.to_numeric(series, errors="coerce").notna().any()


# ---------------------------------------------------------------------------
# 内部ヘルパー関数：カード設計・属性の処理
# ---------------------------------------------------------------------------

def _build_card_design(
    card_ids: List[str],
    attributes: Sequence[Dict[str, Sequence]],
    card_id_colname: str,
) -> pd.DataFrame:
    """カードID と属性・水準の対応テーブルを作成する。"""
    data: Dict[str, list] = {card_id_colname: card_ids}
    for attr_dict in attributes:
        attr_name, levels = list(attr_dict.items())[0]
        data[attr_name] = list(levels)
    return pd.DataFrame(data)


def _normalize_attributes(
    attributes: "pd.DataFrame | Sequence[Dict[str, Sequence]]",
    n_cards: int,
) -> "List[Dict[str, list]]":
    """
    attributes を内部処理用の「辞書のリスト」形式に統一する。

    - pd.DataFrame が渡された場合：列ごとに {列名: 値リスト} の辞書に変換する
    - 辞書のリストが渡された場合：そのまま返す
    """
    if isinstance(attributes, pd.DataFrame):
        if len(attributes) != n_cards:
            raise ValueError(
                f"cards（attributes）の行数 ({len(attributes)}) が "
                f"n_cards ({n_cards}) と一致しません。\n"
                f"cards の行数と n_cards を同じ値にしてください。"
            )
        return [
            {col: list(attributes[col])}
            for col in attributes.columns
        ]
    return list(attributes)


def _check_attributes(
    attributes: "List[Dict[str, list]]",
    n_cards: int,
) -> None:
    """attributes の構造と水準数を検証する。"""
    if not attributes:
        raise ValueError("attributes が空です。少なくとも1つの属性を指定してください。")

    if len(attributes) == 1:
        warnings.warn(
            "属性が1つしかありません。\n"
            "属性が1つの場合、複数属性間のトレードオフが測れないため、\n"
            "支払意思額（WTP）の計算ができません。\n"
            "コンジョイント分析の導入として使う場合は問題ありませんが、\n"
            "本分析では属性を2つ以上にすることを推奨します。",
            UserWarning,
            stacklevel=4,
        )

    for i, attr_dict in enumerate(attributes):
        if not isinstance(attr_dict, dict) or len(attr_dict) != 1:
            raise ValueError(
                f"attributes[{i}] は キー1つの辞書である必要があります。\n"
                f"例：{{\"price\": [6, 10, 6, 10]}}\n"
                f"実際の値：{attr_dict}"
            )
        attr_name, levels = list(attr_dict.items())[0]
        if len(levels) != n_cards:
            raise ValueError(
                f"属性 '{attr_name}' の水準リストの長さ ({len(levels)}) が "
                f"n_cards ({n_cards}) と一致しません。\n"
                f"水準リスト: {list(levels)}"
            )
