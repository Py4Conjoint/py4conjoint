"""
_forms.py
=========
Microsoft Forms / Google Forms の回答ファイルを評点型コンジョイント分析用の
long形式DataFrameに変換する内部モジュール。

公開APIである :func:`forms_to_data` はトップレベル
（``py4conjoint`` パッケージ）から ``import`` できる。
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

def forms_to_data(
    responses_file: str,
    profiles: "pd.DataFrame | Dict[str, Sequence]",
    *,
    n_profiles: Optional[int] = None,
    forms: Literal["microsoft", "google"] = "microsoft",
    respondent_cols: Optional[Dict[str, str]] = None,
    profile_id_prefix: str = "P",
    rating_colname: str = "rating",
    respondent_id_colname: str = "respondent_id",
    profile_id_colname: str = "profile_id",
    out_csv: Optional[str] = None,
) -> pd.DataFrame:
    """
    Microsoft Forms / Google Forms の回答ファイルをlong形式DataFrameに変換する。

    Parameters
    ----------
    responses_file : str
        Forms からダウンロードした回答ファイルのパス。
        Microsoft Forms の場合は .xlsx、Google Forms の場合は .csv。

    profiles : pd.DataFrame または dict
        プロファイル設計を指定する。以下の2形式を受け付ける。

        【形式A：DataFrameをそのまま渡す（推奨）】
            授業で作成した profiles をそのまま渡すことができる。
            行がプロファイル、列が属性に対応する。
            インデックスは ["P1","P2",...] でも整数でも可。

            例：
            profiles = pd.DataFrame({
                "price":  [6, 10, 6, 10],
                "os":     ["android", "apple", "apple", "android"],
                "camera": ["標準", "標準", "高性能", "高性能"],
            }, index=["P1", "P2", "P3", "P4"])

            df = pc.forms_to_data(responses_file, profiles)

        【形式B：辞書】
            属性名をキー、プロファイル順の水準リストを値とする辞書。

            例：
            profiles = {
                "price":  [6, 10, 6, 10],
                "os":     ["android", "apple", "apple", "android"],
                "camera": ["標準", "標準", "高性能", "高性能"],
            }

            df = pc.forms_to_data(responses_file, profiles)

    n_profiles : int, optional
        アンケートで提示したプロファイルの枚数。
        省略時は profiles の水準リストの長さから自動推測する。
        例：4

    forms : {"microsoft", "google"}, default "microsoft"
        使用するFormsの種類を指定する。
        "microsoft" : Microsoft Forms（.xlsx形式）
        "google"    : Google Forms（.csv形式）

    respondent_cols : dict, optional
        回答者属性として残したい列の対応辞書。
        {"CSVの列名": "出力DataFrameの列名"} の形式。
        例：{"性別": "gender", "学年": "year"}
        省略した場合は回答者属性を付与しない。

    profile_id_prefix : str, default "P"
        プロファイルIDの接頭辞。"P" なら P1, P2, P3, P4 となる。

    rating_colname : str, default "rating"
        出力DataFrameの評点列名。

    respondent_id_colname : str, default "respondent_id"
        出力DataFrameの回答者ID列名。

    profile_id_colname : str, default "profile_id"
        出力DataFrameのプロファイルID列名。

    out_csv : str, optional
        変換後のDataFrameをCSVとして保存するパス。
        省略した場合は保存しない。

    Returns
    -------
    pd.DataFrame
        long形式のDataFrame。
        列：respondent_id, profile_id, rating, [回答者属性], [プロファイル属性]

    Raises
    ------
    FileNotFoundError
        responses_file が存在しない場合。
    ValueError
        forms が "microsoft" または "google" 以外の場合。
        属性の水準リストの長さが揃っていない場合。
        評点列が推測されたプロファイル数分見つからない場合。
    """

    # ------------------------------------------------------------------
    # 0. 入力チェック
    # ------------------------------------------------------------------
    if n_profiles is None:
        n_profiles = _infer_n_profiles(profiles)

    if forms not in ("microsoft", "google"):
        raise ValueError(
            f"forms='{forms}' は無効な値です。\n"
            "'microsoft' または 'google' を指定してください。"
        )

    profiles = _normalize_profiles(profiles, n_profiles)
    _check_profiles(profiles, n_profiles)

    csv_path = Path(responses_file)
    if not csv_path.exists():
        raise FileNotFoundError(
            f"ファイルが見つかりません: {responses_file}\n"
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

    rating_cols = _pick_rating_cols(rating_candidate_cols, raw, n_profiles, responses_file)

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
    profile_ids = [f"{profile_id_prefix}{i+1}" for i in range(n_profiles)]
    rating_rename = dict(zip(rating_cols, profile_ids))
    df_wide = df_wide.rename(columns=rating_rename)

    # ------------------------------------------------------------------
    # 5. wide → long 変換
    # ------------------------------------------------------------------
    id_vars = [respondent_id_colname] + respondent_dst_cols
    df_long = df_wide.melt(
        id_vars=id_vars,
        value_vars=profile_ids,
        var_name=profile_id_colname,
        value_name=rating_colname,
    )
    df_long = df_long.sort_values([respondent_id_colname, profile_id_colname])
    df_long = df_long.reset_index(drop=True)

    # ------------------------------------------------------------------
    # 6. プロファイル設計（属性・水準）をマージ
    # ------------------------------------------------------------------
    profile_design = _build_profile_design(profile_ids, profiles, profile_id_colname)
    df_long = df_long.merge(profile_design, on=profile_id_colname)

    # ------------------------------------------------------------------
    # 7. 列順を整理：respondent_id, profile_id, rating, 回答者属性, プロファイル属性
    # ------------------------------------------------------------------
    attr_names = [list(a.keys())[0] for a in profiles]
    col_order = (
        [respondent_id_colname, profile_id_colname, rating_colname]
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
    n_profiles: int,
    csv_path: str,
) -> List[str]:
    """
    評点列を candidates から n_profiles 列分選ぶ。

    優先順位：
    1. 数値型（または数値変換可能）の候補列が n_profiles 個以上ある
       → そのうち右端の n_profiles 列を採用
    2. 候補列全体が n_profiles 個以上ある
       → 右端の n_profiles 列を採用（数値変換できるか確認）
    3. 上記でも取得できなければ ValueError
    """
    numeric_candidates = [
        c for c in candidates
        if pd.api.types.is_numeric_dtype(df[c])
        or _is_coercible_to_numeric(df[c])
    ]

    if len(numeric_candidates) >= n_profiles:
        return numeric_candidates[-n_profiles:]

    if len(candidates) >= n_profiles:
        selected = candidates[-n_profiles:]
        for col in selected:
            if not _is_coercible_to_numeric(df[col]):
                raise ValueError(
                    f"評点列の自動検出に失敗しました。\n"
                    f"列 '{col}' を数値に変換できません。\n"
                    f"ファイルの列構造を確認してください: {csv_path}"
                )
        return selected

    raise ValueError(
        f"評点列が {n_profiles} 列分見つかりませんでした。\n"
        f"評点列の候補: {candidates}\n"
        f"n_profiles={n_profiles} に対して候補が {len(candidates)} 列しかありません。\n"
        f"ファイルの列構造を確認してください: {csv_path}"
    )


def _is_coercible_to_numeric(series: pd.Series) -> bool:
    """pd.to_numeric で変換できるか（NaN以外の値が1つ以上あるか）を確認する。"""
    return pd.to_numeric(series, errors="coerce").notna().any()


# ---------------------------------------------------------------------------
# 内部ヘルパー関数：プロファイル設計・属性の処理
# ---------------------------------------------------------------------------

def _build_profile_design(
    profile_ids: List[str],
    profiles: Sequence[Dict[str, Sequence]],
    profile_id_colname: str,
) -> pd.DataFrame:
    """プロファイルID と属性・水準の対応テーブルを作成する。"""
    data: Dict[str, list] = {profile_id_colname: profile_ids}
    for attr_dict in profiles:
        attr_name, levels = list(attr_dict.items())[0]
        data[attr_name] = list(levels)
    return pd.DataFrame(data)


def _infer_n_profiles(
    profiles: "pd.DataFrame | Dict[str, Sequence]",
) -> int:
    """profiles の水準リスト長からプロファイル数を推測する。"""
    if isinstance(profiles, pd.DataFrame):
        return len(profiles)
    if isinstance(profiles, dict):
        if not profiles:
            raise ValueError("profiles が空です。少なくとも1つの属性を指定してください。")
        return len(next(iter(profiles.values())))
    raise TypeError(
        f"profiles は pd.DataFrame または dict を指定してください。\n"
        f"  受け取った型: {type(profiles).__name__}"
    )


def _normalize_profiles(
    profiles: "pd.DataFrame | Dict[str, Sequence]",
    n_profiles: int,
) -> "List[Dict[str, list]]":
    """
    profiles を内部処理用の「辞書のリスト」形式に統一する。

    - pd.DataFrame → 列ごとに {列名: 値リスト} の辞書に変換する
    - dict → [{属性名: 値リスト}, ...] に変換する
    """
    if isinstance(profiles, pd.DataFrame):
        if len(profiles) != n_profiles:
            raise ValueError(
                f"profiles の行数 ({len(profiles)}) が "
                f"n_profiles ({n_profiles}) と一致しません。"
            )
        return [
            {col: list(profiles[col])}
            for col in profiles.columns
        ]
    if isinstance(profiles, dict):
        return [{k: list(v)} for k, v in profiles.items()]
    raise TypeError(
        f"profiles は pd.DataFrame または dict を指定してください。\n"
        f"  受け取った型: {type(profiles).__name__}"
    )


def _check_profiles(
    profiles: "List[Dict[str, list]]",
    n_profiles: int,
) -> None:
    """profiles の構造と水準数を検証する。"""
    if not profiles:
        raise ValueError("profiles が空です。少なくとも1つの属性を指定してください。")

    if len(profiles) == 1:
        warnings.warn(
            "属性が1つしかありません。\n"
            "属性が1つの場合、複数属性間のトレードオフが測れないため、\n"
            "限界支払意思額（WTP）の計算ができません。\n"
            "コンジョイント分析の導入として使う場合は問題ありませんが、\n"
            "本分析では属性を2つ以上にすることを推奨します。",
            UserWarning,
            stacklevel=3,
        )

    for i, attr_dict in enumerate(profiles):
        if not isinstance(attr_dict, dict) or len(attr_dict) != 1:
            raise ValueError(
                f"profiles[{i}] は キー1つの辞書である必要があります。\n"
                f"例：{{\"price\": [6, 10, 6, 10]}}\n"
                f"実際の値：{attr_dict}"
            )
        attr_name, levels = list(attr_dict.items())[0]
        if len(levels) != n_profiles:
            raise ValueError(
                f"属性 '{attr_name}' の水準リストの長さ ({len(levels)}) が "
                f"n_profiles ({n_profiles}) と一致しません。\n"
                f"水準リスト: {list(levels)}"
            )
