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
import zipfile
from pathlib import Path
from typing import Dict, List, Literal, Optional, Sequence

import pandas as pd

# pandas の行番号列（Unnamed: 0 など）の判定（rating / choice 共通）
from .analysis import _is_index_artifact_column


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
        Microsoft Forms の場合は .xlsx / .csv のどちらでも読み込めるが、
        **.csv を推奨**する（追加パッケージが不要で、ブラウザ上の
        Jupyter でもファイルが壊れないため）。ダウンロードした .xlsx を
        Excel で開き、「CSV UTF-8（コンマ区切り）」で保存し直せばよい。
        Google Forms の場合は .csv。

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
        "microsoft" : Microsoft Forms（.csv 推奨。.xlsx も可）
        "google"    : Google Forms（.csv形式）

        .. note::
            "microsoft" では .xlsx / .csv のどちらも読み込めるが、**.csv を
            推奨**する。.csv なら Excel を読むための追加パッケージ
            （openpyxl など）が不要で、ブラウザ上の Jupyter でもファイルが
            壊れないためである。ダウンロードした .xlsx を Excel で開き、
            「CSV UTF-8（コンマ区切り）」で保存し直してから渡すこと
            （forms="microsoft" のままでよい）。

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
        respondent_cols で指定された列がファイルにない場合。
        評点列が推測されたプロファイル数分見つからない場合。
    """

    # ------------------------------------------------------------------
    # 0. 入力チェック
    # ------------------------------------------------------------------
    # profiles が DataFrame の場合、pandas の行番号列（Unnamed: 0 など）が
    # 混入していれば警告のうえ除外する。profiles.to_csv() を index=False なしで
    # 保存した CSV を読み込むと、index（プロファイルID）がこのような列に
    # なって混入し、属性として出力に紛れ込んでしまうため。
    if isinstance(profiles, pd.DataFrame):
        artifact_cols = [
            c for c in profiles.columns if _is_index_artifact_column(c)
        ]
        if artifact_cols:
            warnings.warn(
                f"profiles に pandas の行番号列とみられる列があります: {artifact_cols}\n"
                "  profiles.to_csv() を index=False なしで保存した CSV を読み込むと、\n"
                "  index（行番号やプロファイルID）がこのような列になって混入します。\n"
                "  属性ではないため除外して処理を続けます。\n"
                "  保存時は profiles.to_csv('profiles.csv', index=False) とするか、\n"
                "  読み込み時に pd.read_csv(..., index_col=0) で index に戻してください。",
                UserWarning,
                stacklevel=2,
            )
            profiles = profiles.drop(columns=artifact_cols)

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

    # forms="microsoft" なのに .xlsx/.xls/.csv 以外の拡張子の場合は警告を出す
    # （.csv は正式にサポートしており、むしろ推奨のため警告しない）
    if forms == "microsoft" and csv_path.suffix.lower() not in (
        ".xlsx",
        ".xls",
        ".csv",
    ):
        warnings.warn(
            f"forms='microsoft' が指定されていますが、\n"
            f"ファイルの拡張子が '{csv_path.suffix}' です。\n"
            "Microsoft Forms のファイルは .xlsx または .csv です。\n"
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
    missing_resp = [c for c in respondent_src_cols if c not in raw.columns]
    if missing_resp:
        raise ValueError(
            f"respondent_cols で指定された列がファイルにありません: {missing_resp}\n"
            f"  ファイルの列: {list(raw.columns)}"
        )

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

    # 評点を数値化する。Forms の出力では評点が文字列（"5" など）で入ることが
    # あり、そのまま fit() に渡すと statsmodels の分かりにくいエラーになる。
    # 数値化できない値（空欄・記号など）は NaN になり、fit() の欠損処理に乗る。
    n_nonnull_before = df_long[rating_colname].notna().sum()
    df_long[rating_colname] = pd.to_numeric(
        df_long[rating_colname], errors="coerce"
    )
    n_coerced = n_nonnull_before - df_long[rating_colname].notna().sum()
    if n_coerced > 0:
        warnings.warn(
            f"評点列に数値へ変換できない値が {n_coerced} 件あり、"
            "欠損（NaN）として扱います。\n"
            "  該当行は fit() の際に分析から除外されます。",
            UserWarning,
            stacklevel=2,
        )

    # プロファイルIDは "P1", "P2", ... の提示順（数値順）で並べる。
    # 文字列のまま並べ替えると P1, P10, P11, P2, … の辞書順になってしまう。
    profile_order = {pid: i for i, pid in enumerate(profile_ids)}
    df_long = df_long.sort_values(
        [respondent_id_colname, profile_id_colname],
        key=lambda s: (
            s.map(profile_order) if s.name == profile_id_colname else s
        ),
    )
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

# Excel の読み込みエンジン。calamine は JupyterLite（Pyodide）に同梱されて
# いて高速だが pandas>=2.2 が必要なため、openpyxl を後続の候補に置く。
_EXCEL_ENGINES = ("calamine", "openpyxl")

# 読み込めなかったときの共通の対処法（CSV で保存し直す手順）。
_EXCEL_TO_CSV_HINT = (
    "\n"
    "  【対処法】\n"
    "  1. このファイルを Excel で開く\n"
    "  2. 「名前を付けて保存」で「CSV UTF-8（コンマ区切り）(*.csv)」を選んで保存する\n"
    "  3. 保存した .csv をこの関数に渡す\n"
    '     （forms="microsoft" のままで .csv を読み込めます）'
)

# ZIP 構造の検査で弾いた場合の案内。.xlsx は ZIP 形式と決まっているので、
# 構造が違えば破損が確定している（断定してよい）。
_CORRUPT_XLSX_MESSAGE = (
    "Excel ファイル（{ext}）を読み込めませんでした。\n"
    "  ファイル：{path}\n"
    "  サイズ：{size} バイト\n"
    "  このファイルは壊れている可能性が高いです。\n"
    "  JupyterLite やブラウザ上の Jupyter では、ファイルを転送するときに\n"
    "  .xlsx が壊れてしまうことがあります。\n" + _EXCEL_TO_CSV_HINT
)

# エンジンが実際に読みにいって失敗した場合の案内。破損とは限らない
# （形式が非対応、そのエンジンが対応しない構造、など）ので断定しない。
_UNREADABLE_EXCEL_MESSAGE = (
    "Excel ファイル（{ext}）を読み込めませんでした。\n"
    "  ファイル：{path}\n"
    "  サイズ：{size} バイト\n"
    "  ファイルが壊れているか、この形式に対応していない可能性があります。\n"
    "  JupyterLite やブラウザ上の Jupyter では、ファイルを転送するときに\n"
    "  {ext} が壊れてしまうことがあります。\n" + _EXCEL_TO_CSV_HINT + "\n"
    "\n"
    "  各エンジンで発生したエラー：\n{read_errors}"
)

# どのエンジンも使えなかった場合の案内（インストール方法）。
_NO_EXCEL_ENGINE_MESSAGE = (
    "Excel ファイル（{ext}）を読み込むには、追加のパッケージが必要です。\n"
    "  次のいずれかをインストールしてください：\n"
    "    pip install py4conjoint[excel]  （openpyxl が入ります）\n"
    "    pip install python-calamine     （高速な代替。pandas>=2.2 が必要です）\n"
    "\n"
    "  追加インストールが難しい場合は、このファイルを Excel で開き\n"
    "  「CSV UTF-8（コンマ区切り）(*.csv)」で保存し直してから、その .csv を\n"
    '  この関数に渡してください（forms="microsoft" のままで読み込めます）。\n'
    "\n"
    "  各エンジンで発生したエラー：\n{engine_errors}"
)


def _read_microsoft_forms(path: Path) -> pd.DataFrame:
    """
    Microsoft Forms の回答ファイルを読み込む。

    .xlsx を想定するが、.csv（BOM付きUTF-8）も受け付ける。

    .xlsx は ZIP 形式なので、読み込む前に :func:`zipfile.is_zipfile` で
    構造を検査する。壊れている場合はエンジンを試さずに、CSV で保存し直す
    方法を案内する（.xls は ZIP 形式ではないため検査しない）。

    構造検査を通ったら、読み込みエンジンを calamine → openpyxl の順に
    試す。あるエンジンが読みに失敗しても、そこで打ち切らず次のエンジンを
    試す（先頭の calamine が失敗しても openpyxl なら読める場合があるため）。
    すべて失敗したときだけ、原因に応じたエラーを出す。
    """
    suffix = path.suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        # .csv の場合（BOM付きUTF-8）
        return pd.read_csv(path, encoding="utf-8-sig")

    if suffix == ".xlsx" and not zipfile.is_zipfile(path):
        raise ValueError(
            _CORRUPT_XLSX_MESSAGE.format(
                ext=path.suffix, path=path, size=path.stat().st_size
            )
        )

    engine_errors: List[str] = []  # エンジンが使えなかった（未インストール等）
    read_errors: List[str] = []  # エンジンは動いたが読めなかった
    last_read_error: Optional[BaseException] = None
    for engine in _EXCEL_ENGINES:
        try:
            return pd.read_excel(path, engine=engine)
        except (ImportError, ValueError) as e:
            # ImportError：エンジンが未インストール
            # ValueError ：pandas<2.2 が engine="calamine" を知らない
            engine_errors.append(f"    - {engine}：{type(e).__name__}: {e}")
        except Exception as e:
            # zipfile.BadZipFile や calamine 由来の例外など。
            # このエンジンでは読めなかっただけかもしれないので、
            # ここでは打ち切らず、記録して次のエンジンを試す。
            read_errors.append(f"    - {engine}：{type(e).__name__}: {e}")
            last_read_error = e

    if read_errors:
        # 少なくとも1つのエンジンが実際に読みにいって失敗している。
        # ファイル側の問題の可能性が高いので、CSV で保存し直す方法を案内する。
        raise ValueError(
            _UNREADABLE_EXCEL_MESSAGE.format(
                ext=path.suffix,
                path=path,
                size=path.stat().st_size,
                read_errors="\n".join(read_errors),
            )
        ) from last_read_error

    raise ImportError(
        _NO_EXCEL_ENGINE_MESSAGE.format(
            ext=path.suffix, engine_errors="\n".join(engine_errors)
        )
    )


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
       （候補が n_profiles を超える場合は、除外した列名を UserWarning で明示する。
       　評点でない数値質問（満足度・年齢など）が混在していると、評点と
       　プロファイルの対応が気づかないままズレる恐れがあるため。）
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
        selected = numeric_candidates[-n_profiles:]
        excluded = numeric_candidates[:-n_profiles]
        if excluded:
            warnings.warn(
                f"数値の候補列が {len(numeric_candidates)} 列見つかったため、"
                f"右端の {n_profiles} 列を評点列として採用しました。\n"
                f"  採用した列: {selected}\n"
                f"  除外した列: {excluded}\n"
                "  除外された列に評点（プロファイルの設問）が含まれていないか、\n"
                "  必ず確認してください。評点でない数値質問（満足度・年齢など）は\n"
                "  respondent_cols 引数で指定すると候補から外れます。",
                UserWarning,
                stacklevel=3,
            )
        return selected

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
