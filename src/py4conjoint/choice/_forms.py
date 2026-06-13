"""
_forms.py（choice 版）
======================
Microsoft Forms / Google Forms の回答ファイルを **選択型コンジョイント分析
（CBC）用の long形式DataFrame** に変換する内部モジュール。

公開APIである :func:`cbc_forms_to_data` は ``py4conjoint.choice`` から
``import`` できる。

前提とするアンケート形式
------------------------
* **1設問 = 1選択セット**。設問数は design の ``n_sets`` と一致させる。
* 各設問は「次のうちどれを選びますか？」のような **択一式（ラジオボタン）**
  で、回答選択肢が選択セット内の代替案に対応する
  （例：「製品A」「製品B」「製品C」）。
* 回答選択肢の表示順は、design の ``alt_id``（1, 2, ...）の順と
  一致させること。``choice_labels=["A", "B", "C"]`` なら
  「A を含む回答」= alt_id 1、「B を含む回答」= alt_id 2、… と解釈する。
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Dict, List, Literal, Optional, Sequence

import pandas as pd

# ファイル読み込み・管理列検出は rating 版の仕組みをそのまま使う
from ..rating._forms import (
    _detect_google_system_cols,
    _detect_microsoft_system_cols,
    _read_google_forms,
    _read_microsoft_forms,
)

# ---------------------------------------------------------------------------
# 公開API
# ---------------------------------------------------------------------------

def cbc_forms_to_data(
    responses_file: str,
    design: pd.DataFrame,
    choice_labels: Sequence[str],
    *,
    forms: Literal["microsoft", "google"] = "microsoft",
    version: int = 1,
    respondent_cols: Optional[Dict[str, str]] = None,
    obs_id_colname: str = "obsID",
    respondent_id_colname: str = "respondent_id",
    alt_colname: str = "alt",
    choice_colname: str = "choice",
    out_csv: Optional[str] = None,
) -> pd.DataFrame:
    """
    Microsoft Forms / Google Forms の回答ファイルを、条件付きロジット推定用の
    long形式DataFrameに変換する。

    前提とするアンケート形式
    ------------------------
    * **1設問 = 1選択セット**。設問の回答選択肢が代替案に対応する
      （例：「製品A」「製品B」「製品C」から1つ選ぶ）。
    * 設問数は ``design`` の設問数（``n_sets``）と一致している必要がある。
    * 回答値から代替案を特定するために ``choice_labels`` を使う。
      回答文字列に含まれるラベルでマッチングする
      （例：``choice_labels=["A", "B", "C"]`` のとき、回答「製品A」は
      ラベル "A" にマッチ → alt_id 1 が選ばれたと解釈）。

    Parameters
    ----------
    responses_file : str
        Forms からダウンロードした回答ファイルのパス。
        Microsoft Forms の場合は .xlsx、Google Forms の場合は .csv。

    design : pd.DataFrame
        :func:`design_choice_sets` の出力DataFrame
        （``version``, ``set_id``, ``alt_id`` + 属性列）。
        アンケートの各設問に提示した代替案の属性・水準が、
        この設計から復元される。

    choice_labels : list of str
        回答選択肢を識別するラベルのリスト（``alt_id`` の順）。
        例：``["A", "B", "C"]``。
        回答文字列がラベルと完全一致するか、ラベルを含む場合にマッチする。
        長さは design の1設問あたりの代替案数（``n_alts``）と
        一致している必要がある。

    forms : {"microsoft", "google"}, default "microsoft"
        使用するFormsの種類を指定する。
        "microsoft" : Microsoft Forms（.xlsx形式）
        "google"    : Google Forms（.csv形式）

    version : int, default 1
        design のどのバージョンを使うか。
        design を ``n_versions >= 2`` で生成した場合は、バージョンごとに
        別のアンケート（別の回答ファイル）になるため、ファイルごとに
        本関数を呼び、対応するバージョン番号を指定すること。

    respondent_cols : dict, optional
        回答者属性として残したい列の対応辞書。
        ``{"ファイル上の列名": "出力DataFrameの列名"}`` の形式。
        例：``{"性別": "gender", "学年": "year"}``
        省略した場合は回答者属性を付与しない。

    obs_id_colname : str, default "obsID"
        出力DataFrameの「回答者×設問」通し番号列の名前。
        :func:`py4conjoint.choice.fit` の ``choice_set_col`` に渡す。

    respondent_id_colname : str, default "respondent_id"
        出力DataFrameの回答者ID列の名前。
        :func:`py4conjoint.choice.fit` の ``respondent_id_col`` に渡す。

    alt_colname : str, default "alt"
        出力DataFrameの代替案番号列の名前。

    choice_colname : str, default "choice"
        出力DataFrameの選択フラグ（0/1）列の名前。

    out_csv : str, optional
        変換後のDataFrameをCSVとして保存するパス。
        省略した場合は保存しない。

    Returns
    -------
    pd.DataFrame
        long形式のDataFrame。1行 = 1回答者の1設問の1代替案。

        列：``obsID``（回答者×設問の通し番号）, ``respondent_id``,
        ``alt``（代替案番号）, ``choice``（選ばれたら1、それ以外0）,
        [回答者属性], + 属性列。

        そのまま :func:`py4conjoint.choice.encode` →
        :func:`py4conjoint.choice.fit` に渡せる::

            df = pcc.cbc_forms_to_data("responses.xlsx", design, ["A", "B", "C"])
            df_coded = pcc.encode(df, reference_levels={"brand": "A社"})
            result = pcc.fit(
                df_coded,
                choice_set_col="obsID",
                respondent_id_col="respondent_id",
            )

    Raises
    ------
    FileNotFoundError
        responses_file が存在しない場合。
    ValueError
        forms が "microsoft" または "google" 以外の場合。
        design に必要な列（version, set_id, alt_id）がない場合。
        指定した version が design に存在しない場合。
        choice_labels の長さが design の代替案数と一致しない場合。
        回答ファイルの設問数が design の設問数（n_sets）と一致しない場合。
        choice_labels のどれにもマッチしない（または複数にマッチする）
        回答値がある場合。

    Warns
    -----
    UserWarning
        未回答（空欄）の設問がある場合。該当する回答者×設問は
        分析から除外される（その旨を件数つきで警告する）。
    """
    # ------------------------------------------------------------------
    # 0. 入力チェック
    # ------------------------------------------------------------------
    if forms not in ("microsoft", "google"):
        raise ValueError(
            f"forms='{forms}' は無効な値です。\n"
            "'microsoft' または 'google' を指定してください。"
        )

    if not isinstance(design, pd.DataFrame):
        raise TypeError(
            f"design は pandas.DataFrame を指定してください"
            f"（design_choice_sets() の出力）。\n"
            f"  受け取った型: {type(design).__name__}"
        )
    required = ["version", "set_id", "alt_id"]
    missing = [c for c in required if c not in design.columns]
    if missing:
        raise ValueError(
            f"design に必要な列がありません: {missing}\n"
            "  design_choice_sets() の出力をそのまま渡してください。"
        )

    versions = sorted(design["version"].unique())
    if version not in versions:
        raise ValueError(
            f"design にバージョン {version} が存在しません。\n"
            f"  存在するバージョン: {versions}\n"
            "  version 引数を確認してください。"
        )
    design_v = design[design["version"] == version]

    attr_names = [c for c in design.columns if c not in required]
    set_ids = sorted(design_v["set_id"].unique())
    n_sets = len(set_ids)
    n_alts = int(design_v.groupby("set_id")["alt_id"].size().iloc[0])

    choice_labels = [str(lb) for lb in choice_labels]
    if len(choice_labels) != n_alts:
        raise ValueError(
            f"choice_labels の長さ ({len(choice_labels)}) が design の"
            f"1設問あたりの代替案数 ({n_alts}) と一致しません。\n"
            f"  choice_labels: {choice_labels}\n"
            "  代替案の数だけラベルを指定してください（alt_id の順）。"
        )

    csv_path = Path(responses_file)
    if not csv_path.exists():
        raise FileNotFoundError(
            f"ファイルが見つかりません: {responses_file}\n"
            "ファイル名とパスを確認してください。"
        )

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
    # ------------------------------------------------------------------
    if forms == "microsoft":
        raw = _read_microsoft_forms(csv_path)
    else:
        raw = _read_google_forms(csv_path)

    # ------------------------------------------------------------------
    # 2. 管理列・回答者属性列を除外して設問列を特定する
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

    non_question_cols = set(system_cols) | set(respondent_src_cols)
    question_cols = [c for c in raw.columns if c not in non_question_cols]

    if len(question_cols) != n_sets:
        raise ValueError(
            f"回答ファイルの設問数 ({len(question_cols)}) が design の"
            f"設問数 n_sets ({n_sets}) と一致しません。\n"
            f"  設問列の候補: {question_cols}\n"
            "  ・設問以外の列（回答者属性など）は respondent_cols で指定して\n"
            "    除外してください。\n"
            "  ・design のバージョン（version 引数）が正しいか確認してください。"
        )

    # ------------------------------------------------------------------
    # 3. 回答値 → 代替案番号（alt_id）のマッチング
    # ------------------------------------------------------------------
    # 設問列の並び順 = design の set_id の昇順、と対応づける
    question_map = dict(zip(question_cols, set_ids))

    n_respondents = len(raw)
    respondent_ids = range(1, n_respondents + 1)

    chosen: Dict[tuple, int] = {}   # (respondent_id, set_id) → alt_id
    n_unanswered = 0
    unmatched: List[str] = []

    for r_idx, resp_id in enumerate(respondent_ids):
        for q_col, set_id in question_map.items():
            value = raw.iloc[r_idx][q_col]
            if pd.isna(value) or str(value).strip() == "":
                n_unanswered += 1
                continue
            alt_id = _match_choice_label(str(value), choice_labels)
            if alt_id is None:
                unmatched.append(str(value))
                continue
            chosen[(resp_id, set_id)] = alt_id

    if unmatched:
        uniq = sorted(set(unmatched))
        raise ValueError(
            f"choice_labels のどれにもマッチしない回答値があります: {uniq}\n"
            f"  choice_labels: {choice_labels}\n"
            "  ・ラベルが回答選択肢の文字列に含まれているか確認してください。\n"
            "  ・複数のラベルにマッチする曖昧な回答値も無効になります。"
        )

    if n_unanswered > 0:
        warnings.warn(
            f"未回答の設問が {n_unanswered} 件ありました。\n"
            "該当する回答者×設問（選択セット）は分析から除外されます。",
            UserWarning,
            stacklevel=2,
        )

    # ------------------------------------------------------------------
    # 4. long形式の組み立て（1行 = 1回答者 × 1設問 × 1代替案）
    # ------------------------------------------------------------------
    design_indexed = design_v.set_index(["set_id", "alt_id"])

    rows = []
    obs_counter = 0
    for r_idx, resp_id in enumerate(respondent_ids):
        for set_id in set_ids:
            key = (resp_id, set_id)
            if key not in chosen:  # 未回答セットは除外
                continue
            obs_counter += 1
            for alt_id in range(1, n_alts + 1):
                row = {
                    obs_id_colname: obs_counter,
                    respondent_id_colname: resp_id,
                    alt_colname: alt_id,
                    choice_colname: int(chosen[key] == alt_id),
                }
                for src, dst in respondent_rename.items():
                    row[dst] = raw.iloc[r_idx][src]
                for attr in attr_names:
                    row[attr] = design_indexed.loc[(set_id, alt_id), attr]
                rows.append(row)

    col_order = (
        [obs_id_colname, respondent_id_colname, alt_colname, choice_colname]
        + list(respondent_rename.values())
        + attr_names
    )
    df_long = pd.DataFrame(rows, columns=col_order)

    # ------------------------------------------------------------------
    # 5. CSV保存（任意）
    # ------------------------------------------------------------------
    if out_csv is not None:
        df_long.to_csv(out_csv, index=False, encoding="utf-8-sig")
        print(f"保存しました: {out_csv}")

    return df_long


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------

def _match_choice_label(
    value: str,
    choice_labels: List[str],
) -> Optional[int]:
    """
    回答文字列 ``value`` を choice_labels とマッチングし、
    対応する代替案番号（1始まりの alt_id）を返す。

    マッチング規則：

    1. ラベルと **完全一致** すればそのラベル。
    2. 完全一致がなければ、``value`` に **含まれる** ラベルを探す。
       ちょうど1つ含まれていればそのラベル。
    3. 0個または2個以上マッチする場合は ``None``（無効）。
    """
    value = value.strip()
    for i, label in enumerate(choice_labels):
        if value == label:
            return i + 1
    hits = [i for i, label in enumerate(choice_labels) if label in value]
    if len(hits) == 1:
        return hits[0] + 1
    return None
