"""
design.py（choice 版）
======================
選択型コンジョイント分析（CBC）の **選択セット設計** を担当するモジュール。

* :func:`design_choice_sets` … CBC用の選択セットを生成する
* :func:`check_design` … 生成した設計の事後診断（バランス・独立性・オーバーラップ）
* :func:`suggest_n_respondents` … Johnson-Orme の経験則による必要回答者数の目安

rating 版（:mod:`py4conjoint.rating.design`）と対称的なAPIを提供する。

>>> import py4conjoint.choice as pcc
>>> design = pcc.design_choice_sets(
...     {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
...     n_sets=8, n_alts=3, seed=42,
... )
>>> pcc.check_design(design)
>>> pcc.suggest_n_respondents(
...     {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
...     n_sets=8, n_alts=3,
... )
"""

from __future__ import annotations

import hashlib
import math
from dataclasses import dataclass
from itertools import product as _itertools_product
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

# 警告の構造化表現・表示ヘルパー・pandas の行番号列（Unnamed: 0 など）の判定は
# rating 版と共通のものを使う
from ..rating.analysis import (
    SEVERITY_ORDER,
    Diagnostic,
    _df_to_string_cjk,
    _is_index_artifact_column,
)

# ---------------------------------------------------------------------------
# 公開API: design_choice_sets 関数
# ---------------------------------------------------------------------------


def design_choice_sets(
    attributes: Dict[str, List[Any]],
    n_sets: int,
    n_alts: int,
    *,
    n_versions: int = 1,
    seed: Optional[int] = None,
    auto_balance: bool = False,
    n_candidates: int = 500,
) -> pd.DataFrame:
    """
    CBC（選択型コンジョイント分析）用の **選択セット** を生成する。

    全属性水準の完全交差（Full factorial）N 個のプロファイル候補から、
    各選択セットに ``n_alts`` 個の代替案をランダムに割り当てる。
    **同一選択セット内に同じプロファイルが重複して入ることはない。**

    Parameters
    ----------
    attributes : dict
        ``{"属性名": [水準1, 水準2, ...]}`` の辞書。
        辞書のキー順が列の順序になる。

        例::

            {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]}

    n_sets : int
        1バージョンあたりの設問数（選択セット数）。
        回答者は1設問につき1つの代替案を選ぶ。

    n_alts : int
        1設問（選択セット）あたりの代替案の数。
        2以上、かつ完全交差の候補数 N 以下である必要がある。

    n_versions : int, default 1
        アンケートのバージョン数。
        複数バージョンを作って回答者を分けると、より多くの
        プロファイル組み合わせをカバーできる（設計の質が上がる）。

    seed : int, optional
        乱数シード（再現性のため）。

    auto_balance : bool, default False
        ``True`` にすると、**バランスの良い設計を自動で選びます**。
        seed を手で 1, 2, 3… と変えて良い設計を探さなくてよくなります。
        内部で ``n_candidates`` 個の設計を作って :func:`check_design` で診断し、
        最もバランスの良いものを1つ返します（既定の ``False`` では従来どおり
        単一のランダム設計を返します）。

        .. note::
            これは「たくさん試した中で最もバランスの良いもの」を選ぶ方法であり、
            数学的な最適計画（D 最適計画）ではありません。
            選ばれた設計の質は :func:`check_design` で確認できます。

    n_candidates : int, default 500
        ``auto_balance=True`` のときに内部で生成・診断する設計の数。
        多いほど良い設計が見つかりやすくなりますが、属性・水準が多い大規模な
        設計では時間がかかります。時間がかかりすぎる場合は値を小さくしてください。
        ``auto_balance=False`` のときは使われません。

    Returns
    -------
    pd.DataFrame
        long形式のDataFrame。1行 = 1つの選択セット内の1つの代替案。

        列：``version``（バージョン番号 1〜）, ``choice_set_id``（設問番号 1〜）,
        ``alt_id``（代替案番号 1〜） + 属性列。

        行数 = ``n_versions × n_sets × n_alts``。

        ``df.attrs["n_candidates"]`` — 完全交差の候補数 N。
        ``df.attrs["design_signature"]`` — この設計を一意に表す署名
        （内容から計算した短いハッシュ。:func:`design_signature` 参照）。

        ``auto_balance=True`` のときは、さらに選定の来歴が入る：

        ``df.attrs["auto_balance"]`` — ``{"n_candidates": 評価した設計の数,
        "n_warnings": 選ばれた設計の警告数, "cv_sum": 全属性の CV の合計}``
        の辞書。どの程度の候補から、どんな質の設計が選ばれたかを後から確認できる。

    Raises
    ------
    ValueError
        ``attributes`` が空または辞書でない場合。
        いずれかの属性の水準数が 2 未満、または水準リストに重複がある場合。
        ``n_sets`` / ``n_alts`` / ``n_versions`` が範囲外の場合。
        ``n_alts`` が完全交差の候補数 N を超える場合
        （セット内の重複を禁止しているため）。

    Notes
    -----
    **ランダム設計について**

    本関数は「完全交差からのランダム割り当て（セット内重複なし）」という
    最も基本的な設計法を使う。設問数 × バージョン数が十分あれば、
    水準バランス・独立性ともに実用上問題のない設計が得られる。
    生成後は必ず :func:`check_design` で品質を確認すること。

    良い設計を得るために seed を手で変えて探すのが面倒なときは、
    ``auto_balance=True`` を使うと、複数候補の中から最もバランスの良い設計を
    自動で選んでくれる（seed 探しが不要になる）。

    .. warning::
        **アンケート作成に使った design と、:func:`forms_to_data` に渡す
        design は完全に同一にすること。**
        属性名・水準・水準の順序・``seed``・``n_sets``・``n_alts``・
        ``n_versions`` が1つでも違うと、**同じ seed でも別の設計**になる。
        たとえば ``{"price": [6, 10]}`` と ``{"price": [10, 6]}`` は
        seed が同じでも各選択セットの中身が変わる。
        design がずれると回答と代替案の対応が静かに食い違い、
        **エラーが出ないまま結果が誤る**（例：本来 正 の係数が 負 に出る）。

        これを避けるため、**design は作成後すぐにファイルへ保存し、分析時は
        作り直さず同じファイルを読み込んで** :func:`forms_to_data` に渡すこと::

            design = pcc.design_choice_sets(attrs, n_sets=8, n_alts=3, seed=42)
            design.to_csv("design.csv", index=False)   # 作成後すぐ保存
            # …アンケート実施…
            design = pd.read_csv("design.csv")          # 分析時は読み込むだけ
            df = pcc.forms_to_data("responses.xlsx", design, ["A", "B", "C"])

        2つの design が同一かは :func:`design_signature` の署名で確認できる。

    Examples
    --------
    >>> design = pcc.design_choice_sets(
    ...     {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
    ...     n_sets=8, n_alts=3, seed=42,
    ... )
    >>> design.head(6)  # 設問1・2の代替案
    """
    # ---------- 入力チェック ----------
    if not isinstance(attributes, dict) or len(attributes) == 0:
        raise ValueError(
            "attributes は空でない辞書を指定してください。\n"
            "  例: {'price': [100, 150, 200], 'brand': ['A社', 'B社']}"
        )
    for attr, levels in attributes.items():
        if len(levels) < 2:
            raise ValueError(
                f"属性 '{attr}' の水準数は 2 以上にしてください（現在: {len(levels)}）。"
            )
        # 重複した水準は完全交差に同一プロファイルを複数作り、
        # 「同一選択セット内に同じプロファイルは入らない」という本関数の
        # 保証が破れるため、ここで弾く。
        if len(set(levels)) != len(levels):
            raise ValueError(
                f"属性 '{attr}' の水準リストに重複があります: {list(levels)}\n"
                "  水準は重複なく指定してください。"
            )
    if n_sets < 1:
        raise ValueError(
            f"n_sets は 1 以上の整数を指定してください（指定値: {n_sets}）。"
        )
    if n_alts < 2:
        raise ValueError(
            f"n_alts は 2 以上の整数を指定してください（指定値: {n_alts}）。\n"
            "  選択セットには比較対象として最低2つの代替案が必要です。"
        )
    if n_versions < 1:
        raise ValueError(
            f"n_versions は 1 以上の整数を指定してください（指定値: {n_versions}）。"
        )
    if auto_balance and n_candidates < 1:
        raise ValueError(
            f"n_candidates は 1 以上の整数を指定してください（指定値: {n_candidates}）。"
        )

    attrs = list(attributes.keys())
    levels_list = [list(attributes[a]) for a in attrs]

    # 完全交差の候補プロファイル
    full = pd.DataFrame(
        [dict(zip(attrs, combo)) for combo in _itertools_product(*levels_list)]
    )
    N = len(full)

    if n_alts > N:
        raise ValueError(
            f"n_alts ({n_alts}) が完全交差の候補数 N ({N}) を超えています。\n"
            "  同一選択セット内のプロファイル重複は禁止しているため、\n"
            f"  n_alts を {N} 以下にするか、属性・水準を増やしてください。"
        )

    # ---------- ランダム割り当て（セット内重複なし） ----------
    if not auto_balance:
        # 従来どおり：単一のランダム設計を生成する（後方互換）。
        rng = np.random.default_rng(seed)
        out = _assign_random_design(full, N, n_versions, n_sets, n_alts, rng)
        out.attrs["n_candidates"] = N
        out.attrs["design_signature"] = design_signature(out)
        return out

    # auto_balance：n_candidates 個の候補を作り、最もバランスの良いものを選ぶ。
    # 候補生成は与えられた seed から決定的に派生させる（同じ seed → 同じ結果）。
    child_seqs = np.random.SeedSequence(seed).spawn(n_candidates)
    best = None  # (cand, n_warnings, cv_sum)
    for cs in child_seqs:
        rng_i = np.random.default_rng(cs)
        cand = _assign_random_design(full, N, n_versions, n_sets, n_alts, rng_i)
        chk = check_design(cand, attributes=attrs)
        n_warn = len(chk.diagnostics)
        cv_sum = float(chk.balance["CV"].sum())
        if _is_better_design(best, n_warn, cv_sum):
            best = (cand, n_warn, cv_sum)

    out, best_warn, best_cv = best
    out.attrs["n_candidates"] = N
    out.attrs["design_signature"] = design_signature(out)
    out.attrs["auto_balance"] = {
        "n_candidates": int(n_candidates),  # 評価した候補設計の数
        "n_warnings": int(best_warn),  # 選ばれた設計の警告数
        "cv_sum": round(float(best_cv), 6),  # 全属性の CV の合計
    }
    return out


# ---------------------------------------------------------------------------
# 公開API: design_signature 関数
# ---------------------------------------------------------------------------


def design_signature(design: pd.DataFrame) -> str:
    """
    選択セット設計（design）の内容から、一意な **署名**（短いハッシュ）を計算する。

    署名は ``version``・``choice_set_id``・``alt_id`` と各属性の **値そのもの**
    から決定的に計算する。したがって：

    * 属性名・水準・**水準の順序**・``n_sets``・``n_alts`` が完全に同一なら
      同じ署名になる（``seed`` を指定して再生成した場合も一致する）。
    * 水準の順序が1つでも違う、属性や設問の中身が違う設計は、別の署名になる。
    * ``seed=None`` で生成した設計は呼ぶたびに中身が変わるため、署名も変わる。
    * pandas の行番号列（``Unnamed: 0`` など。``to_csv()`` を ``index=False``
      なしで保存した CSV の痕跡）は設計の中身ではないため **無視** する。
      index 付きで保存してしまった CSV でも、元の設計と署名が一致する。

    アンケート作成に使った design と、分析時に :func:`forms_to_data` へ渡す
    design が **同一かどうかを確認** するために使う。両者の署名が一致すれば
    同じ設計であり、回答と代替案の対応が正しく解決される。

    Parameters
    ----------
    design : pd.DataFrame
        :func:`design_choice_sets` の出力、または ``choice_set_id``・``alt_id``
        ＋属性列を持つ設計表（CSV から読み込んだものでもよい）。

    Returns
    -------
    str
        12 桁の十六進文字列（内容から計算した署名）。

    Examples
    --------
    >>> d1 = pcc.design_choice_sets({"price": [6, 10], "os": ["a", "b"]},
    ...                             n_sets=4, n_alts=2, seed=1)
    >>> d2 = pcc.design_choice_sets({"price": [10, 6], "os": ["a", "b"]},
    ...                             n_sets=4, n_alts=2, seed=1)
    >>> pcc.design_signature(d1) == pcc.design_signature(d2)
    False
    """
    if not isinstance(design, pd.DataFrame):
        raise TypeError(
            "design は pandas.DataFrame を指定してください"
            f"（受け取った型: {type(design).__name__}）。"
        )
    if "choice_set_id" not in design.columns or "alt_id" not in design.columns:
        raise ValueError(
            "design に choice_set_id・alt_id 列が必要です。\n"
            "  design_choice_sets() の出力、または設計CSV を渡してください。"
        )

    id_cols = [c for c in ("version", "choice_set_id", "alt_id") if c in design.columns]
    # 属性列は順序の影響を受けないよう、列名で並べてから値を取り込む。
    # pandas の行番号列（Unnamed: 0 など。index=False を付けずに保存した
    # CSV の痕跡）は設計の中身ではないため署名から除外する。これにより
    # index 付きで保存してしまった CSV でも、元の設計と署名が一致する。
    attr_cols = sorted(
        c
        for c in design.columns
        if c not in id_cols and not _is_index_artifact_column(c)
    )
    cols = id_cols + attr_cols

    # 行を ID 列で正準順に並べ、列順も固定して値を文字列化（決定的）
    canon = design[cols].sort_values(id_cols).reset_index(drop=True)

    def _plain(v):
        # numpy スカラー（np.int64 など）を Python の値に変換してから repr する。
        # numpy 2.x では repr(np.int64(6)) が 'np.int64(6)' になり（1.x は '6'）、
        # 同じ設計でも numpy のバージョンによって署名が変わってしまうため。
        # 署名は「時間・環境をまたいで design の同一性を確認する」ためのものなので、
        # 環境に依存しない Python 値の repr（'6'）に正規化する。
        return v.item() if hasattr(v, "item") else v

    lines = [
        "|".join(f"{c}={_plain(v)!r}" for c, v in zip(cols, row))
        for row in canon.itertuples(index=False, name=None)
    ]
    payload = "\n".join(lines)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]


# ---------------------------------------------------------------------------
# 公開API: check_design 関数
# ---------------------------------------------------------------------------


@dataclass
class ChoiceDesignCheckResult:
    """
    :func:`check_design` の診断結果を保持するオブジェクト。

    rating 版の :class:`py4conjoint.rating.DesignCheckResult` と
    同様の使い勝手（``summary()`` / ``warnings()`` / ``print()`` で和文表示）。

    Attributes
    ----------
    balance : pd.DataFrame
        各属性の水準出現頻度と変動係数（CV）。
        列: 水準数, 最大出現, 最小出現, CV, 評価
    chi2 : pd.DataFrame
        属性ペアごとのχ²統計量と自由度（独立性の診断）。
        列: 属性1, 属性2, χ², 自由度, χ²/自由度, 評価
    overlap : pd.DataFrame
        属性ごとのセット内オーバーラップ率
        （全代替案が同じ水準を持つ設問の割合）。
        列: オーバーラップ率, 評価
    diagnostics : List[Diagnostic]
        検出された問題の一覧。
    """

    balance: pd.DataFrame
    chi2: pd.DataFrame
    overlap: pd.DataFrame
    diagnostics: List[Diagnostic]

    def summary(self) -> str:
        """診断結果を人間が読みやすい形式で返す。"""
        lines = ["=" * 55, "選択セット設計チェック", "=" * 55]

        lines.append("\n【水準バランス】（CV が小さいほど均等）")
        lines.append(_df_to_string_cjk(self.balance, index=True))

        lines.append("\n【独立性（χ²統計量）】（自由度に対して小さいほど独立）")
        lines.append(_df_to_string_cjk(self.chi2, index=False))

        lines.append(
            "\n【セット内オーバーラップ】"
            "（全代替案が同じ水準になる設問の割合。小さいほど良い）"
        )
        lines.append(_df_to_string_cjk(self.overlap, index=True))

        diags = sorted(
            self.diagnostics, key=lambda d: SEVERITY_ORDER.get(d.severity, 99)
        )
        if diags:
            lines.append("\n【警告】")
            for d in diags:
                lines.append(f"  [{d.severity}] {d.message}")
                lines.append(f"      → {d.recommendation}")
        else:
            lines.append("\n警告はありません。")

        lines.append("=" * 55)
        return "\n".join(lines)

    def warnings(self) -> pd.DataFrame:
        """警告一覧を DataFrame で返す。"""
        if not self.diagnostics:
            return pd.DataFrame(
                columns=["severity", "category", "message", "recommendation"]
            )
        return pd.DataFrame(
            [
                {
                    "severity": d.severity,
                    "category": d.category,
                    "message": d.message,
                    "recommendation": d.recommendation,
                }
                for d in sorted(
                    self.diagnostics, key=lambda d: SEVERITY_ORDER.get(d.severity, 99)
                )
            ]
        )

    def __repr__(self) -> str:
        return self.summary()


def check_design(
    design: pd.DataFrame,
    *,
    attributes: Optional[List[str]] = None,
) -> ChoiceDesignCheckResult:
    """
    選択セット設計の品質をアンケート実施前に診断する。

    :func:`design_choice_sets` の出力（long形式）を渡すことを想定。
    以下の3点を診断する：

    1. **水準バランス** … 各属性の各水準が均等に出現しているか（CV）。
    2. **属性間の独立性** … 2属性の水準組み合わせが偏っていないか（χ²）。
    3. **セット内オーバーラップ** … 同一設問内の全代替案が同じ水準を
       持ってしまう設問の割合。オーバーラップが多い属性は、その設問では
       比較情報を生まない（どれを選んでも同じ水準なので）。

    Parameters
    ----------
    design : pd.DataFrame
        :func:`design_choice_sets` の出力形式のDataFrame
        （``version``, ``choice_set_id``, ``alt_id`` + 属性列）。

    attributes : list of str, optional
        チェック対象の属性名のリスト。
        省略時は ``version`` / ``choice_set_id`` / ``alt_id`` を除く全列を対象とする。

    Returns
    -------
    ChoiceDesignCheckResult

    Notes
    -----
    χ² 統計量の p 値は計算しない（rating 版の ``check_design`` と同じ方針）。
    「χ²/自由度」の比率と定性的な評価記号（◎○△）で判断する。
    """
    if not isinstance(design, pd.DataFrame):
        raise TypeError(
            f"design は pandas.DataFrame を指定してください。\n"
            f"  受け取った型: {type(design).__name__}"
        )
    id_cols = ["version", "choice_set_id", "alt_id"]
    # 既定では ID 列と pandas の行番号列（Unnamed: 0 など）を除いた全列を
    # 属性とみなす（attributes を明示指定した場合はそのまま尊重する）
    attrs = attributes or [
        c
        for c in design.columns
        if c not in id_cols and not _is_index_artifact_column(c)
    ]
    missing = [a for a in attrs if a not in design.columns]
    if missing:
        raise ValueError(
            f"以下の属性が design に存在しません: {missing}\n"
            f"  存在する列: {list(design.columns)}"
        )
    if not attrs:
        raise ValueError(
            "チェック対象の属性列が見つかりません。\n"
            "  design_choice_sets() の出力をそのまま渡すか、\n"
            "  attributes 引数で属性列を指定してください。"
        )

    balance_df = _check_balance(design, attrs)
    chi2_df = _check_chi2(design, attrs)
    overlap_df = _check_overlap(design, attrs)
    diags = _design_diagnostics(design, attrs, balance_df, chi2_df, overlap_df)

    return ChoiceDesignCheckResult(
        balance=balance_df,
        chi2=chi2_df,
        overlap=overlap_df,
        diagnostics=diags,
    )


# ---------------------------------------------------------------------------
# 公開API: suggest_n_respondents 関数
# ---------------------------------------------------------------------------


def suggest_n_respondents(
    attributes: Dict[str, List[Any]],
    *,
    n_sets: int,
    n_alts: int,
) -> pd.DataFrame:
    """
    CBC調査に必要な **回答者数の目安** を Johnson-Orme の経験則で計算する。

    経験則（Johnson & Orme）::

        n ≥ 500 × c / (t × a)

    * ``n`` … 回答者数
    * ``c`` … 最大水準数（全属性のうち最も水準数が多い属性の水準数）
    * ``t`` … 設問数（1人の回答者が答える選択セット数）
    * ``a`` … 1設問あたりの代替案数

    「各水準が（選ばれる機会として）少なくとも500回提示される」ことを
    目安とするルール。主効果のみのモデルを前提とする。

    Parameters
    ----------
    attributes : dict
        ``design_choice_sets()`` と同じ形式の辞書。

    n_sets : int
        1人の回答者が答える設問数（選択セット数）t。

    n_alts : int
        1設問あたりの代替案数 a。

    Returns
    -------
    pd.DataFrame
        属性ごとの必要回答者数の内訳。
        列：``"水準数"``, ``"必要回答者数（目安）"``。
        インデックスは属性名。

        ``df.attrs["n_required"]``  — 全体で必要な回答者数
        （= 最大水準数 c に基づく値。これを満たせば全属性で条件を満たす）

        ``df.attrs["c_max"]``       — 最大水準数 c

        ``df.attrs["n_sets"]``      — 設問数 t

        ``df.attrs["n_alts"]``      — 代替案数 a

    Raises
    ------
    ValueError
        ``attributes`` が空または辞書でない場合。
        いずれかの属性の水準数が 2 未満、または水準リストに重複がある場合。
        ``n_sets`` または ``n_alts`` が範囲外の場合。

    Notes
    -----
    **Johnson-Orme の経験則について**

    Sawtooth Software の創設者 Rich Johnson と Bryan Orme が提案した
    実務上の経験則で、CBC のサンプルサイズ設計で広く使われる。
    「最低限」の目安であり、サブグループ別の分析（男女別など）を
    行う場合はグループごとにこの人数が必要になる。

    Examples
    --------
    >>> pcc.suggest_n_respondents(
    ...     {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]},
    ...     n_sets=8, n_alts=3,
    ... )
    """
    # ---------- 入力バリデーション ----------
    if not isinstance(attributes, dict) or len(attributes) == 0:
        raise ValueError(
            "attributes は空でない辞書を指定してください。\n"
            "  例: {'price': [100, 150, 200], 'brand': ['A社', 'B社']}"
        )
    for attr, lvs in attributes.items():
        if len(lvs) < 2:
            raise ValueError(
                f"属性 '{attr}' の水準数は 2 以上にしてください（現在: {len(lvs)}）。"
            )
        # 重複した水準は最大水準数 c の計算を狂わせるため弾く
        if len(set(lvs)) != len(lvs):
            raise ValueError(
                f"属性 '{attr}' の水準リストに重複があります: {list(lvs)}\n"
                "  水準は重複なく指定してください。"
            )
    if n_sets < 1:
        raise ValueError(
            f"n_sets は 1 以上の整数を指定してください（指定値: {n_sets}）。"
        )
    if n_alts < 2:
        raise ValueError(
            f"n_alts は 2 以上の整数を指定してください（指定値: {n_alts}）。"
        )

    rows = []
    for attr, lvs in attributes.items():
        c = len(lvs)
        n_req = math.ceil(500 * c / (n_sets * n_alts))
        rows.append(
            {
                "属性": attr,
                "水準数": c,
                "必要回答者数（目安）": n_req,
            }
        )
    result = pd.DataFrame(rows).set_index("属性")

    c_max = int(result["水準数"].max())
    n_required = int(result["必要回答者数（目安）"].max())
    result.attrs.update(
        {
            "n_required": n_required,
            "c_max": c_max,
            "n_sets": n_sets,
            "n_alts": n_alts,
        }
    )

    print(
        f"Johnson-Orme の経験則: n ≥ 500 × c / (t × a)\n"
        f"  最大水準数 c = {c_max}, 設問数 t = {n_sets}, 代替案数 a = {n_alts}\n"
        f"  → 必要回答者数の目安: {n_required} 人以上\n"
        "  （サブグループ別に分析する場合はグループごとにこの人数が必要です）"
    )
    return result


# ---------------------------------------------------------------------------
# 内部ヘルパー（auto_balance）
# ---------------------------------------------------------------------------


def _assign_random_design(
    full: pd.DataFrame,
    N: int,
    n_versions: int,
    n_sets: int,
    n_alts: int,
    rng: "np.random.Generator",
) -> pd.DataFrame:
    """完全交差 ``full`` から1つのランダム設計（セット内重複なし）を生成する。

    ``design_choice_sets`` の従来の生成ロジックをそのまま切り出したもの。
    auto_balance の候補生成でも同じロジックを使う。
    """
    frames = []
    for ver in range(1, n_versions + 1):
        for s in range(1, n_sets + 1):
            idx = rng.choice(N, size=n_alts, replace=False)
            block = full.iloc[idx].copy()
            block.insert(0, "version", ver)
            block.insert(1, "choice_set_id", s)
            block.insert(2, "alt_id", range(1, n_alts + 1))
            frames.append(block)
    return pd.concat(frames, ignore_index=True)


def _is_better_design(
    best: Optional[tuple],
    n_warn: int,
    cv_sum: float,
) -> bool:
    """方式D の優先順位で、新候補が現在の最良より良いかを判定する。

    優先順位（小さいほど良い）::

        (警告ゼロなら0・そうでなければ1, 警告数, CV合計)

    * 警告ゼロの候補は、警告のある候補より常に優先される。
    * 警告ゼロが複数あれば CV 合計が小さいものを選ぶ。
    * 警告ゼロが無ければ、警告数が少ない→CV合計が小さい順に選ぶ。

    厳密に「より良い」ときだけ True を返すので、同点では先に評価した候補が
    残る（候補生成順は seed から決定的なので、選択結果も決定的になる）。
    """
    key = (0 if n_warn == 0 else 1, n_warn, cv_sum)
    if best is None:
        return True
    best_key = (0 if best[1] == 0 else 1, best[1], best[2])
    return key < best_key


# ---------------------------------------------------------------------------
# 内部ヘルパー（rating/design.py・rating/analysis.py から流用）
# ---------------------------------------------------------------------------


def _check_balance(design: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """各属性の水準出現頻度と変動係数（CV）を計算する。"""
    rows = []
    for attr in attrs:
        counts = design[attr].value_counts()
        mean = counts.mean()
        cv = float(counts.std() / mean) if mean > 0 else float("inf")
        if cv < 0.05:
            label = "◎"
        elif cv < 0.15:
            label = "○"
        else:
            label = "△"
        rows.append(
            {
                "属性": attr,
                "水準数": len(counts),
                "最大出現": int(counts.max()),
                "最小出現": int(counts.min()),
                "CV": round(cv, 4),
                "評価": label,
            }
        )
    return pd.DataFrame(rows).set_index("属性")


def _check_chi2(design: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """
    属性ペアのχ²統計量と自由度を計算する（scipy不要）。
    p値の代わりに χ²/自由度 の比率と定性的評価を返す。
    """
    rows = []
    for i, a1 in enumerate(attrs):
        for a2 in attrs[i + 1 :]:
            ct = pd.crosstab(design[a1], design[a2]).values.astype(float)
            row_sum = ct.sum(axis=1, keepdims=True)
            col_sum = ct.sum(axis=0, keepdims=True)
            n = ct.sum()
            if n == 0:
                continue
            expected = row_sum @ col_sum / n
            with np.errstate(divide="ignore", invalid="ignore"):
                chi2_val = float(
                    np.where(expected > 0, (ct - expected) ** 2 / expected, 0).sum()
                )
            dof = (ct.shape[0] - 1) * (ct.shape[1] - 1)
            ratio = chi2_val / dof if dof > 0 else float("inf")
            if ratio < 0.1:
                label = "◎"
            elif ratio < 1.0:
                label = "○"
            else:
                label = "△"
            rows.append(
                {
                    "属性1": a1,
                    "属性2": a2,
                    "χ²": round(chi2_val, 4),
                    "自由度": dof,
                    "χ²/自由度": round(ratio, 4),
                    "評価": label,
                }
            )
    return pd.DataFrame(rows)


def _check_overlap(design: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """
    属性ごとのセット内オーバーラップ率を計算する。

    オーバーラップ率 = 「同一選択セット内の全代替案が同じ水準を持つ設問」の割合。
    その設問では当該属性が選択の判断材料にならない（情報を生まない）ため、
    小さいほど良い。
    """
    group_keys = [c for c in ("version", "choice_set_id") if c in design.columns]
    if not group_keys:
        # choice_set_id 等がない場合は全体を1セットとみなす（実用上は起こらない想定）
        group_keys = [design.index]

    nunique = design.groupby(group_keys, sort=False)[attrs].nunique()
    rows = []
    for attr in attrs:
        rate = float((nunique[attr] == 1).mean())
        if rate < 0.1:
            label = "◎"
        elif rate < 0.3:
            label = "○"
        else:
            label = "△"
        rows.append(
            {
                "属性": attr,
                "オーバーラップ率": round(rate, 4),
                "評価": label,
            }
        )
    return pd.DataFrame(rows).set_index("属性")


def _design_diagnostics(
    design: pd.DataFrame,
    attrs: List[str],
    balance_df: pd.DataFrame,
    chi2_df: pd.DataFrame,
    overlap_df: pd.DataFrame,
) -> List[Diagnostic]:
    """設計チェックの結果から Diagnostic のリストを生成する。"""
    diags: List[Diagnostic] = []

    # バランスチェック（rating 版と同じ閾値）
    for attr, row in balance_df.iterrows():
        cv = row["CV"]
        if cv > 0.3:
            diags.append(
                Diagnostic(
                    severity="大",
                    category=f"balance_{attr}",
                    message=f"属性 '{attr}' の水準出現頻度が偏っています（CV={cv:.3f}）。",
                    recommendation=(
                        "設問数（n_sets）またはバージョン数（n_versions）を増やすか、"
                        "seed を変えて再生成してください。"
                    ),
                )
            )
        elif cv > 0.15:
            diags.append(
                Diagnostic(
                    severity="中",
                    category=f"balance_{attr}",
                    message=f"属性 '{attr}' の水準出現頻度にやや偏りがあります（CV={cv:.3f}）。",
                    recommendation="可能であれば設問数・バージョン数を増やしてください。",
                )
            )

    # χ²チェック（rating 版と同じ閾値）
    for _, row in chi2_df.iterrows():
        ratio = row["χ²/自由度"]
        a1, a2 = row["属性1"], row["属性2"]
        if ratio > 1.0:
            diags.append(
                Diagnostic(
                    severity="中",
                    category=f"chi2_{a1}_{a2}",
                    message=(
                        f"'{a1}' と '{a2}' のχ²/自由度={ratio:.3f} > 1.0 で、"
                        "独立性が低い可能性があります。"
                    ),
                    recommendation="2属性の水準組み合わせが偏っていないか確認してください。",
                )
            )

    # オーバーラップチェック
    for attr, row in overlap_df.iterrows():
        rate = row["オーバーラップ率"]
        if rate > 0.5:
            diags.append(
                Diagnostic(
                    severity="大",
                    category=f"overlap_{attr}",
                    message=(
                        f"属性 '{attr}' は設問の {rate:.0%} で全代替案が同じ水準に"
                        "なっています。これらの設問では当該属性が選択の判断材料に"
                        "ならず、推定精度が大きく下がります。"
                    ),
                    recommendation=(
                        "n_alts を減らす、水準数を増やす、または seed を変えて"
                        "再生成し、オーバーラップ率を下げてください。"
                    ),
                )
            )
        elif rate > 0.3:
            diags.append(
                Diagnostic(
                    severity="中",
                    category=f"overlap_{attr}",
                    message=(
                        f"属性 '{attr}' のセット内オーバーラップ率がやや高いです"
                        f"（{rate:.0%}）。"
                    ),
                    recommendation="設問数を増やすか seed を変えて再生成を検討してください。",
                )
            )

    return diags
