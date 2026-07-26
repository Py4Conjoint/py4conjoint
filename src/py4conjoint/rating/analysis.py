"""
analysis.py
============
コンジョイント分析の **回帰実行と結果の保持** を担当するモジュール。

中心となるのは :func:`fit` 関数と :class:`ConjointResult` クラス。

>>> import py4conjoint as pc
>>> df_coded = pc.encode(df, reference_levels={...})
>>> result = pc.fit(df_coded)
>>> result.summary()        # 和文サマリー
>>> result.importance()     # 重要度
>>> result.wtp()            # WTP（限界支払意思額）
>>> result.market_share(products)
>>> pc.check_design(profiles)  # アンケート実施前の直交性チェック
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Union

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from statsmodels.regression.linear_model import RegressionResults

# ---------------------------------------------------------------------------
# pandas の行番号列（index 付き CSV 保存の痕跡）の判定
# （rating / choice で共通利用。choice 側は本モジュールから import する）
# ---------------------------------------------------------------------------

# pandas が index 付きで保存された CSV を読み込んだときに付ける列名
# （to_csv() を index=False なしで保存 → 行番号が "Unnamed: 0" 列になる）
_UNNAMED_COL_RE = re.compile(r"^Unnamed: \d+$")


def _is_index_artifact_column(name: object) -> bool:
    """pandas の行番号列（``Unnamed: 0`` など、index 付き CSV 保存の痕跡）か判定する。

    この形式の列名は pandas が「名前のない列」に機械的に付けるもので、
    設計の属性として現れることは実質ない。設計の中身（属性・水準）では
    ないため、署名・診断・属性復元の対象から除外する。
    """
    return bool(_UNNAMED_COL_RE.match(str(name)))


# ---------------------------------------------------------------------------
# 警告（落とし穴）の構造化表現
# ---------------------------------------------------------------------------

# 重大度ラベル。表示順を決めるために順序を定義する。
SEVERITY_ORDER = {"大": 0, "中": 1, "小": 2}


@dataclass(frozen=True)
class Diagnostic:
    """
    回帰分析で検出された **落とし穴（warning）** の1件を表す不変オブジェクト。

    Attributes
    ----------
    severity : str
        重大度。``"大"`` / ``"中"`` / ``"小"`` のいずれか。
        ``"大"`` のものは :meth:`ConjointResult.summary` に表示される。
    category : str
        警告の種類を識別する英字キー（例：``"r2_low"``, ``"price_insignificant"``,
        ``"wtp_extrapolation"``, ``"price_sign_negative"``, ``"few_respondents"``）。
        プログラム的なフィルタや再帰的な処理に使う。
    message : str
        警告本文（日本語）。何が起きているかの説明。
    recommendation : str
        対処方法の提案（日本語）。
    """

    severity: str
    category: str
    message: str
    recommendation: str

    def to_str(self) -> str:
        """1行の人間可読な文字列に変換する。"""
        return f"[{self.severity}] {self.message} → {self.recommendation}"


# ---------------------------------------------------------------------------
# デザインチェック結果オブジェクト
# ---------------------------------------------------------------------------


@dataclass
class DesignCheckResult:
    """
    check_design() の診断結果を保持するオブジェクト。

    Attributes
    ----------
    balance : pd.DataFrame
        各属性の水準出現頻度と変動係数（CV）。
        列: 水準数, 最大出現, 最小出現, CV, 評価
    correlation : pd.DataFrame
        効果コーディング後の属性間相関行列。
    chi2 : pd.DataFrame
        属性ペアごとのχ²統計量と自由度。
        列: 属性1, 属性2, χ², 自由度, χ²/自由度, 評価
    diagnostics : List[Diagnostic]
        検出された問題の一覧。
    """

    balance: pd.DataFrame
    correlation: pd.DataFrame
    chi2: pd.DataFrame
    diagnostics: List[Diagnostic]

    def summary(self) -> str:
        """診断結果を人間が読みやすい形式で返す。"""
        lines = ["=" * 55, "デザイン直交性チェック", "=" * 55]

        lines.append("\n【水準バランス】（CV が小さいほど均等）")
        lines.append(_df_to_string_cjk(self.balance, index=True))

        lines.append("\n【属性間相関】（0 に近いほど直交）")
        lines.append(self.correlation.to_string())

        lines.append("\n【独立性（χ²統計量）】（自由度に対して小さいほど独立）")
        lines.append(_df_to_string_cjk(self.chi2, index=False))

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


# ---------------------------------------------------------------------------
# 公開API: fit関数
# ---------------------------------------------------------------------------


def fit(
    df: pd.DataFrame,
    *,
    rating: str = "rating",
    encoded_columns: Optional[List[str]] = None,
    reference_levels: Optional[Dict[str, object]] = None,
    price_col: str = "price",
    formula: Optional[str] = None,
    respondent_id_col: str = "respondent_id",
    cluster_se: bool = True,
) -> "ConjointResult":
    """
    コンジョイント分析の回帰モデルを推定し、結果オブジェクトを返す。

    最も簡単な呼び出し方
    --------------------
    符号化済みのDataFrameと評点列だけ渡せば、自動で符号化列を検出して回帰する。

    >>> df_coded = pc.encode(df, reference_levels={"price": 10, ...})
    >>> result = pc.fit(df_coded)

    Parameters
    ----------
    df : pd.DataFrame
        :func:`encode` で符号化済みのDataFrame。

    rating : str, default "rating"
        被説明変数（評点）の列名。
        ``formula`` 指定時は無視され、formula の左辺が使われる。

    encoded_columns : list of str, optional
        説明変数として使う符号化列のリスト。
        省略時は ``reference_levels`` から自動推定するか、
        効果コーディング済みの列（``-1`` と ``1`` を含み、値が
        ``-1, 0, 1`` に収まる数値列）を自動検出する。
        回答者属性などの ``0/1`` 列は自動検出に含まれないため、
        説明変数に加えたい場合はこの引数で明示的に指定する。
        ``formula`` 指定時は無視される。

        .. note::
            回答者属性の列（``respondent_encode`` で作った ``gender_male``
            などの 0/1 列）をここに含めると、その列は :meth:`ConjointResult.importance`
            や :meth:`ConjointResult.wtp` の出力にも1つの「属性」として現れる。
            回答者属性は製品の属性ではないため、その行の重要度・WTP は
            製品属性と同じようには解釈できない点に注意すること。

    reference_levels : dict, optional
        :func:`encode` に渡したのと同じ辞書。
        指定すると、対応する符号化列を確実に取得できる。
        WTP計算時に「価格の元の水準」を復元するためにも使う。

    price_col : str, default "price"
        価格列の名前。WTP計算で使う。
        授業の標準は ``"price"`` だが、別の列名（例：``"値段"``）でも上書き可。

    formula : str, optional
        statsmodels 用の回帰式を直接指定したい場合に使う。
        例：``"rating ~ price_low + os_apple + camera_high"``
        指定した場合、被説明変数は formula の左辺から、説明変数は
        右辺から自動取得される（``rating``・``encoded_columns`` 引数は
        無視される）。通常は省略してよい（自動構築される）。

    respondent_id_col : str, default "respondent_id"
        回答者ID列の名前（:func:`forms_to_data` のデフォルト列名と同じ）。
        クラスタロバスト標準誤差のグループ化と、回答者数の診断
        （``few_respondents``）に使う。

    cluster_se : bool, default True
        ``True`` かつ回答者ID列が存在し回答者が2人以上いる場合、
        回答者IDでグループ化した **クラスタロバスト標準誤差** を使う。
        同じ回答者の複数回答は独立でないため、通常のOLS標準誤差では
        p値が過小（有意に出やすすぎ）になる。
        係数の推定値自体はどちらでも変わらない。
        ``False`` にすると通常のOLS標準誤差を使う。

    Returns
    -------
    ConjointResult
        推定結果と各種解釈メソッドを持つオブジェクト。

    Raises
    ------
    ValueError
        評点列が ``df`` にない場合。
        説明変数が1つも見つからない場合。

    Notes
    -----
    自動的に以下の **落とし穴チェック** を行い、警告を出す：

    * 価格係数が負（＝価格が上がるほど評点が高い）の場合：
      データ品質または符号化方向の誤りを示唆。

    * R² が極端に低い（< 0.20）場合：仮定の妥当性に疑問。

    * 観測数 ÷ 説明変数数が低い場合（重大度：5倍未満=大、10倍未満=中）：
      推定が不安定になる。:func:`suggest_n_profiles` の ``obs_per_predictor`` と
      対応する閾値。

    * 回答者が1人の場合（重大度：大）または2〜4人の場合（重大度：中）：
      個人・少数の好みを集団の傾向と誤解する危険。
      回答者ID列（``respondent_id_col``）がある場合のみ判定する。

    * 回答者ID列が見つからない場合（重大度：中、``independence_assumed``）：
      観測の独立性を仮定した通常の標準誤差を使用するため、
      同一回答者の複数回答が含まれているとp値が過小になる恐れ。
    """
    # ---------- 入力チェック ----------
    if formula is not None:
        # formula 指定時は被説明変数を formula の左辺から取得する
        rating = formula.split("~")[0].strip()
    if rating not in df.columns:
        raise ValueError(
            f"評点列 '{rating}' が DataFrame にありません。\n"
            f"  存在する列: {list(df.columns)}\n"
            f"  rating 引数で正しい列名を指定してください。"
        )

    # ---------- encode() からのメタ情報を取得 ----------
    # encode() が df.attrs["py4conjoint"]["reference_levels"] を残してくれるので、
    # 引数で明示されていない場合はそちらを使う。
    if reference_levels is None:
        meta = df.attrs.get("py4conjoint", {}) if hasattr(df, "attrs") else {}
        reference_levels = meta.get("reference_levels")

    n_before = len(df)
    alias_map: Dict[str, str] = {}

    # 回答者IDでクラスタリングするか（回答者が2人以上いる場合のみ）
    use_cluster = (
        cluster_se
        and respondent_id_col in df.columns
        and df[respondent_id_col].nunique() >= 2
    )

    if formula is not None:
        # ---------- formula 指定時 ----------
        # 説明変数は formula の右辺（推定後の exog 名）から取得する。
        # 自動検出と食い違うと importance()/wtp() が KeyError になるため、
        # formula を唯一の情報源とし、encoded_columns 引数は使わない。
        model = smf.ols(formula, data=df)
        if use_cluster:
            # patsy が NaN 行を落とすため、残った行に対応するグループを渡す
            groups = df.loc[model.data.row_labels, respondent_id_col]
            res: RegressionResults = model.fit(
                cov_type="cluster", cov_kwds={"groups": groups}
            )
        else:
            res = model.fit()
        encoded_columns = [n for n in res.model.exog_names if n != "Intercept"]
        n_dropped = n_before - int(res.nobs)
    else:
        # ---------- 説明変数の決定 ----------
        if encoded_columns is None:
            encoded_columns = _detect_encoded_columns(
                df, rating=rating, reference_levels=reference_levels
            )
        if len(encoded_columns) == 0:
            raise ValueError(
                "符号化済みの説明変数が見つかりませんでした。\n"
                "  encode() で符号化を済ませているか確認してください。\n"
                "  または encoded_columns 引数で明示的に列を指定してください。"
            )

        # 列の存在確認
        missing = [c for c in encoded_columns if c not in df.columns]
        if missing:
            raise ValueError(
                f"指定された符号化列が DataFrame にありません: {missing}\n"
                f"  存在する列: {list(df.columns)}"
            )

        # ---------- 回帰実行 ----------
        # NaN行を落としつつ、何件落としたかを記録（落とし穴チェック用）
        use_cols = [rating] + list(encoded_columns)
        n_dropped = n_before - len(df[use_cols].dropna())

        # 日本語など formula で扱いにくい列名を、内部的に英数字エイリアスへ
        # 一時リネームして回帰し、推定後に係数名を元に戻す。
        # こうすることで `result.params["camera_高性能"]` のように
        # 元の列名でアクセスできるようになる。
        alias_map = {c: f"__pc_var_{i}__" for i, c in enumerate(encoded_columns)}
        rev_map = {v: k for k, v in alias_map.items()}
        rating_alias = "__pc_rating__"
        rev_map[rating_alias] = rating

        # クラスタリングに使う回答者ID列も保持しておく（formula は参照しない）
        model_cols = list(use_cols)
        if use_cluster and respondent_id_col not in model_cols:
            model_cols.append(respondent_id_col)
        df_alias = df[model_cols].rename(columns={**alias_map, rating: rating_alias})
        formula_alias = f"{rating_alias} ~ " + " + ".join(
            alias_map[c] for c in encoded_columns
        )
        model = smf.ols(formula_alias, data=df_alias)
        if use_cluster:
            groups = df_alias.loc[model.data.row_labels, respondent_id_col]
            res = model.fit(cov_type="cluster", cov_kwds={"groups": groups})
        else:
            res = model.fit()
        # 係数名を元に戻す
        res = _rename_result_index(res, rev_map)

    result = ConjointResult(
        ols=res,
        df=df,
        rating=rating,
        encoded_columns=list(encoded_columns),
        reference_levels=reference_levels or {},
        price_col=price_col,
        n_dropped=n_dropped,
        alias_map=alias_map,
        respondent_id_col=respondent_id_col,
        se_type="cluster" if use_cluster else "nonrobust",
    )

    # ---------- 落とし穴の自動検出 ----------
    result._run_diagnostics()

    return result


# ---------------------------------------------------------------------------
# 公開API: check_design関数
# ---------------------------------------------------------------------------


def check_design(
    profiles: pd.DataFrame,
    *,
    attributes: Optional[List[str]] = None,
) -> DesignCheckResult:
    """
    アンケート実施前にプロファイルの直交性を診断する。

    fit() の前、forms_to_data() の前に呼ぶことを想定。
    scipy は不要（numpy・pandas のみで計算）。

    Parameters
    ----------
    profiles : pd.DataFrame
        属性と水準を持つプロファイルのDataFrame。
        例::

            profiles = pd.DataFrame({
                "price":  [6, 10, 6, 10],
                "os":     ["android", "apple", "apple", "android"],
                "camera": ["標準", "標準", "高性能", "高性能"],
            }, index=["P1", "P2", "P3", "P4"])
            pc.check_design(profiles)

    attributes : list of str, optional
        チェック対象の属性名のリスト。省略時は profiles の全列を対象とする。

    Returns
    -------
    DesignCheckResult

    Notes
    -----
    χ² 統計量の p 値は計算しない（scipy 不要とするため）。
    代わりに「χ² / 自由度」の比率と定性的な評価記号（◎○△）で判断する。
    p 値が必要な場合は ``from scipy.stats import chi2_contingency`` を使うこと。
    """
    if not isinstance(profiles, pd.DataFrame):
        raise TypeError(
            f"profiles は pandas.DataFrame を指定してください。\n"
            f"  受け取った型: {type(profiles).__name__}"
        )
    # 既定では pandas の行番号列（Unnamed: 0 など。index=False を付けずに
    # 保存した CSV の痕跡）を除いた全列を属性とみなす。行番号列を属性として
    # 診断すると、パラメータ数が架空に膨らんで insufficient_profiles などの
    # 誤警告が出る。attributes を明示指定した場合はそのまま尊重する。
    attrs = attributes or [
        c for c in profiles.columns if not _is_index_artifact_column(c)
    ]
    missing = [a for a in attrs if a not in profiles.columns]
    if missing:
        raise ValueError(
            f"以下の属性が profiles に存在しません: {missing}\n"
            f"  存在する列: {list(profiles.columns)}"
        )

    balance_df = _check_balance(profiles, attrs)
    corr_df = _check_correlation(profiles, attrs)
    chi2_df = _check_chi2(profiles, attrs)
    diags = _design_diagnostics(profiles, attrs, balance_df, corr_df, chi2_df)

    return DesignCheckResult(
        balance=balance_df,
        correlation=corr_df,
        chi2=chi2_df,
        diagnostics=diags,
    )


# ---------------------------------------------------------------------------
# 結果オブジェクト
# ---------------------------------------------------------------------------


@dataclass
class ConjointResult:
    """
    コンジョイント回帰の推定結果を保持し、解釈メソッドを提供するクラス。

    通常は :func:`fit` 関数経由で生成され、ユーザーが直接インスタンス化する
    ことはない。

    Attributes
    ----------
    ols : statsmodels.regression.linear_model.RegressionResults
        statsmodels の元の結果オブジェクト。
        詳細な統計量を見たい場合は ``result.ols.summary()`` で取得可能。

    df : pd.DataFrame
        分析に使ったデータ。

    rating : str
        評点列の名前。

    encoded_columns : list of str
        説明変数（符号化列）のリスト。

    reference_levels : dict
        ``encode()`` に渡された基準水準の辞書。

    price_col : str
        価格列の名前。WTP計算で使う。

    n_dropped : int
        欠損により分析から除外された行数。

    alias_map : dict, optional
        日本語列名等を内部用エイリアスにリネームしたマップ。
        ``predict`` を呼ぶときに ``products`` をこのマップで変換する。

    respondent_id_col : str
        回答者ID列の名前。クラスタロバスト標準誤差と回答者数診断に使う。

    se_type : str
        標準誤差の種類。``"cluster"``（回答者IDによるクラスタロバスト）
        または ``"nonrobust"``（通常のOLS標準誤差）。
    """

    ols: RegressionResults
    df: pd.DataFrame
    rating: str
    encoded_columns: List[str]
    reference_levels: Dict[str, object]
    price_col: str
    n_dropped: int = 0
    alias_map: Dict[str, str] = field(default_factory=dict)
    respondent_id_col: str = "respondent_id"
    se_type: str = "nonrobust"

    # 内部用：検出された警告（落とし穴）のリスト
    # Diagnostic オブジェクトとして保持する。
    _diagnostics: List[Diagnostic] = field(default_factory=list)
    # 内部用：wtp() の重複登録防止用キーセット（category とは別に per-attribute で管理）
    _warned_keys: Set[str] = field(default_factory=set)

    # ---- 基本情報の取得 ----------------------------------------------------

    @property
    def params(self) -> pd.Series:
        """推定された係数（切片含む）。``result.params`` で取得可能。"""
        return self.ols.params

    @property
    def rsquared(self) -> float:
        """決定係数 R²。"""
        return float(self.ols.rsquared)

    @property
    def n_obs(self) -> int:
        """分析に使った観測数。"""
        return int(self.ols.nobs)

    @property
    def intercept(self) -> float:
        """切片 b0（全水準平均の効用）。"""
        return float(self.params.get("Intercept", np.nan))

    # ---- サマリー ---------------------------------------------------------

    def summary(self, *, slim: bool = True) -> str:
        """
        和文サマリーを返す（``print()`` で表示）。

        Parameters
        ----------
        slim : bool, default True
            ``True`` でコンパクトな和文サマリーを表示。
            ``False`` で statsmodels の詳細な統計表（英語）を表示。

        Returns
        -------
        str
            人間が読みやすい和文サマリー（slim=True）または
            statsmodels の詳細な統計表（slim=False）。

        Examples
        --------
        >>> print(result.summary())
        >>> print(result.summary(slim=False))  # statsmodels 詳細表示
        """
        if not slim:
            return str(self.ols.summary())

        lines: List[str] = []
        lines.append("=" * 60)
        lines.append("コンジョイント分析の結果（和文サマリー）")
        lines.append("=" * 60)
        se_label = (
            f"クラスタロバスト（{self.respondent_id_col}）"
            if self.se_type == "cluster"
            else "通常（観測の独立性を仮定）"
        )
        stat_rows = [
            ("観測数", str(self.n_obs)),
            ("説明変数の数", str(len(self.encoded_columns))),
            ("決定係数 R²", f"{self.rsquared:.4f}"),
            ("自由度修正 R²", f"{self.ols.rsquared_adj:.4f}"),
            ("標準誤差", se_label),
        ]
        if self.n_dropped > 0:
            stat_rows.append(("欠損で除外", f"{self.n_dropped} 行"))
        max_label_w = max(_display_width(lbl) for lbl, _ in stat_rows)
        for label, value in stat_rows:
            lines.append(f"{_ljust_display(label, max_label_w + 1)}: {value}")
        lines.append("")

        # 係数表
        NAME_WIDTH = 25
        STAR_WIDTH = 6  # 「有意性」の表示幅（3文字×2）
        lines.append("【推定された係数（部分効用 part-worth）】")
        params = self.params
        pvals = self.ols.pvalues
        lines.append(
            "  "
            + _ljust_display("変数", NAME_WIDTH)
            + " "
            + _rjust_display("係数", 10)
            + " "
            + _rjust_display("p値", 10)
            + "  "
            + "有意性"
        )
        lines.append(f"  {'-' * NAME_WIDTH} {'-' * 10} {'-' * 10}  {'-' * STAR_WIDTH}")
        for name in params.index:
            coef = params[name]
            p = pvals[name]
            star = _significance_stars(p)
            lines.append(
                f"  {_ljust_display(str(name), NAME_WIDTH)} {coef:>10.4f} {p:>10.4f}  {star}"
            )
        lines.append("")
        lines.append("  有意水準: *** p<0.001  ** p<0.01  * p<0.05  . p<0.1")

        # 警告（落とし穴）— 重大度「大」のみここに表示
        major = [d for d in self._diagnostics if d.severity == "大"]
        minor = [d for d in self._diagnostics if d.severity != "大"]
        if major:
            lines.append("")
            lines.append("【⚠️ 重大な注意事項（落とし穴チェック）】")
            for d in major:
                lines.append(f"  ・[{d.severity}] {d.message}")
                lines.append(f"      → {d.recommendation}")
        if minor:
            lines.append("")
            lines.append(
                f"（その他の注意事項が {len(minor)} 件あります。"
                "result.warnings() で確認できます）"
            )

        lines.append("=" * 60)
        text = "\n".join(lines)
        return text

    def __repr__(self) -> str:  # pragma: no cover
        return self.summary()

    def _repr_html_(self) -> str:
        """Jupyter Notebook 向け HTML 表示。
        セルで ``result`` とだけ入力したときに自動的に使われる。
        ``print(result.summary())`` は従来どおりテキスト表示。
        """
        from html import escape

        _td = "padding:3px 14px 3px 4px;"
        _th = _td + "border-bottom:1px solid #888;font-weight:bold;"
        _tbl = "border-collapse:collapse;margin-bottom:0.8em;"

        def td(txt, align="left"):
            return f'<td style="{_td}text-align:{align};">{escape(str(txt))}</td>'

        def th(txt, align="left"):
            return f'<th style="{_th}text-align:{align};">{escape(str(txt))}</th>'

        p = ["<div>", "<p><strong>コンジョイント分析の結果</strong></p>"]

        # 統計量
        se_label = (
            f"クラスタロバスト（{self.respondent_id_col}）"
            if self.se_type == "cluster"
            else "通常（観測の独立性を仮定）"
        )
        stat_rows = [
            ("観測数", str(self.n_obs)),
            ("説明変数の数", str(len(self.encoded_columns))),
            ("決定係数 R²", f"{self.rsquared:.4f}"),
            ("自由度修正 R²", f"{self.ols.rsquared_adj:.4f}"),
            ("標準誤差", se_label),
        ]
        if self.n_dropped > 0:
            stat_rows.append(("欠損で除外", f"{self.n_dropped} 行"))

        p.append(f'<table style="{_tbl}">')
        for lbl, val in stat_rows:
            p.append(f"<tr>{td(lbl)}{td(val)}</tr>")
        p.append("</table>")

        # 係数テーブル
        p.append("<p><strong>【推定された係数（部分効用 part-worth）】</strong></p>")
        p.append(f'<table style="{_tbl}">')
        p.append(
            "<tr>"
            + th("変数")
            + th("係数", "right")
            + th("p値", "right")
            + th("有意性", "right")
            + "</tr>"
        )

        params = self.params
        pvals = self.ols.pvalues
        for name in params.index:
            c = params[name]
            pv = pvals[name]
            star = _significance_stars(pv)
            p.append(
                "<tr>"
                + td(name)
                + td(f"{c:.4f}", "right")
                + td(f"{pv:.4f}", "right")
                + td(star, "right")
                + "</tr>"
            )
        p.append("</table>")
        p.append(
            '<p style="font-size:0.85em;color:#888;">'
            "有意水準: *** p&lt;0.001&nbsp; ** p&lt;0.01&nbsp;"
            " * p&lt;0.05&nbsp; . p&lt;0.1</p>"
        )

        # 重大警告
        major = [d for d in self._diagnostics if d.severity == "大"]
        minor = [d for d in self._diagnostics if d.severity != "大"]
        if major:
            p.append("<p><strong>⚠️ 重大な注意事項</strong></p><ul>")
            for d in major:
                p.append(
                    f"<li><strong>[{escape(d.severity)}]</strong> "
                    f"{escape(d.message)}<br>"
                    f"&nbsp;&nbsp;→ {escape(d.recommendation)}</li>"
                )
            p.append("</ul>")
        if minor:
            p.append(
                f'<p style="font-size:0.9em;color:#888;">'
                f"その他の注意事項が {len(minor)} 件あります。"
                f"result.warnings() で確認できます。</p>"
            )

        p.append("</div>")
        return "\n".join(p)

    # ---- 警告（落とし穴）の取得 -------------------------------------------

    def warnings(
        self,
        *,
        severity: Optional[Union[str, List[str]]] = None,
        category: Optional[Union[str, List[str]]] = None,
        as_dataframe: bool = True,
    ) -> Union[pd.DataFrame, List[Diagnostic]]:
        """
        検出された落とし穴（警告）の一覧を返す。

        :meth:`summary` には重大度「大」のみが表示される。
        重大度「中」「小」を含むすべての警告を確認したい場合はこのメソッドを使う。

        Parameters
        ----------
        severity : str または list of str, optional
            重大度でフィルタする。``"大"``, ``"中"``, ``"小"`` または
            それらのリスト（例：``["大", "中"]``）。
            省略時はすべての警告を返す。
        category : str または list of str, optional
            カテゴリでフィルタする。
            利用可能な値：``"price_sign_negative"``, ``"r2_low"``,
            ``"obs_per_predictor"``, ``"few_respondents"``,
            ``"independence_assumed"``, ``"price_insignificant"``,
            ``"wtp_extrapolation"``, ``"wtp_price_linear_approx"``。
            省略時はすべて返す。
        as_dataframe : bool, default True
            ``True`` なら ``pd.DataFrame``（列：severity, category, message,
            recommendation）として返す。
            ``False`` なら :class:`Diagnostic` オブジェクトのリストを返す。

        Returns
        -------
        pd.DataFrame または list of Diagnostic

        Examples
        --------
        >>> result.warnings()                         # すべての警告
        >>> result.warnings(severity="大")             # 重大度「大」のみ
        >>> result.warnings(severity=["大", "中"])     # 「大」「中」
        >>> result.warnings(category="wtp_extrapolation")
        """
        diags = list(self._diagnostics)

        # severity フィルタ
        if severity is not None:
            sev_list = [severity] if isinstance(severity, str) else list(severity)
            diags = [d for d in diags if d.severity in sev_list]

        # category フィルタ
        if category is not None:
            cat_list = [category] if isinstance(category, str) else list(category)
            diags = [d for d in diags if d.category in cat_list]

        # 重大度順にソート
        diags = sorted(diags, key=lambda d: SEVERITY_ORDER.get(d.severity, 99))

        if not as_dataframe:
            return diags

        if not diags:
            # 空でも列を持つ DataFrame を返す（呼び出し側で扱いやすくするため）
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
                for d in diags
            ]
        )

    # ---- 重要度 -------------------------------------------------------

    def importance(self, *, as_percent: bool = True) -> pd.DataFrame:
        """
        各属性の **重要度（Relative Importance）** を計算する。

        計算方法
        --------
        各属性について、その属性内での **部分効用の最大値 − 最小値**（効用範囲）
        を求め、全属性の効用範囲の合計に占める割合として重要度を定義する。

        2水準の場合、効用範囲は ``2 × |係数|`` に等しい
        （水準が ``-1`` と ``+1`` なので、差は ``2`` 倍）。

        3水準以上の場合は、各水準の部分効用（基準水準は他のダミー和の符号反転）
        の最大値と最小値の差を使う。

        Parameters
        ----------
        as_percent : bool, default True
            ``True`` ならパーセント（合計100）、``False`` なら比率（合計1）で返す。

        Returns
        -------
        pd.DataFrame
            列：``効用範囲``, ``重要度``。
            インデックスは属性名。``重要度`` の合計は100（または1）になる。

        Notes
        -----
        重要度は **調査で選んだ水準のレンジに依存する相対指標** であり、
        属性の本質的な重要性ではない。
        たとえば価格の水準を 6〜10万円 から 6〜20万円 に広げると、
        価格の効用範囲が広がり、価格の重要度は大きく計算される。
        水準レンジの異なる調査間で重要度を比較してはいけない。

        Examples
        --------
        >>> result.importance()
                効用範囲   重要度
        属性
        price      1.234   45.6
        os         0.567   21.0
        camera     0.901   33.4
        """
        attr_ranges = self._attribute_ranges()
        total = sum(attr_ranges.values())
        if total == 0:
            raise RuntimeError(
                "効用範囲の合計が0です。係数がすべて0の異常なケースです。"
            )
        rows = []
        for attr, rng in attr_ranges.items():
            imp = rng / total
            if as_percent:
                imp *= 100.0
            rows.append({"attribute": attr, "効用範囲": rng, "重要度": imp})
        out = pd.DataFrame(rows).set_index("attribute")
        out.index.name = "属性"
        return out

    # ---- WTP（限界支払意思額） ----------------------------------------------

    def wtp(
        self,
        *,
        price_col: Optional[str] = None,
        method: str = "segment",
        price_segment: Optional[object] = None,
    ) -> pd.DataFrame:
        """
        各非価格属性の **WTP（限界支払意思額、Marginal Willingness to Pay）** を計算する。

        定義
        ----
        WTPは「ある属性を基準水準から非基準水準に変えるとき、回答者が
        最大いくらまで追加で支払ってもよいと思うか」を金額で表す。

        厳密には **限界支払意思額（Marginal Willingness to Pay, MWTP）**
        である。他の属性をすべて一定に保ったまま1つの属性だけを変える
        ときの追加支払額（属性と価格の限界代替率）であり、
        **製品全体に対する支払上限額（総WTP・留保価格）ではない** 点に
        注意。選択型コンジョイントの文献で ``MWTP = -β_attr / β_price``
        と定義されるものと同じ構造を、評点型に適用したもの。

        **2水準価格の場合**

        .. code-block:: python

            wtp_price_factor = -(low_price - high_price) / b_price
            wtp = wtp_price_factor * b_attr

        ``low_price - high_price`` は負の値なので、マイナスを付けて正のスケール係数にする。

        **3水準以上の価格の場合（区間別 WTP）**

        価格を効果コーディングすると各価格水準の効用が独立に推定される。
        これを活かし、``method="segment"``（デフォルト）では **隣接する
        価格水準の区間ごと** に別々の傾き（価格感応度）で WTP を計算する。
        価格帯によって価格感応度が変わるため、WTP も区間ごとに変わるのが
        自然である。戻り値には ``価格区間`` 列が付き、属性 × 区間の行が
        出力される。

        ``method="linear"`` を指定すると、価格水準と部分効用を線形近似
        （``np.polyfit``）した1本の傾きから WTP を計算する（従来方式・
        教材用）。このとき ``wtp_price_linear_approx`` 警告が追加される
        （重大度：中）。価格が2水準のときは区間が1つだけなので、
        ``method`` によらず単一値を返す。

        **非価格属性のWTP（水準数によらず共通）**

        効果コーディングでは基準水準の効用が ``-Σ b_j`` になるため、
        基準水準から水準 k への効用差は ``b_k + Σ b_j``。
        これに「評点1点あたりの金額」（``wtp_price_factor / 2``）を掛けた値を
        WTPとする。

        .. code-block:: python

            wtp_k = (b_k + Σ b_j) * wtp_price_factor / 2

        2水準属性では効用差が ``2 × b`` となるので、
        ``wtp = wtp_price_factor * b_attr`` と同値。

        Parameters
        ----------
        price_col : str, optional
            価格の数値列名。``fit`` で設定した値を上書きしたい場合に使う。
            ``price_col`` には符号化列（``price_0`` など）ではなく、
            元の数値列名（例：``"price"``）を渡すこと。
        method : {"segment", "linear"}, default "segment"
            価格3水準以上のとき、区間別（``"segment"``）で計算するか
            線形近似1本（``"linear"``）で計算するか。上記参照。
        price_segment : str または (low, high), optional
            特定の価格区間の WTP だけを取り出したいときに指定する。
            ラベル文字列（例：``"6〜8"``）または ``(6, 8)`` のタプル。
            区間別 WTP（価格3水準以上かつ ``method="segment"``）のときだけ
            指定できる。それ以外で指定すると ``ValueError`` になる。

        Returns
        -------
        pd.DataFrame
            列：``係数``, ``限界支払意思額``（価格と同じ単位）。
            インデックスは非価格属性の符号化列名。
            3水準以上の非価格属性には K-1 行が出力される。
            区間別（価格3水準以上 × ``method="segment"``）のときは先頭に
            ``価格区間`` 列が付く。

        Raises
        ------
        ValueError
            価格列がデータにない、または価格の符号化列が見つからない場合。
            3水準以上の価格で ``reference_levels`` に価格属性がない場合。

        Notes
        -----
        * 価格列の単位（万円・千円・円など）に依存するので、結果の単位は
          元データと同じ。
        * 2水準価格のみ：価格係数 ``b_price`` が正（価格が低い水準で高評価）
          であることが前提。有意かつ負の場合は ``price_sign_negative`` 警告が出る。
        * ``price_insignificant`` 警告と戻り値の ``attrs["p_price"]`` は、
          2水準価格では価格係数の t 検定、3水準以上では
          「すべての価格係数 = 0」の同時 F 検定の p値で判定する。
        * 計算は「貨幣の限界効用が一定（価格効用が線形）・所得効果なし」を
          仮定している。この仮定の下では補償変分と等価変分が一致し、
          MWTP = 効用差 ÷ 貨幣の限界効用 となる。
        """
        price_col = price_col if price_col is not None else self.price_col
        if price_col is None:
            raise ValueError(
                "価格列が指定されていません。\n"
                "  fit() の price_col 引数か、wtp() の price_col 引数で\n"
                "  価格の数値列名（例: 'price'）を指定してください。"
            )
        if method not in ("segment", "linear"):
            raise ValueError(
                f"method='{method}' は無効です。\n"
                "  'segment'（区間別）または 'linear'（線形近似）を指定してください。"
            )
        if price_col not in self.df.columns:
            raise ValueError(
                f"価格列 '{price_col}' が DataFrame にありません。\n"
                f"  fit() の price_col 引数か、wtp() の price_col 引数で\n"
                f"  正しい列名を指定してください。"
            )

        # 価格の符号化列を構成的に特定（前方一致は使わない）
        price_encoded = self._find_encoded_for(price_col)
        if price_encoded is None:
            raise ValueError(
                f"価格列 '{price_col}' に対応する符号化列が見つかりません。\n"
                f"  encode() で価格属性も符号化しているか確認してください。"
            )

        price_enc_col = price_encoded[0]
        p_price = self._price_pvalue(price_encoded)

        # 価格水準ごとの部分効用を復元（効果コーディング：基準水準 = -Σb）
        levels, util = _price_level_utility_map(
            self.df, self.params, price_col, price_encoded, base_zero=False
        )
        n_levels = len(levels)
        low_price, high_price = levels[0], levels[-1]
        price_range = float(high_price - low_price)

        # ---- 警告①：価格係数の有意性（p ≥ 0.10） ----
        # 重複登録を防ぐ（wtp() は複数回呼ばれる可能性がある）
        already_cats = {d.category for d in self._diagnostics}
        if p_price >= 0.10 and "price_insignificant" not in already_cats:
            test_label = (
                "価格係数" if len(price_encoded) == 1 else "価格係数の同時F検定"
            )
            self._diagnostics.append(
                Diagnostic(
                    severity="中",
                    category="price_insignificant",
                    message=(
                        f"{test_label}（p値 = {p_price:.3f}）が有意水準 0.10 を超えています。"
                        "WTP の計算は価格係数を分母に使うため、"
                        "係数が不確実だと WTP の信頼性も低くなります。"
                    ),
                    recommendation=(
                        "WTP の値は参考程度にとどめ、価格感度については別途検討してください。"
                        "価格の水準数を増やす、または回答者数を増やすことも有効です。"
                    ),
                )
            )

        # 線形近似のスケール係数（attrs と method="linear" 用）。
        # ※「評点1点の金額」は unit_rating_money() = wtp_price_factor / 2 。
        if n_levels == 2:
            b_price = float(self.params[price_enc_col])
            # -(low_price - high_price) / b_price = (high - low) / b_price
            wtp_price_factor = -(low_price - high_price) / b_price
        else:
            wtp_price_factor = self._calc_price_sensitivity(price_col)

        # 隣接する価格水準ごとの区間（傾き）を作る
        segs = _price_segments_from_utilities(levels, util)
        multi_segment = (method == "segment") and (n_levels >= 3)
        # 区間別 WTP でないのに price_segment が指定されたら、黙って無視せず
        # エラーにする（「指定した区間の値が出ている」という誤解を防ぐ）。
        if price_segment is not None and not multi_segment:
            raise ValueError(
                "price_segment は区間別 WTP（価格が3水準以上かつ "
                "method='segment'）のときだけ指定できます。\n"
                f"  現在: 価格 {n_levels} 水準、method='{method}'"
            )
        if method == "linear" and n_levels >= 3:
            cat_key = "wtp_price_linear_approx"
            if cat_key not in already_cats:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="wtp_price_linear_approx",
                        message=(
                            f"価格が {n_levels} 水準ありますが、"
                            "価格効用が線形であると仮定してWTPを計算しています。"
                        ),
                        recommendation=(
                            "価格効用が等間隔でない場合は近似誤差が生じます。"
                            "method='segment' なら価格区間ごとの WTP を確認できます。"
                        ),
                    )
                )
        elif multi_segment and price_segment is not None:
            segs = _select_price_segment(segs, price_segment)

        # 価格属性の符号化列を除外し、属性ごとに「基準水準からの効用差」を金額換算。
        # 効果コーディングでは基準水準の効用 = -Σ b_j なので、
        # 基準→水準k の効用差 = b_k + Σ b_j（2水準属性では 2b に一致する）。
        price_encoded_set = set(price_encoded)
        groups = _group_columns_by_attribute(
            self.encoded_columns, list(self.reference_levels.keys())
        )
        rows = []
        if multi_segment:
            for seg in segs:
                money_per_utility = -1.0 / seg["slope"]
                for attr, cols in groups.items():
                    if all(c in price_encoded_set for c in cols):
                        continue
                    bs = [float(self.params[c]) for c in cols]
                    sum_b = sum(bs)
                    for col, b in zip(cols, bs):
                        rows.append(
                            {
                                "variable": col,
                                "価格区間": seg["label"],
                                "係数": b,
                                "限界支払意思額": (b + sum_b) * money_per_utility,
                            }
                        )
        else:
            # 単一値（2水準価格、または method="linear"）。
            money_per_utility = wtp_price_factor / 2.0  # 評点1点あたりの金額
            for attr, cols in groups.items():
                if all(c in price_encoded_set for c in cols):
                    continue  # 価格属性はWTP出力に含めない
                bs = [float(self.params[c]) for c in cols]
                sum_b = sum(bs)
                for col, b in zip(cols, bs):
                    rows.append(
                        {
                            "variable": col,
                            "係数": b,
                            "限界支払意思額": (b + sum_b) * money_per_utility,
                        }
                    )

        out = pd.DataFrame(rows).set_index("variable")
        out.index.name = "属性（符号化列名）"
        if multi_segment:
            out = out[["価格区間", "係数", "限界支払意思額"]]

        # ---- 警告②：WTP が価格レンジ × 2 を超える（外挿） ----
        for _, row in out.iterrows():
            attr_col = row.name
            wtp_val = float(row["限界支払意思額"])
            threshold = price_range * 2
            cat_key = f"wtp_extrapolation_{attr_col}"
            if abs(wtp_val) > threshold and cat_key not in self._warned_keys:
                self._warned_keys.add(cat_key)
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="wtp_extrapolation",
                        message=(
                            f"{attr_col} の WTP（{wtp_val:.2f}）が"
                            f"価格レンジ（{low_price}〜{high_price}、差 {price_range:.1f}）の"
                            f"2倍（{threshold:.1f}）を大きく超えています。"
                            "これは観測データの範囲を外れた外挿値です。"
                        ),
                        recommendation=(
                            "この WTP 値をそのまま「消費者は X 万円まで払う」と"
                            "解釈するのは危険です。"
                            "「調査した価格レンジ内での相対的な選好」として"
                            "解釈するにとどめてください。"
                        ),
                    )
                )

        # 補足情報を attrs に保持（テストや plot_wtp で使える）
        out.attrs["price_col"] = price_col
        out.attrs["wtp_price_factor"] = wtp_price_factor
        out.attrs["price_low"] = low_price
        out.attrs["price_high"] = high_price
        out.attrs["price_range"] = price_range
        out.attrs["p_price"] = p_price
        out.attrs["method"] = method
        return out

    def unit_rating_money(self, *, price_col: Optional[str] = None) -> float:
        """
        評点1ポイントが何円（または何万円）に相当するかを返す。

        2水準: ``(price_max - price_min) / abs(price_coef * 2)``
        3水準以上: 価格水準と部分効用の線形近似スロープから算出（``wtp()`` の仮定と同じ）。
        単位は価格列の単位と同じ。例えば価格が万円単位なら、戻り値も万円単位。

        Parameters
        ----------
        price_col : str, optional
            価格列名。``fit`` で設定した値を上書きしたい場合に使う。

        Returns
        -------
        float
            評点 1 ポイント相当の金額（価格列と同じ単位）。

        Raises
        ------
        ValueError
            価格列がデータにない場合。
            価格の符号化列が見つからない場合。
        """
        price_col = price_col or self.price_col
        if price_col not in self.df.columns:
            raise ValueError(f"価格列 '{price_col}' が DataFrame にありません。")
        price_encoded = self._find_encoded_for(price_col)
        if not price_encoded:
            raise ValueError(
                f"価格列 '{price_col}' に対応する符号化列が見つかりません。"
            )
        price_levels = sorted(self.df[price_col].dropna().unique().tolist())
        if len(price_levels) == 2:
            b_price = float(self.params[price_encoded[0]])
            price_range = float(price_levels[1] - price_levels[0])
            return price_range / abs(b_price * 2)
        # 3水準以上: wtp_price_factor / 2
        return self._calc_price_sensitivity(price_col) / 2.0

    # ---- 市場シェア予測 ---------------------------------------------------

    def market_share(
        self,
        products: pd.DataFrame,
        *,
        method: str = "logit",
    ) -> pd.Series:
        """
        複数の製品（プロファイル）の **市場シェア** を予測する。

        Parameters
        ----------
        products : pd.DataFrame
            製品ごとの符号化列を含むDataFrame。
            インデックスを製品名にしておくと結果が読みやすい。

            各製品行は ``encoded_columns`` の各列に -1/0/1 の値を持つ必要がある。

        method : {"logit", "max", "share_of_preference"}, default "logit"
            シェア計算方法。

            * ``"logit"`` または ``"share_of_preference"``: ロジット式。
              ``share_i = exp(u_i) / Σ exp(u_j)``
              ノートブックの方法と同じ。最も一般的。
            * ``"max"``: 最大効用ルール。
              最大効用の製品にシェア1、他は0。**仮定**：消費者が完全合理的。

        Returns
        -------
        pd.Series
            製品名 → シェア（0〜1）の Series。合計は1になる。

        Examples
        --------
        >>> products = pd.DataFrame({
        ...     "price_0":  [ 1, -1],   # 製品A: 6万円, 製品B: 10万円
        ...     "os_0":     [-1,  1],   # 製品A: android, 製品B: apple
        ...     "camera_0": [ 1,  1],   # 両製品とも高性能
        ... }, index=["製品A", "製品B"])
        >>> result.market_share(products)
        製品A    0.327
        製品B    0.673
        dtype: float64
        """
        if not isinstance(products, pd.DataFrame):
            raise TypeError("products は pandas.DataFrame で渡してください。")

        # 必要な列がすべて揃っているか確認
        missing = [c for c in self.encoded_columns if c not in products.columns]
        if missing:
            raise ValueError(
                f"products に必要な列がありません: {missing}\n"
                f"  必要な列: {self.encoded_columns}"
            )

        # 内部の formula はエイリアス列名を使っているため、predict 用に変換
        if self.alias_map:
            products_for_predict = products.rename(columns=self.alias_map)
        else:
            products_for_predict = products

        # 効用予測
        u = self.ols.predict(products_for_predict)

        if method in ("logit", "share_of_preference"):
            # 数値安定化のため最大値を引く（softmax）
            u_arr = np.asarray(u, dtype=float)
            u_shift = u_arr - np.max(u_arr)
            ex = np.exp(u_shift)
            share = ex / ex.sum()
        elif method == "max":
            share = np.zeros_like(u, dtype=float)
            share[np.argmax(u)] = 1.0
        else:
            raise ValueError(
                f"method='{method}' は無効です。\n"
                "  'logit', 'share_of_preference', 'max' のいずれかを指定してください。"
            )

        return pd.Series(share, index=products.index, name="market_share")

    # ---- 可視化系のショートカット（plot.py に委譲） ------------------------

    def plot_importance(self, **kwargs: Any) -> Any:
        """
        重要度の棒グラフを描画する。:func:`py4conjoint.plot.plot_importance`
        へのショートカット。

        Returns
        -------
        matplotlib.axes.Axes
        """
        from .plot import plot_importance

        return plot_importance(self, **kwargs)

    def plot_partworth(self, **kwargs: Any) -> Any:
        """
        部分効用（パートワース）の棒グラフを描画する。
        :func:`py4conjoint.plot.plot_partworth` へのショートカット。
        """
        from .plot import plot_partworth

        return plot_partworth(self, **kwargs)

    def plot_wtp(self, **kwargs: Any) -> Any:
        """
        WTPの棒グラフを描画する。:func:`py4conjoint.plot.plot_wtp` へのショートカット。
        """
        from .plot import plot_wtp

        return plot_wtp(self, **kwargs)

    # ---- 内部処理 ---------------------------------------------------------

    def _price_pvalue(self, price_encoded: List[str]) -> float:
        """
        価格の有意性の p値を返す。

        2水準（符号化列が1本）：係数の t 検定の p値。
        3水準以上（符号化列が複数）：「すべての価格係数 = 0」の
        同時 F 検定の p値。先頭列の p値だけでは多水準価格の
        有意性を正しく判定できないため。
        """
        if len(price_encoded) == 1:
            return float(self.ols.pvalues[price_encoded[0]])
        exog_names = list(self.ols.model.exog_names)
        R = np.zeros((len(price_encoded), len(exog_names)))
        for i, c in enumerate(price_encoded):
            R[i, exog_names.index(c)] = 1.0
        return float(self.ols.f_test(R).pvalue)

    def _calc_price_sensitivity(self, price_col: str) -> float:
        """
        wtp_price_factor を返す（WTP = wtp_price_factor * b_attr）。

        3水準以上の価格専用。価格水準と部分効用を線形近似し、
        slope（utility / price）から -2/slope を返す。
        戻り値は正（価格が上がると効用が下がる通常財の場合）。

        呼び出し元: wtp()・unit_rating_money() の 3水準以上ブランチのみ。
        """
        price_encoded = self._find_encoded_for(price_col)
        price_levels = sorted(self.df[price_col].dropna().unique().tolist())

        # 3水準以上: データから各符号化列が対応する価格水準を特定する
        price_util_map: Dict[float, float] = {}
        for enc_col in price_encoded:
            rows_1 = self.df[self.df[enc_col] == 1]
            if len(rows_1) > 0:
                level = float(rows_1[price_col].iloc[0])
                price_util_map[level] = float(self.params[enc_col])

        # 基準水準の効用 = -(他の水準の係数の和)
        bs = np.array([float(self.params[c]) for c in price_encoded])
        ref_price = self.reference_levels.get(price_col)
        if ref_price is None:
            raise ValueError(
                f"価格属性 '{price_col}' の基準水準が reference_levels に見つかりません。\n"
                f"  3水準以上の価格WTPには基準水準が必要です。\n"
                f"  fit() の reference_levels 引数で明示的に指定してください。\n"
                f"  例: pc.fit(df_coded, reference_levels={{'{price_col}': 基準値}})"
            )
        price_util_map[float(ref_price)] = float(-bs.sum())

        price_arr = np.array(price_levels, dtype=float)
        util_arr = np.array([price_util_map[p] for p in price_levels], dtype=float)
        slope, _ = np.polyfit(price_arr, util_arr, 1)
        # slope < 0 for normal goods; wtp_price_factor = -2/slope > 0
        return -2.0 / slope

    def _attribute_ranges(self) -> Dict[str, float]:
        """
        各属性について、部分効用の最大値 − 最小値（効用範囲）を計算する。
        """
        ranges: Dict[str, float] = {}
        groups = _group_columns_by_attribute(
            self.encoded_columns, list(self.reference_levels.keys())
        )
        for attr, cols in groups.items():
            if len(cols) == 1:
                # 2水準: 効用は ±|b|、範囲は 2*|b|
                b = float(self.params[cols[0]])
                ranges[attr] = 2.0 * abs(b)
            else:
                # 3水準以上: 効果コーディング前提
                # 各非基準水準の効用 = b_k
                # 基準水準の効用 = -Σ b_k
                bs = np.array([float(self.params[c]) for c in cols])
                utils = np.append(bs, -bs.sum())
                ranges[attr] = float(utils.max() - utils.min())
        return ranges

    def _find_encoded_for(self, original_col: str) -> Optional[List[str]]:
        """
        元の属性列名に対応する符号化列を、encode の命名規則から
        **構成的に** 特定する。

        数値列の水準数（K）から非基準水準の数（K-1）を求め、
        ``encode()`` の命名規則（デフォルトは ``{attr}_0``, ``{attr}_1`` ...、
        ``suffix_map`` 指定時はそのサフィックス）から想定される列名を構成し、
        ``encoded_columns`` 内の該当列だけを返す。

        ``startswith`` による前方一致は使わないため、``price_range_high``
        のように接頭辞が紛らわしい別属性の列を誤検出しない。
        """
        if original_col not in self.df.columns:
            return None
        levels = list(pd.Series(self.df[original_col].dropna().unique()))
        n_others = len(levels) - 1
        if n_others < 1:
            return None
        meta = self.df.attrs.get("py4conjoint", {}) if hasattr(self.df, "attrs") else {}
        suffix_map = meta.get("suffix_map") or {}
        raw = suffix_map.get(original_col)
        if raw is None:
            suffixes = [str(i) for i in range(n_others)]
        elif isinstance(raw, str):
            suffixes = [raw]
        else:
            suffixes = [str(s) for s in raw]
        names = [f"{original_col}_{s}" for s in suffixes]
        encoded_set = set(self.encoded_columns)
        cols = [n for n in names if n in encoded_set]
        return cols if cols else None

    def _run_diagnostics(self) -> None:
        """
        落とし穴の自動検出を行い、:class:`Diagnostic` のリストとして蓄積する。

        検出される警告
        ---------------
        1. **price_sign_negative**（重大度：中）
           価格係数が負かつ有意（p < 0.10）。符号化方向のミスを示唆。
           有意でない場合は符号がノイズで決まるため発火しない。

        2. **r2_low**（重大度：大）
           R² < 0.20。線形仮定の妥当性に大きな疑問。

        3. **obs_per_predictor**（重大度：大 or 中）
           観測数／説明変数数の比率が低い。
           比率 < 5 で「大」、< 10 で「中」。

        4. **few_respondents**（重大度：大 or 中）
           回答者が1人なら「大」、2〜4人なら「中」。
           ``respondent_id`` 列がある場合のみ判定する。

        5. **independence_assumed**（重大度：中）
           回答者ID列が見つからず、観測の独立性を仮定した通常の標準誤差を
           使用している。同一回答者の複数回答があるとp値が過小になる。

        6. **price_insignificant**（重大度：中、:meth:`wtp` 呼出時）
           価格の p値 ≥ 0.10（2水準は t 検定、3水準以上は同時 F 検定）。
           WTP計算の分母が不確実。

        7. **wtp_extrapolation**（重大度：中、:meth:`wtp` 呼出時）
           ``|WTP| > 価格レンジ × 2``。観測範囲外への外挿。

        8. **wtp_price_linear_approx**（重大度：中、:meth:`wtp` 呼出時）
           価格が 3 水準以上のため、価格効用が線形と仮定して WTP を計算した。
        """
        # 1) 価格係数の符号（有意な場合のみ）
        if self.price_col in self.df.columns:
            price_cols = self._find_encoded_for(self.price_col)
            if price_cols and len(price_cols) == 1:
                b_price = float(self.params[price_cols[0]])
                p_price_diag = float(self.ols.pvalues[price_cols[0]])
                # 有意でない場合は符号がノイズで決まるため偽陽性になりうる
                if b_price < 0 and p_price_diag < 0.10:
                    self._diagnostics.append(
                        Diagnostic(
                            severity="中",
                            category="price_sign_negative",
                            message=(
                                f"価格係数（{price_cols[0]} = {b_price:.4f}）が負です。"
                                "「価格が低い水準で評点が低い」という直感に反します。"
                            ),
                            recommendation=(
                                "符号化方向（reference_levels）が逆になっていないか、"
                                "またはデータ品質に問題がないか確認してください。"
                            ),
                        )
                    )

        # 2) R² が極端に低い（閾値: 0.20）
        if self.rsquared < 0.20:
            self._diagnostics.append(
                Diagnostic(
                    severity="大",
                    category="r2_low",
                    message=(
                        f"R² = {self.rsquared:.3f} が 0.20 未満で、説明力が低いです。"
                    ),
                    recommendation=(
                        "「評点 = 属性の足し算」という線形仮定が当てはまりにくい、"
                        "または回答にノイズが多い可能性があります。"
                        "回答の質、属性の選び方、交互作用の有無を再検討してください。"
                    ),
                )
            )

        # 3) 観測数／説明変数比のチェック
        n_vars = len(self.encoded_columns)
        if n_vars > 0:
            ratio = self.n_obs / n_vars
            if ratio < 5:
                self._diagnostics.append(
                    Diagnostic(
                        severity="大",
                        category="obs_per_predictor",
                        message=(
                            f"観測数（{self.n_obs}）が説明変数数（{n_vars}）の"
                            f"{ratio:.1f}倍しかなく、推定が非常に不安定です。"
                        ),
                        recommendation=(
                            "回答者を増やしてください。"
                            "目安として観測数は説明変数数の10倍以上が望ましいです。"
                        ),
                    )
                )
            elif ratio < 10:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="obs_per_predictor",
                        message=(
                            f"観測数（{self.n_obs}）が説明変数数（{n_vars}）の"
                            f"{ratio:.1f}倍で、やや少なめです。"
                        ),
                        recommendation=(
                            "可能なら回答者を増やしてください。"
                            "目安として観測数は説明変数数の10倍以上が望ましいです。"
                        ),
                    )
                )

        # 4) 回答者数の確認
        if self.respondent_id_col in self.df.columns:
            n_resp = int(self.df[self.respondent_id_col].nunique())
            if n_resp == 1:
                self._diagnostics.append(
                    Diagnostic(
                        severity="大",
                        category="few_respondents",
                        message=(
                            "回答者が1人しかいません。"
                            "個人の好みを集団の傾向と誤解する危険が大きいです。"
                        ),
                        recommendation=(
                            "複数の回答者からデータを集めてください。"
                            "目安として最低でも10人以上が望ましいです。"
                        ),
                    )
                )
            elif n_resp < 5:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="few_respondents",
                        message=(
                            f"回答者が {n_resp} 人と少なめです。"
                            "推定値の不確実性が大きい点に注意してください。"
                        ),
                        recommendation=(
                            "可能ならばさらに回答を集めてください。"
                            "目安として最低でも10人以上が望ましいです。"
                        ),
                    )
                )
        else:
            # 5) 回答者ID列がない → クラスタリングできず独立性を仮定
            self._diagnostics.append(
                Diagnostic(
                    severity="中",
                    category="independence_assumed",
                    message=(
                        f"回答者ID列 '{self.respondent_id_col}' が見つからないため、"
                        "観測の独立性を仮定した通常の標準誤差を使用しています。"
                        "同じ回答者の複数回答が含まれる場合、p値が過小"
                        "（有意に出やすく）になります。"
                    ),
                    recommendation=(
                        "回答者を識別できる列がある場合は、fit() の "
                        "respondent_id_col 引数でその列名を指定してください。"
                    ),
                )
            )


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _rename_result_index(
    res: RegressionResults, rev_map: Dict[str, str]
) -> RegressionResults:
    """
    回帰結果オブジェクトの係数名（params, pvalues, bse, tvalues, conf_int, model.exog_names）
    を ``rev_map`` で置換する。

    Notes
    -----
    statsmodels の ``RegressionResults`` は ``params`` 等を計算プロパティとして
    持つため、戻り値の Series 自体を差し替えることはできない。
    そのため、内部の ``model.data.xnames`` と ``model.data.cov_names`` を
    書き換えることで、以降のすべてのプロパティ（params, pvalues, bse, tvalues,
    conf_int, predict など）に置換が反映される。
    """
    # exog（説明変数行列）の名前
    if hasattr(res.model, "data") and hasattr(res.model.data, "xnames"):
        res.model.data.xnames = [rev_map.get(n, n) for n in res.model.data.xnames]
    if hasattr(res.model, "exog_names"):
        # 一部のバージョンでは exog_names は data.xnames を参照するプロパティ
        # 個別属性として保持されている場合に備えて両方更新を試みる
        try:
            res.model.exog_names[:] = [rev_map.get(n, n) for n in res.model.exog_names]
        except (TypeError, AttributeError):
            pass
    # 内生変数（被説明変数）名
    if hasattr(res.model.data, "ynames"):
        ynames = res.model.data.ynames
        if isinstance(ynames, str):
            res.model.data.ynames = rev_map.get(ynames, ynames)
        elif isinstance(ynames, list):
            res.model.data.ynames = [rev_map.get(n, n) for n in ynames]
    return res


def _significance_stars(p: float) -> str:
    if p < 0.001:
        return "***"
    if p < 0.01:
        return "**"
    if p < 0.05:
        return "*"
    if p < 0.1:
        return "."
    return ""


def _display_width(s: str) -> int:
    """Wide(W)・Fullwidth(F) を2列、それ以外を1列として扱う。"""
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in s)


def _ljust_display(s: str, width: int) -> str:
    """表示幅ベースで左寄せパディングする。"""
    return s + " " * max(0, width - _display_width(s))


def _rjust_display(s: str, width: int) -> str:
    """表示幅ベースで右寄せパディングする。"""
    return " " * max(0, width - _display_width(s)) + s


def _df_to_string_cjk(df: pd.DataFrame, index: bool = True) -> str:
    """全角文字の表示幅を正しく考慮した DataFrame 文字列化。"""
    str_df = df.reset_index() if index else df.copy()
    cols = list(str_df.columns)
    rows_str = [[str(v) for v in row] for row in str_df.itertuples(index=False)]

    col_widths = [
        max(
            _display_width(col),
            max((_display_width(r[i]) for r in rows_str), default=0),
        )
        for i, col in enumerate(cols)
    ]

    def _fmt_row(row_vals: list[str]) -> str:
        parts = [_rjust_display(val, col_widths[i]) for i, val in enumerate(row_vals)]
        return "  ".join(parts)

    lines = [_fmt_row(cols)]
    for row in rows_str:
        lines.append(_fmt_row(row))
    return "\n".join(lines)


def _is_effect_coded_column(s: pd.Series) -> bool:
    """
    効果コーディング済みの列か判定する。

    ``-1`` と ``1`` の両方を含み、値が ``{-1, 0, 1}`` に収まる数値列を
    効果コーディング済みとみなす。``encode()`` の出力では基準水準（-1）と
    対象水準（1）が必ずデータに存在するため、この条件が成り立つ。
    ``respondent_encode`` の出力など 0/1 のみの列は含まれない。
    """
    if not pd.api.types.is_numeric_dtype(s):
        return False
    vals = set(pd.Series(s.dropna().unique()).tolist())
    return {-1, 1}.issubset(vals) and vals.issubset({-1, 0, 1})


def _detect_encoded_columns(
    df: pd.DataFrame,
    *,
    rating: str,
    reference_levels: Optional[Dict[str, object]] = None,
) -> List[str]:
    """
    符号化列を自動検出する。
    優先順位：
    1. reference_levels が与えられていれば、その属性名 + ``"_"`` で始まり、
       かつ効果コーディング済み（``_is_effect_coded_column()``）の列を採用。
       元の属性列・rating列は除外し、複数の属性名に前方一致しても
       1回だけ登録する。
    2. フォールバック：効果コーディング済みの数値列をすべて採用
       （0/1 のみの列＝``respondent_encode`` の出力などは含めない）。
    """
    if reference_levels:
        attr_set = set(reference_levels.keys())
        cols: List[str] = []
        for c in df.columns:
            if c == rating or c in attr_set:
                continue
            if not any(c.startswith(f"{a}_") for a in attr_set):
                continue
            if _is_effect_coded_column(df[c]):
                cols.append(c)
        if cols:
            return cols

    # フォールバック：値の範囲で判定
    return [c for c in df.columns if c != rating and _is_effect_coded_column(df[c])]


def _group_columns_by_attribute(
    encoded_columns: List[str], known_attrs: List[str]
) -> Dict[str, List[str]]:
    """
    符号化列を元の属性名でグルーピングする。
    ``{属性名}_{インデックス}`` という命名規則を前提にする。

    既知の属性リストがあればそれで前方一致を試み、
    なければ ``_`` で分割した先頭部分を属性名とみなす。
    """
    groups: Dict[str, List[str]] = {}
    if known_attrs:
        # 長い名前から優先（部分一致を防ぐ）
        sorted_attrs = sorted(known_attrs, key=len, reverse=True)
        for c in encoded_columns:
            matched = None
            for a in sorted_attrs:
                if c == a or c.startswith(f"{a}_"):
                    matched = a
                    break
            if matched is None:
                # 既知に一致しない場合は最初の "_" で分割
                matched = c.split("_")[0]
            groups.setdefault(matched, []).append(c)
        return groups

    # 既知属性がない場合
    for c in encoded_columns:
        attr = c.split("_")[0]
        groups.setdefault(attr, []).append(c)
    return groups


# ---------------------------------------------------------------------------
# 区間別 WTP の共通ヘルパー（rating / choice で共通利用）
# ---------------------------------------------------------------------------


def _format_price_level(x: float) -> str:
    """価格水準を表示用に整形する（整数なら小数点を出さない）。"""
    xf = float(x)
    return str(int(xf)) if xf == int(xf) else f"{xf:g}"


def _price_level_utility_map(
    df: pd.DataFrame,
    params: pd.Series,
    price_col: str,
    price_encoded: List[str],
    *,
    base_zero: bool,
) -> "tuple[List[float], Dict[float, float]]":
    """
    価格の符号化列から、各価格水準の部分効用を復元する。

    各符号化列が「どの価格水準のダミーか」は、その列が ``1`` になっている
    行の ``price_col`` の値から **データに基づいて** 特定する
    （列名の前方一致には依存しない）。

    Parameters
    ----------
    base_zero : bool
        基準水準の効用を 0 とみなすか（choice のダミーコーディング）。
        ``False`` なら基準水準の効用 = ``-Σ係数``（rating の効果コーディング）。

    Returns
    -------
    (levels, util)
        ``levels``：昇順にソートした価格水準のリスト。
        ``util``：``{価格水準: 部分効用}`` の辞書（基準水準も含む）。
    """
    util: Dict[float, float] = {}
    coefs: List[float] = []
    for c in price_encoded:
        coef = float(params[c])
        coefs.append(coef)
        rows_1 = df[df[c] == 1]
        if len(rows_1) > 0:
            util[float(rows_1[price_col].iloc[0])] = coef
    levels = sorted(float(x) for x in df[price_col].dropna().unique())
    base_util = 0.0 if base_zero else -float(sum(coefs))
    for lv in levels:
        if lv not in util:
            util[lv] = base_util
    return levels, util


def _price_segments_from_utilities(
    levels: List[float], util: Dict[float, float]
) -> List[Dict[str, Any]]:
    """
    昇順の価格水準と各水準の効用から、隣接する価格区間ごとの情報を返す。

    各区間の傾き ``slope = (u(high) − u(low)) / (high − low)``
    （通常財では負）。``money_per_utility = −1 / slope`` は
    「効用1単位あたりの金額」（通常財では正）を表す。
    """
    segs: List[Dict[str, Any]] = []
    for lo, hi in zip(levels[:-1], levels[1:]):
        span = hi - lo
        slope = (util[hi] - util[lo]) / span if span != 0 else float("nan")
        segs.append(
            {
                "low": lo,
                "high": hi,
                "label": f"{_format_price_level(lo)}〜{_format_price_level(hi)}",
                "slope": slope,
            }
        )
    return segs


def _select_price_segment(
    segs: List[Dict[str, Any]], price_segment: Any
) -> List[Dict[str, Any]]:
    """
    ``price_segment`` 引数に一致する価格区間を1つだけ選んで返す。

    ``price_segment`` はラベル文字列（例：``"6〜8"``）または
    ``(low, high)`` のタプル／リストで指定する。
    """
    for seg in segs:
        if isinstance(price_segment, (tuple, list)) and len(price_segment) == 2:
            if float(seg["low"]) == float(price_segment[0]) and float(
                seg["high"]
            ) == float(price_segment[1]):
                return [seg]
        elif str(price_segment) == seg["label"]:
            return [seg]
    labels = [s["label"] for s in segs]
    raise ValueError(
        f"指定された価格区間 {price_segment!r} が見つかりません。\n"
        f"  利用可能な価格区間: {labels}"
    )


# ---------------------------------------------------------------------------
# check_design() の内部ヘルパー
# ---------------------------------------------------------------------------


def _check_balance(profiles: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """各属性の水準出現頻度と変動係数（CV）を計算する。"""
    rows = []
    for attr in attrs:
        counts = profiles[attr].value_counts()
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


def _check_correlation(profiles: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """
    効果コーディング後の属性間相関行列を計算する（scipy不要）。
    各属性の最頻値を基準水準として自動設定する。
    """
    coded_parts = []
    for attr in attrs:
        col = profiles[attr]
        levels = list(col.unique())
        if len(levels) < 2:
            continue
        ref = col.mode().iloc[0]
        others = [lv for lv in levels if lv != ref]
        for i, lv in enumerate(others):

            def _map(v, t=lv, r=ref):
                if v == t:
                    return 1
                if v == r:
                    return -1
                return 0

            s = col.map(_map)
            s.name = f"{attr}__{i}" if len(others) > 1 else attr
            coded_parts.append(s)
    if not coded_parts:
        return pd.DataFrame()
    coded = pd.concat(coded_parts, axis=1)
    return coded.corr().round(4)


def _check_chi2(profiles: pd.DataFrame, attrs: List[str]) -> pd.DataFrame:
    """
    属性ペアのχ²統計量と自由度を計算する（scipy不要）。
    p値の代わりに χ²/自由度 の比率と定性的評価を返す。
    """
    rows = []
    for i, a1 in enumerate(attrs):
        for a2 in attrs[i + 1 :]:
            ct = pd.crosstab(profiles[a1], profiles[a2]).values.astype(float)
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


def _design_diagnostics(
    profiles: pd.DataFrame,
    attrs: List[str],
    balance_df: pd.DataFrame,
    corr_df: pd.DataFrame,
    chi2_df: pd.DataFrame,
) -> List[Diagnostic]:
    """直交性チェックの結果から Diagnostic のリストを生成する。"""
    diags: List[Diagnostic] = []

    # プロファイル数のチェック（最低限の要件：n_profiles >= k+1）
    n_profiles = len(profiles)
    k_params = 1  # 切片
    for attr in attrs:
        k_params += len(profiles[attr].unique()) - 1
    if n_profiles < k_params:
        diags.append(
            Diagnostic(
                severity="大",
                category="insufficient_profiles",
                message=(
                    f"プロファイル数（{n_profiles}）がパラメータ数（{k_params}）より少ないため、"
                    "回帰分析が実行できません。"
                ),
                recommendation=(
                    f"プロファイル数を少なくとも {k_params} 以上にしてください。"
                ),
            )
        )
    elif n_profiles < k_params + 2:
        diags.append(
            Diagnostic(
                severity="中",
                category="few_profiles",
                message=(
                    f"プロファイル数（{n_profiles}）がパラメータ数（{k_params}）に対して"
                    "ほぼ最小限です。"
                ),
                recommendation="プロファイルをさらに追加すると推定の安定性が上がります。",
            )
        )

    # バランスチェック
    for attr, row in balance_df.iterrows():
        cv = row["CV"]
        if cv > 0.3:
            diags.append(
                Diagnostic(
                    severity="大",
                    category=f"balance_{attr}",
                    message=f"属性 '{attr}' の水準出現頻度が偏っています（CV={cv:.3f}）。",
                    recommendation="各水準の出現回数を均等にしてください（バランスの良いデザイン）。",
                )
            )
        elif cv > 0.15:
            diags.append(
                Diagnostic(
                    severity="中",
                    category=f"balance_{attr}",
                    message=f"属性 '{attr}' の水準出現頻度にやや偏りがあります（CV={cv:.3f}）。",
                    recommendation="可能であれば各水準の出現回数を均等に近づけてください。",
                )
            )

    # 相関チェック（同一属性の符号化列ペアはスキップ）
    if not corr_df.empty:
        for col1 in corr_df.columns:
            for col2 in corr_df.columns:
                if col1 >= col2:
                    continue
                # 3水準以上の属性は "attr__i" という列名になる。同一属性内のペアは無視。
                if col1.split("__")[0] == col2.split("__")[0]:
                    continue
                r = abs(float(corr_df.loc[col1, col2]))
                if r > 0.5:
                    diags.append(
                        Diagnostic(
                            severity="大",
                            category=f"correlation_{col1}_{col2}",
                            message=(
                                f"'{col1}' と '{col2}' の相関が高いです（|r|={r:.3f}）。"
                                "パラメータの独立推定が困難になります。"
                            ),
                            recommendation=(
                                "プロファイルの組み合わせを見直し、"
                                "2属性の水準が独立に出現するよう設計してください。"
                            ),
                        )
                    )
                elif r > 0.3:
                    diags.append(
                        Diagnostic(
                            severity="中",
                            category=f"correlation_{col1}_{col2}",
                            message=f"'{col1}' と '{col2}' の相関がやや高いです（|r|={r:.3f}）。",
                            recommendation="可能であればプロファイルの組み合わせを調整してください。",
                        )
                    )

    # χ²チェック
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

    return diags
