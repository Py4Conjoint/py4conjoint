"""
analysis.py
============
コンジョイント分析の **回帰実行と結果の保持** を担当するモジュール。

中心となるのは :func:`fit` 関数と :class:`ConjointResult` クラス。

>>> import py4conjoint as pc
>>> df_coded = pc.encode(df, reference_levels={...})
>>> result = pc.fit(df_coded)
>>> result.summary()        # 和文サマリー
>>> result.importance()     # 相対重要度
>>> result.wtp()            # WTP（支払意思額）
>>> result.market_share(products)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Union

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from statsmodels.regression.linear_model import RegressionResults


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

    encoded_columns : list of str, optional
        説明変数として使う符号化列のリスト。
        省略時は ``reference_levels`` から自動推定するか、
        値が ``-1, 0, 1`` のみを取る列を自動検出する。

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
        通常は省略してよい（自動構築される）。

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

    * 回答者が1人の場合（重大度：大）または2〜4人の場合（重大度：中）：
      個人・少数の好みを集団の傾向と誤解する危険。
    """
    # ---------- 入力チェック ----------
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
    n_before = len(df)
    df_fit = df[use_cols].dropna()
    n_after = len(df_fit)

    # formula が明示指定されていれば、列名はそのまま使う前提
    alias_map: Dict[str, str] = {}
    if formula is not None:
        res: RegressionResults = smf.ols(formula, data=df).fit()
    else:
        # 日本語など formula で扱いにくい列名を、内部的に英数字エイリアスへ
        # 一時リネームして回帰し、推定後に係数名を元に戻す。
        # こうすることで `result.params["camera_高性能"]` のように
        # 元の列名でアクセスできるようになる。
        alias_map = {
            c: f"__pc_var_{i}__" for i, c in enumerate(encoded_columns)
        }
        rev_map = {v: k for k, v in alias_map.items()}
        rating_alias = "__pc_rating__"
        rev_map[rating_alias] = rating

        df_alias = df[use_cols].rename(
            columns={**alias_map, rating: rating_alias}
        )
        formula_alias = (
            f"{rating_alias} ~ " + " + ".join(alias_map[c] for c in encoded_columns)
        )
        res = smf.ols(formula_alias, data=df_alias).fit()
        # 係数名を元に戻す
        res = _rename_result_index(res, rev_map)

    result = ConjointResult(
        ols=res,
        df=df,
        rating=rating,
        encoded_columns=list(encoded_columns),
        reference_levels=reference_levels or {},
        price_col=price_col,
        n_dropped=n_before - n_after,
        alias_map=alias_map,
    )

    # ---------- 落とし穴の自動検出 ----------
    result._run_diagnostics()

    return result


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
    """
    ols: RegressionResults
    df: pd.DataFrame
    rating: str
    encoded_columns: List[str]
    reference_levels: Dict[str, object]
    price_col: str
    n_dropped: int = 0
    alias_map: Dict[str, str] = field(default_factory=dict)

    # 内部用：検出された警告（落とし穴）のリスト
    # Diagnostic オブジェクトとして保持する。
    _diagnostics: List[Diagnostic] = field(default_factory=list)

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
        stat_rows = [
            ("観測数",       str(self.n_obs)),
            ("説明変数の数", str(len(self.encoded_columns))),
            ("決定係数 R²",  f"{self.rsquared:.4f}"),
            ("自由度修正 R²", f"{self.ols.rsquared_adj:.4f}"),
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
            + " " + _rjust_display("係数", 10)
            + " " + _rjust_display("p値", 10)
            + "  " + "有意性"
        )
        lines.append(f"  {'-' * NAME_WIDTH} {'-' * 10} {'-' * 10}  {'-' * STAR_WIDTH}")
        for name in params.index:
            coef = params[name]
            p = pvals[name]
            star = _significance_stars(p)
            lines.append(f"  {name:<{NAME_WIDTH}} {coef:>10.4f} {p:>10.4f}  {star}")
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

        p = ['<div>',
             '<p><strong>コンジョイント分析の結果</strong></p>']

        # 統計量
        stat_rows = [
            ("観測数",       str(self.n_obs)),
            ("説明変数の数", str(len(self.encoded_columns))),
            ("決定係数 R²",  f"{self.rsquared:.4f}"),
            ("自由度修正 R²", f"{self.ols.rsquared_adj:.4f}"),
        ]
        if self.n_dropped > 0:
            stat_rows.append(("欠損で除外", f"{self.n_dropped} 行"))

        p.append(f'<table style="{_tbl}">')
        for lbl, val in stat_rows:
            p.append(f'<tr>{td(lbl)}{td(val)}</tr>')
        p.append('</table>')

        # 係数テーブル
        p.append('<p><strong>【推定された係数（部分効用 part-worth）】</strong></p>')
        p.append(f'<table style="{_tbl}">')
        p.append('<tr>'
                 + th("変数") + th("係数", "right")
                 + th("p値", "right") + th("有意性", "right")
                 + '</tr>')

        params = self.params
        pvals  = self.ols.pvalues
        for name in params.index:
            c    = params[name]
            pv   = pvals[name]
            star = _significance_stars(pv)
            p.append('<tr>'
                     + td(name) + td(f"{c:.4f}", "right")
                     + td(f"{pv:.4f}", "right") + td(star, "right")
                     + '</tr>')
        p.append('</table>')
        p.append('<p style="font-size:0.85em;color:#888;">'
                 '有意水準: *** p&lt;0.001&nbsp; ** p&lt;0.01&nbsp;'
                 ' * p&lt;0.05&nbsp; . p&lt;0.1</p>')

        # 重大警告
        major = [d for d in self._diagnostics if d.severity == "大"]
        minor = [d for d in self._diagnostics if d.severity != "大"]
        if major:
            p.append('<p><strong>⚠️ 重大な注意事項</strong></p><ul>')
            for d in major:
                p.append(f'<li><strong>[{escape(d.severity)}]</strong> '
                         f'{escape(d.message)}<br>'
                         f'&nbsp;&nbsp;→ {escape(d.recommendation)}</li>')
            p.append('</ul>')
        if minor:
            p.append(
                f'<p style="font-size:0.9em;color:#888;">'
                f'その他の注意事項が {len(minor)} 件あります。'
                f'result.warnings() で確認できます。</p>'
            )

        p.append('</div>')
        return '\n'.join(p)

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
            カテゴリでフィルタする。例：``"r2_low"``, ``"price_insignificant"``,
            ``"wtp_extrapolation"``。
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

    # ---- 相対重要度 -------------------------------------------------------

    def importance(self, *, as_percent: bool = True) -> pd.DataFrame:
        """
        各属性の **相対重要度（Relative Importance）** を計算する。

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
            列：``range`` （効用範囲）, ``importance`` （重要度）。
            インデックスは属性名。``importance`` の合計は100（または1）になる。

        Examples
        --------
        >>> result.importance()
                       range  importance
        price          1.234       45.6
        os             0.567       21.0
        camera         0.901       33.4
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
            rows.append({"attribute": attr, "range": rng, "importance": imp})
        out = pd.DataFrame(rows).set_index("attribute")
        out.index.name = "属性"
        return out

    # ---- WTP（支払意思額） ------------------------------------------------

    def wtp(self, *, price_col: Optional[str] = None) -> pd.DataFrame:
        """
        各非価格属性の **WTP（支払意思額、Willingness to Pay）** を計算する。

        定義
        ----
        WTPは「ある属性を基準水準から非基準水準に変えるとき、回答者が
        最大いくらまで追加で支払ってもよいと思うか」を金額で表す。

        ノートブックの式：

        .. code-block:: python

            wtp_price_factor = -(low_price - high_price) / b_price
            wtp = wtp_price_factor * b_attr

        を一般化したもの。``low_price - high_price`` は負の値なので、
        マイナスを付けて正のスケール係数にする。

        Parameters
        ----------
        price_col : str, optional
            価格列名。``fit`` で設定した値を上書きしたい場合に使う。

        Returns
        -------
        pd.DataFrame
            列：``係数``, ``支払意思額``（価格と同じ単位）。
            インデックスは非価格属性の符号化列名。

        Raises
        ------
        ValueError
            価格列がデータにない、または価格の符号化列が見つからない場合。

        Notes
        -----
        * 価格列の単位（万円・千円・円など）に依存するので、結果の単位は
          元データと同じ。
        * 価格係数 ``b_price`` の符号が正常（価格が低い水準で正、高い水準で負）
          であることが前提。逆になっている場合は警告が出る。
        """
        price_col = price_col or self.price_col
        if price_col not in self.df.columns:
            raise ValueError(
                f"価格列 '{price_col}' が DataFrame にありません。\n"
                f"  fit() の price_col 引数か、wtp() の price_col 引数で\n"
                f"  正しい列名を指定してください。"
            )

        # 価格の符号化列を特定
        price_encoded = self._find_encoded_for(price_col)
        if price_encoded is None:
            raise ValueError(
                f"価格列 '{price_col}' に対応する符号化列が見つかりません。\n"
                f"  encode() で価格属性も符号化しているか確認してください。"
            )
        if len(price_encoded) > 1:
            raise NotImplementedError(
                f"価格属性が3水準以上で符号化されています: {price_encoded}\n"
                f"  現状のWTP計算は2水準の価格にのみ対応しています。"
            )

        price_enc_col = price_encoded[0]
        b_price = float(self.params[price_enc_col])
        p_price = float(self.ols.pvalues[price_enc_col])

        # 価格の元の水準を取得
        price_levels = sorted(self.df[price_col].dropna().unique().tolist())
        if len(price_levels) != 2:
            raise NotImplementedError(
                f"価格列 '{price_col}' は {len(price_levels)} 水準あります。\n"
                f"  現状のWTP計算は2水準にのみ対応しています。"
            )
        low_price, high_price = price_levels[0], price_levels[1]
        price_range = float(high_price - low_price)

        # ---- 警告①：価格係数の有意性（p ≥ 0.10） ----
        # 重複登録を防ぐ（wtp() は複数回呼ばれる可能性がある）
        already_cats = {d.category for d in self._diagnostics}
        if p_price >= 0.10 and "price_insignificant" not in already_cats:
            self._diagnostics.append(
                Diagnostic(
                    severity="中",
                    category="price_insignificant",
                    message=(
                        f"価格係数（p値 = {p_price:.3f}）が有意水準 0.10 を超えています。"
                        "WTP の計算は価格係数を分母に使うため、"
                        "係数が不確実だと WTP の信頼性も低くなります。"
                    ),
                    recommendation=(
                        "WTP の値は参考程度にとどめ、価格感度については別途検討してください。"
                        "価格の水準数を増やす、または回答者数を増やすことも有効です。"
                    ),
                )
            )
        price_is_significant = p_price < 0.10

        # WTP計算のスケール係数: price_range / b_price
        # WTP_attr = wtp_price_factor * b_attr として使う。
        # ※「評点1点の金額」は unit_rating_money() = price_range / (2*b_price) であり、
        #   こちらはその2倍（効果コーディングで属性変化時の効用差が 2*b_attr になるため）。
        wtp_price_factor = -(low_price - high_price) / b_price

        rows = []
        for col in self.encoded_columns:
            if col == price_enc_col:
                continue
            b = float(self.params[col])
            wtp_value = wtp_price_factor * b
            rows.append({"variable": col, "係数": b, "支払意思額": wtp_value})

        out = pd.DataFrame(rows).set_index("variable")
        out.index.name = "属性（符号化列名）"

        # ---- 警告②：WTP が価格レンジ × 2 を超える（外挿） ----
        # 価格係数が有意でない場合は「大」、有意な場合は「中」
        for _, row in out.iterrows():
            attr_col = row.name
            wtp_val = float(row["支払意思額"])
            threshold = price_range * 2
            cat_key = f"wtp_extrapolation_{attr_col}"
            if abs(wtp_val) > threshold and cat_key not in already_cats:
                sev = "中"
                self._diagnostics.append(
                    Diagnostic(
                        severity=sev,
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
        return out

    def unit_rating_money(self, *, price_col: Optional[str] = None) -> float:
        """
        評点1ポイントが何円（または何万円）に相当するかを返す。

        計算式: (price_max - price_min) / abs(price_coef * 2)
        単位は価格列の単位と同じ。例えば価格が万円単位なら、戻り値も万円単位。
        """
        price_col = price_col or self.price_col
        if price_col not in self.df.columns:
            raise ValueError(
                f"価格列 '{price_col}' が DataFrame にありません。"
            )
        price_encoded = self._find_encoded_for(price_col)
        if not price_encoded or len(price_encoded) != 1:
            raise ValueError(
                f"価格列 '{price_col}' に対応する符号化列が見つかりません。"
            )
        b_price = float(self.params[price_encoded[0]])
        price_levels = sorted(self.df[price_col].dropna().unique().tolist())
        if len(price_levels) != 2:
            raise ValueError(
                f"価格列 '{price_col}' は {len(price_levels)} 水準あります。2水準のみ対応しています。"
            )
        price_range = float(price_levels[1] - price_levels[0])
        return price_range / abs(b_price * 2)

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
        ...     "price_0":  [1, 0],   # 製品A: 6万円, 製品B: 10万円
        ...     "os_0":     [0, 1],
        ...     "camera_0": [1, 1],
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

    def plot_importance(self, **kwargs):
        """
        相対重要度の棒グラフを描画する。:func:`py4conjoint.plot.plot_importance`
        へのショートカット。

        Returns
        -------
        matplotlib.axes.Axes
        """
        from .plot import plot_importance
        return plot_importance(self, **kwargs)

    def plot_partworth(self, **kwargs):
        """
        部分効用（パートワース）の棒グラフを描画する。
        :func:`py4conjoint.plot.plot_partworth` へのショートカット。
        """
        from .plot import plot_partworth
        return plot_partworth(self, **kwargs)

    def plot_wtp(self, **kwargs):
        """
        WTPの棒グラフを描画する。:func:`py4conjoint.plot.plot_wtp` へのショートカット。
        """
        from .plot import plot_wtp
        return plot_wtp(self, **kwargs)

    # ---- 内部処理 ---------------------------------------------------------

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
        元の属性列名に対応する符号化列を見つける。
        ``encode()`` の命名規則 ``{original}_{インデックス}`` に基づく。
        """
        prefix = f"{original_col}_"
        cols = [c for c in self.encoded_columns if c.startswith(prefix)]
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

        5. **price_insignificant**（重大度：中、:meth:`wtp` 呼出時）
           価格係数の p値 ≥ 0.10。WTP計算の分母が不確実。

        6. **wtp_extrapolation**（重大度：中、:meth:`wtp` 呼出時）
           ``|WTP| > 価格レンジ × 2``。観測範囲外への外挿。
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
        if "回答者ID" in self.df.columns:
            n_resp = int(self.df["回答者ID"].nunique())
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


# ---------------------------------------------------------------------------
# 内部ヘルパー
# ---------------------------------------------------------------------------


def _rename_result_index(res: RegressionResults, rev_map: Dict[str, str]) -> RegressionResults:
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
            res.model.exog_names[:] = [
                rev_map.get(n, n) for n in res.model.exog_names
            ]
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
    import unicodedata
    return sum(
        2 if unicodedata.east_asian_width(c) in ("W", "F") else 1
        for c in s
    )


def _ljust_display(s: str, width: int) -> str:
    """表示幅ベースで左寄せパディングする。"""
    return s + " " * max(0, width - _display_width(s))


def _rjust_display(s: str, width: int) -> str:
    """表示幅ベースで右寄せパディングする。"""
    return " " * max(0, width - _display_width(s)) + s


def _detect_encoded_columns(
    df: pd.DataFrame,
    *,
    rating: str,
    reference_levels: Optional[Dict[str, object]] = None,
) -> List[str]:
    """
    符号化列を自動検出する。
    優先順位：
    1. reference_levels が与えられていれば、その属性名で始まる列を採用
    2. 値が ``{-1, 0, 1}`` の部分集合に収まる数値列を採用
       （ただし元の属性列・rating列は除外）
    """
    if reference_levels:
        cols: List[str] = []
        for attr in reference_levels.keys():
            prefix = f"{attr}_"
            for c in df.columns:
                if c.startswith(prefix):
                    cols.append(c)
        if cols:
            return cols

    # フォールバック：値の範囲で判定
    candidates: List[str] = []
    for c in df.columns:
        if c == rating:
            continue
        s = df[c]
        if not pd.api.types.is_numeric_dtype(s):
            continue
        vals = set(pd.Series(s.dropna().unique()).tolist())
        if vals.issubset({-1, 0, 1}) and vals != {0}:
            candidates.append(c)
    return candidates


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
