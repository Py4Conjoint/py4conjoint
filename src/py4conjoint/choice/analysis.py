"""
analysis.py（choice 版）
========================
選択型コンジョイント分析（CBC）の **条件付きロジット（conditional logit）
推定と結果の保持** を担当するモジュール。

中心となるのは :func:`fit` 関数と :class:`ChoiceConjointResult` クラス。

>>> import py4conjoint.choice as pcc
>>> df_coded = pcc.encode(df, reference_levels={"brand": "dannon"})
>>> result = pcc.fit(
...     df_coded,
...     choice="choice",
...     choice_set_id_col="選択セットID",
...     encoded_columns=["price", "brand_hiland", "brand_yoplait"],
... )
>>> print(result.summary())     # 和文サマリー
>>> result.importance()         # 重要度
>>> result.wtp()                # WTP（限界支払意思額）
>>> result.market_share(products)

実装方針
--------
* 最尤推定は :func:`scipy.optimize.minimize`（BFGS・解析的勾配）による
  **自前実装**。教育目的のため、ブラックボックスにせず計算過程を
  ソースコードで追えるようにしている。
* 標準誤差はヘッセ行列（観測情報行列）の逆行列から計算する。
  回答者ID列がある場合は rating 版と同様の
  **クラスタロバスト標準誤差**（サンドイッチ推定量）を使う。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple, Union

import numpy as np
import pandas as pd
from scipy import stats
from scipy.optimize import minimize

# 表示ヘルパー・警告の構造化表現は rating 版と共通のものを使う
# （表示形式・警告 DataFrame の列構成を rating 版と完全に揃えるため）
from ..rating.analysis import (
    SEVERITY_ORDER,
    Diagnostic,
    _display_width,
    _format_price_level,
    _ljust_display,
    _price_level_utility_map,
    _price_segments_from_utilities,
    _rjust_display,
    _select_price_segment,
    _significance_stars,
)

# ---------------------------------------------------------------------------
# 公開API: fit関数
# ---------------------------------------------------------------------------

def fit(
    df: pd.DataFrame,
    *,
    choice: str = "choice",
    choice_set_id_col: str = "選択セットID",
    encoded_columns: Optional[List[str]] = None,
    reference_levels: Optional[Dict[str, object]] = None,
    price_col: str = "price",
    respondent_id_col: str = "回答者ID",
    cluster_se: bool = True,
) -> ChoiceConjointResult:
    """
    条件付きロジット（conditional logit）モデルを最尤推定し、
    結果オブジェクトを返す。

    モデル
    ------
    選択セット t の代替案 j が選ばれる確率を

    .. code-block:: text

        P(j | t) = exp(x_tj' β) / Σ_k exp(x_tk' β)

    とするモデル。係数 β は「その変数が1単位増えると、代替案の効用が
    どれだけ増えるか」を表す。切片は選択セット内で打ち消し合うため
    存在しない（ブランドダミーなどが実質的な定数の役割を果たす）。

    Parameters
    ----------
    df : pd.DataFrame
        long形式の選択データ。1行が「1つの選択セット内の1つの代替案」。
        :func:`encode` でカテゴリ属性をダミーコーディング済みであること。

    choice : str, default "choice"
        被説明変数の列名。選ばれた代替案が ``1``、それ以外が ``0``。
        各選択セットにはちょうど1つの ``1`` が必要。

    choice_set_id_col : str, default "選択セットID"
        選択セット（質問）を識別する列名。
        同じIDを持つ行が「同時に提示された代替案の組」を表す。
        選択セットごとの代替案数は揃っていなくてもよい。

    encoded_columns : list of str, optional
        説明変数として使う列のリスト（ダミー列と数値列の両方）。
        省略時は :func:`encode` が残したメタ情報からダミー列を取得し、
        さらに ``price_col`` が数値列として存在すれば先頭に加える。
        価格以外の数値変数（例：「特売中」フラグ）を使う場合は
        この引数で明示的に指定すること。

    reference_levels : dict, optional
        :func:`encode` に渡したのと同じ辞書。
        省略時は :func:`encode` が残したメタ情報から取得する。
        ``importance()`` の属性グルーピングに使う。

    price_col : str, default "price"
        価格列の名前。数値（連続）変数として説明変数に入れることを想定。
        WTP計算で使う。

    respondent_id_col : str, default "回答者ID"
        回答者ID列の名前。クラスタロバスト標準誤差と回答者数診断に使う。

    cluster_se : bool, default True
        ``True`` かつ回答者ID列が存在し回答者が2人以上いる場合、
        回答者IDでグループ化した **クラスタロバスト標準誤差** を使う。
        同じ回答者の複数の選択は独立でないため、通常の標準誤差では
        p値が過小（有意に出やすすぎ）になる。
        係数の推定値自体はどちらでも変わらない。

    Returns
    -------
    ChoiceConjointResult
        推定結果と各種解釈メソッドを持つオブジェクト。

    Raises
    ------
    ValueError
        choice列・選択セットID列・説明変数列が ``df`` にない場合。
        choice列に 0/1 以外の値がある場合。
        「ちょうど1つ選ばれている」を満たさない選択セットがある場合。
        代替案が1つしかない選択セットがある場合。

    Notes
    -----
    自動的に以下の **落とし穴チェック** を行い、警告を出す
    （詳細は :meth:`ChoiceConjointResult.warnings`）：

    * ``few_choice_sets``：選択セット数が説明変数数に対して少ない。
    * ``separation``：完全分離の疑い（係数の絶対値が異常に大きい、
      または収束失敗）。
    * ``unbalanced_choices``：特定の代替案位置への選択の極端な偏り。
    * ``price_sign_positive``：価格係数が正かつ有意。
    * ``few_respondents`` / ``independence_assumed``：rating 版と同じ。
    """
    # ---------- 入力チェック ----------
    if not isinstance(df, pd.DataFrame):
        raise TypeError(
            "df は pandas.DataFrame である必要があります。\n"
            f"  受け取った型: {type(df).__name__}"
        )
    for col, label in [(choice, "choice（選択結果）"),
                       (choice_set_id_col, "選択セットID")]:
        if col not in df.columns:
            raise ValueError(
                f"{label} の列 '{col}' が DataFrame にありません。\n"
                f"  存在する列: {list(df.columns)}\n"
                "  choice / choice_set_id_col 引数で正しい列名を指定してください。"
            )

    # ---------- encode() からのメタ情報を取得 ----------
    meta = df.attrs.get("py4conjoint", {}) if hasattr(df, "attrs") else {}
    if reference_levels is None:
        reference_levels = meta.get("reference_levels")

    if encoded_columns is None:
        enc_map: Dict[str, List[str]] = meta.get("encoded_columns") or {}
        encoded_columns = [c for cols in enc_map.values() for c in cols]
        if (
            price_col in df.columns
            and price_col not in enc_map
            and pd.api.types.is_numeric_dtype(df[price_col])
            and price_col not in encoded_columns
        ):
            encoded_columns = [price_col] + encoded_columns
        if not encoded_columns:
            raise ValueError(
                "説明変数が見つかりませんでした。\n"
                "  encode() でダミーコーディングを済ませているか確認するか、\n"
                "  encoded_columns 引数で列名を明示的に指定してください。\n"
                "  例: encoded_columns=['price', 'brand_hiland', 'brand_yoplait']"
            )

    missing = [c for c in encoded_columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"指定された説明変数列が DataFrame にありません: {missing}\n"
            f"  存在する列: {list(df.columns)}"
        )
    for c in encoded_columns:
        if not pd.api.types.is_numeric_dtype(df[c]):
            raise ValueError(
                f"説明変数列 '{c}' が数値ではありません。\n"
                "  カテゴリ属性は encode() でダミーコーディングしてから渡してください。"
            )

    # ---------- 欠損処理（欠損を含む選択セットは丸ごと除外） ----------
    n_before = len(df)
    use_cols = [choice] + list(encoded_columns)
    has_na = df[use_cols].isna().any(axis=1)
    if has_na.any():
        bad_sets = df.loc[has_na, choice_set_id_col].unique()
        df = df[~df[choice_set_id_col].isin(bad_sets)]
    n_dropped = n_before - len(df)
    if len(df) == 0:
        raise ValueError(
            "欠損を除外した結果、分析できる行が残りませんでした。\n"
            "  choice 列・説明変数列の欠損を確認してください。"
        )

    # ---------- choice 列の検証 ----------
    choice_vals = set(pd.Series(df[choice].dropna().unique()).tolist())
    if not choice_vals.issubset({0, 1}):
        raise ValueError(
            f"choice 列 '{choice}' は 0/1 で指定してください。\n"
            f"  見つかった値: {sorted(choice_vals, key=str)}\n"
            "  選ばれた代替案を 1、それ以外を 0 にしてください。"
        )

    # ---------- 選択セットごとに並べ替えて配列化 ----------
    df_sorted = df.sort_values(choice_set_id_col, kind="mergesort").reset_index(drop=True)
    X = df_sorted[encoded_columns].to_numpy(dtype=float)
    y = df_sorted[choice].to_numpy(dtype=float)
    codes = pd.factorize(df_sorted[choice_set_id_col])[0]
    starts = np.flatnonzero(np.r_[True, codes[1:] != codes[:-1]])
    counts = np.diff(np.r_[starts, len(codes)])
    n_sets = len(starts)
    choice_set_ids = df_sorted[choice_set_id_col].to_numpy()[starts]

    # 各選択セットの検証：代替案2つ以上、選択はちょうど1つ
    n_chosen = np.add.reduceat(y, starts)
    too_few = choice_set_ids[counts < 2]
    if len(too_few) > 0:
        raise ValueError(
            f"代替案が1つしかない選択セットがあります: {too_few[:5].tolist()}"
            f"{' ほか' if len(too_few) > 5 else ''}\n"
            "  条件付きロジットには各選択セットに2つ以上の代替案が必要です。\n"
            "  （「選ばない」を許す場合は「購入しない」という代替案を行として追加します）"
        )
    bad_choice = choice_set_ids[n_chosen != 1]
    if len(bad_choice) > 0:
        raise ValueError(
            f"選ばれた代替案がちょうど1つでない選択セットがあります: "
            f"{bad_choice[:5].tolist()}{' ほか' if len(bad_choice) > 5 else ''}\n"
            "  各選択セットで choice 列の 1 はちょうど1行にしてください。"
        )

    # ---------- 最尤推定（scipy.optimize.minimize, BFGS, 解析的勾配） ----------
    k = len(encoded_columns)
    chosen_mask = y == 1

    def _neg_loglik_and_grad(beta: np.ndarray) -> Tuple[float, np.ndarray]:
        v = X @ beta                                   # 各行（代替案）の効用
        # 数値安定化：選択セットごとに最大効用を引く（softmax の定石）
        vmax = np.maximum.reduceat(v, starts)
        vc = v - np.repeat(vmax, counts)
        ev = np.exp(vc)
        denom = np.add.reduceat(ev, starts)            # 選択セットごとの分母
        ll = vc[chosen_mask].sum() - np.log(denom).sum()
        p = ev / np.repeat(denom, counts)              # 各代替案の選択確率
        # 解析的勾配: Σ (x_chosen − Σ_j p_j x_j)
        grad = X[chosen_mask].sum(axis=0) - p @ X
        return -ll, -grad

    opt = minimize(
        _neg_loglik_and_grad,
        np.zeros(k),
        jac=True,
        method="BFGS",
        options={"gtol": 1e-5, "maxiter": 1000},
    )
    beta = opt.x
    loglik = -float(opt.fun)
    null_loglik = -float(np.log(counts).sum())  # β=0（全代替案が等確率）の対数尤度

    # ---------- 標準誤差 ----------
    # 観測情報行列（負の対数尤度のヘッセ行列）を解析的に計算する：
    #   H = Σ_t [ Σ_j p_j x_j x_j' − (Σ_j p_j x_j)(Σ_j p_j x_j)' ]
    v = X @ beta
    vmax = np.maximum.reduceat(v, starts)
    ev = np.exp(v - np.repeat(vmax, counts))
    denom = np.add.reduceat(ev, starts)
    p_hat = ev / np.repeat(denom, counts)
    Xw = X * p_hat[:, None]
    B = np.add.reduceat(Xw, starts, axis=0)            # Σ_j p_j x_j（セットごと）
    H = X.T @ Xw - B.T @ B

    try:
        cov = np.linalg.inv(H)
        hessian_singular = False
    except np.linalg.LinAlgError:
        cov = np.linalg.pinv(H)
        hessian_singular = True

    # クラスタロバスト標準誤差（回答者IDでグループ化）
    use_cluster = (
        cluster_se
        and respondent_id_col in df_sorted.columns
        and df_sorted[respondent_id_col].nunique() >= 2
    )
    if use_cluster:
        resp_per_row = df_sorted[respondent_id_col]
        n_resp_per_set = resp_per_row.groupby(codes).nunique()
        if (n_resp_per_set > 1).any():
            raise ValueError(
                "同じ選択セットに複数の回答者IDが含まれています。\n"
                f"  選択セットID列 '{choice_set_id_col}' は回答者×質問ごとに\n"
                "  一意になるようにしてください。"
            )
        # 選択セットごとのスコア（対数尤度の勾配への寄与）
        scores = X[chosen_mask] - B                    # (n_sets, k)
        resp_of_set = resp_per_row.to_numpy()[starts]
        cluster_codes = pd.factorize(resp_of_set)[0]
        m = int(cluster_codes.max()) + 1
        g = np.zeros((m, k))
        np.add.at(g, cluster_codes, scores)            # クラスタごとのスコア和
        correction = m / (m - 1)                       # 小標本補正
        meat = correction * (g.T @ g)
        cov = cov @ meat @ cov
        se_type = "cluster"
    else:
        se_type = "nonrobust"

    with np.errstate(invalid="ignore"):
        se = np.sqrt(np.diag(cov))
    zvals = np.divide(beta, se, out=np.full(k, np.nan), where=se > 0)
    pvals = 2.0 * stats.norm.sf(np.abs(zvals))

    result = ChoiceConjointResult(
        params=pd.Series(beta, index=encoded_columns, name="係数"),
        bse=pd.Series(se, index=encoded_columns, name="標準誤差"),
        pvalues=pd.Series(pvals, index=encoded_columns, name="p値"),
        df=df_sorted,
        choice=choice,
        choice_set_id_col=choice_set_id_col,
        encoded_columns=list(encoded_columns),
        reference_levels=reference_levels or {},
        price_col=price_col,
        loglik=loglik,
        null_loglik=null_loglik,
        n_obs=len(df_sorted),
        n_sets=n_sets,
        n_dropped=n_dropped,
        respondent_id_col=respondent_id_col,
        se_type=se_type,
        converged=bool(opt.success),
        n_iter=int(opt.nit),
        vcov=cov,
    )

    # 落とし穴チェック用の内部情報（選択された代替案の「セット内位置」の分布）
    positions = np.arange(len(y)) - np.repeat(starts, counts)
    chosen_positions = positions[chosen_mask]
    pos_counts = np.bincount(chosen_positions.astype(int))
    result._position_share = pos_counts / n_sets
    result._hessian_singular = hessian_singular

    # ---------- 落とし穴の自動検出 ----------
    result._run_diagnostics()

    return result


# ---------------------------------------------------------------------------
# 結果オブジェクト
# ---------------------------------------------------------------------------

@dataclass
class ChoiceConjointResult:
    """
    条件付きロジット推定の結果を保持し、解釈メソッドを提供するクラス。

    通常は :func:`fit` 関数経由で生成され、ユーザーが直接インスタンス化する
    ことはない。メソッド名・列名は rating 版の
    :class:`py4conjoint.rating.ConjointResult` と揃えてある。

    Attributes
    ----------
    params : pd.Series
        推定された係数。条件付きロジットに切片はない。
    bse : pd.Series
        標準誤差（``se_type`` 参照）。
    pvalues : pd.Series
        z検定による両側p値。
    df : pd.DataFrame
        分析に使ったデータ（選択セットIDで並べ替え済み）。
    choice : str
        選択結果（0/1）の列名。
    choice_set_id_col : str
        選択セットIDの列名。
    encoded_columns : list of str
        説明変数のリスト。
    reference_levels : dict
        :func:`encode` に渡された基準水準の辞書。
    price_col : str
        価格列の名前。WTP計算で使う。
    loglik : float
        最大化された対数尤度。
    null_loglik : float
        帰無モデル（全代替案が等確率）の対数尤度。
    n_obs : int
        分析に使った行数（代替案の延べ数）。
    n_sets : int
        選択セット数。
    n_dropped : int
        欠損により分析から除外された行数（選択セット単位で除外）。
    respondent_id_col : str
        回答者ID列の名前。
    se_type : str
        標準誤差の種類。``"cluster"``（回答者IDによるクラスタロバスト）
        または ``"nonrobust"``（ヘッセ行列に基づく通常の最尤標準誤差）。
    converged : bool
        最適化が収束したかどうか。
    n_iter : int
        最適化の反復回数。
    """
    params: pd.Series
    bse: pd.Series
    pvalues: pd.Series
    df: pd.DataFrame
    choice: str
    choice_set_id_col: str
    encoded_columns: List[str]
    reference_levels: Dict[str, object]
    price_col: str
    loglik: float
    null_loglik: float
    n_obs: int
    n_sets: int
    n_dropped: int = 0
    respondent_id_col: str = "回答者ID"
    se_type: str = "nonrobust"
    converged: bool = True
    n_iter: int = 0
    # 係数の分散共分散行列（多水準価格の同時 Wald 検定で使う）。
    vcov: Optional[np.ndarray] = None

    # 内部用：検出された警告（落とし穴）のリスト
    _diagnostics: List[Diagnostic] = field(default_factory=list)
    # 内部用：wtp() の重複登録防止用キーセット
    _warned_keys: Set[str] = field(default_factory=set)
    # 内部用：選択された代替案の「セット内位置」の分布（unbalanced_choices 用）
    _position_share: Optional[np.ndarray] = None
    # 内部用：ヘッセ行列が特異だったか（separation の兆候）
    _hessian_singular: bool = False

    # ---- 基本情報の取得 ----------------------------------------------------

    @property
    def pseudo_rsquared(self) -> float:
        """McFadden の擬似決定係数 ``1 − logL / logL0``。"""
        return 1.0 - self.loglik / self.null_loglik

    # ---- サマリー ---------------------------------------------------------

    def summary(self, *, slim: bool = True) -> str:
        """
        和文サマリーを返す（``print()`` で表示）。

        Parameters
        ----------
        slim : bool, default True
            ``True`` でコンパクトな和文サマリーを表示。
            ``False`` で標準誤差・z値を含む詳細な係数表を表示。

        Returns
        -------
        str
            人間が読みやすい和文サマリー。

        Examples
        --------
        >>> print(result.summary())
        >>> print(result.summary(slim=False))  # 標準誤差・z値も表示
        """
        lines: List[str] = []
        lines.append("=" * 60)
        lines.append("選択型コンジョイント分析の結果（和文サマリー）")
        lines.append("=" * 60)
        se_label = (
            f"クラスタロバスト（{self.respondent_id_col}）"
            if self.se_type == "cluster"
            else "通常（選択セット間の独立性を仮定）"
        )
        stat_rows = [
            ("観測数（行数）",   str(self.n_obs)),
            ("選択セット数",     str(self.n_sets)),
            ("説明変数の数",     str(len(self.encoded_columns))),
            ("対数尤度",         f"{self.loglik:.4f}"),
            ("擬似決定係数 R²（McFadden）", f"{self.pseudo_rsquared:.4f}"),
            ("標準誤差",         se_label),
        ]
        if not self.converged:
            stat_rows.append(("収束",     "失敗（結果は信頼できません）"))
        if self.n_dropped > 0:
            stat_rows.append(("欠損で除外", f"{self.n_dropped} 行"))
        max_label_w = max(_display_width(lbl) for lbl, _ in stat_rows)
        for label, value in stat_rows:
            lines.append(f"{_ljust_display(label, max_label_w + 1)}: {value}")
        lines.append("")

        # 係数表
        NAME_WIDTH = 25
        STAR_WIDTH = 6
        lines.append("【推定された係数（部分効用 part-worth）】")
        if slim:
            lines.append(
                "  "
                + _ljust_display("変数", NAME_WIDTH)
                + " " + _rjust_display("係数", 10)
                + " " + _rjust_display("p値", 10)
                + "  " + "有意性"
            )
            lines.append(f"  {'-' * NAME_WIDTH} {'-' * 10} {'-' * 10}  {'-' * STAR_WIDTH}")
            for name in self.params.index:
                coef = self.params[name]
                p = self.pvalues[name]
                star = _significance_stars(p)
                lines.append(
                    f"  {_ljust_display(str(name), NAME_WIDTH)} {coef:>10.4f} {p:>10.4f}  {star}"
                )
        else:
            lines.append(
                "  "
                + _ljust_display("変数", NAME_WIDTH)
                + " " + _rjust_display("係数", 10)
                + " " + _rjust_display("標準誤差", 10)
                + " " + _rjust_display("z値", 10)
                + " " + _rjust_display("p値", 10)
                + "  " + "有意性"
            )
            lines.append(
                f"  {'-' * NAME_WIDTH} {'-' * 10} {'-' * 10} {'-' * 10} {'-' * 10}  {'-' * STAR_WIDTH}"
            )
            for name in self.params.index:
                coef = self.params[name]
                se = self.bse[name]
                z = coef / se if se > 0 else float("nan")
                p = self.pvalues[name]
                star = _significance_stars(p)
                lines.append(
                    f"  {_ljust_display(str(name), NAME_WIDTH)} {coef:>10.4f}"
                    f" {se:>10.4f} {z:>10.4f} {p:>10.4f}  {star}"
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
        return "\n".join(lines)

    def __repr__(self) -> str:  # pragma: no cover
        return self.summary()

    def _repr_html_(self) -> str:
        """Jupyter Notebook 向け HTML 表示（rating 版と同じスタイル）。"""
        from html import escape

        _td = "padding:3px 14px 3px 4px;"
        _th = _td + "border-bottom:1px solid #888;font-weight:bold;"
        _tbl = "border-collapse:collapse;margin-bottom:0.8em;"

        def td(txt, align="left"):
            return f'<td style="{_td}text-align:{align};">{escape(str(txt))}</td>'

        def th(txt, align="left"):
            return f'<th style="{_th}text-align:{align};">{escape(str(txt))}</th>'

        se_label = (
            f"クラスタロバスト（{self.respondent_id_col}）"
            if self.se_type == "cluster"
            else "通常（選択セット間の独立性を仮定）"
        )
        stat_rows = [
            ("観測数（行数）",   str(self.n_obs)),
            ("選択セット数",     str(self.n_sets)),
            ("説明変数の数",     str(len(self.encoded_columns))),
            ("対数尤度",         f"{self.loglik:.4f}"),
            ("擬似決定係数 R²（McFadden）", f"{self.pseudo_rsquared:.4f}"),
            ("標準誤差",         se_label),
        ]
        if not self.converged:
            stat_rows.append(("収束", "失敗（結果は信頼できません）"))
        if self.n_dropped > 0:
            stat_rows.append(("欠損で除外", f"{self.n_dropped} 行"))

        p = ['<div>',
             '<p><strong>選択型コンジョイント分析の結果</strong></p>']
        p.append(f'<table style="{_tbl}">')
        for lbl, val in stat_rows:
            p.append(f'<tr>{td(lbl)}{td(val)}</tr>')
        p.append('</table>')

        p.append('<p><strong>【推定された係数（部分効用 part-worth）】</strong></p>')
        p.append(f'<table style="{_tbl}">')
        p.append('<tr>'
                 + th("変数") + th("係数", "right")
                 + th("p値", "right") + th("有意性", "right")
                 + '</tr>')
        for name in self.params.index:
            c = self.params[name]
            pv = self.pvalues[name]
            star = _significance_stars(pv)
            p.append('<tr>'
                     + td(name) + td(f"{c:.4f}", "right")
                     + td(f"{pv:.4f}", "right") + td(star, "right")
                     + '</tr>')
        p.append('</table>')
        p.append('<p style="font-size:0.85em;color:#888;">'
                 '有意水準: *** p&lt;0.001&nbsp; ** p&lt;0.01&nbsp;'
                 ' * p&lt;0.05&nbsp; . p&lt;0.1</p>')

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
            それらのリスト。省略時はすべての警告を返す。
        category : str または list of str, optional
            カテゴリでフィルタする。
            利用可能な値：``"few_choice_sets"``, ``"separation"``,
            ``"unbalanced_choices"``, ``"price_sign_positive"``,
            ``"few_respondents"``, ``"independence_assumed"``,
            ``"price_insignificant"``, ``"wtp_extrapolation"``。
            省略時はすべて返す。
        as_dataframe : bool, default True
            ``True`` なら ``pd.DataFrame``（列：severity, category, message,
            recommendation）として返す。
            ``False`` なら :class:`Diagnostic` オブジェクトのリストを返す。

        Returns
        -------
        pd.DataFrame または list of Diagnostic
        """
        diags = list(self._diagnostics)
        if severity is not None:
            sev_list = [severity] if isinstance(severity, str) else list(severity)
            diags = [d for d in diags if d.severity in sev_list]
        if category is not None:
            cat_list = [category] if isinstance(category, str) else list(category)
            diags = [d for d in diags if d.category in cat_list]
        diags = sorted(diags, key=lambda d: SEVERITY_ORDER.get(d.severity, 99))

        if not as_dataframe:
            return diags
        if not diags:
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

        * ダミーコーディングした属性：基準水準の効用は 0、
          非基準水準の効用は各係数なので、
          効用範囲は ``max(0, 係数...) − min(0, 係数...)``。
        * 数値（連続）属性（価格など）：効用範囲は
          ``|係数| × （データ上の最大値 − 最小値）``。

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
        水準レンジの異なる調査間で重要度を比較してはいけない。
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
        price_segment: Optional[Any] = None,
    ) -> pd.DataFrame:
        """
        各非価格変数の **WTP（限界支払意思額、Marginal Willingness to Pay）**
        を計算する。

        定義
        ----
        条件付きロジットでは、価格係数 ``β_price`` が「貨幣1単位の効用」
        （の符号反転）を表すため、

        .. code-block:: text

            MWTP = −β_attr / β_price

        で計算する。ダミー変数なら「基準水準からその水準に変えるときに
        追加で支払ってもよい金額」、数値変数なら「その変数1単位あたりの
        支払意思額」を表す。

        価格列の指定（rating 版と統一）
        --------------------------------
        ``price_col`` には **数値（6, 10 など）が入った数値列のラベル**
        （例：``"price"``）を渡す。価格を ``encode()`` でダミーコーディング
        した場合（``price_6`` など）も、``price_col`` にはダミー列名ではなく
        元の数値列名を渡すこと。どの符号化列が価格かは、数値列の水準と
        ``encode()`` の命名規則から構成的に特定する（``startswith`` による
        前方一致は使わないため ``price_range_high`` のような別属性の列を
        誤検出しない）。

        区間別 WTP（method 引数）
        --------------------------
        価格をダミーコーディングすると、各価格水準の効用が独立に推定される。
        これを活かし、価格が3水準以上のときは **隣接する価格水準の区間ごと**
        に別々の傾き（価格感応度）で WTP を計算する（``method="segment"``、
        デフォルト）。価格帯によって価格感応度が変わるため、WTP も区間ごとに
        変わるのが自然である。

        * ``method="segment"``（デフォルト）：区間別。価格が3水準以上のとき、
          戻り値に ``価格区間`` 列が付き、属性 × 区間の行が出力される。
        * ``method="linear"``：価格効用が線形だと仮定し、1本の傾きで計算する
          （従来方式・教材用）。
        * 価格が数値（線形）変数として説明変数に入っている場合、または価格が
          2水準のときは、区間が1つだけなので ``method`` によらず単一値を返す。

        Parameters
        ----------
        price_col : str, optional
            価格の数値列名。``fit`` で設定した値を上書きしたい場合に使う。
        method : {"segment", "linear"}, default "segment"
            区間別か線形近似か。上記参照。
        price_segment : str または (low, high), optional
            特定の価格区間の WTP だけを取り出したいときに指定する。
            ラベル文字列（例：``"6〜8"``）または ``(6, 8)`` のタプル。

        Returns
        -------
        pd.DataFrame
            列：``係数``, ``限界支払意思額``（価格と同じ単位）。
            区間別（3水準以上 × ``method="segment"``）のときは先頭に
            ``価格区間`` 列が付く。インデックスは価格以外の説明変数名。

        Raises
        ------
        ValueError
            価格列が指定されていない場合、または価格に対応する説明変数が
            見つからない場合。``method`` が不正な場合。

        Notes
        -----
        * 価格列の単位（万円・千円・円など）に依存するので、結果の単位は
          元データと同じ。
        * 価格係数 ``β_price`` が負（価格が高いほど選ばれにくい）である
          ことが前提。正かつ有意の場合は ``price_sign_positive`` 警告が
          :func:`fit` 時に出ている。
        * 計算は「貨幣の限界効用が一定（価格効用が線形）・所得効果なし」を
          仮定している。
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

        # ---- 価格の説明変数を特定する ----
        # (A) 数値（線形）変数として encoded_columns に入っている場合
        # (B) ダミーコーディングされ price_col の各水準ダミーが入っている場合
        numeric_price = (
            price_col in self.encoded_columns
            and price_col in self.df.columns
            and pd.api.types.is_numeric_dtype(self.df[price_col])
        )
        price_encoded = (
            None if numeric_price else self._price_dummy_columns(price_col)
        )
        if not numeric_price and not price_encoded:
            raise ValueError(
                f"価格列 '{price_col}' に対応する説明変数が見つかりません。\n"
                "  価格は数値（線形）変数として encoded_columns に含めるか、\n"
                "  encode() でダミーコーディングして fit() してください。\n"
                "  （ダミーの場合も price_col には元の数値列名を渡します。）"
            )

        # ---- 価格水準と各水準の効用、価格レンジ、価格区間 ----
        if numeric_price:
            price_set = {price_col}
            b_price = float(self.params[price_col])
            p_price = float(self.pvalues[price_col])
            price_vals = self.df[price_col].dropna()
            levels = sorted(float(x) for x in price_vals.unique())
            util = {}
            # 数値線形は区間が1本（傾き = β_price）のみ
            segs = [{
                "low": float(price_vals.min()),
                "high": float(price_vals.max()),
                "label": (
                    f"{_format_price_level(price_vals.min())}〜"
                    f"{_format_price_level(price_vals.max())}"
                ),
                "slope": b_price,
            }]
        else:
            price_set = set(price_encoded)
            p_price = self._price_pvalue(price_encoded)
            levels, util = _price_level_utility_map(
                self.df, self.params, price_col, price_encoded, base_zero=True
            )
            segs = _price_segments_from_utilities(levels, util)

        low_price, high_price = levels[0], levels[-1]
        price_range = float(high_price - low_price)

        # ---- 価格係数の有意性の警告（p ≥ 0.10） ----
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
                        "選択セット数を増やす、または価格レンジを広げることも有効です。"
                    ),
                )
            )

        # ---- 区間別か単一値かを決める ----
        multi_segment = (method == "segment") and (len(segs) >= 2)
        if method == "linear" and len(segs) >= 2:
            # 線形近似：全水準を1本の傾きにまとめ直す
            slope_lin = float(np.polyfit(
                np.array(levels, dtype=float),
                np.array([util[lv] for lv in levels], dtype=float),
                1,
            )[0])
            segs = [{
                "low": low_price, "high": high_price,
                "label": (
                    f"{_format_price_level(low_price)}〜"
                    f"{_format_price_level(high_price)}"
                ),
                "slope": slope_lin,
            }]
            cat_key = "wtp_price_linear_approx"
            if cat_key not in already_cats:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="wtp_price_linear_approx",
                        message=(
                            f"価格が {len(levels)} 水準ありますが、"
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

        # ---- WTP を計算（MWTP = -β_attr / 区間の傾き） ----
        rows = []
        for seg in segs:
            money_per_utility = -1.0 / seg["slope"]
            for col in self.encoded_columns:
                if col in price_set:
                    continue
                b = float(self.params[col])
                row = {"variable": col, "係数": b,
                       "限界支払意思額": b * money_per_utility}
                if multi_segment:
                    row["価格区間"] = seg["label"]
                rows.append(row)
        out = pd.DataFrame(rows).set_index("variable")
        out.index.name = "属性（符号化列名）"
        if multi_segment:
            out = out[["価格区間", "係数", "限界支払意思額"]]

        # ---- 警告：WTP が価格レンジ × 2 を超える（外挿） ----
        if price_range > 0:
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
                                "この WTP 値をそのまま「消費者は X 円まで払う」と"
                                "解釈するのは危険です。"
                                "「調査した価格レンジ内での相対的な選好」として"
                                "解釈するにとどめてください。"
                            ),
                        )
                    )

        out.attrs["price_col"] = price_col
        out.attrs["p_price"] = p_price
        out.attrs["price_low"] = low_price
        out.attrs["price_high"] = high_price
        out.attrs["price_range"] = price_range
        out.attrs["method"] = method
        return out

    def _price_dummy_columns(self, price_col: str) -> Optional[List[str]]:
        """
        価格列に対応する **ダミー符号化列** を構成的に特定する。

        ``encode()`` が ``df.attrs`` に残したメタ情報（属性→符号化列）を
        優先し、なければ数値水準から ``{price_col}_{水準}`` を構成して
        ``encoded_columns`` と照合する。``startswith`` の前方一致は使わない
        ため、``price_range_high`` のような別属性の列を誤検出しない。
        """
        meta = (
            self.df.attrs.get("py4conjoint", {}) if hasattr(self.df, "attrs") else {}
        )
        enc_map = meta.get("encoded_columns") or {}
        mapped = enc_map.get(price_col)
        if mapped:
            cols = [c for c in mapped if c in self.encoded_columns]
            return cols or None
        if price_col not in self.df.columns:
            return None
        levels = list(pd.Series(self.df[price_col].dropna().unique()))
        names = {f"{price_col}_{lv}" for lv in levels}
        cols = [c for c in self.encoded_columns if c in names]
        return cols or None

    def _price_pvalue(self, price_encoded: List[str]) -> float:
        """
        価格の有意性の p値を返す。

        ダミー符号化列が1本（2水準）なら係数の z 検定の p値。
        複数（3水準以上）なら「すべての価格係数 = 0」の同時 Wald 検定
        （χ² 近似）の p値を返す。先頭列の p値だけでは多水準価格の
        有意性を正しく判定できないため。
        """
        if len(price_encoded) == 1:
            return float(self.pvalues[price_encoded[0]])
        idx = [self.encoded_columns.index(c) for c in price_encoded]
        beta = self.params.to_numpy()[idx]
        if self.vcov is not None:
            V = self.vcov[np.ix_(idx, idx)]
            try:
                wald = float(beta @ np.linalg.solve(V, beta))
                return float(stats.chi2.sf(wald, len(idx)))
            except np.linalg.LinAlgError:
                pass
        # フォールバック：最小 p 値の Bonferroni 調整（保守的）
        ps = [float(self.pvalues[c]) for c in price_encoded]
        return float(min(1.0, min(ps) * len(ps)))

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
            製品ごとの説明変数列を含むDataFrame。
            インデックスを製品名にしておくと結果が読みやすい。
            各製品行は ``encoded_columns`` の各列に値を持つ必要がある
            （ダミー列は 0/1、価格などの数値列は実際の値）。

        method : {"logit", "max", "share_of_preference"}, default "logit"
            シェア計算方法。

            * ``"logit"`` または ``"share_of_preference"``: ロジット式。
              ``share_i = exp(u_i) / Σ exp(u_j)``。
              条件付きロジットの選択確率そのもので、最も自然な方法。
            * ``"max"``: 最大効用ルール。
              最大効用の製品にシェア1、他は0。**仮定**：消費者が完全合理的。

        Returns
        -------
        pd.Series
            製品名 → シェア（0〜1）の Series。合計は1になる。

        Examples
        --------
        >>> products = pd.DataFrame({
        ...     "price":         [100, 150],
        ...     "brand_hiland":  [1, 0],
        ...     "brand_yoplait": [0, 1],
        ... }, index=["製品A", "製品B"])
        >>> result.market_share(products)
        """
        if not isinstance(products, pd.DataFrame):
            raise TypeError("products は pandas.DataFrame で渡してください。")

        missing = [c for c in self.encoded_columns if c not in products.columns]
        if missing:
            raise ValueError(
                f"products に必要な列がありません: {missing}\n"
                f"  必要な列: {self.encoded_columns}"
            )

        u = products[self.encoded_columns].to_numpy(dtype=float) @ self.params.to_numpy()

        if method in ("logit", "share_of_preference"):
            # 数値安定化のため最大値を引く（softmax）
            u_shift = u - np.max(u)
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

    # ---- 可視化 -----------------------------------------------------------

    def plot_importance(self, **kwargs: Any) -> Any:
        """
        重要度の棒グラフを描画する。
        :func:`py4conjoint.choice.plot.plot_importance` へのショートカット。

        Returns
        -------
        matplotlib.axes.Axes
        """
        from .plot import plot_importance
        return plot_importance(self, **kwargs)

    def plot_partworth(self, **kwargs: Any) -> Any:
        """
        部分効用（パートワース）の棒グラフを描画する。
        :func:`py4conjoint.choice.plot.plot_partworth` へのショートカット。
        """
        from .plot import plot_partworth
        return plot_partworth(self, **kwargs)

    def plot_wtp(self, **kwargs: Any) -> Any:
        """
        WTPの棒グラフを描画する。
        :func:`py4conjoint.choice.plot.plot_wtp` へのショートカット。
        """
        from .plot import plot_wtp
        return plot_wtp(self, **kwargs)

    # ---- 内部処理 ---------------------------------------------------------

    def _attribute_ranges(self) -> Dict[str, float]:
        """
        各属性について、部分効用の最大値 − 最小値（効用範囲）を計算する。

        * ダミーコーディングした属性（``reference_levels`` に登録あり）：
          水準の効用 = {0（基準水準）} ∪ {各ダミー係数}。
        * それ以外の列（価格などの数値変数）：``|係数| × 観測レンジ``。
        """
        known = sorted(self.reference_levels.keys(), key=len, reverse=True)
        groups: Dict[str, List[str]] = {}
        for c in self.encoded_columns:
            matched = None
            for a in known:
                if c.startswith(f"{a}_"):
                    matched = a
                    break
            if matched is None:
                matched = c  # 数値変数はその列自体を属性とみなす
            groups.setdefault(matched, []).append(c)

        ranges: Dict[str, float] = {}
        for attr, cols in groups.items():
            if attr in self.reference_levels:
                utils = np.append(
                    [float(self.params[c]) for c in cols], 0.0  # 基準水準の効用
                )
                ranges[attr] = float(utils.max() - utils.min())
            else:
                col = cols[0]
                vals = self.df[col].dropna()
                ranges[attr] = abs(float(self.params[col])) * float(
                    vals.max() - vals.min()
                )
        return ranges

    def _run_diagnostics(self) -> None:
        """
        落とし穴の自動検出を行い、:class:`Diagnostic` のリストとして蓄積する。

        検出される警告
        ---------------
        1. **separation**（重大度：大）
           完全分離の疑い。最適化が収束しなかった、係数の絶対値が
           異常に大きい（> 10）、またはヘッセ行列が特異。
           ある変数が選択結果をほぼ完全に予測している可能性。

        2. **few_choice_sets**（重大度：大 or 中）
           選択セット数／説明変数数の比率が低い。
           比率 < 5 で「大」、< 10 で「中」。

        3. **unbalanced_choices**（重大度：中）
           選択が特定の代替案位置（1番目・2番目…）に極端に偏っている
           （最大シェア ≥ 80%）。回答者がよく考えずに同じ位置を
           選んでいる可能性。

        4. **price_sign_positive**（重大度：中）
           価格係数が正かつ有意（p < 0.10）。
           「価格が高いほど選ばれやすい」という直感に反する。

        5. **few_respondents**（重大度：大 or 中）
           回答者が1人なら「大」、2〜4人なら「中」。
           回答者ID列がある場合のみ判定する。

        6. **independence_assumed**（重大度：中）
           回答者ID列が見つからず、選択セット間の独立性を仮定した
           標準誤差を使用している。

        7. **price_insignificant** / **wtp_extrapolation**
           （重大度：中、:meth:`wtp` 呼出時）rating 版と同じ。
        """
        # 1) 完全分離の疑い
        max_abs_beta = float(np.nanmax(np.abs(self.params.to_numpy())))
        if (not self.converged) or max_abs_beta > 10 or self._hessian_singular:
            reasons = []
            if not self.converged:
                reasons.append("最適化が収束しませんでした")
            if max_abs_beta > 10:
                reasons.append(
                    f"係数の絶対値が異常に大きいです（最大 {max_abs_beta:.1f}）"
                )
            if self._hessian_singular:
                reasons.append("ヘッセ行列が特異です（標準誤差が計算できません）")
            self._diagnostics.append(
                Diagnostic(
                    severity="大",
                    category="separation",
                    message=(
                        "完全分離の疑いがあります（" + "、".join(reasons) + "）。"
                        "ある変数が選択結果をほぼ完全に予測している可能性があり、"
                        "係数と標準誤差は信頼できません。"
                    ),
                    recommendation=(
                        "選択セット数を増やす、水準の組み合わせを見直す、"
                        "または該当する変数を説明変数から外すことを検討してください。"
                    ),
                )
            )

        # 2) 選択セット数／説明変数比のチェック
        n_vars = len(self.encoded_columns)
        if n_vars > 0:
            ratio = self.n_sets / n_vars
            if ratio < 5:
                self._diagnostics.append(
                    Diagnostic(
                        severity="大",
                        category="few_choice_sets",
                        message=(
                            f"選択セット数（{self.n_sets}）が説明変数数（{n_vars}）の"
                            f"{ratio:.1f}倍しかなく、推定が非常に不安定です。"
                        ),
                        recommendation=(
                            "回答者または1人あたりの質問数を増やしてください。"
                            "目安として選択セット数は説明変数数の10倍以上が望ましいです。"
                        ),
                    )
                )
            elif ratio < 10:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="few_choice_sets",
                        message=(
                            f"選択セット数（{self.n_sets}）が説明変数数（{n_vars}）の"
                            f"{ratio:.1f}倍で、やや少なめです。"
                        ),
                        recommendation=(
                            "可能なら回答者または1人あたりの質問数を増やしてください。"
                            "目安として選択セット数は説明変数数の10倍以上が望ましいです。"
                        ),
                    )
                )

        # 3) 特定の代替案位置への選択の偏り
        if self._position_share is not None and len(self._position_share) >= 2:
            max_pos = int(np.argmax(self._position_share))
            max_share = float(self._position_share[max_pos])
            if max_share >= 0.8:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="unbalanced_choices",
                        message=(
                            f"選択の {max_share * 100:.0f}% が選択セット内の "
                            f"{max_pos + 1} 番目の代替案に集中しています。"
                            "回答者がよく考えずに同じ位置を選び続けている"
                            "（ストレートライニング）可能性があります。"
                        ),
                        recommendation=(
                            "代替案の提示順をランダム化しているか、"
                            "回答の質に問題がないかを確認してください。"
                        ),
                    )
                )

        # 4) 価格係数の符号（有意な場合のみ）
        if self.price_col in self.encoded_columns:
            b_price = float(self.params[self.price_col])
            p_price = float(self.pvalues[self.price_col])
            if b_price > 0 and p_price < 0.10:
                self._diagnostics.append(
                    Diagnostic(
                        severity="中",
                        category="price_sign_positive",
                        message=(
                            f"価格係数（{self.price_col} = {b_price:.4f}）が正です。"
                            "「価格が高いほど選ばれやすい」という直感に反します。"
                        ),
                        recommendation=(
                            "価格列の値や単位に誤りがないか、"
                            "またはデータ品質に問題がないか確認してください。"
                        ),
                    )
                )

        # 5) 回答者数の確認 / 6) 回答者ID列がない場合
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
            self._diagnostics.append(
                Diagnostic(
                    severity="中",
                    category="independence_assumed",
                    message=(
                        f"回答者ID列 '{self.respondent_id_col}' が見つからないため、"
                        "選択セット間の独立性を仮定した標準誤差を使用しています。"
                        "同じ回答者の複数の選択が含まれる場合、p値が過小"
                        "（有意に出やすく）になります。"
                    ),
                    recommendation=(
                        "回答者を識別できる列がある場合は、fit() の "
                        "respondent_id_col 引数でその列名を指定してください。"
                    ),
                )
            )
