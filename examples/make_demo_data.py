"""教材デモ用の合成回答データ生成スクリプト（py4conjoint.choice）

このスクリプトは、overview_choice.ipynb で使うデモ用の「回答データ」を
合成的に生成する。生成されるのは Google Forms 形式の CSV で、実際の
学生アンケートの代わりにノートブックの動作確認・教材に用いる。

⚠️ 重要な注意
----------------
- ここで作るのは「実際の調査データ」ではなく、あらかじめ仮定した選好構造
  （下記 TRUE_* / UTIL_PRICE）から確率的に生成した合成データである。
  研究や実証の根拠には使えない。あくまで教材・動作確認用。
- 設計（design）は py4conjoint の design_choice_sets() が出力した純正の
  CSV（version 列付き）を読み込んで使う。このスクリプト内で設計を作り直さない。
  これにより「設計とデータの不整合」（回答と代替案の対応ずれ）を防ぐ。
  → 設計とデータは必ず同じ design に基づくこと（README / docstring 参照）。

入力（同じ examples/ ディレクトリに置く想定）
----------------------------------------------
- design_smartphone_cbc.csv     : 2水準設計（price 6/10, seed=1096）
- design_smartphone_3price.csv  : 3水準設計（price 6/8/10, seed=643）

出力
----
- responses_smartphone_google_30.csv  : 2水準・回答者30名
- responses_smartphone_3price_40.csv  : 3水準・回答者40名（価格の非線形選好）

使い方
------
    python make_demo_data.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).parent

# 表示用の OS 表記（設計データ内部の "apple"/"android" → 学生向けの見せ方）
OS_DISPLAY = {"apple": "Apple (iOS)", "android": "Android"}
LABEL_MAP = {1: "製品A", 2: "製品B", 3: "製品C"}

COL_GENDER = "【回答者属性１】あなたの性別を教えてください。"
COL_OS = "【回答者属性２】現在使っているスマートフォンのOSはどちらですか？"


def _question_text(design: pd.DataFrame, set_id: int) -> str:
    """設計から1設問ぶんの設問文（Forms にそのまま貼れる長文）を作る。"""
    g = design[design["choice_set_id"] == set_id].sort_values("alt_id")
    lines = [
        f"【設問{set_id}】次の3つのうち、最も購入したいスマートフォンはどれですか？",
        "",
    ]
    for _, r in g.iterrows():
        prod = LABEL_MAP[int(r["alt_id"])]
        lines.append(
            f"・{prod}：価格 {r['price']}万円 ｜ "
            f"OS {OS_DISPLAY[r['os']]} ｜ カメラ {r['camera']}"
        )
    return "\n".join(lines)


def _simulate_choices(
    design: pd.DataFrame,
    n_resp: int,
    util_os_apple: float,
    util_camera_high: float,
    sd_os: float,
    sd_camera: float,
    seed: int,
    util_price: dict[int, float] | None = None,
    price_coef: float | None = None,
    sd_price_coef: float = 0.0,
) -> pd.DataFrame:
    """仮定した選好構造（＋回答者ごとの個人差）から選択を確率的に生成する。

    各回答者ごとに係数を正規分布で揺らし、各設問内で条件付きロジット
    （softmax）に従って1つの代替案を選ぶ。

    価格の効用は2通りの指定ができる：
    - price_coef を渡す（線形）：効用 = price_coef × 価格（万円）。
      2水準デモで使用（価格に比例した素直な選好）。
    - util_price を渡す（水準ごと）：各価格水準に独立した効用を与える。
      3水準デモで使用（価格帯ごとに感応度が変わる非線形を表現）。
    """
    rng = np.random.default_rng(seed)
    n_sets = int(design["choice_set_id"].nunique())
    rows = []
    for _ in range(n_resp):
        if price_coef is not None:  # 線形（2水準デモ）
            pc = price_coef + rng.normal(0, sd_price_coef)
            price_util = lambda pr: pc * pr  # noqa: E731
        else:  # 水準ごと（3水準デモ・非線形）
            up = {k: v + rng.normal(0, 0.1) for k, v in util_price.items()}
            price_util = lambda pr: up[int(pr)]  # noqa: E731
        uo = util_os_apple + rng.normal(0, sd_os)
        uc = util_camera_high + rng.normal(0, sd_camera)
        choices = {}
        for s in range(1, n_sets + 1):
            g = design[design["choice_set_id"] == s].sort_values("alt_id")
            u = np.array(
                [
                    price_util(r["price"])
                    + uo * (r["os"] == "apple")
                    + uc * (r["camera"] == "高性能")
                    for _, r in g.iterrows()
                ]
            )
            p = np.exp(u - u.max())
            p /= p.sum()
            pick = rng.choice(g["alt_id"].to_numpy(), p=p)
            choices[s] = LABEL_MAP[int(pick)]
        gender = rng.choice(
            ["男性", "女性", "その他・回答しない"], p=[0.55, 0.42, 0.03]
        )
        used_os = rng.choice(
            ["Apple (iOS)", "Android", "その他・わからない"], p=[0.5, 0.47, 0.03]
        )
        rows.append(
            {"gender": gender, "used_os": used_os, **{f"Q{s}": choices[s] for s in choices}}
        )
    return pd.DataFrame(rows)


def _to_google_forms_csv(
    design: pd.DataFrame,
    resp: pd.DataFrame,
    out_path: Path,
    ts_start: str,
    ts_seed: int,
) -> None:
    """合成した選択を Google Forms 形式の CSV（BOMなしUTF-8）で書き出す。"""
    n_sets = int(design["choice_set_id"].nunique())
    rng = np.random.default_rng(ts_seed)
    base = pd.Timestamp(ts_start)
    ts = sorted(base + pd.Timedelta(minutes=int(m)) for m in rng.integers(0, 800, len(resp)))

    out = pd.DataFrame()
    out["タイムスタンプ"] = [t.strftime("%Y/%m/%d %H:%M:%S") for t in ts]
    out[COL_GENDER] = resp["gender"].to_numpy()
    out[COL_OS] = resp["used_os"].to_numpy()
    for s in range(1, n_sets + 1):
        out[_question_text(design, s)] = resp[f"Q{s}"].to_numpy()

    # Google Forms と同じ BOMなしUTF-8 で保存
    out.to_csv(out_path, index=False, encoding="utf-8")
    print(f"  書き出し: {out_path.name}（回答者 {len(out)}名）")


def main() -> None:
    # ---- 2水準（price 6/10）：30名・価格に比例した線形の選好 ----
    design2 = pd.read_csv(HERE / "design_smartphone_cbc.csv")
    resp2 = _simulate_choices(
        design2,
        n_resp=30,
        price_coef=-0.3,          # 価格 1万円あたり効用 -0.3（線形）
        sd_price_coef=0.05,
        util_os_apple=0.8,
        util_camera_high=0.6,
        sd_os=0.4,
        sd_camera=0.3,
        seed=20260615,
    )
    _to_google_forms_csv(
        design2,
        resp2,
        HERE / "responses_smartphone_google_30.csv",
        ts_start="2026-06-15 09:00:00",
        ts_seed=77,
    )

    # ---- 3水準（price 6/8/10）：40名・価格の非線形選好（8→10 で感応度が高い）----
    design3 = pd.read_csv(HERE / "design_smartphone_3price.csv")
    resp3 = _simulate_choices(
        design3,
        n_resp=40,
        util_price={6: 0.0, 8: -0.4, 10: -1.5},  # 区間別 WTP を実演するため非線形
        util_os_apple=0.8,
        util_camera_high=0.6,
        sd_os=0.4,
        sd_camera=0.3,
        seed=20260616,
    )
    _to_google_forms_csv(
        design3,
        resp3,
        HERE / "responses_smartphone_3price_40.csv",
        ts_start="2026-06-16 09:00:00",
        ts_seed=55,
    )

    print("完了。設計（design_*.csv）は py4conjoint 純正のものを使用しています。")


if __name__ == "__main__":
    main()
