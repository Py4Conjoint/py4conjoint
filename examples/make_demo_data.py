"""教材デモ用の合成データ生成スクリプト（py4conjoint）

examples/ のノートブック（rating 用・choice 用）で使うデモデータを、この
スクリプト1本ですべて生成する。設計（design_*.csv）も回答データ
（responses_*.csv）も、ここで作る。

⚠️ 重要な注意
----------------
- ここで作るのは「実際の調査データ」ではなく、あらかじめ仮定した選好構造
  （下記 TRUE_* 定数）から確率的に生成した合成データである。
  研究や実証の根拠には使えない。あくまで教材・動作確認用。
- 設計は py4conjoint 自身の ``design_profiles()`` / ``design_choice_sets()``
  で作り、その設計をそのまま使って回答を合成する。設計とデータは必ず同じ
  design に基づくこと（回答と代替案の対応ずれを防ぐため）。

「同じ100人が4通りの聞き方に答える」設定
------------------------------------------
回答者の生成（:func:`_make_respondents`）と回答のシミュレーション
（:func:`_simulate_ratings` / :func:`_simulate_choices`）を分離してある。
100人分の個人別部分効用と回答者属性は一度だけ生成し、4つのデータすべてで
同じものを使う。つまり「同じ100人・同じ選好・4通りの聞き方」という設定に
なっており、ノートブックの付録で評点型と選択型を比較できる。

価格の効用は水準ごとに与えており（TRUE_UTIL_PRICE）、意図的に **非線形**
にしてある（6→8 の下落幅より 8→10 の下落幅が大きい）。区間別 WTP に意味を
持たせるためである。2price 版は、この効用のうち {6, 10} の部分だけを使う。

出力（すべて examples/ 直下）
-------------------------------
設計

- design_rating_2price.csv  : 評点型・price 2水準・4プロファイル
- design_rating_3price.csv  : 評点型・price 3水準・6プロファイル
- design_choice_2price.csv  : 選択型・price 2水準・6設問 × 3選択肢
- design_choice_3price.csv  : 選択型・price 3水準・6設問 × 3選択肢

回答データ（Microsoft Forms 形式の CSV / BOM付きUTF-8）

- responses_rating_2price.csv
- responses_rating_3price.csv
- responses_choice_2price.csv
- responses_choice_3price.csv

回答データは4つとも ``forms="microsoft"`` で読み込める。管理列（ID／開始時刻
／完了時刻／メール／名前／最終変更時刻）の構成は、実際の Microsoft Forms の
出力である tests/data/forms_cbc_smartphone_real.xlsx に合わせてある。ただし
設問の列名については、実際の Forms 出力に付く末尾の改行は付けていない
（ノートブックで ``respondent_cols`` に列名を書き写しやすくするため）。

使い方
------
    python make_demo_data.py

生成後、8ファイルが実際に読めることをスクリプト自身が検証する
（:func:`_verify`）。「生成はできたが読めない」という事故を防ぐため。
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

# リポジトリを clone しただけの状態（未インストール）でも動くようにする
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import py4conjoint.choice as pcc
import py4conjoint.rating as pcr

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# 属性と水準（4つのデータで共通。変わるのは price の水準数だけ）
# ---------------------------------------------------------------------------
PRICE_LEVELS_2 = [6, 10]           # 単位：万円
PRICE_LEVELS_3 = [6, 8, 10]
OS_LEVELS = ["android", "apple"]
CAMERA_LEVELS = ["標準", "高性能"]


def _attributes(price_levels: list[int]) -> dict[str, list]:
    """price の水準だけを差し替えた属性辞書を返す。"""
    return {"price": price_levels, "os": OS_LEVELS, "camera": CAMERA_LEVELS}


# ---------------------------------------------------------------------------
# 設計のサイズ
# ---------------------------------------------------------------------------
# 評点型：D 最適解が同時に水準バランスも満たす、最小のプロファイル数。
# 2水準 → 4（8候補から4個を選ぶ D 最適解は2通りで、どちらも price 2/2、
# os 2/2、camera 2/2 の完全バランス・完全直交）。
# 3水準 → 6（D 最適解がすべて price 2/2/2、os 3/3、camera 3/3）。
# suggest_n_profiles() の推奨値（2水準 → 6、3水準 → 8）とは意図的に異なる。
# 推奨値では D 最適性と水準バランスが両立せず、check_design() が [大] の
# バランス警告を出すため、教材ではバランスを優先した。
N_PROFILES_2PRICE = 4
N_PROFILES_3PRICE = 6

# 選択型：既存の design_smartphone_cbc.csv と同じ構成
N_SETS = 6
N_ALTS = 3

N_RESPONDENTS = 100

# ---------------------------------------------------------------------------
# 乱数シード（再現性のため、すべてここに集約する）
# ---------------------------------------------------------------------------
SEED_RESPONDENTS = 20260701       # 100人の個人差（4データで共通）

# 設計の seed は check_design() の警告が少なくなるものを選んである。
# 特に price の水準バランスを優先した（WTP の説明で使う属性のため）。
# rating 2price だけは理由が別で、この seed だと design_profiles() の出力が
# 授業で学生に手作りさせている設計（price 6/10 × os android/apple ×
# camera 標準/高性能 の半分実施要因計画）と同じプロファイル集合になる。
# ノートブックで「手で作った設計と同じものが選ばれた」と示せるようにするため。
SEED_DESIGN_RATING_2PRICE = 4
SEED_DESIGN_RATING_3PRICE = 643
SEED_DESIGN_CHOICE_2PRICE = 180
SEED_DESIGN_CHOICE_3PRICE = 94

SEED_ANSWER_RATING_2PRICE = 20260711   # 評点の誤差項
SEED_ANSWER_RATING_3PRICE = 20260712
SEED_ANSWER_CHOICE_2PRICE = 20260713   # 選択の確率的な引き
SEED_ANSWER_CHOICE_3PRICE = 20260714

# 回答時刻の散らばり。4ファイルで別々の値にする（同じ値だと4ファイルの
# 回答時刻がそろってしまうが、エラーにはならず気づきにくいため、
# 計算式ではなくファイルごとの定数にしてある）。
SEED_TIMESTAMP_RATING_2PRICE = 20260721
SEED_TIMESTAMP_RATING_3PRICE = 20260722
SEED_TIMESTAMP_CHOICE_2PRICE = 20260723
SEED_TIMESTAMP_CHOICE_3PRICE = 20260724

# ---------------------------------------------------------------------------
# 仮定した「真の」選好構造（⚠️ 合成データの前提。実データではない）
# ---------------------------------------------------------------------------
# 価格の効用は水準ごとに指定する。6→8 は -0.35、8→10 は -1.15 と
# 下落幅を変えてあり、価格に対して非線形になっている。
TRUE_UTIL_PRICE = {6: 0.0, 8: -0.35, 10: -1.50}
SD_UTIL_PRICE = 0.15               # 価格効用の個人差（水準ごとに独立）

TRUE_UTIL_OS_APPLE = 0.8           # android を基準にした apple の効用
SD_UTIL_OS = 0.4

TRUE_UTIL_CAMERA_HIGH = 0.6        # 標準を基準にした高性能の効用
SD_UTIL_CAMERA = 0.3

# ---------------------------------------------------------------------------
# 評点の尺度（変更しやすいようにここにまとめる）
# ---------------------------------------------------------------------------
RATING_MIN = 1                     # 評点は 1〜10 の整数
RATING_MAX = 10
RATING_CENTER = 5.5                # 効用0のときの平均的な評点
RATING_GAIN = 2.0                  # 効用1あたり何点動くか
RATING_SD_PERSON = 0.6             # 「辛口／甘口」の個人差
RATING_SD_NOISE = 0.8              # 1回答ごとの誤差

# ---------------------------------------------------------------------------
# 回答者属性（Forms の設問として聞く想定）
# ---------------------------------------------------------------------------
GENDER_CHOICES = ["男性", "女性", "その他・回答しない"]
GENDER_PROBS = [0.55, 0.42, 0.03]
USED_OS_CHOICES = ["Apple (iOS)", "Android", "その他・わからない"]
USED_OS_PROBS = [0.50, 0.47, 0.03]

# 表示用の OS 表記（設計内部の "apple"/"android" → 学生向けの見せ方）
OS_DISPLAY = {"apple": "Apple (iOS)", "android": "Android"}

# 選択型の選択肢ラベル（pcc.forms_to_data の choice_labels と一致させる）
CHOICE_LABELS = ["製品A", "製品B", "製品C"]

# ---------------------------------------------------------------------------
# Microsoft Forms の列名
# ---------------------------------------------------------------------------
# 管理列。tests/data/forms_cbc_smartphone_real.xlsx（実際の Forms 出力）と
# 同じ構成・同じ順序にしてある。
MS_SYSTEM_COLS = ["ID", "開始時刻", "完了時刻", "メール", "名前", "最終変更時刻"]

# 回答者属性の設問
COL_Q_GENDER = "あなたの性別を教えてください。"
COL_Q_USED_OS = "現在使っているスマートフォンのOSはどちらですか？"


# ---------------------------------------------------------------------------
# 1. 回答者（100人）を一度だけ作る
# ---------------------------------------------------------------------------

def _make_respondents(n: int = N_RESPONDENTS, seed: int = SEED_RESPONDENTS) -> pd.DataFrame:
    """100人分の個人別部分効用と回答者属性を生成する。

    4つのデータすべてでこの同じ回答者を使うため、ここでしか乱数を引かない
    （:func:`_simulate_ratings` / :func:`_simulate_choices` は個人差を
    内部で引かない）。

    価格の効用は 3水準 {6, 8, 10} すべてについて持たせる。2price 版は
    そのうち {6, 10} の分だけを参照する。

    Returns
    -------
    pd.DataFrame
        1行 = 1人。列：
        ``u_price_6`` / ``u_price_8`` / ``u_price_10``（価格水準ごとの効用）、
        ``u_os_apple``、``u_camera_high``、``rating_bias``（評点の辛口／甘口）、
        ``gender``、``used_os``
    """
    rng = np.random.default_rng(seed)
    data = {
        f"u_price_{level}": mean + rng.normal(0, SD_UTIL_PRICE, n)
        for level, mean in TRUE_UTIL_PRICE.items()
    }
    data["u_os_apple"] = TRUE_UTIL_OS_APPLE + rng.normal(0, SD_UTIL_OS, n)
    data["u_camera_high"] = TRUE_UTIL_CAMERA_HIGH + rng.normal(0, SD_UTIL_CAMERA, n)
    data["rating_bias"] = rng.normal(0, RATING_SD_PERSON, n)
    data["gender"] = rng.choice(GENDER_CHOICES, size=n, p=GENDER_PROBS)
    data["used_os"] = rng.choice(USED_OS_CHOICES, size=n, p=USED_OS_PROBS)
    return pd.DataFrame(data)


def _utility(person: pd.Series, price: int, os_level: str, camera: str) -> float:
    """ある回答者にとっての、1つの製品案の効用。"""
    return (
        person[f"u_price_{int(price)}"]
        + person["u_os_apple"] * (os_level == "apple")
        + person["u_camera_high"] * (camera == "高性能")
    )


# ---------------------------------------------------------------------------
# 2. 設計をつくる
# ---------------------------------------------------------------------------

def _make_rating_design(price_levels: list[int], n_profiles: int, seed: int) -> pd.DataFrame:
    """評点型のプロファイル設計を py4conjoint で作る。"""
    return pcr.design_profiles(_attributes(price_levels), n_profiles, seed=seed)


def _make_choice_design(price_levels: list[int], seed: int) -> pd.DataFrame:
    """選択型の選択セット設計を py4conjoint で作る（6設問 × 3選択肢）。"""
    return pcc.design_choice_sets(
        _attributes(price_levels), n_sets=N_SETS, n_alts=N_ALTS, seed=seed
    )


# ---------------------------------------------------------------------------
# 3. 回答をシミュレートする
# ---------------------------------------------------------------------------

def _simulate_ratings(
    profiles: pd.DataFrame, respondents: pd.DataFrame, *, seed: int
) -> pd.DataFrame:
    """各回答者が各プロファイルに付ける評点（1〜10 の整数）を生成する。

    効用を線形変換して評点尺度に載せ、個人差（辛口／甘口）と誤差を足して
    四捨五入し、``RATING_MIN``〜``RATING_MAX`` に収める。

    Returns
    -------
    pd.DataFrame
        1行 = 1人、列 = ``Q1``…``Q{n_profiles}``（プロファイルの提示順）
    """
    rng = np.random.default_rng(seed)
    rows = []
    for _, person in respondents.iterrows():
        answers = {}
        for i, (_, prof) in enumerate(profiles.iterrows(), start=1):
            u = _utility(person, prof["price"], prof["os"], prof["camera"])
            score = (
                RATING_CENTER
                + RATING_GAIN * u
                + person["rating_bias"]
                + rng.normal(0, RATING_SD_NOISE)
            )
            answers[f"Q{i}"] = int(np.clip(round(score), RATING_MIN, RATING_MAX))
        rows.append(answers)
    return pd.DataFrame(rows)


def _simulate_choices(
    design: pd.DataFrame, respondents: pd.DataFrame, *, seed: int
) -> pd.DataFrame:
    """各回答者が各選択セットで選ぶ代替案を条件付きロジットで生成する。

    設問ごとに、代替案の効用を softmax で選択確率に変換して1つ引く。

    Returns
    -------
    pd.DataFrame
        1行 = 1人、列 = ``Q1``…``Q{n_sets}``（値は ``CHOICE_LABELS`` の文字列）
    """
    rng = np.random.default_rng(seed)
    set_ids = sorted(design["choice_set_id"].unique())
    groups = {
        s: design[design["choice_set_id"] == s].sort_values("alt_id") for s in set_ids
    }
    rows = []
    for _, person in respondents.iterrows():
        answers = {}
        for i, s in enumerate(set_ids, start=1):
            g = groups[s]
            u = np.array(
                [
                    _utility(person, r["price"], r["os"], r["camera"])
                    for _, r in g.iterrows()
                ]
            )
            p = np.exp(u - u.max())
            p /= p.sum()
            pick = rng.choice(g["alt_id"].to_numpy(), p=p)
            answers[f"Q{i}"] = CHOICE_LABELS[int(pick) - 1]
        rows.append(answers)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 4. Microsoft Forms 形式の CSV に書き出す
# ---------------------------------------------------------------------------

def _rating_question_text(profiles: pd.DataFrame, i: int) -> str:
    """評点型の設問文（Forms にそのまま貼れる形）を作る。"""
    r = profiles.iloc[i - 1]
    return (
        f"【製品案{i}】この製品をどれくらい買いたいと思いますか？"
        f"（{RATING_MIN}〜{RATING_MAX} の整数でお答えください）\n"
        f"・価格 {r['price']}万円 ｜ OS {OS_DISPLAY[r['os']]} ｜ カメラ {r['camera']}"
    )


def _choice_question_text(design: pd.DataFrame, set_id: int) -> str:
    """選択型の設問文（Forms にそのまま貼れる形）を作る。"""
    g = design[design["choice_set_id"] == set_id].sort_values("alt_id")
    lines = [
        f"【設問{set_id}】次の3つのうち、最も購入したいスマートフォンはどれですか？",
        "",
    ]
    for label, (_, r) in zip(CHOICE_LABELS, g.iterrows()):
        lines.append(
            f"・{label}：価格 {r['price']}万円 ｜ "
            f"OS {OS_DISPLAY[r['os']]} ｜ カメラ {r['camera']}"
        )
    return "\n".join(lines)


def _write_microsoft_forms_csv(
    out_path: Path,
    question_texts: list[str],
    answers: pd.DataFrame,
    respondents: pd.DataFrame,
    *,
    start: str,
    seed: int,
) -> None:
    """合成した回答を Microsoft Forms 形式の CSV（BOM付きUTF-8）で書き出す。

    列の順序は実際の Forms 出力に合わせる：
    管理列（ID／開始時刻／完了時刻／メール／名前／最終変更時刻）→ 設問 →
    回答者属性の設問。
    """
    # 設問文が重複すると pandas が列名に .1 を付けてしまい、設計との対応が
    # 分かりにくくなる。設計が異なれば設問文も異なるはずなので、ここで確認する。
    if len(set(question_texts)) != len(question_texts):
        raise ValueError(f"設問文が重複しています: {out_path.name}")

    n = len(respondents)
    rng = np.random.default_rng(seed)
    base = pd.Timestamp(start)
    starts = sorted(base + pd.Timedelta(minutes=int(m)) for m in rng.integers(0, 800, n))
    durations = rng.integers(60, 400, n)  # 回答所要時間（秒）

    out = pd.DataFrame()
    out["ID"] = range(1, n + 1)
    out["開始時刻"] = [t.strftime("%Y-%m-%d %H:%M:%S") for t in starts]
    out["完了時刻"] = [
        (t + pd.Timedelta(seconds=int(d))).strftime("%Y-%m-%d %H:%M:%S")
        for t, d in zip(starts, durations)
    ]
    out["メール"] = "anonymous"
    out["名前"] = ""            # 匿名回答では空欄（実際の Forms 出力と同じ）
    out["最終変更時刻"] = ""
    for text, col in zip(question_texts, answers.columns):
        out[text] = answers[col].to_numpy()
    out[COL_Q_GENDER] = respondents["gender"].to_numpy()
    out[COL_Q_USED_OS] = respondents["used_os"].to_numpy()

    # Microsoft Forms の CSV と同じ BOM付きUTF-8 で保存
    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"  書き出し: {out_path.name}（回答者 {n}名・設問 {len(question_texts)}問）")


def _save_design(design: pd.DataFrame, out_path: Path) -> None:
    """設計を CSV で保存する。

    ``index=False`` にするのは、プロファイルIDが列として混入して属性と
    誤認されるのを防ぐため（``forms_to_data()`` はプロファイルの提示順を
    行の順序で解釈する）。
    """
    design.to_csv(out_path, index=False, encoding="utf-8")
    print(f"  書き出し: {out_path.name}（{len(design)}行）")


# ---------------------------------------------------------------------------
# 5. 生成物の検証
# ---------------------------------------------------------------------------

RESPONDENT_COLS = {COL_Q_GENDER: "gender", COL_Q_USED_OS: "used_os"}


def _verify() -> None:
    """生成した回答データが実際に forms_to_data() で読めることを確認する。

    「生成はできたが読めない」という事故を防ぐための検証。行数と回答者数が
    期待どおりかまで確認する。
    """
    print("\n検証：生成したファイルを forms_to_data() で読み込みます")

    checks = [
        ("responses_rating_2price.csv", "design_rating_2price.csv",
         "rating", N_RESPONDENTS * N_PROFILES_2PRICE),
        ("responses_rating_3price.csv", "design_rating_3price.csv",
         "rating", N_RESPONDENTS * N_PROFILES_3PRICE),
        ("responses_choice_2price.csv", "design_choice_2price.csv",
         "choice", N_RESPONDENTS * N_SETS * N_ALTS),
        ("responses_choice_3price.csv", "design_choice_3price.csv",
         "choice", N_RESPONDENTS * N_SETS * N_ALTS),
    ]

    for resp_name, design_name, kind, expected_rows in checks:
        design = pd.read_csv(HERE / design_name)
        if kind == "rating":
            df = pcr.forms_to_data(
                str(HERE / resp_name), design, respondent_cols=RESPONDENT_COLS
            )
        else:
            df = pcc.forms_to_data(
                str(HERE / resp_name), design, CHOICE_LABELS,
                respondent_cols=RESPONDENT_COLS,
            )
        n_resp = df["respondent_id"].nunique()
        if len(df) != expected_rows:
            raise AssertionError(
                f"{resp_name} の行数が期待と違います：{len(df)}行"
                f"（期待：{expected_rows}行）"
            )
        if n_resp != N_RESPONDENTS:
            raise AssertionError(
                f"{resp_name} の回答者数が期待と違います：{n_resp}名"
                f"（期待：{N_RESPONDENTS}名）"
            )
        print(f"  OK: {resp_name}（{len(df)}行・回答者 {n_resp}名）")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> None:
    respondents = _make_respondents()
    print(f"回答者を生成しました（{len(respondents)}名。4データで共通）")

    # ---- 評点型 ----
    print("\n評点型（rating）")
    for tag, price_levels, n_profiles, seed_design, seed_answer, seed_timestamp, start in [
        ("2price", PRICE_LEVELS_2, N_PROFILES_2PRICE,
         SEED_DESIGN_RATING_2PRICE, SEED_ANSWER_RATING_2PRICE,
         SEED_TIMESTAMP_RATING_2PRICE, "2026-06-15 09:00:00"),
        ("3price", PRICE_LEVELS_3, N_PROFILES_3PRICE,
         SEED_DESIGN_RATING_3PRICE, SEED_ANSWER_RATING_3PRICE,
         SEED_TIMESTAMP_RATING_3PRICE, "2026-06-16 09:00:00"),
    ]:
        design = _make_rating_design(price_levels, n_profiles, seed_design)
        _save_design(design, HERE / f"design_rating_{tag}.csv")
        answers = _simulate_ratings(design, respondents, seed=seed_answer)
        _write_microsoft_forms_csv(
            HERE / f"responses_rating_{tag}.csv",
            [_rating_question_text(design, i) for i in range(1, n_profiles + 1)],
            answers,
            respondents,
            start=start,
            seed=seed_timestamp,
        )

    # ---- 選択型 ----
    print("\n選択型（choice）")
    for tag, price_levels, seed_design, seed_answer, seed_timestamp, start in [
        ("2price", PRICE_LEVELS_2, SEED_DESIGN_CHOICE_2PRICE,
         SEED_ANSWER_CHOICE_2PRICE, SEED_TIMESTAMP_CHOICE_2PRICE,
         "2026-06-17 09:00:00"),
        ("3price", PRICE_LEVELS_3, SEED_DESIGN_CHOICE_3PRICE,
         SEED_ANSWER_CHOICE_3PRICE, SEED_TIMESTAMP_CHOICE_3PRICE,
         "2026-06-18 09:00:00"),
    ]:
        design = _make_choice_design(price_levels, seed_design)
        _save_design(design, HERE / f"design_choice_{tag}.csv")
        answers = _simulate_choices(design, respondents, seed=seed_answer)
        set_ids = sorted(design["choice_set_id"].unique())
        _write_microsoft_forms_csv(
            HERE / f"responses_choice_{tag}.csv",
            [_choice_question_text(design, s) for s in set_ids],
            answers,
            respondents,
            start=start,
            seed=seed_timestamp,
        )

    try:
        _verify()
    except Exception as e:  # noqa: BLE001
        print("\n検証に失敗しました。生成したファイルは読み込めません。")
        print(f"  {type(e).__name__}: {e}")
        sys.exit(1)

    print("\n完了。8ファイルを生成し、読み込みまで確認しました。")


if __name__ == "__main__":
    main()
