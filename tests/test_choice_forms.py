"""choice/_forms.py（cbc_forms_to_data）のテスト。

Microsoft Forms（.xlsx）/ Google Forms（.csv）の模擬回答ファイルを
tmp_path に生成して変換を検証する。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np
import pandas as pd
import pytest

import py4conjoint.choice as pcc

ATTRS = {"price": [100, 150, 200], "brand": ["A社", "B社", "C社"]}
N_SETS = 4
N_ALTS = 3
LABELS = ["A", "B", "C"]


@pytest.fixture(scope="module")
def design():
    return pcc.design_choice_sets(ATTRS, n_sets=N_SETS, n_alts=N_ALTS, seed=42)


def _make_microsoft_xlsx(path, answers, extra_cols=None):
    """Microsoft Forms 形式の模擬 .xlsx を作る。

    answers : list of list — 回答者ごとの設問回答（「製品A」等）のリスト。
    """
    n = len(answers)
    data = {
        "ID": range(1, n + 1),
        "Start time": ["2026-06-01 10:00"] * n,
        "Completion time": ["2026-06-01 10:05"] * n,
        "Email": ["anonymous"] * n,
        "Name": [""] * n,
    }
    if extra_cols:
        data.update(extra_cols)
    for q in range(N_SETS):
        data[f"Q{q+1}. どの製品を選びますか？"] = [a[q] for a in answers]
    pd.DataFrame(data).to_excel(path, index=False)


def _make_google_csv(path, answers):
    n = len(answers)
    data = {"タイムスタンプ": ["2026/06/01 10:00"] * n}
    for q in range(N_SETS):
        data[f"Q{q+1}. どの製品を選びますか？"] = [a[q] for a in answers]
    pd.DataFrame(data).to_csv(path, index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------------------
# 正常系
# ---------------------------------------------------------------------------

def test_microsoft_happy_path(tmp_path, design):
    answers = [
        ["製品A", "製品B", "製品C", "製品A"],
        ["製品B", "製品B", "製品A", "製品C"],
    ]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)

    df = pcc.cbc_forms_to_data(str(f), design, LABELS)

    # 形状：2回答者 × 4設問 × 3代替案
    assert len(df) == 2 * N_SETS * N_ALTS
    assert list(df.columns) == ["choice_set_id", "respondent_id", "alt", "choice",
                                "price", "brand"]
    # choice_set_id は回答者×設問の通し番号（1〜8）、各 choice_set_id に3行
    assert sorted(df["choice_set_id"].unique()) == list(range(1, 9))
    assert (df.groupby("choice_set_id").size() == N_ALTS).all()
    # 各選択セットでちょうど1つ選ばれている
    assert (df.groupby("choice_set_id")["choice"].sum() == 1).all()
    # 回答者1の設問1は「製品A」→ alt 1 が choice=1
    first = df[(df["respondent_id"] == 1) & (df["choice_set_id"] == 1)]
    assert first.loc[first["alt"] == 1, "choice"].iloc[0] == 1
    # 属性は design と一致する（設問1・alt1 の水準）
    d11 = design[(design["choice_set_id"] == 1) & (design["alt_id"] == 1)].iloc[0]
    assert first.loc[first["alt"] == 1, "price"].iloc[0] == d11["price"]
    assert first.loc[first["alt"] == 1, "brand"].iloc[0] == d11["brand"]


def test_google_happy_path(tmp_path, design):
    answers = [["製品A", "製品B", "製品C", "製品A"]]
    f = tmp_path / "responses.csv"
    _make_google_csv(f, answers)
    df = pcc.cbc_forms_to_data(str(f), design, LABELS, forms="google")
    assert len(df) == 1 * N_SETS * N_ALTS
    assert (df.groupby("choice_set_id")["choice"].sum() == 1).all()


def test_respondent_cols(tmp_path, design):
    answers = [
        ["製品A", "製品B", "製品C", "製品A"],
        ["製品B", "製品B", "製品A", "製品C"],
    ]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers, extra_cols={"性別": ["男", "女"]})
    df = pcc.cbc_forms_to_data(
        str(f), design, LABELS, respondent_cols={"性別": "gender"}
    )
    assert "gender" in df.columns
    assert set(df.loc[df["respondent_id"] == 1, "gender"]) == {"男"}
    assert set(df.loc[df["respondent_id"] == 2, "gender"]) == {"女"}


def test_end_to_end_encode_fit(tmp_path, design):
    """cbc_forms_to_data → encode → fit がそのまま流れることを確認する。"""
    rng = np.random.default_rng(0)
    answers = [
        [f"製品{rng.choice(LABELS)}" for _ in range(N_SETS)]
        for _ in range(30)
    ]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    df = pcc.cbc_forms_to_data(str(f), design, LABELS)
    df_coded = pcc.encode(df, reference_levels={"brand": "A社"})
    result = pcc.fit(
        df_coded,
        choice_set_id_col="choice_set_id",
        respondent_id_col="respondent_id",
    )
    assert result.n_sets == 30 * N_SETS
    assert "price" in result.params.index


# ---------------------------------------------------------------------------
# エラー・警告
# ---------------------------------------------------------------------------

def test_n_sets_mismatch_error(tmp_path):
    """設問数と design の n_sets が一致しない場合はエラー。"""
    design_small = pcc.design_choice_sets(ATTRS, n_sets=2, n_alts=3, seed=0)
    answers = [["製品A", "製品B", "製品C", "製品A"]]  # 4設問分
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    with pytest.raises(ValueError, match="設問数.*一致しません"):
        pcc.cbc_forms_to_data(str(f), design_small, LABELS)


def test_unmatched_answer_error(tmp_path, design):
    answers = [["製品A", "製品X", "製品C", "製品A"]]  # 製品X は無効
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    with pytest.raises(ValueError, match="マッチしない回答値"):
        pcc.cbc_forms_to_data(str(f), design, LABELS)


def test_unanswered_warning_and_drop(tmp_path, design):
    answers = [
        ["製品A", None, "製品C", "製品A"],  # 設問2が未回答
        ["製品B", "製品B", "製品A", "製品C"],
    ]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    with pytest.warns(UserWarning, match="未回答の設問が 1 件"):
        df = pcc.cbc_forms_to_data(str(f), design, LABELS)
    # 未回答の選択セットは除外される：(2人×4設問 − 1) × 3代替案
    assert len(df) == (2 * N_SETS - 1) * N_ALTS
    # choice_set_id は欠番なく連番
    assert sorted(df["choice_set_id"].unique()) == list(range(1, 2 * N_SETS))


def test_choice_labels_length_error(tmp_path, design):
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, [["製品A", "製品B", "製品C", "製品A"]])
    with pytest.raises(ValueError, match="choice_labels の長さ"):
        pcc.cbc_forms_to_data(str(f), design, ["A", "B"])


def test_invalid_version_error(tmp_path, design):
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, [["製品A", "製品B", "製品C", "製品A"]])
    with pytest.raises(ValueError, match="バージョン 2 が存在しません"):
        pcc.cbc_forms_to_data(str(f), design, LABELS, version=2)


def test_file_not_found(design):
    with pytest.raises(FileNotFoundError, match="ファイルが見つかりません"):
        pcc.cbc_forms_to_data("no_such_file.xlsx", design, LABELS)


def test_invalid_forms_error(tmp_path, design):
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, [["製品A", "製品B", "製品C", "製品A"]])
    with pytest.raises(ValueError, match="無効な値"):
        pcc.cbc_forms_to_data(str(f), design, LABELS, forms="yahoo")
