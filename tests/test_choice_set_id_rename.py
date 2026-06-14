"""選択セット識別子を choice_set_id 系に統一した改名の検証。

- design_choice_sets() と cbc_forms_to_data() が **同じ列名** choice_set_id を
  出力すること（生成側・設計側で名前がねじれていないこと）。
- fit() が choice_set_id_col 引数を受け取って動くこと。
- 旧名（set_id 出力・obsID 出力・choice_set_col 引数）がもう存在しないこと。
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


def _make_microsoft_xlsx(path, answers):
    n = len(answers)
    data = {
        "ID": range(1, n + 1),
        "Start time": ["2026-06-01 10:00"] * n,
        "Completion time": ["2026-06-01 10:05"] * n,
        "Email": ["anonymous"] * n,
        "Name": [""] * n,
    }
    for q in range(N_SETS):
        data[f"Q{q+1}. どの製品を選びますか？"] = [a[q] for a in answers]
    pd.DataFrame(data).to_excel(path, index=False)


# ---------------------------------------------------------------------------
# 設計側と生成側で同じ列名 choice_set_id を出力する
# ---------------------------------------------------------------------------

def test_design_outputs_choice_set_id(design):
    assert "choice_set_id" in design.columns
    # 旧名 set_id は残っていない（alt_id は別概念なので残る）
    assert "set_id" not in design.columns
    assert "alt_id" in design.columns


def test_forms_outputs_choice_set_id(tmp_path, design):
    answers = [
        ["製品A", "製品B", "製品C", "製品A"],
        ["製品B", "製品B", "製品A", "製品C"],
    ]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    df = pcc.cbc_forms_to_data(str(f), design, LABELS)
    assert "choice_set_id" in df.columns
    # 旧名 obsID は残っていない
    assert "obsID" not in df.columns


def test_design_and_forms_use_same_set_identifier(tmp_path, design):
    """設計・回答データの選択セット識別子の列名が一致する（ねじれがない）。"""
    answers = [["製品A", "製品B", "製品C", "製品A"]]
    f = tmp_path / "responses.xlsx"
    _make_microsoft_xlsx(f, answers)
    df = pcc.cbc_forms_to_data(str(f), design, LABELS)
    set_id_col = "choice_set_id"
    assert set_id_col in design.columns
    assert set_id_col in df.columns


# ---------------------------------------------------------------------------
# fit() は choice_set_id_col を受け取る
# ---------------------------------------------------------------------------

def _toy_choice_df(seed=0):
    rng = np.random.default_rng(seed)
    rows = []
    for t in range(300):
        pr = rng.choice([100, 200], 2)
        br = rng.choice(["A", "B"], 2)
        v = np.array([-0.01 * p + (0.5 if b == "B" else 0.0)
                      for p, b in zip(pr, br)])
        prob = np.exp(v - v.max())
        prob = prob / prob.sum()
        ch = (rng.random() < prob.cumsum()).argmax()
        for j in range(2):
            rows.append({"choice_set_id": t, "choice": int(j == ch),
                         "price": int(pr[j]), "brand": br[j]})
    return pd.DataFrame(rows)


def test_fit_accepts_choice_set_id_col():
    df = _toy_choice_df()
    dc = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id")
    assert result.converged
    assert result.choice_set_id_col == "choice_set_id"


# ---------------------------------------------------------------------------
# 旧名はもう存在しない（使うとエラーになる）
# ---------------------------------------------------------------------------

def test_old_choice_set_col_arg_rejected():
    df = _toy_choice_df()
    dc = pcc.encode(df, reference_levels={"brand": "A"})
    # 旧引数名 choice_set_col は廃止されており TypeError になる
    with pytest.raises(TypeError):
        pcc.fit(dc, choice="choice", choice_set_col="choice_set_id")


def test_result_has_no_old_attribute():
    df = _toy_choice_df()
    dc = pcc.encode(df, reference_levels={"brand": "A"})
    result = pcc.fit(dc, choice="choice", choice_set_id_col="choice_set_id")
    # 旧フィールド名 choice_set_col は残っていない
    assert not hasattr(result, "choice_set_col")
