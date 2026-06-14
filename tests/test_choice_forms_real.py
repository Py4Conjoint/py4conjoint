"""実 Microsoft Forms 出力での cbc_forms_to_data() 検証。

リハーサルで回収した実ファイル（3人分の回答）を使い、
- 列名が長文（改行・全角空白・\\xa0 を含む）でも設問列を検出できること
- 性別・利用OS の属性質問が混在しても設問列から除外されること
- version 列を持たない手作りの設計CSVをそのまま渡せること
- 3人分の選択が実回答と一致すること
を確認する。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd
import pytest

import py4conjoint.choice as pcc

DATA_DIR = Path(__file__).resolve().parent / "data"
RESPONSES = DATA_DIR / "forms_cbc_smartphone_real.xlsx"
GOOGLE_RESPONSES = DATA_DIR / "forms_cbc_smartphone_google.csv"
DESIGN_CSV = DATA_DIR / "design_smartphone_cbc.csv"

LABELS = ["製品A", "製品B", "製品C"]
N_RESP = 3
N_SETS = 6
N_ALTS = 3

# 実回答から手作業で確認した正解（alt_id: 製品A=1, 製品B=2, 製品C=3）。
# choice_set_id は回答者1の設問1〜6 → 1〜6、回答者2 → 7〜12、回答者3 → 13〜18。
EXPECTED_CHOSEN_ALT = {
    1: 2, 2: 3, 3: 3, 4: 3, 5: 1, 6: 2,        # 回答者1
    7: 1, 8: 3, 9: 3, 10: 3, 11: 2, 12: 2,     # 回答者2
    13: 1, 14: 2, 15: 2, 16: 1, 17: 2, 18: 1,  # 回答者3
}

# Google Forms 版の正解表（同じ design・choice_labels で照合する）。
# 回答者1: Q1=A, Q2=B, Q3=B, Q4=C, Q5=B, Q6=A
# 回答者2: Q1=B, Q2=C, Q3=C, Q4=B, Q5=B, Q6=A
# 回答者3: Q1=B, Q2=A, Q3=C, Q4=A, Q5=A, Q6=B
EXPECTED_CHOSEN_ALT_GOOGLE = {
    1: 1, 2: 2, 3: 2, 4: 3, 5: 2, 6: 1,        # 回答者1
    7: 2, 8: 3, 9: 3, 10: 2, 11: 2, 12: 1,     # 回答者2
    13: 2, 14: 1, 15: 3, 16: 1, 17: 1, 18: 2,  # 回答者3
}


@pytest.fixture(scope="module")
def design():
    return pd.read_csv(DESIGN_CSV)


def test_design_csv_has_no_version_column(design):
    """設計CSVは version 列を持たない（手作り設計の想定）。"""
    assert "version" not in design.columns
    assert list(design.columns) == ["choice_set_id", "alt_id", "price", "os", "camera"]


def test_real_forms_runs_and_shape(design):
    """実ファイルが例外なく long 形式に変換され、形状が正しい。"""
    df = pcc.cbc_forms_to_data(str(RESPONSES), design, LABELS, forms="microsoft")

    # (a) 行数 = 3人 × 6設問 × 3代替案 = 54
    assert len(df) == N_RESP * N_SETS * N_ALTS == 54
    # 出力列：choice_set_id, respondent_id, alt, choice + 属性列
    assert list(df.columns) == [
        "choice_set_id", "respondent_id", "alt", "choice", "price", "os", "camera"
    ]
    # (b) choice の合計 = 3 × 6 = 18（各回答者×設問でちょうど1つ choice=1）
    assert df["choice"].sum() == N_RESP * N_SETS == 18
    assert (df.groupby("choice_set_id")["choice"].sum() == 1).all()
    # (c) choice_set_id は 18 通り（3人 × 6設問）
    assert sorted(df["choice_set_id"].unique()) == list(range(1, 19))
    assert (df.groupby("choice_set_id").size() == N_ALTS).all()


def test_real_forms_choices_match_answer_key(design):
    """(d) 3人分の選択が実回答の正解表と一致する。"""
    df = pcc.cbc_forms_to_data(str(RESPONSES), design, LABELS, forms="microsoft")

    chosen = (
        df[df["choice"] == 1]
        .set_index("choice_set_id")["alt"]
        .to_dict()
    )
    assert chosen == EXPECTED_CHOSEN_ALT


def test_real_forms_attributes_match_design(design):
    """選ばれた代替案の属性が設計CSVの水準と一致する。"""
    df = pcc.cbc_forms_to_data(str(RESPONSES), design, LABELS, forms="microsoft")
    design_indexed = design.set_index(["choice_set_id", "alt_id"])

    # choice_set_id 1 = 回答者1の設問1。製品B(alt2)が選ばれている。
    row = df[(df["choice_set_id"] == 1) & (df["alt"] == 2)].iloc[0]
    assert row["choice"] == 1
    expected = design_indexed.loc[(1, 2)]
    assert row["price"] == expected["price"]
    assert row["os"] == expected["os"]
    assert row["camera"] == expected["camera"]


def test_real_forms_keep_attribute_questions_via_respondent_cols(design):
    """性別・利用OS の属性質問を respondent_cols で回答者属性として残せる。"""
    raw = pd.read_excel(RESPONSES)
    gender_col = next(c for c in raw.columns if c.startswith("あなたの性別"))
    os_col = next(c for c in raw.columns if c.startswith("現在使っている"))

    df = pcc.cbc_forms_to_data(
        str(RESPONSES),
        design,
        LABELS,
        forms="microsoft",
        respondent_cols={gender_col: "性別", os_col: "利用OS"},
    )
    # 設問列の検出は壊れず、形状は同じ
    assert len(df) == 54
    assert "性別" in df.columns and "利用OS" in df.columns
    # 回答者1=女性/Apple、回答者2=男性/Apple、回答者3=男性/Android
    assert set(df.loc[df["respondent_id"] == 1, "性別"]) == {"女性"}
    assert set(df.loc[df["respondent_id"] == 3, "利用OS"]) == {"Android"}


def test_real_forms_pipeline_to_encode(design):
    """実データが encode までそのまま流れる（fit に渡せる形になる）。"""
    df = pcc.cbc_forms_to_data(str(RESPONSES), design, LABELS, forms="microsoft")
    df_coded = pcc.encode(df, reference_levels={"os": "android", "camera": "標準"})
    # ダミー列が追加され、choice_set_id/choice はそのまま残る
    assert "os_apple" in df_coded.columns
    assert "camera_高性能" in df_coded.columns
    assert {"choice_set_id", "choice", "price"}.issubset(df_coded.columns)


# ---------------------------------------------------------------------------
# Google Forms（CSV・BOMなしUTF-8）版
# ---------------------------------------------------------------------------

def test_google_forms_runs_and_shape(design):
    """Google Forms の CSV（BOMなしUTF-8）が同じ design・labels で変換できる。"""
    df = pcc.cbc_forms_to_data(
        str(GOOGLE_RESPONSES), design, LABELS, forms="google"
    )
    # 出力54行（3人 × 6設問 × 3代替案）、列も Microsoft 版と同じ
    assert len(df) == N_RESP * N_SETS * N_ALTS == 54
    assert list(df.columns) == [
        "choice_set_id", "respondent_id", "alt", "choice", "price", "os", "camera"
    ]
    # choice 合計 = 18、各 choice_set_id でちょうど1つ choice=1、choice_set_id は 18 通り
    assert df["choice"].sum() == N_RESP * N_SETS == 18
    assert (df.groupby("choice_set_id")["choice"].sum() == 1).all()
    assert sorted(df["choice_set_id"].unique()) == list(range(1, 19))


def test_google_forms_choices_match_answer_key(design):
    """Google 版の選択が正解表と一致する（【設問N】接頭辞の有無に依存しない）。"""
    df = pcc.cbc_forms_to_data(
        str(GOOGLE_RESPONSES), design, LABELS, forms="google"
    )
    chosen = df[df["choice"] == 1].set_index("choice_set_id")["alt"].to_dict()
    assert chosen == EXPECTED_CHOSEN_ALT_GOOGLE


def test_google_forms_keep_attribute_questions(design):
    """Google 版でも属性質問（性別・利用OS）を respondent_cols で残せる。"""
    raw = pd.read_csv(GOOGLE_RESPONSES, encoding="utf-8-sig")
    gender_col = next(c for c in raw.columns if "性別" in c)
    os_col = next(c for c in raw.columns if "OS" in c)

    df = pcc.cbc_forms_to_data(
        str(GOOGLE_RESPONSES),
        design,
        LABELS,
        forms="google",
        respondent_cols={gender_col: "性別", os_col: "利用OS"},
    )
    assert len(df) == 54
    assert "性別" in df.columns and "利用OS" in df.columns
    assert set(df.loc[df["respondent_id"] == 1, "性別"]) == {"男性"}
    assert set(df.loc[df["respondent_id"] == 1, "利用OS"]) == {"Apple (iOS)"}
