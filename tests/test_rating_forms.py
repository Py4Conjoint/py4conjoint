"""rating の forms_to_data()（回答ファイル読み込み）の頑健性テスト。

- 評点列の自動検出：数値候補が n_profiles を超えるとき、除外した列を
  UserWarning で明示する（評点でない数値質問の混在による静かなズレ対策）。
- 評点の数値化：文字列の評点（"5" など）は数値化され、数値化できない値は
  警告のうえ NaN として扱われる（fit() の欠損処理に乗る）。
- 出力の行順：profile_id は提示順（P1, P2, …, P10, …）で並ぶ
  （文字列の辞書順 P1, P10, P11, P2, … にならない）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd
import pytest

import py4conjoint.rating as pcr

PROFILES = {
    "price":  [6, 10, 6, 10],
    "os":     ["android", "apple", "apple", "android"],
    "camera": ["標準", "標準", "高性能", "高性能"],
}

_SYSTEM = {
    "ID": [1, 2, 3],
    "Start time": ["2026-06-01 10:00"] * 3,
    "Completion time": ["2026-06-01 10:05"] * 3,
    "Email": ["anonymous"] * 3,
    "Name": [""] * 3,
}


def _write_xlsx(path, extra_cols):
    """管理列 + 指定の設問列で Microsoft Forms 形式の xlsx を作る。"""
    pd.DataFrame({**_SYSTEM, **extra_cols}).to_excel(path, index=False)


# ---------------------------------------------------------------------------
# 評点列の自動検出：余分な数値列があると警告する
# ---------------------------------------------------------------------------

def test_warns_when_extra_numeric_column_present(tmp_path):
    """数値候補が n_profiles を超えると、除外した列名を警告で明示する。"""
    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        "Q1. 製品案1の評価": [5, 6, 7],
        "Q2. 製品案2の評価": [3, 4, 5],
        "Q3. 製品案3の評価": [6, 6, 7],
        "Q4. 製品案4の評価": [4, 5, 4],
        "総合満足度（数値）": [9, 2, 5],   # 評点でない数値質問
    })
    with pytest.warns(UserWarning, match="除外した列"):
        pcr.forms_to_data(str(f), PROFILES)


def test_no_warning_when_extra_numeric_column_in_respondent_cols(tmp_path):
    """余分な数値列を respondent_cols で指定すれば候補から外れ、警告は出ない。"""
    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        "Q1. 製品案1の評価": [5, 6, 7],
        "Q2. 製品案2の評価": [3, 4, 5],
        "Q3. 製品案3の評価": [6, 6, 7],
        "Q4. 製品案4の評価": [4, 5, 4],
        "総合満足度（数値）": [9, 2, 5],
    })
    import warnings as _w
    with _w.catch_warnings(record=True) as caught:
        _w.simplefilter("always")
        df = pcr.forms_to_data(
            str(f), PROFILES, respondent_cols={"総合満足度（数値）": "satisfaction"}
        )
    # 「除外した列」の警告が出ないこと（依存ライブラリの無関係な警告は許容）
    assert not [x for x in caught if "除外した列" in str(x.message)]
    assert "satisfaction" in df.columns
    # 評点は Q1〜Q4 が正しく対応する（回答者1: P1=5, P2=3, P3=6, P4=4）
    r1 = df[df["respondent_id"] == 1].set_index("profile_id")["rating"]
    assert r1.loc["P1"] == 5 and r1.loc["P4"] == 4


# ---------------------------------------------------------------------------
# 評点の数値化（文字列の評点・変換できない値）
# ---------------------------------------------------------------------------

def test_string_ratings_are_coerced_to_numeric(tmp_path):
    """文字列の評点（"5" など）は数値化され、そのまま fit() まで通る。"""
    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        "Q1. 製品案1の評価": ["5", "6", "7"],
        "Q2. 製品案2の評価": ["3", "4", "5"],
        "Q3. 製品案3の評価": ["6", "6", "7"],
        "Q4. 製品案4の評価": ["4", "5", "4"],
    })
    df = pcr.forms_to_data(str(f), PROFILES)
    assert pd.api.types.is_numeric_dtype(df["rating"])
    assert float(df.loc[
        (df["respondent_id"] == 1) & (df["profile_id"] == "P1"), "rating"
    ].iloc[0]) == 5.0


def test_uncoercible_rating_becomes_nan_with_warning(tmp_path):
    """数値化できない評点（"x" など）は警告のうえ NaN になる。"""
    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        "Q1. 製品案1の評価": ["5", "6", "7"],
        "Q2. 製品案2の評価": ["3", "x", "5"],   # 回答者2の Q2 が不正値
        "Q3. 製品案3の評価": ["6", "6", "7"],
        "Q4. 製品案4の評価": ["4", "5", "4"],
    })
    with pytest.warns(UserWarning, match="数値へ変換できない"):
        df = pcr.forms_to_data(str(f), PROFILES)
    bad = df[(df["respondent_id"] == 2) & (df["profile_id"] == "P2")]
    assert bad["rating"].isna().all()
    # 他の評点は無事
    assert df["rating"].notna().sum() == 11


# ---------------------------------------------------------------------------
# profiles の行番号列（Unnamed: 0）：警告のうえ属性から除外
# ---------------------------------------------------------------------------

def test_warns_and_drops_pandas_index_column_in_profiles(tmp_path):
    """index=False を付け忘れた profiles CSV でも、警告のうえ正しく動く。

    行番号列（Unnamed: 0。保存された index の P1, P2, … ラベル）は属性でない
    ため、出力 DataFrame に混入しない。
    """
    profiles = pd.DataFrame(PROFILES, index=["P1", "P2", "P3", "P4"])
    csv = tmp_path / "profiles.csv"
    profiles.to_csv(csv)                      # index=False を付け忘れたケース
    loaded = pd.read_csv(csv)
    assert "Unnamed: 0" in loaded.columns

    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        f"Q{i+1}. 製品案{i+1}の評価": [5, 6, 4] for i in range(4)
    })
    with pytest.warns(UserWarning, match="index=False"):
        df = pcr.forms_to_data(str(f), loaded)
    # 行番号列は属性として混入しない
    assert "Unnamed: 0" not in df.columns
    # 属性列は本来のものだけ
    assert set(df.columns) == {"respondent_id", "profile_id", "rating",
                               "price", "os", "camera"}


# ---------------------------------------------------------------------------
# 出力の行順：profile_id は提示順（数値順）
# ---------------------------------------------------------------------------

def test_profile_order_is_numeric_with_10plus_profiles(tmp_path):
    """n_profiles >= 10 でも P1, P2, …, P10, … の提示順で並ぶ。"""
    n_profiles = 12
    profiles = {
        "price": [6, 10] * 6,
        "os": (["android"] * 6 + ["apple"] * 6),
    }
    f = tmp_path / "responses.xlsx"
    _write_xlsx(f, {
        f"Q{i+1}. 製品案{i+1}の評価": [((i + r) % 7) + 1 for r in range(3)]
        for i in range(n_profiles)
    })
    df = pcr.forms_to_data(str(f), profiles)
    expected = [f"P{i+1}" for i in range(n_profiles)]
    for rid in (1, 2, 3):
        got = df[df["respondent_id"] == rid]["profile_id"].tolist()
        assert got == expected, f"回答者{rid}の行順が提示順でない: {got[:5]}…"
