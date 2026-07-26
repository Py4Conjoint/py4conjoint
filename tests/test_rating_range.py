"""rating の forms_to_data() の rating_range 引数のテスト。

rating_range は choice の forms_to_data() の choice_labels に相当する引数で、
評点列を「位置（右端の n_profiles 列）」ではなく「値の内容」から同定する。

ここで確認すること：
- 未指定なら従来どおりの挙動（後方互換）。年齢の列が評点より後ろにあると、
  評点とプロファイルの対応が1つずれる（rating_range はこれを解決する）。
- 指定すると、値がすべて範囲外の列（年齢など）が候補から外れ、対応が正しくなる。
- 採用した評点列に範囲外の値があれば ValueError（入力ミスの検出）。
- 曖昧さが残っていない正常系では警告を出さない。
"""

import sys
import warnings
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd
import pytest

import py4conjoint.rating as pcr

PROFILES = pd.DataFrame(
    {
        "price": [6, 10, 6, 10],
        "os": ["android", "apple", "apple", "android"],
        "camera": ["標準", "標準", "高性能", "高性能"],
    },
    index=["P1", "P2", "P3", "P4"],
)

N_PROFILES = 4
RATING_COLS = [
    f"【製品案{i}】はどれくらい欲しいですか" for i in range(1, N_PROFILES + 1)
]
AGE_COL = "あなたの年齢を教えてください"

# 1行 = 1回答者の評点（10段階評価）。年齢は評点の列より後ろに置く。
# fmt: off
RATINGS = [
    [9, 7, 3, 9],    # 回答者1
    [5, 6, 8, 4],    # 回答者2
    [10, 1, 2, 6],   # 回答者3
]
# fmt: on
AGES = [20, 21, 19]


def _write_csv(path, ratings=None, ages=None, extra_cols=None):
    """管理列 + 評点4列 + 年齢列（+ 追加列）の Microsoft Forms 形式 CSV を作る。"""
    ratings = RATINGS if ratings is None else ratings
    ages = AGES if ages is None else ages
    n = len(ratings)
    data = {
        "ID": range(1, n + 1),
        "開始時刻": ["2026-06-01 10:00"] * n,
        "完了時刻": ["2026-06-01 10:05"] * n,
        "メール": ["anonymous"] * n,
        "名前": [""] * n,
    }
    for i, col in enumerate(RATING_COLS):
        data[col] = [r[i] for r in ratings]
    data[AGE_COL] = ages
    if extra_cols:
        data.update(extra_cols)
    pd.DataFrame(data).to_csv(path, index=False, encoding="utf-8-sig")
    return str(path)


def _ratings_of(df, respondent_id=1):
    """指定した回答者の評点を提示順（P1〜P4）で取り出す。"""
    return df[df["respondent_id"] == respondent_id]["rating"].tolist()


# ---------------------------------------------------------------------------
# 1. rating_range を指定しない場合は従来どおり（後方互換）
# ---------------------------------------------------------------------------


def test_without_rating_range_keeps_existing_behavior(tmp_path):
    """未指定なら従来の挙動のまま：右端4列が採用され、対応が1つずれる。

    年齢列が評点列より後ろにあるため、評点は [9, 7, 3, 9] ではなく
    [7, 3, 9, 20]（年齢が評点として混入）になる。これが rating_range で
    解決したい問題そのものであり、既定では挙動を変えないことを固定する。
    """
    f = _write_csv(tmp_path / "responses.csv")
    with pytest.warns(UserWarning, match="除外した列"):
        df = pcr.forms_to_data(f, PROFILES)
    assert _ratings_of(df) == [7, 3, 9, 20]


# ---------------------------------------------------------------------------
# 2. rating_range を指定すると年齢列が外れ、対応が正しくなる
# ---------------------------------------------------------------------------


def test_rating_range_excludes_out_of_range_column(tmp_path):
    """rating_range=(1, 10) で年齢列が候補から外れ、評点が正しく並ぶ。"""
    f = _write_csv(tmp_path / "responses.csv")
    df = pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))
    assert _ratings_of(df, 1) == [9, 7, 3, 9]
    assert _ratings_of(df, 2) == [5, 6, 8, 4]
    assert _ratings_of(df, 3) == [10, 1, 2, 6]
    # 年齢は評点にもプロファイル属性にも現れない
    assert AGE_COL not in df.columns


# ---------------------------------------------------------------------------
# 3. 採用した評点列に範囲外の値があれば ValueError（入力ミスの検出）
# ---------------------------------------------------------------------------


def test_rating_range_detects_out_of_range_value(tmp_path):
    """評点列に 99 が1つ混入していると ValueError。列名と値を示す。"""
    ratings = [row[:] for row in RATINGS]
    ratings[1][1] = 99  # 回答者2の【製品案2】
    f = _write_csv(tmp_path / "responses.csv", ratings=ratings)

    with pytest.raises(ValueError) as excinfo:
        pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))

    message = str(excinfo.value)
    assert RATING_COLS[1] in message
    assert "99" in message
    # 元ファイルを直せるように、該当する回答者も示す
    assert "respondent_id" in message


def test_out_of_range_value_column_is_not_dropped(tmp_path):
    """入力ミスが1つあるだけの列は候補から外さない（外すと検品が働かない）。

    「1つでも範囲外なら除外」にすると、正当な評点列が落ちて候補不足になり、
    入力ミスがエラーとして現れなくなってしまう。
    """
    ratings = [row[:] for row in RATINGS]
    ratings[0][3] = 99
    f = _write_csv(tmp_path / "responses.csv", ratings=ratings)

    with pytest.raises(ValueError, match="範囲外の値"):
        pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))


# ---------------------------------------------------------------------------
# 5. 欠損（無回答）があっても 2. と 3. が正しく動く
# ---------------------------------------------------------------------------


def test_rating_range_with_missing_values(tmp_path):
    """無回答（空欄）があっても、列の同定は正しく行われる。"""
    ratings = [row[:] for row in RATINGS]
    ratings[1][1] = None  # 回答者2が【製品案2】に無回答
    f = _write_csv(tmp_path / "responses.csv", ratings=ratings)

    df = pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))
    assert _ratings_of(df, 1) == [9, 7, 3, 9]
    assert pd.isna(_ratings_of(df, 2)[1])


def test_rating_range_with_missing_values_still_detects_out_of_range(tmp_path):
    """無回答が混じっていても、範囲外の値は検出される（欠損は検品の対象外）。"""
    ratings = [row[:] for row in RATINGS]
    ratings[1][1] = None
    ratings[2][0] = 99
    f = _write_csv(tmp_path / "responses.csv", ratings=ratings)

    with pytest.raises(ValueError) as excinfo:
        pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))
    assert "99" in str(excinfo.value)


def test_all_missing_column_is_kept_as_candidate(tmp_path):
    """全員が無回答の列は判定できないので、候補から外さない。"""
    ratings = [[9, None, 3, 9], [5, None, 8, 4], [10, None, 2, 6]]
    f = _write_csv(tmp_path / "responses.csv", ratings=ratings)

    df = pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))
    assert _ratings_of(df, 1)[0] == 9
    assert pd.isna(_ratings_of(df, 1)[1])
    assert _ratings_of(df, 1)[2:] == [3, 9]


# ---------------------------------------------------------------------------
# 6. 候補不足：評点列が丸ごと範囲外になる指定
# ---------------------------------------------------------------------------


def test_rating_range_too_far_from_actual_scale(tmp_path):
    """rating_range=(50, 60) だと候補が全滅し、尺度違いを示す ValueError。"""
    f = _write_csv(tmp_path / "responses.csv")

    with pytest.raises(ValueError) as excinfo:
        pcr.forms_to_data(f, PROFILES, rating_range=(50, 60))

    message = str(excinfo.value)
    assert "rating_range が実際の評点尺度と合っていない可能性" in message
    # 除外した列とその実際の値域を示す（どこがおかしいか分かるように）
    assert RATING_COLS[0] in message
    assert "値域: 5 〜 10" in message  # 【製品案1】の実際の値（9, 5, 10）
    assert "値域: 19 〜 21" in message  # 年齢列も外れる


# ---------------------------------------------------------------------------
# 7. 範囲が狭すぎる場合は「候補不足」ではなく「検品エラー」になる
# ---------------------------------------------------------------------------


def test_narrow_rating_range_raises_value_check_error_not_shortage(tmp_path):
    """1〜10 の評点に rating_range=(1, 5) を指定した場合の設計上の挙動。

    評点列は範囲内の値（3 や 5）も持つため「値がすべて範囲外」には該当せず、
    候補から外れない。したがって候補不足ではなく、段階2の検品エラーになる。
    これは意図した挙動である（1つでも範囲外なら除外する設計にすると、
    入力ミスが1つあるだけの正当な評点列まで落ちてしまうため）。
    エラー文には尺度違いの可能性も併記しているので、どちらの原因でも
    利用者は次にすべきことが分かる。
    """
    f = _write_csv(tmp_path / "responses.csv")

    with pytest.raises(ValueError) as excinfo:
        pcr.forms_to_data(f, PROFILES, rating_range=(1, 5))

    message = str(excinfo.value)
    assert "範囲外の値" in message
    assert "rating_range が実際の\n  評点尺度と合っていない可能性" in message


# ---------------------------------------------------------------------------
# 8. rating_range そのものの指定ミス
# ---------------------------------------------------------------------------


def test_rating_range_not_a_pair(tmp_path):
    """数値ひとつだけを渡すと、日本語で指定方法を示す ValueError。"""
    f = _write_csv(tmp_path / "responses.csv")
    with pytest.raises(ValueError, match="2つの数値で指定してください"):
        pcr.forms_to_data(f, PROFILES, rating_range=5)


def test_rating_range_reversed(tmp_path):
    """(最大値, 最小値) の順で渡すと、順序を指摘する ValueError。"""
    f = _write_csv(tmp_path / "responses.csv")
    with pytest.raises(ValueError, match="最小値.*が最大値"):
        pcr.forms_to_data(f, PROFILES, rating_range=(10, 1))


# ---------------------------------------------------------------------------
# 9. 正常系では警告を出さない／曖昧さが残るときだけ出す
# ---------------------------------------------------------------------------


def test_no_warning_when_rating_range_resolves_ambiguity(tmp_path):
    """年齢列があっても、絞り込みでちょうど n_profiles 列になれば警告なし。

    年齢や学年を聞くアンケートは一般的なので、この正常系で毎回警告が出ると
    本当に危険な警告への感度が下がる。
    """
    f = _write_csv(tmp_path / "responses.csv")

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        df = pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))

    assert [str(w.message) for w in caught] == []
    assert _ratings_of(df) == [9, 7, 3, 9]


def test_warns_when_ambiguity_remains_after_filtering(tmp_path):
    """絞り込んでも候補が n_profiles を超える場合は、除外した列を知らせる。

    範囲内の値をとる数値設問（満足度など）が残ると位置で選ぶしかないため、
    候補から外した列（年齢）も確認の手がかりとして示す。
    """
    f = _write_csv(
        tmp_path / "responses.csv",
        extra_cols={"全体の満足度（1〜5）": [4, 2, 5]},
    )

    with pytest.warns(UserWarning) as record:
        pcr.forms_to_data(f, PROFILES, rating_range=(1, 10))

    messages = "\n".join(str(w.message) for w in record)
    assert "評点列ではないと判断しました" in messages
    assert AGE_COL in messages
