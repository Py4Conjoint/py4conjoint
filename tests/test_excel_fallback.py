"""Excel 読み込みのフォールバックとエラー案内の検証。

v0.5.0 で openpyxl を必須依存から外し、読み込みエンジンを
calamine → openpyxl の順に試すカスケードに変更した。ここでは

- 壊れた .xlsx で、破損の性質に応じた日本語の案内が出ること
- 先頭のエンジンが失敗しても次のエンジンにフォールバックすること
- エンジンが1つも無いときはインストール方法を案内すること
- .csv 経路（推奨経路）が警告なしで通ること

を検証する。rating / choice の両方で確認する項目は、
pcr / pcc の対称性を担保する目的なので必ず両方書く。
"""
import importlib.util
import sys
import warnings
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd
import pytest

import py4conjoint.choice as pcc
import py4conjoint.rating as pcr
import py4conjoint.rating._forms as rating_forms

DATA_DIR = Path(__file__).resolve().parent / "data"
REAL_XLSX = DATA_DIR / "forms_cbc_smartphone_real.xlsx"
DESIGN_CSV = DATA_DIR / "design_smartphone_cbc.csv"

# rating 用のプロファイル設計（4案）
PROFILES = {
    "price": [6, 10, 6, 10],
    "os": ["android", "apple", "apple", "android"],
    "camera": ["標準", "標準", "高性能", "高性能"],
}

# choice 用（実 fixture と同じ設計・ラベル）
LABELS = ["製品A", "製品B", "製品C"]

# 導入されている読み込みエンジン。どのエンジンが使えるかで通る経路が変わる
# （JupyterLite は calamine のみ、pip install py4conjoint[excel] は openpyxl のみ）。
HAS_OPENPYXL = importlib.util.find_spec("openpyxl") is not None
HAS_CALAMINE = importlib.util.find_spec("python_calamine") is not None


@pytest.fixture(scope="module")
def design():
    return pd.read_csv(DESIGN_CSV)


def _make_ms_csv(path: Path, question_cols: dict) -> Path:
    """Microsoft Forms 形式の模擬回答を .csv で作る（管理列つき）。"""
    n = len(next(iter(question_cols.values())))
    system = {
        "ID": range(1, n + 1),
        "Start time": ["2026-06-01 10:00"] * n,
        "Completion time": ["2026-06-01 10:05"] * n,
        "Email": ["anonymous"] * n,
        "Name": [""] * n,
    }
    pd.DataFrame({**system, **question_cols}).to_csv(
        path, index=False, encoding="utf-8-sig"
    )
    return path


# ---------------------------------------------------------------------------
# 1. ZIP 構造の検査で弾かれる破損（_CORRUPT_XLSX_MESSAGE）
# ---------------------------------------------------------------------------

def test_rating_broken_not_zip_error(broken_not_zip_xlsx):
    """rating：ZIP でない .xlsx は、破損を断定した案内で弾かれる。"""
    with pytest.raises(ValueError) as exc:
        pcr.forms_to_data(str(broken_not_zip_xlsx), PROFILES)
    msg = str(exc.value)
    assert "壊れている可能性が高い" in msg
    assert "CSV" in msg


def test_choice_broken_not_zip_error(broken_not_zip_xlsx, design):
    """choice：rating と同じ案内が出る（サブパッケージ間の対称性）。"""
    with pytest.raises(ValueError) as exc:
        pcc.forms_to_data(str(broken_not_zip_xlsx), design, LABELS)
    msg = str(exc.value)
    assert "壊れている可能性が高い" in msg
    assert "CSV" in msg


# ---------------------------------------------------------------------------
# 2. ZIP は正常だが中身が壊れている（_UNREADABLE_EXCEL_MESSAGE）
# ---------------------------------------------------------------------------

def _assert_unreadable_excel_message(msg: str) -> None:
    """_UNREADABLE_EXCEL_MESSAGE 経由であることと、内訳の中身を確認する。

    どのエンジンが内訳に載るかは環境によって変わるため、導入されている
    エンジンだけを条件付きで確認する。エンジンが未導入だと read_errors
    ではなく engine_errors に入り、この経路自体を通らないため。
    """
    assert "対応していない可能性" in msg
    assert "CSV" in msg
    # 実際に読みにいったエンジンの内訳が載る
    assert "各エンジンで発生したエラー" in msg
    if HAS_CALAMINE:
        assert "calamine" in msg
    if HAS_OPENPYXL:
        assert "openpyxl" in msg


@pytest.mark.skipif(
    not (HAS_OPENPYXL or HAS_CALAMINE),
    reason="エンジンが1つも無いと読み込みに到達せず ImportError になる",
)
def test_rating_broken_zip_ok_error(broken_zip_ok_xlsx):
    """rating：エンジンが読みにいって失敗した場合は断定せず、内訳を出す。"""
    with pytest.raises(ValueError) as exc:
        pcr.forms_to_data(str(broken_zip_ok_xlsx), PROFILES)
    _assert_unreadable_excel_message(str(exc.value))


@pytest.mark.skipif(
    not (HAS_OPENPYXL or HAS_CALAMINE),
    reason="エンジンが1つも無いと読み込みに到達せず ImportError になる",
)
def test_choice_broken_zip_ok_error(broken_zip_ok_xlsx, design):
    """choice：rating と同じ案内が出る（サブパッケージ間の対称性）。"""
    with pytest.raises(ValueError) as exc:
        pcc.forms_to_data(str(broken_zip_ok_xlsx), design, LABELS)
    _assert_unreadable_excel_message(str(exc.value))


# ---------------------------------------------------------------------------
# 3. カスケード：先頭のエンジンが失敗しても次へ進む
# ---------------------------------------------------------------------------

def test_falls_back_to_openpyxl_when_calamine_raises(monkeypatch):
    """calamine が RuntimeError を投げても openpyxl で読めれば成功する。

    JupyterLite では calamine が第一候補になるため、ここで打ち切ると
    「openpyxl なら読めるファイル」を破損と誤診してしまう。
    """
    pytest.importorskip("openpyxl")
    if not REAL_XLSX.exists():
        pytest.skip(f"実ファイルがありません: {REAL_XLSX}")

    real_read_excel = pd.read_excel
    tried = []

    def fake_read_excel(path, engine=None, **kwargs):
        tried.append(engine)
        if engine == "calamine":
            raise RuntimeError("calamine: unsupported workbook structure")
        return real_read_excel(path, engine=engine, **kwargs)

    monkeypatch.setattr(rating_forms.pd, "read_excel", fake_read_excel)

    df = rating_forms._read_microsoft_forms(REAL_XLSX)

    assert tried == ["calamine", "openpyxl"]
    pd.testing.assert_frame_equal(df, real_read_excel(REAL_XLSX, engine="openpyxl"))


# ---------------------------------------------------------------------------
# 4. エンジンが1つも無い場合はインストール方法を案内する
# ---------------------------------------------------------------------------

def test_import_error_when_no_engine_available(monkeypatch, real_xlsx):
    """全エンジンが ImportError なら、導入方法を示す ImportError を出す。"""

    def always_missing(path, engine=None, **kwargs):
        raise ImportError(f"Missing optional dependency for engine={engine!r}")

    monkeypatch.setattr(rating_forms.pd, "read_excel", always_missing)

    with pytest.raises(ImportError) as exc:
        rating_forms._read_microsoft_forms(real_xlsx)
    msg = str(exc.value)
    assert "py4conjoint[excel]" in msg
    assert "python-calamine" in msg


# ---------------------------------------------------------------------------
# 5. .csv 経路（推奨経路）は警告なしで通る
# ---------------------------------------------------------------------------

def test_rating_csv_with_microsoft_emits_no_warning(tmp_path):
    """rating：forms="microsoft" に .csv を渡しても警告は出ない。"""
    f = _make_ms_csv(
        tmp_path / "responses.csv",
        {f"Q{i + 1}. 製品案{i + 1}の評価": [5, 6, 4] for i in range(4)},
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        df = pcr.forms_to_data(str(f), PROFILES)
    assert not [w for w in caught if "forms='microsoft'" in str(w.message)]
    assert len(df) == 3 * 4


def test_choice_csv_with_microsoft_emits_no_warning(tmp_path, design):
    """choice：rating と同じく .csv でも警告なし（対称性）。"""
    n_sets = design["choice_set_id"].nunique()
    f = _make_ms_csv(
        tmp_path / "responses.csv",
        {f"Q{i + 1}. どの製品を選びますか？": ["製品A", "製品B", "製品C"]
         for i in range(n_sets)},
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        df = pcc.forms_to_data(str(f), design, LABELS)
    assert not [w for w in caught if "forms='microsoft'" in str(w.message)]
    assert (df.groupby("choice_set_id")["choice"].sum() == 1).all()


# ---------------------------------------------------------------------------
# 6. エンジン間で読み取り結果が一致する
#    （test_choice_forms_real.py から移動）
# ---------------------------------------------------------------------------

def _normalize_newlines(df: pd.DataFrame) -> pd.DataFrame:
    """セル内改行を CRLF から LF に揃える（列名・値の両方）。

    古い python-calamine（0.4.0 以前）はセル内改行を CRLF のまま返すため、
    openpyxl（LF に正規化する）と列名が一致しない。Python 3.9 では
    python-calamine 0.8.2 が requires_python >=3.10 のため 0.4.0 しか
    入らず、この差異が出る。Pyodide 同梱の 0.6.2 では発生しない。

    ここで見たいのは回答データの中身がエンジン間で一致することなので、
    この既知の差異は正規化して比較する。
    """

    def fix(v):
        return v.replace("\r\n", "\n") if isinstance(v, str) else v

    out = df.rename(columns=fix)
    # 値側にも同じ正規化をかけておく（将来 CRLF を含むセルが来た場合のため）
    for col in out.columns:
        if pd.api.types.is_string_dtype(out[col]):
            out[col] = out[col].apply(fix)
    return out


def test_excel_engines_agree_on_real_file():
    """calamine と openpyxl が同じ .xlsx から同じ DataFrame を返す。

    _read_microsoft_forms() は calamine → openpyxl の順にエンジンを試すため、
    どちらが使われるかは環境（JupyterLite では calamine のみ）で変わる。
    エンジンの違いで読み取り結果がずれる回帰を検出するためのテスト。
    改行コードの既知の差異だけは正規化して比較する
    （:func:`_normalize_newlines` の説明を参照）。
    """
    pytest.importorskip("openpyxl")
    pytest.importorskip("python_calamine")

    df_openpyxl = pd.read_excel(REAL_XLSX, engine="openpyxl")
    df_calamine = pd.read_excel(REAL_XLSX, engine="calamine")
    pd.testing.assert_frame_equal(
        _normalize_newlines(df_openpyxl), _normalize_newlines(df_calamine)
    )
