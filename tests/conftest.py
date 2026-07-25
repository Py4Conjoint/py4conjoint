"""テスト共通のフィクスチャ。

壊れた .xlsx は性質の異なる2パターンを用意する。
コード上も別の分岐を通るため、両方を検証できるようにしておく。

- broken_not_zip_xlsx : ZIP 構造の検査（zipfile.is_zipfile）で弾かれる
- broken_zip_ok_xlsx  : ZIP としては正常で、エンジンが実際に読みにいって失敗する

いずれも実行時に tmp_path へ生成する（壊れたバイナリをリポジトリに
置かずに済み、生成手順がコードとして読める）。生成結果が意図と違うと
テストが無意味になるため、フィクスチャ内で構造をアサートしている。
"""
import zipfile
from pathlib import Path

import pytest

DATA_DIR = Path(__file__).resolve().parent / "data"
REAL_XLSX = DATA_DIR / "forms_cbc_smartphone_real.xlsx"

# 実 fixture には xl/worksheets/_rels/sheet1.xml.rels もあるため、
# 差し替え対象はパスまで含めて指定する。
_SHEET1_PATH = "xl/worksheets/sheet1.xml"


@pytest.fixture
def broken_not_zip_xlsx(tmp_path: Path) -> Path:
    """ZIP 構造の検査で弾かれる .xlsx。

    先頭は ZIP のシグネチャ（PK\\x03\\x04）だが、中央ディレクトリが
    ないため zipfile.is_zipfile() は False になる。ブラウザ経由の
    ファイル転送で先頭だけ残るような破損を模したもの。
    """
    p = tmp_path / "broken_not_zip.xlsx"
    p.write_bytes(b"PK\x03\x04" + b"\x00" * 512)
    assert not zipfile.is_zipfile(p)
    return p


@pytest.fixture
def broken_zip_ok_xlsx(tmp_path: Path) -> Path:
    """ZIP としては正常だが、中身の XML が壊れている .xlsx。

    実ファイルの xl/worksheets/sheet1.xml だけを b"<broken" に
    差し替えて書き出す。is_zipfile() は True になるため、
    読み込みエンジンが実際に読みにいって失敗する経路を通る。
    """
    if not REAL_XLSX.exists():
        pytest.skip(f"実ファイルがありません: {REAL_XLSX}")

    p = tmp_path / "broken_zip_ok.xlsx"
    replaced = 0
    with zipfile.ZipFile(REAL_XLSX) as src, zipfile.ZipFile(p, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.endswith(_SHEET1_PATH):
                data = b"<broken"
                replaced += 1
            dst.writestr(item, data)

    # 0件だと正常なファイルができてしまい、テストが無意味になる
    assert replaced == 1, f"sheet1.xml の差し替え件数が {replaced} 件です"
    # ZIP としては正常であることが、このフィクスチャの要件
    assert zipfile.is_zipfile(p)
    return p


@pytest.fixture
def real_xlsx() -> Path:
    """実 Microsoft Forms 出力の .xlsx（読み込み専用）。"""
    if not REAL_XLSX.exists():
        pytest.skip(f"実ファイルがありません: {REAL_XLSX}")
    return REAL_XLSX
