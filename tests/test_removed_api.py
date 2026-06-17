"""トップレベルAPI廃止（v0.4.0）の検証テスト。

- 旧API名（pc.fit など）へのアクセスが日本語の AttributeError になること
- __version__ や rating サブパッケージなど正当な属性は通ること
"""
import re
import sys
from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _dist_version
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

import py4conjoint

# v0.3.x までトップレベルに存在した旧API名
OLD_API_NAMES = [
    "forms_to_conjoint_data",
    "design_profiles",
    "suggest_n_profiles",
    "encode",
    "auto_reference_levels",
    "fit",
    "ConjointResult",
    "check_design",
    "DesignCheckResult",
    "plot_importance",
    "plot_partworth",
    "plot_wtp",
]


@pytest.mark.parametrize("name", OLD_API_NAMES)
def test_old_api_raises_japanese_attribute_error(name):
    """旧API名へのアクセスは移行案内付きの AttributeError になる"""
    with pytest.raises(AttributeError) as excinfo:
        getattr(py4conjoint, name)
    msg = str(excinfo.value)
    assert f"py4conjoint.{name} は v0.4.0 で廃止されました。" in msg
    assert "`import py4conjoint.rating as pcr` を使ってください。" in msg


def test_unknown_attribute_raises_plain_attribute_error():
    """旧APIでない未知の属性は通常の AttributeError になる"""
    with pytest.raises(AttributeError) as excinfo:
        py4conjoint.no_such_attribute
    assert "廃止" not in str(excinfo.value)


def test_version_is_accessible():
    """__version__ は正当な属性としてアクセスでき、妥当なバージョン文字列である。

    特定のバージョン番号をハードコードせず（リリースのたびにテストが壊れない）、
    (1) 文字列であること、(2) セマンティックバージョニング形式であること、
    (3) 配布メタデータ（pyproject.toml 由来）と一致すること、を検証する。
    (3) は __init__.py の __version__ と pyproject.toml のバージョンが
    食い違ったまま公開される事故を防ぐ。
    """
    v = py4conjoint.__version__
    assert isinstance(v, str)
    # 例: "0.4.0" / "0.4.0a1" / "1.2.3rc2" など（先頭が X.Y.Z）
    assert re.match(r"^\d+\.\d+\.\d+", v), f"バージョン形式が不正です: {v!r}"
    # 配布メタデータと一致すること（パッケージが導入済みのときのみ検証）
    try:
        dist_v = _dist_version("py4conjoint")
    except PackageNotFoundError:
        pytest.skip("py4conjoint が未インストールのためメタデータ照合をスキップ")
    assert v == dist_v, (
        f"__init__.py の __version__ ({v!r}) と配布メタデータ "
        f"({dist_v!r}) が一致しません"
    )


def test_rating_subpackage_is_accessible():
    """py4conjoint.rating はサブパッケージとしてアクセスできる"""
    assert callable(py4conjoint.rating.fit)
    assert callable(py4conjoint.rating.encode)


def test_rating_import_as_alias():
    """推奨の import 形式（import py4conjoint.rating as pcr）が動く"""
    import py4conjoint.rating as pcr

    assert callable(pcr.fit)
