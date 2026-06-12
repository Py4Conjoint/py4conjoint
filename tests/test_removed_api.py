"""トップレベルAPI廃止（v0.4.0）の検証テスト。

- 旧API名（pc.fit など）へのアクセスが日本語の AttributeError になること
- __version__ や rating サブパッケージなど正当な属性は通ること
"""
import sys
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
    """__version__ は正当な属性としてアクセスできる"""
    assert isinstance(py4conjoint.__version__, str)


def test_rating_subpackage_is_accessible():
    """py4conjoint.rating はサブパッケージとしてアクセスできる"""
    assert callable(py4conjoint.rating.fit)
    assert callable(py4conjoint.rating.encode)


def test_rating_import_as_alias():
    """推奨の import 形式（import py4conjoint.rating as pcr）が動く"""
    import py4conjoint.rating as pcr

    assert callable(pcr.fit)
