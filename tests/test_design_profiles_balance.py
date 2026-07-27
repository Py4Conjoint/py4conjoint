"""rating の design_profiles(auto_balance=True) のテスト。

auto_balance の契約は次の2つ。

  (a) すべての属性について、水準の出現回数の最大と最小の差が 1 以下
  (b) (a) を満たす設計の中で det(X'X) が最大

(b) の照合には、**実装とは独立に総当たりで正解を求める** 補助関数
（:func:`_brute_force_best`）を使う。実装の出力を正解として固定すると、
実装が間違ったときにテストも一緒に間違うため。
"""

import sys
from itertools import combinations, product
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import numpy as np
import pandas as pd
import pytest

import py4conjoint.rating as pcr
from py4conjoint.rating import design as design_module


# ---------------------------------------------------------------------------
# 独立実装：総当たりで「制約なしの最良」と「均等制約下の最良」を求める
# ---------------------------------------------------------------------------


def _effect_matrix(rows, levels_list):
    """効果コーディング設計行列（先頭列 = 切片、基準水準は各リストの先頭）。"""
    cols = [np.ones(len(rows))]
    for a, levels in enumerate(levels_list):
        ref = levels[0]
        for lv in levels[1:]:
            cols.append(
                np.array(
                    [
                        1.0 if r[a] == lv else (-1.0 if r[a] == ref else 0.0)
                        for r in rows
                    ]
                )
            )
    return np.column_stack(cols)


def _is_balanced(rows, levels_list):
    for a, levels in enumerate(levels_list):
        counts = [sum(1 for r in rows if r[a] == lv) for lv in levels]
        if max(counts) - min(counts) > 1:
            return False
    return True


def _brute_force_best(levels_list, n):
    """(制約なしの最大 det, 均等制約下の最大 det) を総当たりで求める。"""
    cand = list(product(*levels_list))
    best_all, best_bal = 0.0, None
    for combo in combinations(range(len(cand)), n):
        rows = [cand[i] for i in combo]
        X = _effect_matrix(rows, levels_list)
        det = float(np.linalg.det(X.T @ X))
        best_all = max(best_all, det)
        if _is_balanced(rows, levels_list):
            best_bal = det if best_bal is None else max(best_bal, det)
    return best_all, best_bal


def _attrs_dict(levels_list):
    """levels_list を design_profiles の attribute_levels 形式にする。"""
    return {f"attr{i}": list(levels) for i, levels in enumerate(levels_list)}


def _level_counts(profiles, attr):
    return profiles[attr].value_counts().tolist()


def _det_of(profiles, attribute_levels):
    """返された設計の det(X'X) を、独立実装で計算し直す。"""
    levels_list = list(attribute_levels.values())
    rows = [
        tuple(profiles[a].iloc[i] for a in attribute_levels)
        for i in range(len(profiles))
    ]
    X = _effect_matrix(rows, levels_list)
    return float(np.linalg.det(X.T @ X))


# 変更5 の照合表（属性構成, n_profiles）
REFERENCE_CASES = [
    ([[0, 1], [0, 1], [0, 1]], 4),
    ([[0, 1], [0, 1], [0, 1]], 6),
    ([[0, 1, 2], [0, 1], [0, 1]], 6),
    ([[0, 1, 2], [0, 1], [0, 1]], 8),
]


# ---------------------------------------------------------------------------
# 1. auto_balance=False では従来どおり
# ---------------------------------------------------------------------------


def test_auto_balance_false_is_unchanged():
    """既定（False）では結果も attrs も従来のまま。"""
    attrs = {"price": [6, 10], "os": ["android", "apple"], "camera": ["標準", "高性能"]}
    before = pcr.design_profiles(attrs, 6, seed=42)
    after = pcr.design_profiles(attrs, 6, auto_balance=False, seed=42)
    pd.testing.assert_frame_equal(before, after)
    # 来歴は付かない
    assert "auto_balance" not in after.attrs
    assert set(before.attrs) == {"d_efficiency", "n_candidates", "det_xpx"}


# ---------------------------------------------------------------------------
# 2. 総当たりの正解と一致すること
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("levels_list, n", REFERENCE_CASES)
def test_matches_brute_force_optimum(levels_list, n):
    """均等制約下の最良解（総当たりで求めた正解）と det が一致する。"""
    expected_all, expected_bal = _brute_force_best(levels_list, n)
    assert expected_bal is not None  # 前提：この構成では実現可能

    attrs = _attrs_dict(levels_list)
    out = pcr.design_profiles(attrs, n, auto_balance=True, seed=0)
    info = out.attrs["auto_balance"]

    assert info["method"] == "exhaustive"  # 小さいので厳密解の経路
    assert info["balanced"] is True
    assert info["det_xpx"] == pytest.approx(expected_bal)
    assert info["det_xpx_unconstrained"] == pytest.approx(expected_all)
    assert info["det_ratio"] == pytest.approx(expected_bal / expected_all)


@pytest.mark.parametrize(
    "levels_list, n, expected_ratio",
    [
        ([[0, 1], [0, 1], [0, 1]], 4, 1.0),
        ([[0, 1], [0, 1], [0, 1]], 6, 0.75),
        ([[0, 1, 2], [0, 1], [0, 1]], 6, 1.0),
        ([[0, 1, 2], [0, 1], [0, 1]], 8, 0.9375),
    ],
)
def test_reference_ratios(levels_list, n, expected_ratio):
    """比が変更5の照合表（100% / 75% / 100% / 93.75%）と一致する。"""
    out = pcr.design_profiles(_attrs_dict(levels_list), n, auto_balance=True, seed=0)
    assert out.attrs["auto_balance"]["det_ratio"] == pytest.approx(expected_ratio)


# ---------------------------------------------------------------------------
# 3. 返る設計は必ず (a) を満たす
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "attrs, n",
    [
        ({"price": [6, 10], "os": ["a", "b"], "camera": ["x", "y"]}, 4),
        ({"price": [6, 10], "os": ["a", "b"], "camera": ["x", "y"]}, 6),
        ({"price": [6, 8, 10], "os": ["a", "b"], "camera": ["x", "y"]}, 6),
        ({"price": [6, 8, 10], "os": ["a", "b"], "camera": ["x", "y"]}, 8),
        ({"price": [6, 8, 10], "os": ["a", "b"], "camera": ["x", "y", "z"]}, 9),
        # 大きめ（発見的探索の経路に入る）
        (
            {
                "price": [6, 8, 10, 12],
                "brand": ["A", "B", "C"],
                "camera": ["低", "中", "高"],
                "os": ["a", "b"],
            },
            12,
        ),
    ],
)
def test_returned_design_is_balanced(attrs, n):
    """どの属性も「最大出現回数 − 最小出現回数 ≤ 1」になっている。"""
    out = pcr.design_profiles(attrs, n, auto_balance=True, seed=3)
    assert len(out) == n
    for attr in attrs:
        counts = _level_counts(out, attr)
        # 出現しない水準は value_counts に現れないので補う
        counts += [0] * (len(attrs[attr]) - len(counts))
        assert max(counts) - min(counts) <= 1, (attr, counts)
    assert out.attrs["auto_balance"]["balanced"] is True


def test_large_case_uses_exchange_method():
    """候補の選び方が閾値を超えると、発見的探索（exchange）に切り替わる。"""
    attrs = {
        "price": [6, 8, 10, 12],
        "brand": ["A", "B", "C"],
        "camera": ["低", "中", "高"],
        "os": ["a", "b"],
    }
    out = pcr.design_profiles(attrs, 12, auto_balance=True, seed=3)
    assert out.attrs["auto_balance"]["method"] == "exchange"


# ---------------------------------------------------------------------------
# 4. attrs に記録された比が、実際の det の比と一致する
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("levels_list, n", REFERENCE_CASES)
def test_attrs_ratio_matches_actual_dets(levels_list, n):
    """det_xpx は返した設計の実際の det であり、比もそれと整合する。"""
    attrs = _attrs_dict(levels_list)
    out = pcr.design_profiles(attrs, n, auto_balance=True, seed=0)
    info = out.attrs["auto_balance"]

    actual_det = _det_of(out, attrs)
    assert info["det_xpx"] == pytest.approx(actual_det)
    assert out.attrs["det_xpx"] == pytest.approx(actual_det)
    assert info["det_ratio"] == pytest.approx(
        info["det_xpx"] / info["det_xpx_unconstrained"]
    )


def test_full_factorial_records_history():
    """n_profiles == N（完全交差）でも来歴が入り、比は 1.0。"""
    attrs = {"price": [6, 10], "os": ["a", "b"], "camera": ["x", "y"]}
    out = pcr.design_profiles(attrs, 8, auto_balance=True)
    info = out.attrs["auto_balance"]
    assert info["method"] == "full_factorial"
    assert info["balanced"] is True
    assert info["det_ratio"] == 1.0


# ---------------------------------------------------------------------------
# 5. check_design の [大] バランス警告が消えること（今回の目的）
# ---------------------------------------------------------------------------


def test_auto_balance_removes_check_design_balance_warning():
    """2×2×2 を 6 プロファイルで設計すると、既定ではバランス [大] が出る。

    auto_balance=True にするとそれが消える。ただし契約は水準バランスだけで、
    属性間相関については何も約束していないため、相関の指摘は残りうる。
    """
    attrs = {"price": [6, 10], "os": ["android", "apple"], "camera": ["標準", "高性能"]}

    default = pcr.design_profiles(attrs, 6, seed=1)
    diags = pcr.check_design(default).diagnostics
    assert any(d.severity == "大" and d.category.startswith("balance") for d in diags)

    balanced = pcr.design_profiles(attrs, 6, auto_balance=True, seed=1)
    diags_b = pcr.check_design(balanced).diagnostics
    assert not any(d.category.startswith("balance") for d in diags_b)
    # バランスの評価はすべて ◎（CV = 0）
    assert (pcr.check_design(balanced).balance["CV"] == 0).all()


# ---------------------------------------------------------------------------
# 6. バランスを満たす設計が得られない場合は、例外ではなく警告
# ---------------------------------------------------------------------------

ATTRS_FOR_FALLBACK = {
    "price": [6, 10],
    "os": ["android", "apple"],
    "camera": ["標準", "高性能"],
}


def test_warns_instead_of_raising_when_exhaustive_finds_none(monkeypatch):
    """総当たりで解なしのとき：「存在しません」と断定してよい経路。

    完全交差から相異なるプロファイルを選ぶ限り、(a) が実現不可能な構成は
    見つかっていない（属性構成24通り × n=1〜N の456通りを走査して0件）。
    そのため防御的に残しているこの経路は、内部関数を差し替えて通す。
    """
    original = design_module._exhaustive_search

    def no_balanced(*args, **kwargs):
        _bal_idx, _bal_det, best_idx, best_det = original(*args, **kwargs)
        return None, -np.inf, best_idx, best_det

    monkeypatch.setattr(design_module, "_exhaustive_search", no_balanced)

    with pytest.warns(UserWarning, match="存在しません") as record:
        out = pcr.design_profiles(ATTRS_FOR_FALLBACK, 6, auto_balance=True, seed=1)

    # 例外にはせず、制約なしの最良解を返す
    assert len(out) == 6
    info = out.attrs["auto_balance"]
    assert info["balanced"] is False
    assert info["method"] == "exhaustive"
    # バランスを満たせなかったので比は定義しない（1.0 だと「損失なし」に読める）
    assert info["det_ratio"] is None
    # どの属性で均等にできなかったかを示す（どれになるかは最良解しだい）
    message = str(record[0].message)
    assert any(attr in message for attr in ATTRS_FOR_FALLBACK)


def test_warns_without_asserting_nonexistence_in_exchange(monkeypatch):
    """発見的探索で解なしのとき：「存在しない」と断定してはいけない経路。"""
    monkeypatch.setattr(design_module, "_EXHAUSTIVE_MAX_COMBINATIONS", 0)
    monkeypatch.setattr(design_module, "_balanced_exchange_run", lambda *a, **k: None)

    with pytest.warns(UserWarning, match="見つけられませんでした") as record:
        out = pcr.design_profiles(ATTRS_FOR_FALLBACK, 6, auto_balance=True, seed=1)

    assert len(out) == 6
    info = out.attrs["auto_balance"]
    assert info["balanced"] is False
    assert info["method"] == "exchange"
    assert info["det_ratio"] is None
    message = str(record[0].message)
    assert "存在しません" not in message
    assert "網羅的ではない" in message
