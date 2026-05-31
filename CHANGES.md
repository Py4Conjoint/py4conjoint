# Changelog

All notable changes to this project will be documented in this file.

---

## [0.2.2] - 2026-05-31

### Changed
- 「カード」→「プロファイル」に全面統一
  - `card_id_prefix` → `profile_id_prefix`、`card_id_colname` → `profile_id_colname`、
    `n_cards` → `n_profiles`、内部関数 `_build_card_design` → `_build_profile_design` など
- `model_result` 属性を `ols` に改名
- `encode()` の `respondent_encode` で `[zero_level, suffix]` リスト形式をサポート
- `result.wtp()` の列名を日本語化（`coef` → `係数`、`wtp` → `支払意思額`）
- `result` をセルに入力したときの Jupyter Notebook 表示を HTML（`_repr_html_()`）に変更
  （有意性列が右揃えになり、見やすく）

### Fixed
- `warnings()` の改善
  - `price_sign_negative`：p 値 ≥ 0.10 の場合は符号がノイズ起因のため発火しないよう修正
  - `wtp_extrapolation`：重大度を常に「中」に統一（`price_insignificant` と重複するため）
  - `obs_per_predictor` 警告を新規追加（観測数／説明変数数 < 5 → 大、< 10 → 中）
- `plot_wtp()` が `wtp()` の列名変更（`"wtp"` → `"支払意思額"`）に追従できていなかった
  バグを修正（`KeyError: 'wtp'` が発生していた）

### Added
- `examples/overview_os.ipynb`：全公開APIを実データ（`responses_os.csv`）で動作確認するノートブック

---

## [0.2.1] - 2026-05-30

### Fixed
- `summary(slim=False)` が未実装だった問題を修正（statsmodels の詳細統計表を返すように）
- `fit()` および `_run_diagnostics()` のドキュメントで `few_respondents` の重大度説明が不正確だった点を修正
  （「5人未満→大」→「1人→大、2〜4人→中、5人以上は警告なし」）
- README の係数名・サマリー出力例を現行の列名形式（`price_0`, `os_0`, `camera_0`）に統一

### Changed
- `examples/overview.ipynb` を全公開APIを網羅する単一ノートブックに統合

### Tests
- `test_diagnostics_price_insignificant` を確定的な直交バランスデザインに変更
- `test_few_respondents_major/minor/no_warning`, `test_price_sign_negative`,
  `test_summary_slim_false`, `test_auto_reference_levels`, `test_market_share_max`,
  `test_encode_drop_original`, `test_encode_inplace`, `test_encode_binary_suffix_map`,
  `test_importance_ratio`, `test_wtp_attrs` を追加（計12テスト追加）

---

## [0.2.0] - 2026-05-17

### Added
- `encode()`: 効果コーディング（-1/1）を自動化
- `fit()`: OLS回帰を実行し `ConjointResult` を返す
- `ConjointResult` クラス：以下のメソッド・プロパティを提供
  - `params`, `rsquared`, `n_obs`, `intercept`（プロパティ）
  - `summary()`: 和文サマリー
  - `warnings()`: 落とし穴の一覧（severity/categoryフィルタ対応）
  - `importance()`: 相対重要度
  - `wtp()`: 支払意思額（WTP）
  - `unit_rating_money()`: 評点1点の金額換算（float）
  - `market_share()`: 市場シェア予測（logit/max）
  - `plot_importance()`, `plot_partworth()`, `plot_wtp()`: 可視化
- 落とし穴の自動検出（`r2_low`, `price_sign_negative`, `few_respondents`,
  `price_insignificant`, `wtp_extrapolation`）
- `auto_reference_levels()`: 基準水準の自動推測
- CI（GitHub Actions）にテストジョブを追加

### Changed
- パッケージ構成を単一ファイルから4モジュール構成に変更
  （`_forms.py`, `encoding.py`, `analysis.py`, `plot.py`）
- 依存関係に `numpy`, `statsmodels`, `matplotlib`, `openpyxl` を追加

---

## [0.1.2] - 2026-04-06
- 引数名 `responses_csv` を `responses_file` に変更した（`.xlsx` にも対応するため）。


## [0.1.1] - 2026-04-06

### Added
- `forms` 引数を追加。`"microsoft"`（デフォルト）と `"google"` を指定できる。
  - `"microsoft"` : Microsoft Forms の `.xlsx` または `.csv`（BOM付きUTF-8）を読み込む。`.xlsx` の読み込みには `openpyxl` が必要。
  - `"google"` : Google Forms の `.csv`（UTF-8 / BOM付きUTF-8）を読み込む。
- Microsoft Forms 用の管理列検出パターン（`_MICROSOFT_SYSTEM_PATTERNS`）を追加。
- Google Forms 用の管理列検出パターン（`_GOOGLE_SYSTEM_PATTERNS`）を追加。
- `forms` 引数に無効な値を渡した場合に `ValueError` を発生させるようにした。
- `openpyxl` が未インストールの場合に日本語のインストール案内を含む `ImportError` を発生させるようにした。
- `forms="microsoft"` を指定しているにもかかわらず `.xlsx`/`.xls` 以外の拡張子のファイルを渡した場合に `UserWarning` を発生させるようにした。処理は続行する。

### Changed
- 管理列の検出処理を `_detect_system_cols()` として共通化し、`_detect_microsoft_system_cols()` と `_detect_google_system_cols()` から呼び出す構造に変更した。
- ファイルが見つからない場合のエラーメッセージを `"CSVファイルが見つかりません"` から `"ファイルが見つかりません"` に変更した（`.xlsx` にも対応するため）。

---

## [0.1.0] - 2026-04-06

### Added
- `forms_to_conjoint_data()` 関数を実装。Google Forms の回答 CSV を評点型コンジョイント分析用の long 形式 DataFrame に変換する。
- `attributes` 引数に `pd.DataFrame`（形式A）と辞書のリスト（形式B）の2形式を受け付ける。
- `cards`（`pd.DataFrame`）をそのまま `attributes` に渡せる `_normalize_attributes()` を実装。
- `n_cards` と `attributes` の整合性チェック（`_check_attributes()`）を実装。
- 属性が1つのみの場合に `UserWarning` を発生させる（WTP計算不可の旨を通知）。
- Google Forms の管理列（タイムスタンプ・メールアドレス等）を自動検出して除外する `_detect_forms_system_cols()` を実装。
- 評点列を右端の数値列から自動検出する `_pick_rating_cols()` を実装。
- BOM 付き UTF-8 の CSV を正常に読み込めるよう `encoding="utf-8-sig"` を使用。
- `out_csv` 引数で変換後の DataFrame を CSV として保存できる機能を追加。
- プロファイル ID の接頭辞を `card_id_prefix` 引数で変更できる（デフォルト: `"P"`）。
- `responses_csv` が存在しない場合に `FileNotFoundError` を発生させる。
- `pyproject.toml`、`README.md`、`LICENSE`（MIT）、`.gitignore` を整備。
- GitHub Actions による PyPI への手動デプロイワークフロー（`publish.yml`）を追加。
