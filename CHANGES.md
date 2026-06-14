# Changelog

All notable changes to this project will be documented in this file.

---

## [0.4.0] - Unreleased

### Added
- `choice/` サブパッケージを追加（選択型コンジョイント分析・CBC）。条件付きロジット（conditional logit）を `scipy.optimize` による自前実装で推定する（教育目的のため透明性優先）。`fit()`・`encode()`・`ChoiceConjointResult`（`summary()` / `importance()` / `wtp()` / `market_share()` / `warnings()`）を提供し、メソッド名・日本語列名は `rating` 版と統一。
- `scipy>=1.8` を明示的依存に追加（`choice/` の最尤推定に使用。従来も statsmodels 経由で間接的に必要だった）。
- `choice/plot.py` を追加。`plot_importance()`・`plot_partworth()`・`plot_wtp()` を rating 版と同一のタイトル・日本語ラベルで提供し、`ChoiceConjointResult` のメソッドとしても呼び出せる。`plot_partworth()` はダミーコーディングの基準水準（係数0）を「{水準名}（基準）」として明示的に表示する。
- `choice/design.py` を追加。`design_choice_sets()`（完全交差からのランダム割り当てによるCBC用選択セット生成。セット内のプロファイル重複は禁止、`n_versions` 対応）、`check_design()`（水準バランス・属性間独立性・セット内オーバーラップ率を診断する `ChoiceDesignCheckResult` を返す。和文サマリー付き）、`suggest_n_respondents()`（Johnson-Orme の経験則 n ≥ 500c/(t×a) による必要回答者数の目安）。
- `cbc_forms_to_data()` を追加（`choice/_forms.py`）。Microsoft Forms / Google Forms の回答ファイル（1設問=1選択セット、回答選択肢=代替案）を、`design_choice_sets()` の出力と `choice_labels` によるマッチングで条件付きロジット推定用の long 形式（`choice_set_id`・`respondent_id`・`alt`・`choice` + 属性列）に変換する。設問数の不一致・未マッチ回答値は日本語エラー、未回答は警告のうえ該当選択セットを除外。
- `examples/overview_choice.ipynb` を追加。choice の全公開APIを網羅する教材ノートブック（設計→診断→必要回答者数→ヨーグルトデータでの推定→解釈→可視化→rating 版との違い）。
- `rating`・`choice` サブパッケージから `__version__` を参照可能にした（`pcr.__version__` / `pcc.__version__`）。
- **区間別 WTP** を `rating`・`choice` 双方の `wtp()` に追加。価格を符号化（rating は効果コーディング、choice はダミーコーディング）すると各価格水準の効用が独立に推定されるため、価格が3水準以上のときは **隣接する価格水準の区間ごと** に別々の傾き（価格感応度）で WTP を計算する（`method="segment"`、デフォルト）。戻り値に `価格区間` 列が付き、属性 × 区間の行を返す。`method="linear"` で従来の線形近似1本（単一値）も選べる（教材用）。特定区間だけを取り出す `price_segment` 引数（ラベル文字列または `(low, high)` タプル）も追加。価格2水準のときは区間が1つなので従来どおり単一値を返す。
- README を更新。rating / choice の2節構成にし、choice のインストール〜クイックスタートを追加。コード例を `import py4conjoint.rating as pcr` 形式に更新。

### Changed
- 既存モジュール（`_forms.py`・`design.py`・`encoding.py`・`analysis.py`・`plot.py`）を `rating/` サブパッケージへ移動。評点型コンジョイント分析は `import py4conjoint.rating as pcr` で利用する。
- `cbc_forms_to_data()`：実 Microsoft Forms 出力での **設問列の検出を回答値ベースに変更**。設問の列名は依存せず（実ファイルでは列名が選択肢を含む長文になり、改行・全角空白・`\xa0` を含む）、回答値が `choice_labels` に一致する列だけを設問列として検出する。性別・利用OS などの属性質問の列は回答値が一致しないため自動的に除外される（列名ベースの候補数が `n_sets` と一致しない場合に発動。一致する場合は従来どおり）。
- `cbc_forms_to_data()`：`design` の `version` 列を **オプション扱い**に変更。`version` 列を持たない手作りの設計表（`choice_set_id`・`alt_id` + 属性列のCSVなど）をそのまま渡せるようになった（`version` 列が無い場合は設計全体を単一バージョンとして扱う）。
- `cbc_forms_to_data()` の docstring に、設問文・属性質問内の水準表記（例：`"Apple (iOS)"`）と `design` の水準表記（例：`"apple"`）は一致していなくてよいこと、ただし属性質問の回答を分析に使う場合は利用者側で正規化が必要なことを明記。
- **価格列の指定を `rating`・`choice` で統一**。両者とも `price_col` には「数値（6, 10 など）が入った数値列のラベル」（例：`"price"`）を渡す。choice で価格をダミーコーディングした場合（`price_6` など）も `price_col` には元の数値列名を渡せばよくなった（従来は数値線形列のみ対応）。どの符号化列が価格かは、数値列の水準と `encode()` の命名規則から **構成的に特定** する。`startswith` による前方一致を使わないため、`price_range_high` のような接頭辞が紛らわしい別属性の列を誤検出しない。
- `rating`・`choice` の WTP 計算ロジック・列名（`価格区間`）・警告（`wtp_price_linear_approx` は `method="linear"` のときのみ）を共通化。0.3.0 で追加した価格3水準以上の線形近似は `method="linear"` に整理し、デフォルトは区間別（`method="segment"`）にした。
- `plot_wtp()`（`rating`・`choice` 共通）を `wtp()` の表と **常に一致** するように変更。価格3水準以上のデフォルト（`method="segment"`）では、価格区間ごとに色分けした **グループ化棒グラフ**（横軸=属性、凡例=価格区間）を描く。価格2水準・`method="linear"` のときは従来どおり属性1本の棒グラフ（`method="linear"` ではタイトルに「線形近似」と明示）。`plot_wtp()` に `method`・`price_segment` 引数を追加。従来は内部で `method="linear"` 固定だった。
- **選択セット識別子の名前を `choice_set_id` に統一**。生成側・設計側・推定側でばらばらだった列名・引数名を揃えた：`cbc_forms_to_data()` の出力列 `obsID` → `choice_set_id`（引数 `obs_id_colname` → `choice_set_id_colname`）、`design_choice_sets()` の出力列 `set_id` → `choice_set_id`、`fit()` の引数 `choice_set_col` → `choice_set_id_col`（`respondent_id_col` と対称的な `_id_col` 形式）。代替案の識別子 `alt_id` は別概念のため変更なし。**後方互換なし**（旧名のエイリアスや `DeprecationWarning` は用意しない）。`examples/overview_choice.ipynb` に、1人の回答者（`respondent_id`）が複数の選択セット（`choice_set_id`）に回答する階層構造の説明を追加。なお外部検証用の `tests/data/yogurt.csv` は logitr 由来の慣習列名 `obsID` のまま維持し、テスト内で `choice_set_id` にリネームして `fit()` に渡す。
- `openpyxl` の下限を `>=3.0` から `>=3.1.5` に引き上げ。古い 3.1.0 系で学生環境のファイル読み込みが不安定になる問題を避けるため。

### Fixed
- `cbc_forms_to_data()`：実 Microsoft Forms の回答ファイル（設問列が長文・改行・全角空白・`\xa0` を含み、性別・利用OS などの属性質問が混在する）で、設問列の自動検出が失敗していた問題を修正。実ファイル（3人分の回答）での回帰テスト `tests/test_choice_forms_real.py` を追加。
- `cbc_forms_to_data(forms="google")`：実 Google Forms の回答 CSV（BOMなしUTF-8、設問列名に `【設問N】` などの識別子接頭辞、属性質問の混在）でも正しく変換できることを確認。読み込みは共通ヘルパー経由で `encoding="utf-8-sig"` を使うため BOMなし・BOM付きの両方に対応する。設問列の検出は回答値ベースのため接頭辞の有無に依存しない（設問の順序は列の出現順に従う）。実 Google ファイル（3人分）での検証テストを `tests/test_choice_forms_real.py` に追加。docstring と `examples/overview_choice.ipynb` に、Google の CSV は BOMなしUTF-8 のため Excel で開くと文字化けするが py4conjoint は正しく読める旨を注記。

### Removed
- トップレベルAPI（`pc.fit`・`pc.encode` 等）を廃止。旧API名へアクセスすると日本語の `AttributeError` で `py4conjoint.rating` への移行を案内する（`__version__` などの正当な属性は従来どおり）。

---

## [0.3.0] - 2026-06-02

### Added
- `design_profiles()` 関数を追加（`design.py` 新規）。D最適計画法（Fedorov交換アルゴリズム）により全水準の完全交差からM個のプロファイルを選択する。任意の属性数・任意の水準数・混在水準数に対応。numpy のみで実装（追加依存なし）。
- `suggest_n_profiles()` 関数を追加（`design.py`）。属性・水準数と予定回答者数から、統計的最低限・Orme の経験則・観測数条件に基づく推奨プロファイル数を計算する。
- `fit()` に `respondent_id_col`・`cluster_se` 引数を追加。回答者ID列がある場合、デフォルトで**回答者IDによるクラスタロバスト標準誤差**を使用する（同一回答者の複数回答は独立でないため、通常のOLS標準誤差ではp値が過小になる。係数の推定値は変わらない）。`summary()` に標準誤差の種別を表示。
- `independence_assumed` 警告（重大度：中）を追加。回答者ID列が見つからず観測の独立性を仮定した標準誤差を使っている場合に通知する。
- `importance()` の docstring に、重要度が調査で選んだ水準レンジに依存する相対指標である旨の注記を追加。
- `encode()` に `suffix_map` 引数を追加。2水準には `str`、3水準以上には `List[str]` を渡すことで列名サフィックスを指定可能。
- `check_design()` 関数と `DesignCheckResult` クラスを追加。アンケート実施前にプロファイルの直交性・バランス・独立性を診断する（scipy不要）。
- `wtp()` の価格3水準以上対応。線形近似（仮定）を用いてWTPを計算し、`wtp_price_linear_approx` 警告を自動追加。
- `unit_rating_money()` の価格3水準以上対応。

### Changed
- `binary_suffix_map` 引数を非推奨化（`DeprecationWarning` を出して `suffix_map` への移行を促す）。後方互換のため引数は残す。
- `_encode_multi()` にサフィックス引数を追加。
- `wtp()`・`plot_wtp()` のドキュメントに、計算値が厳密には限界支払意思額（MWTP：他属性一定のまま1属性を変えるときの追加支払額＝属性と価格の限界代替率）であり、製品全体に対する支払上限額（総WTP）ではないことを明記。
- `wtp()` の列名を `支払意思額` → `限界支払意思額` に変更（MWTPであることを明確化）。`plot_wtp()` の軸ラベル・タイトルも追従。
- `importance()` の列名を `range` → `効用範囲`、`importance` → `重要度` に変更（`wtp()` の列名と同様に日本語化）。
- `plot_importance()` のデフォルトタイトル・X軸ラベルを「属性の重要度」「重要度（%）」に、`plot_wtp()` を「属性の限界支払意思額」「限界支払意思額」に変更。

### Fixed
- `wtp()`：3水準以上の非価格属性のWTPが定義（基準水準からの効用差の金額換算）と一致しない値を返すバグを修正。属性ごとに `wtp_k = (b_k + Σb_j) × wtp_price_factor / 2` で計算するように変更。2水準属性の結果は従来と同一。
- `suggest_n_profiles()`：`max_burden` がパラメータ数 p を下回る場合に、回帰分析が実行不能な推奨値（< p）を警告なしで返すバグを修正。推奨値を p まで引き上げ、`UserWarning` を出すように変更。
- `wtp()`：3水準以上の価格では先頭の符号化列の p値のみで `price_insignificant` を判定していたのを、全価格係数の同時F検定の p値（`attrs["p_price"]`）に変更。2水準価格は従来どおり t 検定。
- `encode()`：attrs のネスト辞書（`reference_levels` 等）が入力 DataFrame と共有され、encode 済みデータを再 encode すると入力側の attrs まで書き換わる副作用を修正。
- `fit()`：符号化列の自動検出を改善。属性名が別の列名の接頭辞になっている場合の誤検出・重複登録を防止し、効果コーディング済み（-1/1 を含む）の列のみを採用するように変更（0/1 の回答者属性列は `encoded_columns` で明示指定する）。
- `fit(formula=...)`：formula 指定時に説明変数の自動検出と食い違い `importance()`/`wtp()` が `KeyError` になる問題を修正。被説明変数・説明変数を formula から取得するように変更。
- ビルド要件を `setuptools>=77` に引き上げ（PEP 639 の SPDX ライセンス文字列 `license = "MIT"` に必要）。
- `test_e2e_real_data` が環境固有の存在しないパス（`/mnt/user-data/uploads/test.xlsx`）を参照して常にスキップされていた問題を修正。リポジトリ内の `examples/responses_os.csv` を使うように変更。

---

## [0.2.3] - 2026-05-31

### Fixed
- `wtp()` を複数回呼んだとき `wtp_extrapolation` 警告が重複登録されるバグを修正
- `summary()` の係数表で全角文字を含む変数名があっても列がズレないよう修正

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
