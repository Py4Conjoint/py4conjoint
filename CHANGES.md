# Changelog

All notable changes to this project will be documented in this file.

---

## [0.4.1] - 2026-07-04

### Changed
- `wtp()` / `plot_wtp()`（`rating`・`choice` 共通）：**区間別 WTP でない状況で `price_segment` を指定すると日本語の `ValueError`** を出すようにした（価格2水準・数値線形価格・`method="linear"` のとき）。従来は黙って無視されるため「指定した区間の値が出ている」と誤解する恐れがあった。区間別（価格3水準以上かつ `method="segment"`）での挙動は不変。
- rating の `fit()` docstring（`encoded_columns`）に、回答者属性の 0/1 列（`respondent_encode` の出力）を説明変数に含めた場合、`importance()`・`wtp()` の出力にもその列が「属性」として現れるが、製品属性と同じようには解釈できない旨の注意書きを追加。
- README の依存表の `openpyxl` を `≥ 3.1.5` に更新（pyproject と整合）。pyproject の `description` を「評点型・選択型コンジョイント分析…」に更新（choice 追加を反映）。
- テスト：`__version__` の検証を特定のバージョン文字列とのハードコード比較ではなく、(1) semver 形式であること、(2) 配布メタデータ（`importlib.metadata.version`、pyproject 由来）と一致すること、の検証に変更。リリースでバージョンを上げてもテストが壊れず、`__init__.py` と pyproject のバージョン乖離は CI が確実に検知する。
- `forms_to_data()`・`check_design()`（`rating`・`choice` 共通）と `design_signature()`（choice）：**pandas の行番号列（`Unnamed: 0` など）を警告付きで無視**するようにした。design / profiles を `to_csv()` で保存するとき `index=False` を付け忘れると、読み込んだ CSV に行番号（rating では P1, P2, … のプロファイルID）が `Unnamed: 0` 列として混入する。従来は無言で「属性」として扱われ、(1) `forms_to_data()` の出力に属性として混入、(2) choice では署名が元の設計と食い違い「同じ設計のはずなのに署名が合わない」という混乱を招き、(3) rating の `check_design()` ではプロファイル数と同数の水準を持つ架空の属性としてパラメータ数が膨らみ、正常な設計に `insufficient_profiles`（分析不能）などの**誤警告**が大量に出ていた。行番号は設計の中身ではないため、`forms_to_data()` は日本語警告（`index=False` を付けて保存する案内つき）のうえ属性から除外し、`check_design()`・`design_signature()` は黙って無視する。これにより **`index=False` を忘れて保存した CSV でも正しく動き、choice の署名も元の設計と一致する**（`to_csv()` は pandas のメソッドなので保存側の既定は変えられず、読み込み側で無害化する方針）。
- `design_profiles()` / `suggest_n_profiles()`（rating）と `design_choice_sets()` / `suggest_n_respondents()`（choice）：**水準リストに重複がある属性（例：`{"price": [100, 150, 100]}` のようなタイプミス）を日本語の `ValueError` で弾く**ようにした。従来はエラーも警告も出ずに設計が静かに壊れていた：choice では完全交差に同一プロファイルが複数作られ、「同一選択セット内に同じプロファイルは入らない」という `design_choice_sets()` の保証が破れる（重複ありの実験では 200 設問中 73 設問に同一の代替案ペアが混入。どちらを選んでも同じ製品＝その設問は情報を生まない）。rating では候補数 N とパラメータ数 p が架空に膨らみ、`d_efficiency` や `n_profiles` の下限チェックが誤った値になっていた。
- rating の `forms_to_data()`：**`respondent_cols` で指定した列名がファイルに無い場合、日本語の `ValueError`** で列名の確認を促すようにした。従来は pandas の生の `KeyError`（英語：`"['列名'] not in index"`）で原因が分かりにくかった。choice 版 `forms_to_data()` には同じチェックが既にあり、rating・choice で挙動が揃った。
- choice の `forms_to_data()`：**`choice_labels` に重複がある場合（例：`["A", "A", "C"]`）、日本語の `ValueError`** を即座に出すようにした。重複を許すと回答とのマッチングが曖昧になり、従来は後段で「choice_labels のどれにもマッチしない回答値があります」という真因（ラベル重複）から遠いエラーになっていた（回答が重複ラベルと完全一致する稀なケースでは、無言で先頭の代替案に誤って割り当てられる）。

### Fixed
- choice の `fit()`（条件付きロジットの最尤推定）が、大標本でまれに収束判定に失敗する問題を修正。最適化の目的関数を「合計」負の対数尤度から「1選択セットあたりの平均」に変更した。合計だと標本が大きいほど目的関数の絶対値が巨大になり、最適点近傍で BFGS の直線探索が浮動小数点の精度落ちを起こして `converged=False`（＝完全分離の誤警告）になることがあった（環境差で再現が不安定。CI の `test_synthetic_recovery` が推定値は正しいのに収束失敗で落ちていた）。平均化しても最小化の解は不変のため、係数・標準誤差・対数尤度・McFadden R² などの数値結果は一切変わらず、大標本でも収束判定が安定し、収束許容値 `gtol` が標本サイズに依らず一定の意味を持つようになる。
- rating の `encode()`：**3水準以上の属性で欠損（NaN）が 0 に符号化されてしまう問題を修正**。`Series.map` が既定で NaN にも変換関数を適用するため、欠損行の全ダミー列が 0（効果コーディングでは「全水準の平均」を意味する値）になり、`fit()` の欠損除外をすり抜けて回帰に静かに混入していた。`na_action="ignore"` で欠損を欠損のまま残すよう修正し、2水準（NaN 保持）・choice 版（`pd.NA` 保持）と挙動を統一した。
- choice の `design_signature()`：**署名が numpy のバージョンに依存する問題を修正**。セル値（numpy スカラー）の `repr` をそのままハッシュしていたため、numpy 2.x（`repr(np.int64(6))` → `'np.int64(6)'`）と 1.x（`'6'`）で同一設計の署名が食い違った。署名は「アンケート作成時と分析時（時間・環境をまたぐ）の design 同一性確認」が目的なので、ハッシュ前に Python の値へ正規化（`.item()`）して環境非依存にした（numpy 1.26 と 2.4 で同一署名になることを確認済み）。**注意**：numpy 2.x 環境でこれまでに記録した署名は今回の修正で値が変わる（numpy 1.x 環境の署名とは一致する）。`examples/overview_choice.ipynb` は再実行済み。
- rating の `forms_to_data()`：**評点列の自動検出が無警告でズレ得る問題への対策**。数値の候補列が `n_profiles` を超える場合（評点でない数値質問——満足度・年齢など——が混在し `respondent_cols` で指定されていない場合）に、従来は右端の n 列を黙って採用して評点とプロファイルの対応が静かに崩れることがあった。採用列・除外列を明示する `UserWarning` を出すようにした。あわせて (1) 評点列を melt 後に `pd.to_numeric(errors="coerce")` で数値化（Forms 出力で評点が文字列 `"5"` の場合も `fit()` まで通る。数値化できない値は件数つきで警告して NaN → 既存の欠損処理に乗る）、(2) 出力の行順を `profile_id` の**提示順**（P1, P2, …, P10, …）に修正（従来は文字列の辞書順で `n_profiles ≥ 10` のとき P1, P10, P11, P2, … と並んだ。データの対応自体は正しく、行順のみの問題）。
- `examples/` の3ノートブック（`overview.ipynb`・`overview_os.ipynb`・`overview_choice.ipynb`）の出力セルに残っていた旧バージョン表記 `0.4.0a1` を `0.4.0` に更新（コードは不変、表示のみ）。

---

## [0.4.0] - 2026-06-17

### Added
- `choice/` サブパッケージを追加（選択型コンジョイント分析・CBC）。条件付きロジット（conditional logit）を `scipy.optimize` による自前実装で推定する（教育目的のため透明性優先）。`fit()`・`encode()`・`ChoiceConjointResult`（`summary()` / `importance()` / `wtp()` / `market_share()` / `warnings()`）を提供し、メソッド名・和文の表示ラベルは `rating` 版と統一。
- `scipy>=1.8` choice/の最尤推定でscipy.optimizeを直接使用するため追加 
- `choice/plot.py` を追加。`plot_importance()`・`plot_partworth()`・`plot_wtp()` を rating 版と同一のタイトル・日本語ラベルで提供し、`ChoiceConjointResult` のメソッドとしても呼び出せる。`plot_partworth()` はダミーコーディングの基準水準（係数0）を「{水準名}（基準）」として明示的に表示する。
- `choice/design.py` を追加。`design_choice_sets()`（完全交差からのランダム割り当てによるCBC用選択セット生成。セット内のプロファイル重複は禁止、`n_versions` 対応）、`check_design()`（水準バランス・属性間独立性・セット内オーバーラップ率を診断する `ChoiceDesignCheckResult` を返す。和文サマリー付き）、`suggest_n_respondents()`（Johnson-Orme の経験則 n ≥ 500c/(t×a) による必要回答者数の目安）。
- `forms_to_data()` を追加（`choice/_forms.py`）。Microsoft Forms / Google Forms の回答ファイル（1設問=1選択セット、回答選択肢=代替案）を、`design_choice_sets()` の出力と `choice_labels` によるマッチングで条件付きロジット推定用の long 形式（`respondent_id`・`choice_set_id`・`alt`・`choice` + 属性列）に変換する。設問数の不一致・未マッチ回答値は日本語エラー、未回答は警告のうえ該当選択セットを除外。
- `examples/overview_choice.ipynb` を追加。choice の全公開APIを網羅する教材ノートブック（スマートフォンを題材に、設計→診断→必要回答者数→Google Forms 読み込み→推定→解釈→可視化→価格3水準の区間別 WTP→rating 版との違い）。
- **`design_signature()`（choice）を追加**。選択セット設計（design）の内容（`version`・`choice_set_id`・`alt_id` + 各属性の値）から決定的に計算する署名（12桁の十六進ハッシュ）。属性名・水準・**水準の順序**・`n_sets`・`n_alts` が完全に同一なら同じ署名になり、順序が1つでも違えば別の署名になる（`seed` を指定して再生成した場合や、CSV に保存して読み込み直した場合も内容が同じなら一致する。`seed=None` は呼ぶたびに中身が変わるため署名も変わる）。アンケート作成に使った design と、分析時に `forms_to_data()` へ渡す design が同一かを**確認する**ために使う。`design_choice_sets()` の出力 `df.attrs["design_signature"]` にも自動付与し、`forms_to_data()` は使った design の署名を出力 `df.attrs["design_signature"]` に引き継ぐ。
- **`design_choice_sets()`（choice）に `auto_balance` を追加**（バランスの良い設計を自動で選ぶ）。`auto_balance=True` にすると、内部で `n_candidates` 個（既定 500）のランダム設計を生成し、`check_design()` の診断で最もバランスの良いものを1つ選んで返す。選び方（方式D）は「①警告ゼロの候補を優先 → ②その中で全属性の CV 合計が最小 → ③警告ゼロが無ければ警告数が最少（同数なら CV 合計が最小）」で、既存の `check_design()` のロジックをそのまま再利用する（新しい診断基準は作らない）。これにより学生が良い設計を得るために `seed` を手で 1, 2, 3… と探す必要がなくなる。**数学的な最適計画（D 最適計画）ではなく**「多数の候補から最もバランスの良いものを選ぶ」方法のため、引数名は `optimize` ではなく `auto_balance` とした。`seed` を指定すれば候補生成は決定的に派生し、同じ `seed`・引数なら必ず同じ設計（同じ `design_signature`）を返す（`seed=None` は毎回変わる）。選定の来歴を `df.attrs["auto_balance"]`（`{"n_candidates", "n_warnings", "cv_sum"}`）に残す。**既定は `auto_balance=False`** で従来どおりの単一ランダム生成（完全な後方互換）。大規模設計では時間がかかるため `n_candidates` で調整できる。なお rating 側の `design_profiles()` は元々 D 最適計画（Fedorov 交換・多スタート）で最適化済みのため、同種の `auto_balance` は追加しない（seed 探しの問題が無い）。
- `rating`・`choice` サブパッケージから `__version__` を参照可能にした（`pcr.__version__` / `pcc.__version__`）。
- **区間別 WTP** を `rating`・`choice` 双方の `wtp()` に追加。価格を符号化（rating は効果コーディング、choice はダミーコーディング）すると各価格水準の効用が独立に推定されるため、価格が3水準以上のときは **隣接する価格水準の区間ごと** に別々の傾き（価格感応度）で WTP を計算する（`method="segment"`、デフォルト）。戻り値に `価格区間` 列が付き、属性 × 区間の行を返す。`method="linear"` で従来の線形近似1本（単一値）も選べる（教材用）。特定区間だけを取り出す `price_segment` 引数（ラベル文字列または `(low, high)` タプル）も追加。価格2水準のときは区間が1つなので従来どおり単一値を返す。
- README を更新。rating / choice の2節構成にし、choice のインストール〜クイックスタートを追加。コード例を `import py4conjoint.rating as pcr` 形式に更新。

### Changed
- 既存モジュール（`_forms.py`・`design.py`・`encoding.py`・`analysis.py`・`plot.py`）を `rating/` サブパッケージへ移動。評点型コンジョイント分析は `import py4conjoint.rating as pcr` で利用する。
- `forms_to_data()`：実 Microsoft Forms 出力での **設問列の検出を回答値ベースに変更**。設問の列名は依存せず（実ファイルでは列名が選択肢を含む長文になり、改行・全角空白・`\xa0` を含む）、回答値が `choice_labels` に一致する列だけを設問列として検出する。性別・利用OS などの属性質問の列は回答値が一致しないため自動的に除外される（列名ベースの候補数が `n_sets` と一致しない場合に発動。一致する場合は従来どおり）。
- `forms_to_data()`：`design` の `version` 列を **オプション扱い**に変更。`version` 列を持たない手作りの設計表（`choice_set_id`・`alt_id` + 属性列のCSVなど）をそのまま渡せるようになった（`version` 列が無い場合は設計全体を単一バージョンとして扱う）。
- `forms_to_data()` の docstring に、設問文・属性質問内の水準表記（例：`"Apple (iOS)"`）と `design` の水準表記（例：`"apple"`）は一致していなくてよいこと、ただし属性質問の回答を分析に使う場合は利用者側で正規化が必要なことを明記。
- **価格列の指定を `rating`・`choice` で統一**。両者とも `price_col` には「数値（6, 10 など）が入った数値列のラベル」（例：`"price"`）を渡す。choice で価格をダミーコーディングした場合（`price_6` など）も `price_col` には元の数値列名を渡せばよくなった（従来は数値線形列のみ対応）。どの符号化列が価格かは、数値列の水準と `encode()` の命名規則から **構成的に特定** する。`startswith` による前方一致を使わないため、`price_range_high` のような接頭辞が紛らわしい別属性の列を誤検出しない。
- `rating`・`choice` の WTP 計算ロジック・列名（`価格区間`）・警告（`wtp_price_linear_approx` は `method="linear"` のときのみ）を共通化。0.3.0 で追加した価格3水準以上の線形近似は `method="linear"` に整理し、デフォルトは区間別（`method="segment"`）にした。
- `plot_wtp()`（`rating`・`choice` 共通）を `wtp()` の表と **常に一致** するように変更。価格3水準以上のデフォルト（`method="segment"`）では、価格区間ごとに色分けした **グループ化棒グラフ**（横軸=属性、凡例=価格区間）を描く。価格2水準・`method="linear"` のときは従来どおり属性1本の棒グラフ（`method="linear"` ではタイトルに「線形近似」と明示）。`plot_wtp()` に `method`・`price_segment` 引数を追加。従来は内部で `method="linear"` 固定だった。
- **選択セット識別子の名前を `choice_set_id` に統一**。生成側・設計側・推定側でばらばらだった列名・引数名を揃えた：`forms_to_data()` の出力列 `obsID` → `choice_set_id`（引数 `obs_id_colname` → `choice_set_id_colname`）、`design_choice_sets()` の出力列 `set_id` → `choice_set_id`、`fit()` の引数 `choice_set_col` → `choice_set_id_col`（`respondent_id_col` と対称的な `_id_col` 形式）。代替案の識別子 `alt_id` は別概念のため変更なし。**後方互換なし**（旧名のエイリアスや `DeprecationWarning` は用意しない）。`examples/overview_choice.ipynb` に、1人の回答者（`respondent_id`）が複数の選択セット（`choice_set_id`）に回答する階層構造の説明を追加。なお外部検証用の `tests/data/yogurt.csv` は logitr 由来の慣習列名 `obsID` のまま維持し、テスト内で `choice_set_id` にリネームして `fit()` に渡す。
- **Forms 変換関数を `forms_to_data()` に統一改名**。rating の `forms_to_conjoint_data()` と choice の `forms_to_data()`（旧 `cbc_forms_to_data()`）を、両サブパッケージで同名の `forms_to_data()` に揃えた（`pcr.forms_to_data()` / `pcc.forms_to_data()`）。あわせて rating の第2引数 `attributes` を `profiles` に改名し（渡すのは属性名だけでなく属性×水準のプロファイル集合のため）、内部ヘルパー `_normalize_attributes` / `_check_attributes` も `_normalize_profiles` / `_check_profiles` に統一。さらに rating の出力 DataFrame の共通列名を英語化して choice と揃えた：`回答者ID` → `respondent_id`、`プロファイルID` → `profile_id`（`fit()` の `respondent_id_col` の既定値も `respondent_id` に追従）。これにより rating（`respondent_id`, `profile_id`, `rating`）と choice（`respondent_id`, `choice_set_id`, `choice`, `alt`）の共通列が英語で揃う。**後方互換なし**（旧名のエイリアスや `DeprecationWarning` は用意しない）。概念が異なる引数・列名（choice の `design` 引数、`profile_id` と `choice_set_id`）は無理に同名化していない。
- **choice の出力列順を rating に統一（`respondent_id` を先頭に）**。`forms_to_data()`（choice）の出力列順を `choice_set_id, respondent_id, alt, choice, …` から `respondent_id, choice_set_id, alt, choice, …` に変更し、rating（`respondent_id, profile_id, rating, …`）と同じ「回答者を先頭にする」列順にそろえた。データの階層構造（回答者 → 設問（選択セット）→ 代替案 → 選択）を左から順に並べる方が直感的なため。`respondent_id` と `choice_set_id` の値そのものは不変。
- **`fit()` の「選択セット数」表示に内訳（回答者数 × 設問数/人）を追加**。`summary()`・HTML 表示で、選択セット総数を `選択セット数（回答者数 × 設問数/人）: 180（30 × 6/人）` のように内訳付きで表示する。各回答者は1つの版に答えるため、版が複数（`n_versions > 1`）でも「回答者数 × 設問数/人 = 選択セット総数」が常に成り立つ。回答者ID列が無い、または回答者ごとに設問数が異なる（不正な）データでは、誤解を招かないよう内訳を付けず総数のみを表示する。
- **`forms_to_data()`（choice）に design の構造チェックを追加**（design とデータの対応ずれによる「静かな誤り」対策）。選択セットごとの代替案数が揃っていない design、`alt_id` が `1..n_alts` を網羅しない design を、先頭セットを盲信せず日本語 `ValueError` で弾く（`version` 引数の指定ミスや設計表の破損を検出）。正常な design では一切警告を出さない。あわせて `design_choice_sets()`・`forms_to_data()` の docstring と `examples/overview_choice.ipynb`・README に、**アンケート作成に使った design と分析時の design を完全に同一にする**こと（水準の順序が違うと同じ seed でも別設計になり結果が静かに誤る）と、**design は作成後すぐ保存し分析時は読み込んで使う**推奨ワークフローを明記。ノートブックは `design_choice_sets()` を2回呼ばず、保存した CSV を読み込んで `forms_to_data()` に渡す構成にした。
- **`examples/overview_choice.ipynb` を題材「スマートフォン」に全面統一**。従来は設計デモ＝架空ブランド・推定＝ヨーグルト・WTP＝架空3水準が混在していたものを、価格・OS・カメラのスマートフォン例に一本化した。第1部＝価格2水準（30名）で設計→診断→必要回答者数→Google Forms 読み込み→符号化→推定→解釈（重要度・WTP・市場シェア・警告）→可視化までの基本フロー、第2部＝価格3水準（40名）で区間別 WTP（価格帯ごとに感応度が変わる非線形）を学ぶ構成。回答データは `examples/make_demo_data.py` が生成する合成データ（`responses_smartphone_*.csv`）で、設計（`design_smartphone_*.csv`）を読み込んで対応ずれが起きないよう作成する。推奨ワークフロー（design は保存→読込で同一に保ち、署名 `design_signature` で確認）を全編で実践。**ヨーグルトデータはノートブックから外し、条件付きロジットの外部検証はテスト（`tests/`）専用**とした（ノートブックには検証の所在を1行だけ補足）。
- **`plot_partworth()`（`rating`・`choice` 共通）の基準水準表示を改善**（歯抜けの解消）。基準水準を「◇（ひし形）マーカー＋『（基準）』ラベル」で**比較の起点**として明示し、0 の位置に点線の基準線を引く。`choice`（ダミーコーディング）は基準水準の部分効用が 0 のため従来は棒が描かれず**歯抜け（データ欠損のような見た目）**になっていたが、◇ マーカーを 0 の位置に置くことで解消。`rating`（効果コーディング）は基準水準が −Σb の**実値を持つので通常の棒**で描き（元々歯抜けにはならない）、その棒の先端に同じ ◇ マーカーを重ねて「ここが基準水準」と分かるようにした（基準の棒は控えめな濃さにして効果の大きさと区別）。両版で表示方針を対称にしつつ、符号化方式の違い（0 か −Σb か）を正しく反映する。属性ごとに色分けして、どの水準がどの属性かが分かるようにした。
- **`plot_wtp()`（`rating`・`choice` 共通）の数値ラベル配置を改善**。数値ラベルを棒の**外側**に置く（正の WTP＝右伸びは右・左揃え、負の WTP＝左伸びは左・右揃え）方針を徹底し、ラベルが縦軸（0 線・属性ラベル）や枠と重ならないよう **x 軸範囲（`xlim`）に余白を確保**する。区間別（グループ化縦棒）でも各棒のラベルを棒の外側に少し離して置き、回転ラベルが切れないよう上下の余白を広げた。
- `openpyxl` の下限を `>=3.0` から `>=3.1.5` に引き上げ。古い 3.1.0 系で学生環境のファイル読み込みが不安定になる問題を避けるため。

### Fixed
- choice の `fit()` の既定列名を `forms_to_data()` の出力列名に整合させた。`choice_set_id_col` の既定 `"選択セットID"` → `"choice_set_id"`、`respondent_id_col` の既定 `"回答者ID"` → `"respondent_id"`。従来は `forms_to_data()` が英語列（`choice_set_id`・`respondent_id`）を出力する一方、`fit()` の既定が日本語のままで不一致だったため、`pcc.fit(df_coded)` を列名引数なしで呼ぶと選択セット列が見つからずエラーになったり、回答者ID列が認識されずクラスタロバスト標準誤差が効かず `independence_assumed` 警告が出たりしていた。これで `forms_to_data() → encode() → fit()` の標準的な流れが列名引数なしでそのまま動く（rating 側の `respondent_id` 既定とも揃う）。エラーメッセージ内の表示ラベル（「選択セットID」など和文）は従来どおり。`tests/test_forms_to_data_rename.py` に列名引数なしでクラスタSEが適用される検証を追加。
- `forms_to_data()`：実 Microsoft Forms の回答ファイル（設問列が長文・改行・全角空白・`\xa0` を含み、性別・利用OS などの属性質問が混在する）で、設問列の自動検出が失敗していた問題を修正。実ファイル（3人分の回答）での回帰テスト `tests/test_choice_forms_real.py` を追加。
- `forms_to_data(forms="google")`：実 Google Forms の回答 CSV（BOMなしUTF-8、設問列名に `【設問N】` などの識別子接頭辞、属性質問の混在）でも正しく変換できることを確認。読み込みは共通ヘルパー経由で `encoding="utf-8-sig"` を使うため BOMなし・BOM付きの両方に対応する。設問列の検出は回答値ベースのため接頭辞の有無に依存しない（設問の順序は列の出現順に従う）。実 Google ファイル（3人分）での検証テストを `tests/test_choice_forms_real.py` に追加。docstring と `examples/overview_choice.ipynb` に、Google の CSV は BOMなしUTF-8 のため Excel で開くと文字化けするが py4conjoint は正しく読める旨を注記。

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
