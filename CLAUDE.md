- 全テストがパスすることを各ステップで確認
- git pushはしない（既存のdeny設定通り）

## choice サブパッケージの設計原則
- rating/ と対称的なAPI（fit, encode, summary, importance, wtp, market_share）
- メソッド名・列名（日本語）は rating 版と統一
- 依存は pandas/numpy/statsmodels/matplotlib/openpyxl のみ。追加しない
- MNL推定は scipy.optimize による自前実装（教育目的のため透明性優先）
- エラーメッセージ・警告・docstringはすべて日本語

