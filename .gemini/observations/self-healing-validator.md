# self-healing-validator 知見 (L4: Project)

## 仮想環境の強制 (Environment)
- **RED**: `ModuleNotFoundError: 'openpyxl'` などのエラーが発生した。
- **GREEN**: プロジェクト内の Python スクリプトを実行する際は、必ず事前に `source .venv/bin/activate` を実行し、仮想環境を有効化しなければならない。

## 多数決ロジックの限界 (Algorithm)
- **RED**: 全ての行が破損している場合、多数決（Majority Vote）では正解を特定できず修復に失敗する。
- **GREEN**: 多数決が成立しない（正解パターン数 < 2）場合は、直ちに修復を諦め、ユーザーに「1行手動で直してオートフィルする」か「マスターテンプレートから復元する（将来機能）」ことを提案せよ。
