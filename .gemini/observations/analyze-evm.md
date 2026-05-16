# EVM分析エンジンの知見 (L4: Project)

## テスト時の物理的制約
- **規律**: `EVMAnalyst` の初期化時に `WBSConfigManager` がエクセルの物理構造（ZipFile）を読み込むため、テスト環境（`pytest`等）では `tmp_path` を用いて物理的な `.xlsx` ファイルを生成しなければならない。
- **背景**: 単なる DataFrame のモックでは `openpyxl.load_workbook` が失敗し、`BadZipFile` エラーが発生する。

## 柔軟なマッピングへの対応
- **規律**: 分析実行前に必ず `WBSConfigManager.load_or_infer()` を呼び出し、現在のエクセル構造とマッピング情報の整合性を確認すること。
- **背景**: ユーザーがカラム名や位置を変更した場合、マッピングを更新しなければ誤った計算（CPI/SPI等）を算出するリスクがある。
