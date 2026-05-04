#!/bin/bash
# ==============================================================================
# WBS整合性チェック 実行ラッパースクリプト
# 
# 使い方: ./scripts/check_wbs.sh [Excelファイルへのパス]
# 例: ./scripts/check_wbs.sh projects/test_project/hoge_wbs_evm.xlsx
# ==============================================================================

# スクリプトの場所を基準にプロジェクトルートを特定
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# プロジェクトルートに移動
cd "$PROJECT_ROOT"

# 仮想環境のチェックと有効化
if [ ! -d ".venv" ]; then
    echo "ERROR: .venv ディレクトリが見つかりません。セットアップを確認してください。"
    exit 1
fi

source .venv/bin/activate

# PYTHONPATHの設定（自作モジュールのインポート用）
export PYTHONPATH=$PYTHONPATH:.

# Pythonスクリプトの実行
python3 scripts/check_wbs_integrity.py "$@"
