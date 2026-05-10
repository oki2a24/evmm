import os
import subprocess
import shutil
import pytest

def test_cli_initializes_context_automatically(tmp_path):
    """
    CLI からスクリプトを直接実行した際、docs/context.md が自動生成されることを検証する。
    """
    # 1. テスト環境の準備
    project_dir = tmp_path / "cli_test_pj"
    project_dir.mkdir()
    
    # テンプレート Excel をコピー (中身は空でも良いが、ファイルが存在する必要がある)
    excel_path = project_dir / "test_wbs.xlsx"
    shutil.copy("tests/data/wbs_template_for_testing.xlsx", excel_path)
    
    context_md = project_dir / "docs" / "context.md"
    
    # 実行前は存在しない
    assert not context_md.exists()
    
    # 2. CLI 経由で実行
    # PYTHONPATH を設定して、scripts モジュールをインポート可能にする
    env = os.environ.copy()
    env["PYTHONPATH"] = f".:{env.get('PYTHONPATH', '')}"
    
    result = subprocess.run(
        ["./.venv/bin/python3", "scripts/analyze_evm.py", str(excel_path)],
        capture_output=True,
        text=True,
        env=env
    )
    
    # 実行が成功していることを確認
    assert result.returncode == 0
    
    # 3. 検証 (RED になるはず: 現状の CLI は ensure_project_context を呼んでいないため)
    assert context_md.exists(), "CLI 実行後に context.md が生成されていません"
    assert "# Project Context" in context_md.read_text()
