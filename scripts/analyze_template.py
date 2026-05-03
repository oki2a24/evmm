import pandas as pd
import json

"""
WBS/EVM テンプレート分析スクリプト

このスクリプトは、ユーザーから提供されたエクセルファイル（wbs_ebm_format.xlsx）の
構造（シート名、列名、サンプルデータ）を解析し、AIが理解できる形式で出力します。
"""

def analyze_excel(file_path):
    """
    指定されたエクセルファイルを解析して、その構造をJSON形式で出力する。
    
    Args:
        file_path (str): 解析対象のエクセルファイルのパス
    """
    try:
        # エクセルファイルを読み込み（全シートを対象とする）
        xl = pd.ExcelFile(file_path)
        analysis_result = {}

        for sheet_name in xl.sheet_names:
            # 各シートをデータフレームとして読み込む
            # header=None にして、結合セルや複雑なヘッダー構造をそのまま取得することを試みる
            df = xl.parse(sheet_name, header=None)
            
            # シートの構造情報を抽出
            analysis_result[sheet_name] = {
                "data": df.head(20).values.tolist(),   # 冒頭20行のデータ（ヘッダー構造把握のため多めに取得）
                "shape": df.shape                      # シートの行列サイズ
            }
        
        # 解析結果を、日本語を維持したまま整形して表示
        print(json.dumps(analysis_result, ensure_ascii=False, indent=2))

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    # ユーザーから提供されたファイルを指定して解析を実行
    analyze_excel("wbs_ebm_format.xlsx")
