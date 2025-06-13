#!/usr/bin/env python3
"""
Excel to PDF 自動化ツール - メインエントリーポイント
"""

import sys
import os
from pathlib import Path

# excel_to_pdfモジュールをインポート
from excel_to_pdf import ExcelToWordPDFConverter


def main():
    """メインエントリーポイント"""
    print("Excel to PDF 自動化ツール")
    print("-" * 40)
    
    # コマンドライン引数をチェック
    if len(sys.argv) < 2:
        print("使い方: python main.py <Excelファイル> [出力ディレクトリ]")
        print("例: python main.py sample.xlsx")
        print("例: python main.py sample.xlsx ./output")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    # ファイルの存在確認
    if not Path(excel_file).exists():
        print(f"エラー: Excelファイルが見つかりません: {excel_file}")
        sys.exit(1)
    
    try:
        # コンバーターを初期化
        converter = ExcelToWordPDFConverter()
        
        # 変換を実行
        print(f"\n処理中: {excel_file}")
        word_path, pdf_path = converter.convert(excel_file, output_dir)
        
        print("\n✅ 変換が完了しました!")
        print(f"📄 Word: {word_path}")
        print(f"📑 PDF: {pdf_path}")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 
