#!/usr/bin/env python3
"""
Excel to PDF 自動化ツール - メインエントリーポイント
"""

import sys
import os
from pathlib import Path

# excel_to_pdfモジュールをインポート
from excel_to_pdf import ExcelToWordPDFConverter


def select_sheet_interactive(sheets):
    """対話形式でシートを選択する"""
    print("\n利用可能なシート:")
    print("-" * 40)
    for i, sheet in enumerate(sheets, 1):
        print(f"{i}. {sheet}")
    print(f"{len(sheets) + 1}. すべてのシートを変換")
    print("-" * 40)
    
    while True:
        try:
            choice = input(f"シートを選択してください (1-{len(sheets) + 1}): ")
            choice_num = int(choice)
            
            if 1 <= choice_num <= len(sheets):
                return sheets[choice_num - 1]
            elif choice_num == len(sheets) + 1:
                return None  # すべてのシートを変換
            else:
                print(f"❌ 1から{len(sheets) + 1}の範囲で入力してください。")
        except ValueError:
            print("❌ 数値を入力してください。")
        except KeyboardInterrupt:
            print("\n\n処理を中断しました。")
            sys.exit(0)


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
        
        # シート一覧を取得
        sheets = converter.get_sheet_names(excel_file)
        
        if not sheets:
            print("❌ Excelファイルからシートを読み取れませんでした。")
            sys.exit(1)
        
        # 対話形式でシートを選択
        selected_sheet = select_sheet_interactive(sheets)
        
        if selected_sheet is None:
            # すべてのシートを変換
            print("\n🔄 すべてのシートを変換します...")
            for sheet_name in sheets:
                print(f"\n処理中: {excel_file} - シート: {sheet_name}")
                word_path, pdf_path = converter.convert(excel_file, output_dir, sheet_name)
                print(f"✅ 完了: {sheet_name}")
                print(f"  📄 Word: {word_path}")
                print(f"  📑 PDF: {pdf_path}")
        else:
            # 選択されたシートのみを変換
            print(f"\n処理中: {excel_file} - シート: {selected_sheet}")
            word_path, pdf_path = converter.convert(excel_file, output_dir, selected_sheet)
            
            print("\n✅ 変換が完了しました!")
            print(f"📄 Word: {word_path}")
            print(f"📑 PDF: {pdf_path}")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 