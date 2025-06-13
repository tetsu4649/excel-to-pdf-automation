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


def select_columns_interactive():
    """対話形式で列を選択する"""
    print("\n列の選択:")
    print("-" * 40)
    print("1. B列のみ（デフォルト）")
    print("2. 特定の列を選択")
    print("3. すべての列")
    print("-" * 40)
    
    while True:
        try:
            choice = input("選択してください (1-3) [デフォルト: 1]: ").strip()
            
            if choice == "" or choice == "1":
                return ["B"]
            elif choice == "2":
                columns = input("列を入力してください（カンマ区切り、例: A,B,D）: ").strip()
                if columns:
                    # 列文字をリストに変換し、大文字に統一
                    return [col.strip().upper() for col in columns.split(",") if col.strip()]
                else:
                    print("❌ 列を入力してください。")
            elif choice == "3":
                return ["ALL"]
            else:
                print("❌ 1から3の範囲で入力してください。")
        except KeyboardInterrupt:
            print("\n\n処理を中断しました。")
            sys.exit(0)


def select_output_mode():
    """出力モードを選択する"""
    print("\nPDF出力モード:")
    print("-" * 40)
    print("1. 通常（セルの色や罫線を含む）")
    print("2. テキストのみ（シンプルな表示）")
    print("-" * 40)
    
    while True:
        try:
            choice = input("選択してください (1-2) [デフォルト: 2]: ").strip()
            
            if choice == "" or choice == "2":
                return True  # text_only = True
            elif choice == "1":
                return False  # text_only = False
            else:
                print("❌ 1から2の範囲で入力してください。")
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
        print("例: python main.py sample.xlsm")  # .xlsmもサポート
        print("\nヒント: ターミナルにExcelファイルをドラッグ&ドロップできます!")
        print("\nサポートされている形式: .xlsx, .xls, .xlsm")
        sys.exit(1)
    
    # ファイルパスを取得（引用符を削除）
    excel_file = sys.argv[1].strip('"').strip("'")
    output_dir = sys.argv[2].strip('"').strip("'") if len(sys.argv) > 2 else None
    
    # ファイルの存在確認
    if not Path(excel_file).exists():
        print(f"エラー: Excelファイルが見つかりません: {excel_file}")
        sys.exit(1)
    
    # ファイル拡張子の確認
    file_ext = Path(excel_file).suffix.lower()
    if file_ext not in ['.xlsx', '.xls', '.xlsm']:
        print(f"エラー: サポートされていないファイル形式です: {file_ext}")
        print("サポートされている形式: .xlsx, .xls, .xlsm")
        sys.exit(1)
    
    try:
        # 出力モードを選択
        text_only = select_output_mode()
        
        # コンバーターを初期化（まず列選択なしで）
        converter = ExcelToWordPDFConverter(text_only=text_only)
        
        # シート一覧を取得
        sheets = converter.get_sheet_names(excel_file)
        
        if not sheets:
            print("❌ Excelファイルからシートを読み取れませんでした。")
            sys.exit(1)
        
        # 対話形式でシートを選択
        selected_sheet = select_sheet_interactive(sheets)
        
        # シートが選択された場合は列選択も行う
        if selected_sheet is not None:
            selected_columns = select_columns_interactive()
            # コンバーターを再初期化（列選択を含む）
            converter = ExcelToWordPDFConverter(text_only=text_only, selected_columns=selected_columns)
        
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
            if converter.selected_columns != ["ALL"]:
                print(f"選択された列: {', '.join(converter.selected_columns)}")
            word_path, pdf_path = converter.convert(excel_file, output_dir, selected_sheet)
            
            print("\n✅ 変換が完了しました!")
            print(f"📄 Word: {word_path}")
            print(f"📑 PDF: {pdf_path}")
        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()