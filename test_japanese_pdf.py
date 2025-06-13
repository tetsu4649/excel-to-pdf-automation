#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""日本語PDF出力のテストスクリプト"""

import sys
from pathlib import Path
from excel_to_pdf import ExcelToWordPDFConverter
import openpyxl

def create_japanese_test_excel():
    """日本語テストデータを含むExcelファイルを作成"""
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # 日本語データを設定
    test_data = [
        ["商品名", "価格", "在庫数", "備考"],
        ["りんご（青森産）", "¥300", "150個", "新鮮な青森産のりんごです"],
        ["みかん（愛媛産）", "¥250", "200個", "甘くてジューシー"],
        ["ぶどう（山梨産）", "¥800", "50房", "高級品種のシャインマスカット"],
        ["日本酒「大吟醸」", "¥3,500", "30本", "香り高い純米大吟醸"],
        ["緑茶（静岡産）", "¥1,200", "80袋", "深蒸し茶100g入り"],
    ]
    
    # データをワークシートに書き込む
    for row_idx, row_data in enumerate(test_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # ファイルを保存
    test_file = "test_japanese.xlsx"
    wb.save(test_file)
    print(f"テストファイルを作成しました: {test_file}")
    
    return test_file

def main():
    """メイン処理"""
    print("日本語PDF出力テストを開始します...")
    
    # テスト用Excelファイルを作成
    test_file = create_japanese_test_excel()
    
    try:
        # コンバーターを初期化
        converter = ExcelToWordPDFConverter()
        
        # 変換を実行
        word_path, pdf_path = converter.convert(test_file)
        
        print("\n変換完了!")
        print(f"Word: {word_path}")
        print(f"PDF: {pdf_path}")
        print("\nPDFファイルを開いて日本語が正しく表示されているか確認してください。")
        
    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()