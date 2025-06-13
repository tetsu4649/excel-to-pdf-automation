#!/usr/bin/env python3
"""
シート選択機能のテスト
"""

from excel_to_pdf import ExcelToWordPDFConverter

def test_sheet_listing():
    """シート一覧の取得をテスト"""
    converter = ExcelToWordPDFConverter()
    
    # サンプルExcelファイルがあるか確認
    import os
    if not os.path.exists("sample_data.xlsx"):
        print("sample_data.xlsxが見つかりません。create_sample_excel.pyを実行してください。")
        return
    
    # シート一覧を取得
    sheets = converter.get_sheet_names("sample_data.xlsx")
    print(f"見つかったシート: {sheets}")
    
    # 各シートの最初の数行を表示
    for sheet_name in sheets:
        print(f"\n--- シート: {sheet_name} ---")
        data = converter.read_excel("sample_data.xlsx", sheet_name)
        for i, row in enumerate(data[:3]):  # 最初の3行のみ表示
            print(f"行{i+1}: {row}")
    
    print("\n✅ シート選択機能のテストが完了しました。")

if __name__ == "__main__":
    test_sheet_listing()