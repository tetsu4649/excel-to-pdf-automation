#!/usr/bin/env python3
"""
Excel to PDF変換のテストスクリプト
"""

import os
import sys
from pathlib import Path

# プロジェクトのルートディレクトリをPythonパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    # 必要なモジュールをインポート
    from create_sample_excel import create_sample_excel
    from excel_to_pdf import ExcelToWordPDFConverter
    
    print("テストを開始します...")
    print("-" * 50)
    
    # Step 1: サンプルExcelファイルを作成
    print("\n1. サンプルExcelファイルを作成中...")
    excel_file = create_sample_excel()
    print(f"   ✅ 作成完了: {excel_file}")
    
    # Step 2: Excel to PDF変換を実行
    print("\n2. Excel → Word → PDF変換を実行中...")
    converter = ExcelToWordPDFConverter()
    
    # 出力ディレクトリを作成
    output_dir = project_root / "output"
    output_dir.mkdir(exist_ok=True)
    
    # 変換を実行
    word_path, pdf_path = converter.convert(excel_file, str(output_dir))
    
    # Step 3: 結果を確認
    print("\n3. 変換結果:")
    print(f"   📊 入力Excel: {excel_file}")
    print(f"   📄 出力Word: {word_path}")
    print(f"   📑 出力PDF: {pdf_path}")
    
    # ファイルの存在確認
    if Path(word_path).exists() and Path(pdf_path).exists():
        print("\n✅ すべてのファイルが正常に作成されました！")
        print(f"\n出力ファイルは以下のディレクトリに保存されています:")
        print(f"   {output_dir}")
    else:
        print("\n❌ ファイルの作成に失敗しました。")
        sys.exit(1)
    
except ImportError as e:
    print(f"\n❌ モジュールのインポートエラー: {e}")
    print("\n以下のコマンドで必要なパッケージをインストールしてください:")
    print("pip install -r requirements.txt")
    sys.exit(1)
    
except Exception as e:
    print(f"\n❌ エラーが発生しました: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)