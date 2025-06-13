# Excel to PDF 自動化ツール

ExcelファイルをWord経由でPDF形式に変換するPythonベースの自動化ツールです。

## 説明

このプロジェクトは、Excelスプレッドシートの内容をWordドキュメントにコピーし、最終的にPDF形式で出力する自動化ソリューションを提供します。ExcelデータからPDFレポートを生成するプロセスを効率化するために設計されています。

## 機能

- Excelファイル（.xlsx、.xls）の内容を読み取り
- WordドキュメントへのExcelデータの転送（テーブル形式）
- Word経由でPDF形式に変換
- バッチ処理機能
- コマンドラインインターフェース

## インストール

```bash
# リポジトリをクローン
git clone https://github.com/tetsu4649/excel-to-pdf-automation.git

# プロジェクトディレクトリに移動
cd excel-to-pdf-automation

# 必要な依存関係をインストール
pip install -r requirements.txt
```

## 使い方

### 基本的な使い方

```bash
# Excelファイルを指定して変換
python main.py sample.xlsx

# 出力ディレクトリを指定して変換
python main.py sample.xlsx ./output
```

### 高度な使い方

```bash
# excel_to_pdf.pyを直接使用
python excel_to_pdf.py input.xlsx -o ./output
```

### サンプルファイルでテスト

```bash
# サンプルExcelファイルを作成
python create_sample_excel.py

# テストスクリプトを実行
python test_conversion.py
```

## プロジェクト構造

```
excel-to-pdf-automation/
├── main.py                 # メインエントリーポイント
├── excel_to_pdf.py        # Excel→Word→PDF変換の核となるモジュール
├── create_sample_excel.py  # サンプルExcelファイル作成スクリプト
├── test_conversion.py      # 変換機能のテストスクリプト
├── requirements.txt        # Python依存関係
└── README.md              # このファイル
```

## 必要条件

- Python 3.7以上
- 必要なPythonパッケージ:
  - openpyxl==3.1.2 (Excel操作用)
  - python-docx==1.1.0 (Word文書作成用)
  - reportlab==4.1.0 (PDF生成用)
  - Pillow==10.3.0 (画像処理用)

## 動作の仕組み

1. **Excel読み取り**: `openpyxl`を使用してExcelファイルからデータを読み取ります
2. **Word文書作成**: `python-docx`を使用してWordドキュメントを作成し、Excelデータをテーブル形式で挿入します
3. **PDF変換**: `reportlab`を使用してPDFファイルを生成します

## 注意事項

- 大きなExcelファイルの処理には時間がかかる場合があります
- 日本語を含むファイルも正しく処理されます
- 複雑な書式設定やグラフは現在サポートされていません

## コントリビューション

貢献を歓迎します！お気軽にプルリクエストを送信してください。

## ライセンス

このプロジェクトは、リポジトリオーナーが指定する条件に基づいてライセンスされています。