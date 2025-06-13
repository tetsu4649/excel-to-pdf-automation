#!/usr/bin/env python3
"""
Excel to PDF 簡単変換ツール - よりシンプルなインターフェース
"""

import sys
import os
from pathlib import Path
from excel_to_pdf import ExcelToWordPDFConverter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading


class SimpleExcelToPDFGUI:
    """シンプルなGUIインターフェース"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel to PDF 変換ツール")
        self.root.geometry("650x650")
        self.root.resizable(False, False)
        
        # ファイルパス保存用
        self.excel_file = None
        self.output_dir = None
        self.converter = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """UIをセットアップ"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, text="Excel to PDF 変換ツール", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # ファイル選択セクション
        file_frame = ttk.LabelFrame(main_frame, text="1. Excelファイルを選択", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="ファイルが選択されていません (.xlsx, .xls, .xlsm対応)", foreground="gray")
        self.file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        ttk.Button(file_frame, text="ファイルを選択", command=self.select_file).grid(row=0, column=1, sticky=tk.E)
        
        # ドラッグ&ドロップエリア
        drop_frame = ttk.Frame(file_frame, relief=tk.RIDGE, borderwidth=2)
        drop_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        drop_label = ttk.Label(drop_frame, text="またはここにファイルをドラッグ&ドロップ", 
                              foreground="gray", padding="20")
        drop_label.grid(row=0, column=0)
        
        # 出力先選択セクション
        output_frame = ttk.LabelFrame(main_frame, text="2. 出力先を選択（オプション）", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.output_label = ttk.Label(output_frame, text="元のファイルと同じフォルダに保存", foreground="gray")
        self.output_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        ttk.Button(output_frame, text="フォルダを選択", command=self.select_output_dir).grid(row=0, column=1, sticky=tk.E)
        
        # PDF出力設定セクション
        pdf_frame = ttk.LabelFrame(main_frame, text="3. PDF出力設定", padding="10")
        pdf_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.text_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(pdf_frame, text="テキストのみ（色やセル装飾なし）", 
                       variable=self.text_only_var).grid(row=0, column=0, sticky=tk.W)
        
        # シート選択セクション
        sheet_frame = ttk.LabelFrame(main_frame, text="4. 変換するシートを選択", padding="10")
        sheet_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.sheet_var = tk.StringVar(value="all")
        ttk.Radiobutton(sheet_frame, text="すべてのシートを変換", variable=self.sheet_var, 
                       value="all", command=self.on_sheet_option_change).grid(row=0, column=0, sticky=tk.W)
        
        ttk.Radiobutton(sheet_frame, text="特定のシートを選択:", variable=self.sheet_var, 
                       value="selected", command=self.on_sheet_option_change).grid(row=1, column=0, sticky=tk.W)
        
        self.sheet_combo = ttk.Combobox(sheet_frame, state="disabled", width=30)
        self.sheet_combo.grid(row=1, column=1, padx=(10, 0))
        
        # 列選択フレーム
        column_frame = ttk.Frame(sheet_frame)
        column_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(column_frame, text="列の選択:").grid(row=0, column=0, sticky=tk.W)
        
        self.column_var = tk.StringVar(value="B")
        ttk.Radiobutton(column_frame, text="B列のみ（デフォルト）", variable=self.column_var,
                       value="B").grid(row=0, column=1, padx=(10, 0))
        ttk.Radiobutton(column_frame, text="すべての列", variable=self.column_var,
                       value="ALL").grid(row=0, column=2, padx=(10, 0))
        ttk.Radiobutton(column_frame, text="カスタム:", variable=self.column_var,
                       value="custom").grid(row=0, column=3, padx=(10, 0))
        
        self.custom_columns_entry = ttk.Entry(column_frame, width=15)
        self.custom_columns_entry.grid(row=0, column=4, padx=(5, 0))
        self.custom_columns_entry.insert(0, "A,B,C")
        
        # 初期状態では列選択は無効
        self.column_frame_widgets = column_frame.winfo_children()
        for widget in self.column_frame_widgets:
            widget.config(state="disabled")
        
        # 変換ボタン
        self.convert_button = ttk.Button(main_frame, text="変換開始", command=self.convert, 
                                        state="disabled", style="Accent.TButton")
        self.convert_button.grid(row=5, column=0, columnspan=3, pady=(20, 10))
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.progress.grid_remove()  # 初期は非表示
        
        # ステータスラベル
        self.status_label = ttk.Label(main_frame, text="", foreground="green")
        self.status_label.grid(row=7, column=0, columnspan=3)
        
    def on_sheet_option_change(self, *args):
        """シート選択オプションが変更されたときの処理"""
        if self.sheet_var.get() == "selected":
            self.sheet_combo.config(state="readonly")
            # 列選択も有効化
            for widget in self.column_frame_widgets:
                widget.config(state="normal")
        else:
            self.sheet_combo.config(state="disabled")
            # 列選択も無効化
            for widget in self.column_frame_widgets:
                widget.config(state="disabled")
            
    def select_file(self):
        """ファイル選択ダイアログを表示"""
        filename = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
        )
        
        if filename:
            self.excel_file = filename
            self.file_label.config(text=Path(filename).name, foreground="black")
            self.convert_button.config(state="normal")
            self.status_label.config(text="")
            
            # シート一覧を取得
            self.update_sheet_list()
            
    def update_sheet_list(self):
        """選択されたExcelファイルからシート一覧を取得"""
        if self.excel_file:
            try:
                temp_converter = ExcelToWordPDFConverter()
                sheets = temp_converter.get_sheet_names(self.excel_file)
                self.sheet_combo['values'] = sheets
                if sheets:
                    self.sheet_combo.current(0)
            except Exception as e:
                messagebox.showerror("エラー", f"シート一覧の取得に失敗しました: {e}")
                
    def select_output_dir(self):
        """出力先フォルダ選択ダイアログを表示"""
        dirname = filedialog.askdirectory(title="出力先フォルダを選択")
        
        if dirname:
            self.output_dir = dirname
            self.output_label.config(text=dirname, foreground="black")
    
    def get_selected_columns(self):
        """選択された列を取得"""
        if self.sheet_var.get() == "all":
            return ["ALL"]  # すべてのシートの場合は全列
        
        column_choice = self.column_var.get()
        if column_choice == "B":
            return ["B"]
        elif column_choice == "ALL":
            return ["ALL"]
        elif column_choice == "custom":
            custom = self.custom_columns_entry.get().strip()
            if custom:
                return [col.strip().upper() for col in custom.split(",") if col.strip()]
            else:
                return ["B"]  # デフォルト
        return ["B"]
            
    def convert(self):
        """変換処理を実行"""
        if not self.excel_file:
            messagebox.showerror("エラー", "Excelファイルを選択してください")
            return
            
        # UIを無効化
        self.convert_button.config(state="disabled")
        self.progress.grid()
        self.progress.start()
        self.status_label.config(text="変換中...", foreground="blue")
        
        # 別スレッドで変換処理を実行
        thread = threading.Thread(target=self.do_convert)
        thread.start()
        
    def do_convert(self):
        """実際の変換処理（別スレッド）"""
        try:
            text_only = self.text_only_var.get()
            selected_columns = self.get_selected_columns()
            
            # コンバーターを作成
            self.converter = ExcelToWordPDFConverter(text_only=text_only, selected_columns=selected_columns)
            
            if self.sheet_var.get() == "all":
                # すべてのシートを変換
                sheets = self.converter.get_sheet_names(self.excel_file)
                results = []
                for i, sheet in enumerate(sheets):
                    self.root.after(0, self.update_status, f"変換中... ({i+1}/{len(sheets)})")
                    word_path, pdf_path = self.converter.convert(self.excel_file, self.output_dir, sheet)
                    results.append((sheet, pdf_path))
                
                # 完了メッセージ
                self.root.after(0, self.conversion_complete, results, True)
            else:
                # 選択されたシートのみ変換
                selected_sheet = self.sheet_combo.get()
                word_path, pdf_path = self.converter.convert(self.excel_file, self.output_dir, selected_sheet)
                self.root.after(0, self.conversion_complete, [(selected_sheet, pdf_path)], False)
                
        except Exception as e:
            self.root.after(0, self.conversion_error, str(e))
            
    def update_status(self, message):
        """ステータスメッセージを更新"""
        self.status_label.config(text=message, foreground="blue")
        
    def conversion_complete(self, results, is_multiple):
        """変換完了時の処理"""
        self.progress.stop()
        self.progress.grid_remove()
        self.convert_button.config(state="normal")
        
        if is_multiple:
            message = f"✅ {len(results)}個のシートの変換が完了しました！\n\n"
            for sheet, pdf_path in results:
                message += f"• {sheet} → {Path(pdf_path).name}\n"
        else:
            sheet, pdf_path = results[0]
            message = f"✅ 変換が完了しました！\n\nPDF: {Path(pdf_path).name}"
            
        self.status_label.config(text="変換完了！", foreground="green")
        messagebox.showinfo("完了", message)
        
    def conversion_error(self, error_message):
        """変換エラー時の処理"""
        self.progress.stop()
        self.progress.grid_remove()
        self.convert_button.config(state="normal")
        self.status_label.config(text="エラーが発生しました", foreground="red")
        messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{error_message}")
        
    def run(self):
        """GUIを起動"""
        self.root.mainloop()


def simple_cli_mode():
    """簡単なCLIモード（ファイルパスだけで変換）"""
    print("Excel to PDF 簡単変換ツール")
    print("-" * 40)
    
    # 引数がある場合はそれを使用
    if len(sys.argv) > 1:
        excel_file = sys.argv[1].strip('"').strip("'")
    else:
        # 対話形式でファイルパスを取得
        print("Excelファイルのパスを入力してください")
        print("（ヒント: ファイルをドラッグ&ドロップできます）")
        print("サポート形式: .xlsx, .xls, .xlsm")
        excel_file = input("ファイルパス: ").strip('"').strip("'")
    
    # ファイルの存在確認
    if not Path(excel_file).exists():
        print(f"❌ エラー: ファイルが見つかりません: {excel_file}")
        return
    
    # ファイル拡張子の確認
    file_ext = Path(excel_file).suffix.lower()
    if file_ext not in ['.xlsx', '.xls', '.xlsm']:
        print(f"❌ エラー: サポートされていないファイル形式です: {file_ext}")
        print("サポートされている形式: .xlsx, .xls, .xlsm")
        return
    
    try:
        # PDF出力モードを選択
        print("\nPDF出力モード:")
        print("1. テキストのみ（シンプル・デフォルト）")
        print("2. 通常（セルの色や罫線を含む）")
        mode_choice = input("選択 (1-2) [デフォルト: 1]: ").strip()
        text_only = mode_choice != "2"
        
        converter = ExcelToWordPDFConverter(text_only=text_only)
        
        # シート一覧を取得
        sheets = converter.get_sheet_names(excel_file)
        
        if len(sheets) == 1:
            # シートが1つだけの場合は列選択を確認
            print(f"\nシート: {sheets[0]}")
            print("\n列の選択:")
            print("1. B列のみ（デフォルト）")
            print("2. すべての列")
            print("3. カスタム")
            col_choice = input("選択 (1-3) [デフォルト: 1]: ").strip()
            
            if col_choice == "2":
                selected_columns = ["ALL"]
            elif col_choice == "3":
                custom = input("列を入力（カンマ区切り、例: A,B,D）: ").strip()
                selected_columns = [col.strip().upper() for col in custom.split(",") if col.strip()] or ["B"]
            else:
                selected_columns = ["B"]
            
            # コンバーターを再作成
            converter = ExcelToWordPDFConverter(text_only=text_only, selected_columns=selected_columns)
            
            print(f"\n🔄 変換中: {sheets[0]}")
            word_path, pdf_path = converter.convert(excel_file, None, sheets[0])
            print(f"\n✅ 変換完了!")
            print(f"📑 PDF: {pdf_path}")
        else:
            # 複数シートの場合は選択肢を表示
            print(f"\n{len(sheets)}個のシートが見つかりました。")
            print("すべて変換しますか？ (Y/n): ", end="")
            choice = input().strip().lower()
            
            if choice == '' or choice == 'y':
                # すべて変換
                print("\n🔄 すべてのシートを変換します...")
                for sheet_name in sheets:
                    print(f"  • {sheet_name} を変換中...")
                    word_path, pdf_path = converter.convert(excel_file, None, sheet_name)
                print(f"\n✅ {len(sheets)}個のシートの変換が完了しました!")
            else:
                # 個別選択
                print("\n変換するシートを番号で選択してください:")
                for i, sheet in enumerate(sheets, 1):
                    print(f"  {i}. {sheet}")
                
                while True:
                    try:
                        num = int(input("番号: "))
                        if 1 <= num <= len(sheets):
                            selected_sheet = sheets[num - 1]
                            
                            # 列選択
                            print("\n列の選択:")
                            print("1. B列のみ（デフォルト）")
                            print("2. すべての列")
                            print("3. カスタム")
                            col_choice = input("選択 (1-3) [デフォルト: 1]: ").strip()
                            
                            if col_choice == "2":
                                selected_columns = ["ALL"]
                            elif col_choice == "3":
                                custom = input("列を入力（カンマ区切り、例: A,B,D）: ").strip()
                                selected_columns = [col.strip().upper() for col in custom.split(",") if col.strip()] or ["B"]
                            else:
                                selected_columns = ["B"]
                            
                            # コンバーターを再作成
                            converter = ExcelToWordPDFConverter(text_only=text_only, selected_columns=selected_columns)
                            
                            print(f"\n🔄 変換中: {selected_sheet}")
                            word_path, pdf_path = converter.convert(excel_file, None, selected_sheet)
                            print(f"\n✅ 変換完了!")
                            print(f"📑 PDF: {pdf_path}")
                            break
                        else:
                            print("❌ 無効な番号です")
                    except ValueError:
                        print("❌ 数値を入力してください")
                        
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")


def main():
    """メインエントリーポイント"""
    # GUIモードかCLIモードかを判定
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        # 明示的にGUIモードを指定
        app = SimpleExcelToPDFGUI()
        app.run()
    elif len(sys.argv) > 1 and sys.argv[1] == "--help":
        print("Excel to PDF 簡単変換ツール")
        print("\n使い方:")
        print("  python easy_converter.py              # 対話形式で変換")
        print("  python easy_converter.py file.xlsx    # 指定ファイルを変換") 
        print("  python easy_converter.py file.xlsm    # .xlsmファイルも対応")
        print("  python easy_converter.py --gui        # GUIモードで起動")
        print("  python easy_converter.py --help       # このヘルプを表示")
        print("\n機能:")
        print("  • .xlsx, .xls, .xlsm ファイルに対応")
        print("  • テキストのみのPDF出力（色やセル装飾なし）")
        print("  • 列選択機能（デフォルトはB列）")
    else:
        # デフォルトはCLIモード
        try:
            # tkinterが利用可能か確認
            import tkinter
            # 環境変数でGUIを無効化できる
            if os.environ.get('NO_GUI') != '1':
                # 引数がなければGUIを起動
                if len(sys.argv) == 1:
                    app = SimpleExcelToPDFGUI()
                    app.run()
                else:
                    # 引数があればCLIモード
                    simple_cli_mode()
            else:
                simple_cli_mode()
        except ImportError:
            # tkinterが利用不可の場合はCLIモード
            simple_cli_mode()


if __name__ == "__main__":
    main()