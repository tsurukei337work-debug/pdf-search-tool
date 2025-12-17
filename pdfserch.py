import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import configparser
import os
import re
import time
import threading
import webbrowser
import platform
from typing import List, Dict, Optional, Any
from datetime import datetime

# 外部ライブラリ
import pandas as pd
from pypdf import PdfReader

# =============================================================================
# 2. ini設定の読み書き関数 / クラス
# =============================================================================
class ConfigManager:
    """設定ファイル(config.ini)の管理を行うクラス"""
    
    DEFAULT_CONFIG = {
        'Settings': {
            'target_folder': '',
            'search_keyword': '',
            'file_extension': '.pdf',
            'context_length': '30'  # 検索ヒット時の文字数
        }
    }
    
    def __init__(self, filename: str = 'config.ini'):
        self.filename = filename
        self.config = configparser.ConfigParser()

    def load_config(self) -> Dict[str, Any]:
        """設定を読み込む。ファイルがない場合はデフォルトを返す"""
        if not os.path.exists(self.filename):
            return self.DEFAULT_CONFIG['Settings'].copy()
        
        try:
            self.config.read(self.filename, encoding='utf-8')
            if 'Settings' in self.config:
                return dict(self.config['Settings'])
            else:
                return self.DEFAULT_CONFIG['Settings'].copy()
        except Exception as e:
            print(f"Config load error: {e}")
            return self.DEFAULT_CONFIG['Settings'].copy()

    def save_config(self, settings: Dict[str, str]) -> None:
        """設定を保存する"""
        self.config['Settings'] = settings
        try:
            with open(self.filename, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except Exception as e:
            raise IOError(f"設定ファイルの保存に失敗しました: {e}")

# =============================================================================
# 3. メイン処理関数 main_processor (ロジッククラス)
# =============================================================================
class SearchLogic:
    """検索処理の実装"""
    
    def __init__(self):
        self.is_running = False
        self.cancel_flag = False

    def get_pdf_files(self, folder_path: str) -> List[str]:
        """指定フォルダ以下のPDFファイルを再帰的に取得"""
        pdf_files = []
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        return pdf_files

    def search_in_pdf(self, file_path: str, keyword: str, context_len: int = 30) -> List[Dict]:
        """単一PDF内の検索実行"""
        results = []
        try:
            reader = PdfReader(file_path)
            # 暗号化されている場合の簡易チェック
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except:
                    return [{"error": "Encrypted/Password Protected"}]

            num_pages = len(reader.pages)
            
            for i, page in enumerate(reader.pages):
                try:
                    text = page.extract_text()
                    if text:
                        # 改行を削除して検索しやすくする
                        clean_text = text.replace('\n', '')
                        
                        # 大文字小文字を区別しない検索
                        matches = [m.start() for m in re.finditer(re.escape(keyword), clean_text, re.IGNORECASE)]
                        
                        for match_idx in matches:
                            start = max(0, match_idx - 10)
                            end = min(len(clean_text), match_idx + len(keyword) + context_len)
                            snippet = clean_text[start:end]
                            
                            results.append({
                                "file_name": os.path.basename(file_path),
                                "file_path": file_path,
                                "page": i + 1,
                                "context": "..." + snippet + "...",
                                "error": None
                            })
                            # 1ページに複数ヒットしても、とりあえず1つ見つかればそのページはヒットとする場合はここでbreak
                except Exception as e:
                    # ページ読み込みエラーはログに残すが処理は継続
                    pass
                    
        except Exception as e:
            return [{"error": str(e)}]
            
        return results

# =============================================================================
# 4. GUIクラス 作成とイベント処理
# =============================================================================
class PDFSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF文字列検索ツール")
        self.geometry("900x650")
        
        # クラスの初期化
        self.config_manager = ConfigManager()
        self.logic = SearchLogic()
        self.results_df = pd.DataFrame() # 結果保持用
        
        # GUI変数の初期化
        self.var_folder_path = tk.StringVar()
        self.var_keyword = tk.StringVar()
        self.var_status = tk.StringVar(value="準備完了")
        self.var_progress = tk.DoubleVar(value=0)
        
        # GUI構築
        self._create_widgets()
        
        # 設定読み込み
        self._load_settings_to_gui()

    def _create_widgets(self):
        """ウィジェットの配置"""
        # --- メインフレーム ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- 設定・入力エリア ---
        input_frame = ttk.LabelFrame(main_frame, text="検索設定", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # フォルダ選択
        ttk.Label(input_frame, text="検索対象フォルダ:").grid(row=0, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.var_folder_path, width=60).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(input_frame, text="参照", command=self._browse_folder).grid(row=0, column=2)
        
        # キーワード
        ttk.Label(input_frame, text="検索キーワード:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(input_frame, textvariable=self.var_keyword, width=60).grid(row=1, column=1, padx=5, sticky="ew", pady=5)
        
        # 設定ボタン群
        settings_btn_frame = ttk.Frame(input_frame)
        settings_btn_frame.grid(row=2, column=1, sticky="e", pady=5)
        ttk.Button(settings_btn_frame, text="設定読込", command=self._load_settings_to_gui).pack(side=tk.LEFT, padx=2)
        ttk.Button(settings_btn_frame, text="設定保存", command=self._save_current_settings).pack(side=tk.LEFT, padx=2)

        input_frame.columnconfigure(1, weight=1)

        # --- アクションエリア ---
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_run = ttk.Button(action_frame, text="検索実行", command=self._start_search_thread)
        self.btn_run.pack(side=tk.LEFT, padx=5)
        
        self.btn_cancel = ttk.Button(action_frame, text="キャンセル", command=self._cancel_search, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        
        self.btn_save_log = ttk.Button(action_frame, text="ログ保存(Excel)", command=self._save_log, state=tk.DISABLED)
        self.btn_save_log.pack(side=tk.RIGHT, padx=5)

        # --- 進捗エリア ---
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.var_progress, maximum=100)
        self.progress_bar.pack(fill=tk.X)
        
        ttk.Label(progress_frame, textvariable=self.var_status).pack(anchor="w")

        # --- 結果表示エリア (Treeview) ---
        result_frame = ttk.LabelFrame(main_frame, text="検索結果 (ダブルクリックで開く)", padding="5")
        result_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("file_name", "page", "context", "full_path")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        
        self.tree.heading("file_name", text="ファイル名")
        self.tree.heading("page", text="ページ")
        self.tree.heading("context", text="検出箇所 (プレビュー)")
        self.tree.heading("full_path", text="フルパス") # 非表示または後ろに配置
        
        self.tree.column("file_name", width=200)
        self.tree.column("page", width=50, anchor="center")
        self.tree.column("context", width=400)
        self.tree.column("full_path", width=0, stretch=False) # 隠しカラム的に使う

        # スクロールバー
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # イベントバインド
        self.tree.bind("<Double-1>", self._on_item_double_click)

    # --- イベントハンドラ ---
    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.var_folder_path.set(folder)

    def _load_settings_to_gui(self):
        settings = self.config_manager.load_config()
        self.var_folder_path.set(settings.get('target_folder', ''))
        self.var_keyword.set(settings.get('search_keyword', ''))
        messagebox.showinfo("設定", "設定ファイルを読み込みました。")

    def _save_current_settings(self):
        settings = {
            'target_folder': self.var_folder_path.get(),
            'search_keyword': self.var_keyword.get()
        }
        try:
            self.config_manager.save_config(settings)
            messagebox.showinfo("設定", "現在の設定を保存しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _start_search_thread(self):
        # 入力チェック
        folder = self.var_folder_path.get()
        keyword = self.var_keyword.get()
        
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("入力エラー", "有効なフォルダを選択してください。")
            return
        if not keyword:
            messagebox.showwarning("入力エラー", "検索キーワードを入力してください。")
            return

        # UI状態更新
        self._toggle_ui_state(processing=True)
        self.logic.cancel_flag = False
        
        # 既存結果のクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results_df = pd.DataFrame()

        # スレッド起動
        thread = threading.Thread(target=self._process_search, args=(folder, keyword))
        thread.daemon = True
        thread.start()

    def _process_search(self, folder: str, keyword: str):
        try:
            self.var_status.set("ファイルリストを取得中...")
            pdf_files = self.logic.get_pdf_files(folder)
            total_files = len(pdf_files)
            
            if total_files == 0:
                self.after(0, lambda: messagebox.showinfo("検索終了", "PDFファイルが見つかりませんでした。"))
                self.after(0, lambda: self._toggle_ui_state(processing=False))
                return

            all_results = []
            
            for idx, file_path in enumerate(pdf_files):
                # キャンセルチェック
                if self.logic.cancel_flag:
                    self.var_status.set("検索を中断しました。")
                    break
                
                # 進捗更新
                progress = (idx / total_files) * 100
                self.var_progress.set(progress)
                self.var_status.set(f"検索中 ({idx+1}/{total_files}): {os.path.basename(file_path)}")
                
                # 検索処理
                file_results = self.logic.search_in_pdf(file_path, keyword)
                
                # 結果処理
                for res in file_results:
                    if "error" in res and res["error"]:
                        # エラーログ（コンソール出力またはログリストへの追加）
                        print(f"Skipped {file_path}: {res['error']}")
                        continue
                    
                    all_results.append(res)
                    # Treeviewへの追加（メインスレッドで実行）
                    self.after(0, self._add_result_to_tree, res)

                # CPU負荷を考慮し少しウェイト（必要に応じて）
                time.sleep(0.01)

            self.var_progress.set(100)
            
            # DataFrameへ変換して保持
            self.results_df = pd.DataFrame(all_results)
            
            msg = "検索が完了しました。" if not self.logic.cancel_flag else "検索が中断されました。"
            self.var_status.set(msg)
            self.after(0, lambda: messagebox.showinfo("完了", msg))

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("エラー", f"予期せぬエラーが発生しました: {e}"))
        finally:
            self.after(0, lambda: self._toggle_ui_state(processing=False))

    def _add_result_to_tree(self, result: Dict):
        """検索結果をTreeviewに追加"""
        values = (
            result['file_name'],
            result['page'],
            result['context'],
            result['file_path']
        )
        self.tree.insert("", "end", values=values)

    def _cancel_search(self):
        if messagebox.askyesno("確認", "検索を中断しますか？"):
            self.logic.cancel_flag = True

    def _toggle_ui_state(self, processing: bool):
        """処理中のボタン活性/非活性切り替え"""
        if processing:
            self.btn_run.config(state=tk.DISABLED)
            self.btn_cancel.config(state=tk.NORMAL)
            self.btn_save_log.config(state=tk.DISABLED)
            self.progress_bar.config(mode='determinate')
        else:
            self.btn_run.config(state=tk.NORMAL)
            self.btn_cancel.config(state=tk.DISABLED)
            # 結果があれば保存ボタンを有効化
            if not self.results_df.empty:
                self.btn_save_log.config(state=tk.NORMAL)
            else:
                self.btn_save_log.config(state=tk.DISABLED)

    def _on_item_double_click(self, event):
        """行ダブルクリックでファイルを開く"""
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        item = self.tree.item(selected_item)
        file_path = item['values'][3] # full_path
        page_num = item['values'][1]

        try:
            # Webブラウザ経由で開くと、多くの環境でページ指定(pdf#page=N)が効きやすい
            file_url = f"file:///{os.path.abspath(file_path).replace(os.sep, '/')}#page={page_num}"
            webbrowser.open(file_url)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルを開けませんでした: {e}")

    def _save_log(self):
        """結果をExcelに保存"""
        if self.results_df.empty:
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="検索結果を保存"
        )
        
        if file_path:
            try:
                # Excel出力用に列を整理
                output_df = self.results_df[["file_name", "page", "context", "file_path"]]
                output_df.to_excel(file_path, index=False, engine='openpyxl')
                messagebox.showinfo("保存完了", f"ログを保存しました:\n{file_path}")
            except Exception as e:
                messagebox.showerror("保存エラー", f"保存に失敗しました: {e}")

# =============================================================================
# 5. if __name__ == "__main__": でGUIを起動
# =============================================================================
if __name__ == "__main__":
    app = PDFSearchApp()
    app.mainloop()