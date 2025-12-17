import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import configparser
import os
import re
import time
import threading
import webbrowser
import concurrent.futures  # 並列処理用
from typing import List, Dict, Any

# 外部ライブラリ
import pandas as pd
from pypdf import PdfReader

# =============================================================================
# 2. ini設定の読み書き関数 / クラス
# =============================================================================
class ConfigManager:
    DEFAULT_CONFIG = {
        'Settings': {
            'target_folder': '',
            'search_keyword': '',
            'context_length': '30'
        }
    }
    
    def __init__(self, filename: str = 'config.ini'):
        self.filename = filename
        self.config = configparser.ConfigParser()

    def load_config(self) -> Dict[str, Any]:
        if not os.path.exists(self.filename):
            return self.DEFAULT_CONFIG['Settings'].copy()
        try:
            self.config.read(self.filename, encoding='utf-8')
            if 'Settings' in self.config:
                return dict(self.config['Settings'])
            else:
                return self.DEFAULT_CONFIG['Settings'].copy()
        except Exception:
            return self.DEFAULT_CONFIG['Settings'].copy()

    def save_config(self, settings: Dict[str, str]) -> None:
        self.config['Settings'] = settings
        with open(self.filename, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

# =============================================================================
# 3. メイン処理関数 (並列処理ロジック)
# =============================================================================
class SearchLogic:
    def __init__(self):
        self.cancel_flag = False

    def get_pdf_files(self, folder_path: str) -> List[str]:
        pdf_files = []
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        return pdf_files

    def search_in_pdf(self, file_path: str, keyword: str, context_len: int = 30) -> List[Dict]:
        """pypdfを使用した検索処理"""
        results = []
        try:
            if self.cancel_flag: return []

            reader = PdfReader(file_path)
            # 暗号化ファイル対応
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except:
                    return [{"error": "Password Protected"}]

            for i, page in enumerate(reader.pages):
                if self.cancel_flag: break
                try:
                    text = page.extract_text()
                    if text:
                        clean_text = text.replace('\n', '')
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
                            # 1ページに複数ヒットしても良いが、高速化のためbreakを入れても良い
                except:
                    pass
        except Exception as e:
            return [{"error": str(e)}]
            
        return results

# =============================================================================
# 4. GUIクラス
# =============================================================================
class PDFSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF文字列検索ツール (並列処理版)")
        self.geometry("900x650")
        
        self.config_manager = ConfigManager()
        self.logic = SearchLogic()
        self.results_df = pd.DataFrame()
        
        self.var_folder_path = tk.StringVar()
        self.var_keyword = tk.StringVar()
        self.var_status = tk.StringVar(value="準備完了")
        self.var_progress = tk.DoubleVar(value=0)
        
        self._create_widgets()
        self._load_settings_to_gui()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 入力エリア
        input_frame = ttk.LabelFrame(main_frame, text="検索設定", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(input_frame, text="フォルダ:").grid(row=0, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.var_folder_path).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(input_frame, text="参照", command=self._browse_folder).grid(row=0, column=2)
        
        ttk.Label(input_frame, text="キーワード:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(input_frame, textvariable=self.var_keyword).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=2, column=1, sticky="e")
        ttk.Button(btn_frame, text="設定読込", command=self._load_settings_to_gui).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="設定保存", command=self._save_current_settings).pack(side=tk.LEFT, padx=2)
        
        input_frame.columnconfigure(1, weight=1)

        # アクションエリア
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        self.btn_run = ttk.Button(action_frame, text="検索実行", command=self._start_search_thread)
        self.btn_run.pack(side=tk.LEFT, padx=5)
        self.btn_cancel = ttk.Button(action_frame, text="キャンセル", command=self._cancel_search, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT, padx=5)
        self.btn_save_log = ttk.Button(action_frame, text="ログ保存", command=self._save_log, state=tk.DISABLED)
        self.btn_save_log.pack(side=tk.RIGHT, padx=5)

        # 進捗
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.var_progress)
        self.progress_bar.pack(fill=tk.X, pady=5)
        ttk.Label(main_frame, textvariable=self.var_status).pack(anchor="w")

        # 結果表示
        columns = ("file_name", "page", "context", "full_path")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        self.tree.heading("file_name", text="ファイル名")
        self.tree.heading("page", text="ページ")
        self.tree.heading("context", text="検出箇所")
        self.tree.column("file_name", width=200)
        self.tree.column("page", width=50, anchor="center")
        self.tree.column("context", width=400)
        self.tree.column("full_path", width=0, stretch=False)
        
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<Double-1>", self._on_item_double_click)

    def _browse_folder(self):
        f = filedialog.askdirectory()
        if f: self.var_folder_path.set(f)

    def _load_settings_to_gui(self):
        s = self.config_manager.load_config()
        self.var_folder_path.set(s.get('target_folder', ''))
        self.var_keyword.set(s.get('search_keyword', ''))

    def _save_current_settings(self):
        self.config_manager.save_config({
            'target_folder': self.var_folder_path.get(),
            'search_keyword': self.var_keyword.get()
        })
        messagebox.showinfo("設定", "保存しました。")

    def _start_search_thread(self):
        folder = self.var_folder_path.get()
        keyword = self.var_keyword.get()
        if not folder or not keyword:
            messagebox.showwarning("エラー", "フォルダとキーワードを入力してください。")
            return
            
        self._toggle_ui_state(True)
        self.logic.cancel_flag = False
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        threading.Thread(target=self._process_search, args=(folder, keyword), daemon=True).start()

    def _process_search(self, folder, keyword):
        try:
            self.var_status.set("ファイルリスト取得中...")
            pdf_files = self.logic.get_pdf_files(folder)
            total = len(pdf_files)
            
            if total == 0:
                self.after(0, lambda: messagebox.showinfo("完了", "ファイルがありません。"))
                self.after(0, lambda: self._toggle_ui_state(False))
                return

            self.var_status.set(f"検索開始: {total}件...")
            processed_count = 0
            
            # 並列処理 (ネットワーク負荷を考慮し同時実行数は5程度に制限)
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                future_to_file = {executor.submit(self.logic.search_in_pdf, p, keyword): p for p in pdf_files}
                
                for future in concurrent.futures.as_completed(future_to_file):
                    if self.logic.cancel_flag: break
                    
                    try:
                        results = future.result()
                        for res in results:
                            if not res.get("error"):
                                self.after(0, self._add_result, res)
                    except Exception:
                        pass
                    
                    processed_count += 1
                    self.var_progress.set((processed_count / total) * 100)
                    self.var_status.set(f"検索中 ({processed_count}/{total})")

            # 結果データの保存用DataFrame作成
            self._update_results_df_from_tree()

            msg = "検索完了" if not self.logic.cancel_flag else "中断されました"
            self.after(0, lambda: messagebox.showinfo("完了", msg))

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("エラー", str(e)))
        finally:
            self.after(0, lambda: self._toggle_ui_state(False))

    def _add_result(self, res):
        self.tree.insert("", "end", values=(res['file_name'], res['page'], res['context'], res['file_path']))

    def _update_results_df_from_tree(self):
        # GUIスレッドで実行する必要があるためafterで呼ぶか、完了後に実行
        rows = []
        for child in self.tree.get_children():
            val = self.tree.item(child)['values']
            rows.append({"file_name": val[0], "page": val[1], "context": val[2], "file_path": val[3]})
        self.results_df = pd.DataFrame(rows)

    def _toggle_ui_state(self, processing):
        self.btn_run.config(state=tk.DISABLED if processing else tk.NORMAL)
        self.btn_cancel.config(state=tk.NORMAL if processing else tk.DISABLED)
        self.btn_save_log.config(state=tk.DISABLED if processing else tk.NORMAL)
        if not processing and not self.results_df.empty:
            self.btn_save_log.config(state=tk.NORMAL)

    def _cancel_search(self):
        if messagebox.askyesno("確認", "中断しますか？"):
            self.logic.cancel_flag = True

    def _on_item_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        path = self.tree.item(sel)['values'][3]
        page = self.tree.item(sel)['values'][1]
        try:
            webbrowser.open(f"file:///{os.path.abspath(path).replace(os.sep, '/')}#page={page}")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _save_log(self):
        if self.results_df.empty: return
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if f:
            try:
                # Treeviewからの再取得が必要な場合があるため念のため更新
                self._update_results_df_from_tree()
                self.results_df.to_excel(f, index=False)
                messagebox.showinfo("保存", "保存しました。")
            except Exception as e:
                messagebox.showerror("エラー", str(e))

if __name__ == "__main__":
    app = PDFSearchApp()
    app.mainloop()