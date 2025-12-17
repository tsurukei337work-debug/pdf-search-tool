import os
import re
import sys
import zlib
import time
import queue
import mmap
import typing as t
import webbrowser
import traceback
import configparser
import threading
import concurrent.futures
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from openpyxl import Workbook


# ========== 2. ini設定の読み書き関数 ==========
def load_ini(ini_path: str) -> dict:
    config = configparser.ConfigParser()
    settings = {
        "target_folder": "",
        "target_file": "",
        "search_text": "",
        "recursive": True,
        "case_sensitive": False,
        "use_regex": False,
    }
    if os.path.exists(ini_path):
        try:
            config.read(ini_path, encoding="utf-8")
            s = config["SETTINGS"]
            settings["target_folder"] = s.get("target_folder", "")
            settings["target_file"] = s.get("target_file", "")
            settings["search_text"] = s.get("search_text", "")
            settings["recursive"] = s.getboolean("recursive", True)
            settings["case_sensitive"] = s.getboolean("case_sensitive", False)
            settings["use_regex"] = s.getboolean("use_regex", False)
        except Exception:
            pass
    return settings


def save_ini(ini_path: str, settings: dict) -> None:
    config = configparser.ConfigParser()
    config["SETTINGS"] = {
        "target_folder": settings.get("target_folder", ""),
        "target_file": settings.get("target_file", ""),
        "search_text": settings.get("search_text", ""),
        "recursive": str(settings.get("recursive", True)),
        "case_sensitive": str(settings.get("case_sensitive", False)),
        "use_regex": str(settings.get("use_regex", False)),
    }
    with open(ini_path, "w", encoding="utf-8") as f:
        config.write(f)


# ========== PDFテキスト抽出（純Python最適化版） ==========
# 事前コンパイル済みパターンで速度を稼ぐ
_RE_OBJ = re.compile(rb"\b(\d+)\s+0\s+obj\b")
_RE_IS_PAGE = re.compile(rb"/Type\s*/Page\b")
_RE_CONTENTS_SINGLE = re.compile(rb"/Contents\s+(\d+)\s+0\s+R")
_RE_CONTENTS_ARRAY = re.compile(rb"/Contents\s*\[(.*?)\]", flags=re.S)
_RE_INDIRECT = re.compile(rb"(\d+)\s+0\s+R")
_RE_BT_ET = re.compile(rb"BT(.*?)ET", flags=re.S)
_RE_PAREN_STRING = re.compile(rb"\((?:\\.|[^\\])*?\)", flags=re.S)


def _pdf_unescape_string(b: bytes) -> str:
    out = bytearray()
    i = 0
    n = len(b)
    while i < n:
        c = b[i]
        if c == 92 and i + 1 < n:  # "\"
            i += 1
            esc = b[i]
            if esc in (40, 41, 92):  # ( ) \
                out.append(esc)
                i += 1
            elif esc == 110:  # n
                out.append(10)
                i += 1
            elif esc == 114:  # r
                out.append(13)
                i += 1
            elif esc == 116:  # t
                out.append(9)
                i += 1
            elif esc == 98:  # b
                out.append(8)
                i += 1
            elif esc == 102:  # f
                out.append(12)
                i += 1
            elif 48 <= esc <= 55:  # octal up to 3 digits
                oct_digits = [esc]
                i += 1
                for _ in range(2):
                    if i < n and 48 <= b[i] <= 55:
                        oct_digits.append(b[i])
                        i += 1
                    else:
                        break
                try:
                    out.append(int(bytes(oct_digits), 8))
                except Exception:
                    pass
            else:
                out.append(esc)
                i += 1
        else:
            out.append(c)
            i += 1
    try:
        return out.decode("utf-8")
    except Exception:
        return out.decode("latin-1", errors="ignore")


def _extract_streams_from_object(obj_bytes: bytes) -> t.List[bytes]:
    streams: t.List[bytes] = []
    pos = 0
    while True:
        s_idx = obj_bytes.find(b"stream", pos)
        if s_idx == -1:
            break
        s_idx_end = s_idx + 6
        # EOL調整
        if s_idx_end < len(obj_bytes) and obj_bytes[s_idx_end] in (10, 13):
            if obj_bytes[s_idx_end] == 13 and s_idx_end + 1 < len(obj_bytes) and obj_bytes[s_idx_end + 1] == 10:
                s_idx_end += 2
            else:
                s_idx_end += 1
        e_idx = obj_bytes.find(b"endstream", s_idx_end)
        if e_idx == -1:
            break
        raw = obj_bytes[s_idx_end:e_idx]
        header = obj_bytes[:s_idx]
        data = raw
        if b"/FlateDecode" in header:
            # 2パターンで伸長を試みる
            try:
                data = zlib.decompress(raw)
            except Exception:
                try:
                    data = zlib.decompress(raw, -15)
                except Exception:
                    data = raw
        streams.append(data)
        pos = e_idx + 9
    return streams


def _extract_text_from_content_stream(data: bytes) -> str:
    # BT..ET 範囲から () 文字列のみ抽出
    chunks: t.List[str] = []
    for m in _RE_BT_ET.finditer(data):
        sec = m.group(1)
        for sm in _RE_PAREN_STRING.finditer(sec):
            s = sm.group(0)
            inner = s[1:-1]
            chunks.append(_pdf_unescape_string(inner))
    return "".join(chunks)


def extract_text_per_page_fast(pdf_path: str) -> t.Dict[int, str]:
    # メモリマップで高速にスキャン
    with open(pdf_path, "rb") as f:
        with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            data = mm[:]

    # オブジェクト境界
    objects: t.Dict[int, bytes] = {}
    for m in _RE_OBJ.finditer(data):
        obj_id = int(m.group(1))
        start = m.end()
        end = data.find(b"endobj", start)
        if end == -1:
            continue
        objects[obj_id] = data[start:end]

    # PageとContents参照抽出
    pages: t.List[int] = []
    page_contents_map: t.Dict[int, t.List[int]] = {}
    for oid, body in objects.items():
        if _RE_IS_PAGE.search(body):
            pages.append(oid)
            contents: t.List[int] = []
            m_single = _RE_CONTENTS_SINGLE.search(body)
            if m_single:
                contents.append(int(m_single.group(1)))
            else:
                m_arr = _RE_CONTENTS_ARRAY.search(body)
                if m_arr:
                    refs = _RE_INDIRECT.findall(m_arr.group(1))
                    contents.extend([int(r) for r in refs])
            page_contents_map[oid] = contents

    result: t.Dict[int, str] = {}
    page_index = 1
    for page_oid in pages:
        texts: t.List[str] = []
        for cid in page_contents_map.get(page_oid, []):
            obj_bytes = objects.get(cid)
            if not obj_bytes:
                continue
            for stream in _extract_streams_from_object(obj_bytes):
                ttxt = _extract_text_from_content_stream(stream)
                if ttxt:
                    texts.append(ttxt)
        result[page_index] = "\n".join(texts)
        page_index += 1
    return result


# ========== 3. メイン処理関数 ==========
class ProcessorOptions(t.TypedDict):
    recursive: bool
    case_sensitive: bool
    use_regex: bool


class SearchResult(t.TypedDict, total=False):
    file: str
    page: int
    snippet: str
    error: str
    trace: str


def _compile_pattern(search_text: str, case_sensitive: bool, use_regex: bool) -> t.Optional[re.Pattern]:
    if not use_regex:
        return None
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(search_text, flags)


def _find_matches_in_text(text: str, search_text: str, case_sensitive: bool, patt: t.Optional[re.Pattern]) -> t.List[t.Tuple[int, int]]:
    spans: t.List[t.Tuple[int, int]] = []
    if patt is not None:
        for m in patt.finditer(text):
            spans.append((m.start(), m.end()))
        return spans
    if case_sensitive:
        idx = text.find(search_text)
        L = len(search_text)
        while idx != -1:
            spans.append((idx, idx + L))
            idx = text.find(search_text, idx + 1)
    else:
        tgt = search_text.lower()
        src = text.lower()
        L = len(tgt)
        idx = src.find(tgt)
        while idx != -1:
            spans.append((idx, idx + L))
            idx = src.find(tgt, idx + 1)
    return spans


def _make_snippet(text: str, span: t.Tuple[int, int], context: int = 30) -> str:
    s, e = span
    start = max(0, s - context)
    end = min(len(text), e + context)
    snippet = text[start:end].replace("\n", " ").replace("\r", " ")
    if start > 0:
        snippet = "..." + snippet
    if end < len(text):
        snippet = snippet + "..."
    return snippet


def _worker_search_file(
    fpath: str,
    search_text: str,
    case_sensitive: bool,
    patt: t.Optional[re.Pattern],
    cancel_event: threading.Event,
) -> t.Tuple[t.List[SearchResult], t.List[SearchResult]]:
    res: t.List[SearchResult] = []
    err: t.List[SearchResult] = []

    if cancel_event.is_set():
        return res, err

    try:
        pages_text = extract_text_per_page_fast(fpath)
        for page_no, page_text in pages_text.items():
            if cancel_event.is_set():
                break
            if not page_text:
                continue
            spans = _find_matches_in_text(page_text, search_text, case_sensitive, patt)
            for sp in spans:
                res.append(
                    {
                        "file": fpath,
                        "page": page_no,
                        "snippet": _make_snippet(page_text, sp),
                    }
                )
    except Exception as e:
        err.append(
            {
                "file": fpath,
                "error": f"{type(e).__name__}: {e}",
                "trace": traceback.format_exc(),
            }
        )
    return res, err


def main_processor(
    files: t.List[str],
    search_text: str,
    options: ProcessorOptions,
    progress_cb: t.Callable[[int, int, str], None],
    cancel_flag: t.Callable[[], bool],
    log_cb: t.Callable[[str], None],
) -> t.Tuple[t.List[SearchResult], t.List[SearchResult]]:
    total = len(files)
    results: t.List[SearchResult] = []
    errors: t.List[SearchResult] = []

    patt = _compile_pattern(search_text, options.get("case_sensitive", False), options.get("use_regex", False))

    # 並列数はCPUコア×2を上限に調整
    max_workers = max(2, min(32, (os.cpu_count() or 4) * 2))

    done = 0
    progress_cb(0, total, "検索開始")

    # スレッド並列（zlibやreでGIL解放が多く有利）
    cancel_event = threading.Event()

    def _cancel_check() -> bool:
        if cancel_flag():
            cancel_event.set()
            return True
        return False

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        future_map = {
            ex.submit(_worker_search_file, f, search_text, options.get("case_sensitive", False), patt, cancel_event): f
            for f in files
        }
        for fut in concurrent.futures.as_completed(future_map):
            if _cancel_check():
                log_cb("ユーザーによりキャンセルされました")
                break
            fpath = future_map[fut]
            try:
                res, err = fut.result()
                if res:
                    results.extend(res)
                if err:
                    errors.extend(err)
            except Exception as e:
                errors.append(
                    {
                        "file": fpath,
                        "error": f"{type(e).__name__}: {e}",
                        "trace": traceback.format_exc(),
                    }
                )
            finally:
                done += 1
                progress_cb(done, total, f"進捗: {done}/{total}")

    return results, errors


# ========== 4. GUIクラス ==========
class PdfSearchApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF文字列検索ツール（高速化・並列版）")
        self.geometry("980x640")
        self.minsize(900, 560)

        self.settings_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "settings.ini")

        self.var_folder = tk.StringVar()
        self.var_file = tk.StringVar()
        self.var_search = tk.StringVar()
        self.var_recursive = tk.BooleanVar(value=True)
        self.var_case = tk.BooleanVar(value=False)
        self.var_regex = tk.BooleanVar(value=False)
        self.var_status = tk.StringVar(value="準備完了")

        self._thread: t.Optional[threading.Thread] = None
        self._cancel_event = threading.Event()
        self._ui_queue: "queue.Queue[t.Tuple[str, t.Any]]" = queue.Queue()
        self._results: t.List[SearchResult] = []
        self._errors: t.List[SearchResult] = []

        self._build_ui()
        self._load_settings_to_ui()

        self.after(100, self._drain_ui_queue)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}

        frm_top = ttk.LabelFrame(self, text="入力と設定")
        frm_top.pack(fill="x", **pad)

        row1 = ttk.Frame(frm_top)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="対象フォルダ").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_folder).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row1, text="参照", command=self._browse_folder).pack(side="left")

        row2 = ttk.Frame(frm_top)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="対象ファイル").pack(side="left")
        ttk.Entry(row2, textvariable=self.var_file).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row2, text="参照", command=self._browse_file).pack(side="left")

        row3 = ttk.Frame(frm_top)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="検索文字列").pack(side="left")
        ttk.Entry(row3, textvariable=self.var_search).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Checkbutton(row3, text="サブフォルダも検索", variable=self.var_recursive).pack(side="left", padx=6)
        ttk.Checkbutton(row3, text="大文字小文字を区別", variable=self.var_case).pack(side="left", padx=6)
        ttk.Checkbutton(row3, text="正規表現を使用", variable=self.var_regex).pack(side="left", padx=6)

        row4 = ttk.Frame(frm_top)
        row4.pack(fill="x", **pad)
        ttk.Button(row4, text="設定読込", command=self._on_load_settings).pack(side="left")
        ttk.Button(row4, text="設定保存", command=self._on_save_settings).pack(side="left")
        ttk.Button(row4, text="ログ保存", command=self._on_save_log).pack(side="left", padx=12)
        ttk.Button(row4, text="実行", command=self._on_run).pack(side="right")
        self.btn_cancel = ttk.Button(row4, text="キャンセル", command=self._on_cancel, state="disabled")
        self.btn_cancel.pack(side="right", padx=6)

        row5 = ttk.Frame(self)
        row5.pack(fill="x", **pad)
        self.prog = ttk.Progressbar(row5, orient="horizontal", mode="determinate")
        self.prog.pack(fill="x", expand=True, side="left")
        ttk.Label(row5, textvariable=self.var_status, anchor="w").pack(side="left", padx=8)

        frm_res = ttk.LabelFrame(self, text="検索結果（左ダブルクリックで該当ページを開く）")
        frm_res.pack(fill="both", expand=True, **pad)

        columns = ("file", "page", "snippet")
        self.tree = ttk.Treeview(frm_res, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("file", text="ファイル")
        self.tree.heading("page", text="ページ")
        self.tree.heading("snippet", text="抜粋")
        self.tree.column("file", width=360, anchor="w")
        self.tree.column("page", width=60, anchor="center")
        self.tree.column("snippet", width=520, anchor="w")
        vsb = ttk.Scrollbar(frm_res, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frm_res, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.bind("<Double-1>", self._on_open_item)

    def _browse_folder(self) -> None:
        d = filedialog.askdirectory(title="フォルダを選択")
        if d:
            self.var_folder.set(d)

    def _browse_file(self) -> None:
        f = filedialog.askopenfilename(
            title="PDFファイルを選択",
            filetypes=[("PDF", "*.pdf"), ("すべてのファイル", "*.*")],
        )
        if f:
            self.var_file.set(f)

    # ---------- 設定 ----------
    def _on_load_settings(self) -> None:
        path = filedialog.askopenfilename(
            title="設定ファイルを選択",
            filetypes=[("INI ファイル", "*.ini"), ("すべてのファイル", "*.*")],
            initialfile=os.path.basename(self.settings_path),
            initialdir=os.path.dirname(self.settings_path),
        )
        if not path:
            return
        try:
            settings = load_ini(path)
            self.var_folder.set(settings.get("target_folder", ""))
            self.var_file.set(settings.get("target_file", ""))
            self.var_search.set(settings.get("search_text", ""))
            self.var_recursive.set(bool(settings.get("recursive", True)))
            self.var_case.set(bool(settings.get("case_sensitive", False)))
            self.var_regex.set(bool(settings.get("use_regex", False)))
            self.var_status.set("設定を読み込みました")
        except Exception as e:
            messagebox.showerror("エラー", f"設定の読み込みに失敗しました: {e}")

    def _on_save_settings(self) -> None:
        path = filedialog.asksaveasfilename(
            title="設定を保存",
            defaultextension=".ini",
            filetypes=[("INI ファイル", "*.ini")],
            initialfile=os.path.basename(self.settings_path),
            initialdir=os.path.dirname(self.settings_path),
        )
        if not path:
            return
        try:
            settings = {
                "target_folder": self.var_folder.get().strip(),
                "target_file": self.var_file.get().strip(),
                "search_text": self.var_search.get().strip(),
                "recursive": self.var_recursive.get(),
                "case_sensitive": self.var_case.get(),
                "use_regex": self.var_regex.get(),
            }
            save_ini(path, settings)
            self.var_status.set("設定を保存しました")
        except Exception as e:
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")

    def _load_settings_to_ui(self) -> None:
        s = load_ini(self.settings_path)
        self.var_folder.set(s.get("target_folder", ""))
        self.var_file.set(s.get("target_file", ""))
        self.var_search.set(s.get("search_text", ""))
        self.var_recursive.set(bool(s.get("recursive", True)))
        self.var_case.set(bool(s.get("case_sensitive", False)))
        self.var_regex.set(bool(s.get("use_regex", False)))

    # ---------- 実行フロー ----------
    def _collect_files(self) -> t.List[str]:
        files: t.List[str] = []
        folder = self.var_folder.get().strip()
        fpath = self.var_file.get().strip()
        recursive = self.var_recursive.get()

        if folder and os.path.isdir(folder):
            if recursive:
                for root, _, fnames in os.walk(folder):
                    for n in fnames:
                        if n.lower().endswith(".pdf"):
                            files.append(os.path.join(root, n))
            else:
                for n in os.listdir(folder):
                    if n.lower().endswith(".pdf"):
                        files.append(os.path.join(folder, n))

        if fpath and os.path.isfile(fpath) and fpath.lower().endswith(".pdf"):
            if fpath not in files:
                files.append(fpath)

        # 小さいファイルを先に処理（体感速度向上）
        try:
            files.sort(key=lambda p: os.path.getsize(p))
        except Exception:
            pass
        return files

    def _validate_inputs(self) -> bool:
        search = self.var_search.get().strip()
        if not search:
            messagebox.showwarning("入力不足", "検索文字列を入力してください。")
            return False
        files = self._collect_files()
        if not files:
            messagebox.showwarning("入力不足", "PDFファイルのあるフォルダまたはファイルを指定してください。")
            return False
        return True

    def _clear_results(self) -> None:
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._results = []
        self._errors = []

    def _on_run(self) -> None:
        if not self._validate_inputs():
            return

        self._clear_results()
        files = self._collect_files()
        self.prog.configure(value=0, maximum=len(files))
        self.var_status.set("1. 設定読み込みと入力チェック: OK")
        self.update_idletasks()

        opts: ProcessorOptions = {
            "recursive": self.var_recursive.get(),
            "case_sensitive": self.var_case.get(),
            "use_regex": self.var_regex.get(),
        }
        search_text = self.var_search.get().strip()

        self._cancel_event.clear()
        self.btn_cancel.configure(state="normal")
        self.var_status.set("2. データ読み込みと処理を開始（高速並列）")
        print(f"[INFO] 検索開始: files={len(files)} search='{search_text}' options={opts}")

        def progress_cb(done: int, total: int, msg: str) -> None:
            # 進捗更新はUIスレッドで
            self._ui_queue.put(("progress", (done, total, msg)))

        def cancel_flag() -> bool:
            return self._cancel_event.is_set()

        def log_cb(msg: str) -> None:
            self._ui_queue.put(("log", msg))

        def worker() -> None:
            try:
                results, errors = main_processor(files, search_text, opts, progress_cb, cancel_flag, log_cb)
                self._ui_queue.put(("result_all", (results, errors)))
            except Exception as e:
                self._ui_queue.put(("fatal", f"致命的エラー: {e}\n{traceback.format_exc()}"))

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()

    def _on_cancel(self) -> None:
        if self._thread and self._thread.is_alive():
            self._cancel_event.set()
            self.btn_cancel.configure(state="disabled")
            self.var_status.set("キャンセル要求を送信しました...")

    # ---------- 結果/イベント ----------
    def _on_open_item(self, event: tk.Event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        vals = self.tree.item(item_id, "values")
        if not vals or len(vals) < 2:
            return
        fpath = vals[0]
        try:
            page = int(vals[1])
        except Exception:
            page = 1
        self._open_pdf_at_page(fpath, page)

    def _open_pdf_at_page(self, fpath: str, page: int) -> None:
        try:
            url = Path(fpath).resolve().as_uri() + f"#page={page}"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("エラー", f"PDFを開けませんでした: {e}")

    def _append_result_row(self, row: SearchResult) -> None:
        self.tree.insert("", "end", values=(row["file"], row["page"], row["snippet"]))
        self._results.append(row)

    def _append_error(self, err: SearchResult) -> None:
        self._errors.append(err)

    def _drain_ui_queue(self) -> None:
        try:
            # バッチ処理でUI更新を抑制
            batch_results: t.List[SearchResult] = []
            while True:
                kind, payload = self._ui_queue.get_nowait()
                if kind == "progress":
                    done, total, msg = payload
                    self.prog.configure(maximum=total, value=done)
                    self.var_status.set(msg)
                elif kind == "log":
                    print(f"[LOG] {payload}")
                elif kind == "result_all":
                    results, errors = payload
                    batch_results.extend(results)
                    for r in batch_results:
                        self._append_result_row(r)
                    for e in errors:
                        self._append_error(e)
                    self.btn_cancel.configure(state="disabled")
                    self.var_status.set("3. 出力と完了メッセージ表示")
                    messagebox.showinfo("完了", "検索が完了しました。")
                    print(f"[INFO] 検索完了: ヒット件数={len(results)} エラー件数={len(errors)}")
                    batch_results.clear()
                elif kind == "fatal":
                    self.btn_cancel.configure(state="disabled")
                    messagebox.showerror("致命的エラー", str(payload))
                    self.var_status.set("エラー終了")
        except queue.Empty:
            pass
        # 100ms周期で更新
        self.after(100, self._drain_ui_queue)

    # ---------- ログ保存 ----------
    def _on_save_log(self) -> None:
        if not self._results and not self._errors:
            messagebox.showinfo("情報", "保存するログがありません。")
            return
        path = filedialog.asksaveasfilename(
            title="ログを保存",
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
            initialfile=f"pdf_search_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )
        if not path:
            return
        try:
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "results"
            ws1.append(["file", "page", "snippet"])
            for r in self._results:
                ws1.append([r.get("file", ""), r.get("page", ""), r.get("snippet", "")])

            ws2 = wb.create_sheet("errors")
            ws2.append(["file", "error", "trace"])
            for e in self._errors:
                ws2.append([e.get("file", ""), e.get("error", ""), e.get("trace", "")])

            wb.save(path)
            self.var_status.set(f"ログを保存しました: {path}")
            print(f"[INFO] ログ保存: {path}")
        except Exception as e:
            messagebox.showerror("エラー", f"ログ保存に失敗しました: {e}")


# ========== 5. GUI起動 ==========
if __name__ == "__main__":
    app = PdfSearchApp()
    app.mainloop()