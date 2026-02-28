# -*- coding: utf-8 -*-
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

try:
    import noticeforge_best as core
except Exception as e:
    messagebox.showerror("エラー", f"noticeforge_best.py が読み込めません。\n{type(e).__name__}: {e}")
    sys.exit(1)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

APP_TITLE = "NoticeForge v6.0（法令・通知・マニュアル3層対応）"

HELP_RIGHT = """ここだけ読めばOK

【1】入力フォルダにファイルを入れる。
　フォルダ名で自動判別されます:
　・「法令」→ 消防法・政令・規則等
　・「マニュアル」→ 社内手順書等
　・それ以外 → 通知として処理

　推奨フォルダ構成:
　  data/
　  ├ 法令/      ← 消防法・政令・規則
　  ├ 通知/      ← 消防庁通知
　  └ マニュアル/ ← 社内手順書

【2】画像スキャンPDFがある場合のみ
　　「OCRを実行」にチェック。

【3】「処理開始」を押す。

【NotebookLMに入れるもの】
出力フォルダにある:
・ NotebookLM用_○○.txt（全て）
　→ そのままNotebookLMへアップロード

【OCRで処理したファイル】
00_処理ログ.txt で「OCR品質」の
低いファイルを確認してください。
OCRの読取ミスがNotebookLMに
そのまま入ってしまいます。

【途中で止めるには】
「■ 止める」ボタンを押してください。
"""

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1040x760")
        self.minsize(980, 700)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.use_ocr = tk.BooleanVar(value=False)

        self._busy = False
        self._stop_event: threading.Event = threading.Event()
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(header, text="NoticeForge v6.0", font=ctk.CTkFont(size=26, weight="bold")).pack(pady=(16, 4))
        ctk.CTkLabel(header, text="法令・通知・マニュアル3層対応 → NotebookLM専用データ自動生成", text_color="gray").pack(pady=(0, 16))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        step = ctk.CTkFrame(main)
        step.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        step.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(step, text="① 入力フォルダ（法令・通知・マニュアルが入っているフォルダ）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))
        self.in_entry = ctk.CTkEntry(step, textvariable=self.input_dir)
        self.in_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.in_btn = ctk.CTkButton(step, text="フォルダを選ぶ…", width=160, command=self.pick_input)
        self.in_btn.grid(row=1, column=2, padx=12, pady=(0, 12))

        ctk.CTkLabel(step, text="② 出力フォルダ（結果の保存先）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 6))
        self.out_entry = ctk.CTkEntry(step, textvariable=self.output_dir)
        self.out_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.out_btn = ctk.CTkButton(step, text="変更…", width=160, command=self.pick_output, fg_color="#6b7280")
        self.out_btn.grid(row=3, column=2, padx=12, pady=(0, 12))

        opt_frame = ctk.CTkFrame(main, fg_color="transparent")
        opt_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.chk_ocr = ctk.CTkCheckBox(opt_frame, text="スキャンPDFにOCR（文字認識）を実行する ※テキスト埋め込みPDFは自動で通常読取", variable=self.use_ocr, font=ctk.CTkFont(weight="bold"))
        self.chk_ocr.pack(side="left", padx=12)

        # ③ 処理開始ボタン ＋ ■ 止めるボタン（横並び）
        btn_row = ctk.CTkFrame(main, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        btn_row.grid_columnconfigure(0, weight=3)
        btn_row.grid_columnconfigure(1, weight=1)

        self.run_btn = ctk.CTkButton(
            btn_row, text="③ 処理開始（出力先はリセットされます）",
            height=48, font=ctk.CTkFont(size=16, weight="bold"),
            command=self.start
        )
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.stop_btn = ctk.CTkButton(
            btn_row, text="■ 止める",
            height=48, font=ctk.CTkFont(size=14, weight="bold"),
            command=self.stop_processing,
            fg_color="#dc2626", hover_color="#b91c1c",
            state="disabled"
        )
        self.stop_btn.grid(row=0, column=1, sticky="ew")

        st = ctk.CTkFrame(main)
        st.grid(row=3, column=0, sticky="nsew")
        st.grid_columnconfigure(0, weight=1)
        st.grid_rowconfigure(2, weight=1)

        self.progress = ctk.CTkProgressBar(st, height=14)
        self.progress.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        self.progress.set(0)

        self.status = ctk.CTkLabel(st, text="準備完了。入力フォルダを選んで「処理開始」を押してください。", text_color="gray", anchor="w", justify="left")
        self.status.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))

        self.log = ctk.CTkTextbox(st)
        self.log.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.log.insert("end", "ログがここに出ます。\n")

        side = ctk.CTkFrame(self)
        side.grid(row=1, column=1, rowspan=3, sticky="nsew", padx=(10, 20), pady=10)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(side, text="使い方", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        self.help = ctk.CTkTextbox(side, width=320)
        self.help.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.help.insert("end", HELP_RIGHT)
        self.help.configure(state="disabled")

        btns = ctk.CTkFrame(side, fg_color="transparent")
        btns.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 12))
        btns.grid_columnconfigure(0, weight=1)

        self.open_out = ctk.CTkButton(btns, text="出力フォルダを開く", command=self.open_output, fg_color="#2563eb")
        self.open_out.grid(row=0, column=0, sticky="ew")
        self.open_out.configure(state="disabled")

    def append_log(self, s: str):
        self.log.insert("end", s + "\n")
        self.log.see("end")

    def set_busy(self, busy: bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.run_btn.configure(state=state)
        self.in_btn.configure(state=state)
        self.out_btn.configure(state=state)
        self.in_entry.configure(state=state)
        self.out_entry.configure(state=state)
        self.chk_ocr.configure(state=state)
        # 止めるボタンは処理中のみ有効
        self.stop_btn.configure(
            state="normal" if busy else "disabled",
            text="■ 止める"
        )

    def stop_processing(self):
        """処理停止リクエストを送る"""
        self._stop_event.set()
        self.stop_btn.configure(state="disabled", text="停止中…")
        self.status.configure(text="停止リクエスト中… 現在のファイル処理が完了次第停止します。", text_color="#f59e0b")
        self.append_log("[STOP] 停止リクエストを送信しました。")

    def pick_input(self):
        p = filedialog.askdirectory(title="入力フォルダを選択")
        if p:
            self.input_dir.set(p)
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(p, "output_noticeforge"))

    def pick_output(self):
        p = filedialog.askdirectory(title="出力フォルダを選択")
        if p:
            self.output_dir.set(p)

    def start(self):
        indir = self.input_dir.get()
        outdir = self.output_dir.get()
        if not indir or not os.path.isdir(indir):
            messagebox.showwarning("確認", "入力フォルダが選択されていません。")
            return
        if not outdir:
            messagebox.showwarning("確認", "出力フォルダが選択されていません。")
            return

        ans = messagebox.askyesno("確認", f"出力フォルダ ({os.path.basename(outdir)}) の内容を再構築（上書き）します。\nよろしいですか？\n※元のファイルは消えません。")
        if not ans:
            return

        os.makedirs(outdir, exist_ok=True)
        self._stop_event = threading.Event()  # 新しいイベントでリセット
        self.set_busy(True)
        self.open_out.configure(state="disabled")
        self.progress.set(0)
        self.status.configure(text="開始準備中…", text_color="gray")
        self.append_log("=== 処理開始 ===")

        t = threading.Thread(
            target=self._worker,
            args=(indir, outdir, self.use_ocr.get(), self._stop_event),
            daemon=True
        )
        t.start()

    def _worker(self, indir: str, outdir: str, do_ocr: bool, stop_event: threading.Event):
        try:
            def cb(curr: int, total: int, fn: str, status_msg: str = ""):
                if stop_event.is_set():
                    return
                msg = f"[{curr}/{total}] {fn} {status_msg}"
                self.after(0, lambda: self._progress(curr, total, msg))

            cfg = dict(core.DEFAULTS)
            cfg["use_ocr"] = do_ocr
            total, needs, detail = core.process_folder(indir, outdir, cfg, cb, stop_event)

            if stop_event.is_set():
                msg = f"処理を途中で停止しました。処理済み: {total}件 / 要確認: {needs}件\n途中結果は出力フォルダで確認できます。"
                self.after(0, lambda: self._done(msg, False, outdir, stopped=True))
            else:
                detail_msg = f"\n内訳: {detail}" if detail else ""
                msg = f"完了しました。総数: {total} / 要確認: {needs}{detail_msg}\nNotebookLM用データを作成しました。詳細は 00_処理ログ.txt で確認できます。"
                self.after(0, lambda: self._done(msg, False, outdir))
        except PermissionError as pe:
            self.after(0, lambda: self._done(str(pe), True, outdir))
        except Exception as e:
            msg = f"致命的なエラー: {type(e).__name__}: {e}"
            self.after(0, lambda: self._done(msg, True, outdir))

    def _progress(self, curr: int, total: int, msg: str):
        if total > 0:
            self.progress.set(curr / total)
        self.status.configure(text=f"処理中… {msg}", text_color="gray")
        self.append_log(msg)

    def _done(self, msg: str, is_error: bool, outdir: str, stopped: bool = False):
        self.set_busy(False)
        if is_error:
            self.progress.set(0)
            self.status.configure(text=msg, text_color="#ff5555")
            self.append_log("[ERROR] " + msg)
            messagebox.showerror("処理失敗", msg)
        elif stopped:
            self.progress.set(0)
            self.status.configure(text=msg, text_color="#f59e0b")
            self.append_log("[STOPPED] " + msg)
            self.open_out.configure(state="normal")
            messagebox.showinfo("停止しました", msg)
        else:
            self.progress.set(1)
            self.status.configure(text=msg, text_color="#00aa00")
            self.append_log("[DONE] " + msg)
            self.open_out.configure(state="normal")
            messagebox.showinfo("処理完了", msg)

    def open_output(self):
        p = self.output_dir.get()
        if p and sys.platform.startswith("win"): os.startfile(p)

if __name__ == "__main__":
    app = App()
    app.mainloop()
