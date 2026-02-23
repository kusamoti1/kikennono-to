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

APP_TITLE = "NoticeForge v4.0（NotebookLM最適化・OCR統合版）"

HELP_RIGHT = """ここだけ読めばOK

【1】通知を入力フォルダに入れる。
【2】古い画像PDFがある場合は
　　「OCRを実行」にチェックを入れる。
【3】「処理開始」を押す。
※出力先は毎回リセット（上書き）されます。

【NotebookLMに入れるもの】
出力フォルダにある
・ 00_統合目次.md
・ NotebookLM用_統合データ_〇〇.txt
だけをNotebookLMに入れればOKです！
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
        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(header, text="NoticeForge v4.0", font=ctk.CTkFont(size=26, weight="bold")).pack(pady=(16, 4))
        ctk.CTkLabel(header, text="危険物通知 → NotebookLM用合体テキストを全自動生成", text_color="gray").pack(pady=(0, 16))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        step = ctk.CTkFrame(main)
        step.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        step.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(step, text="① 入力フォルダ（通知が入っているフォルダ）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))
        self.in_entry = ctk.CTkEntry(step, textvariable=self.input_dir)
        self.in_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.in_btn = ctk.CTkButton(step, text="フォルダを選ぶ…", width=160, command=self.pick_input)
        self.in_btn.grid(row=1, column=2, padx=12, pady=(0, 12))

        ctk.CTkLabel(step, text="② 出力フォルダ（結果の保存先）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 6))
        self.out_entry = ctk.CTkEntry(step, textvariable=self.output_dir)
        self.out_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.out_btn = ctk.CTkButton(step, text="変更…", width=160, command=self.pick_output, fg_color="#6b7280")
        self.out_btn.grid(row=3, column=2, padx=12, pady=(0, 12))

        # OCRオプション
        opt_frame = ctk.CTkFrame(main, fg_color="transparent")
        opt_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.chk_ocr = ctk.CTkCheckBox(opt_frame, text="読めないPDFにOCR（画像文字認識）を実行する ※Tesseract必須・時間がかかります", variable=self.use_ocr, font=ctk.CTkFont(weight="bold"))
        self.chk_ocr.pack(side="left", padx=12)

        self.run_btn = ctk.CTkButton(main, text="③ 処理開始（出力先はリセットされます）", height=48, font=ctk.CTkFont(size=16, weight="bold"), command=self.start)
        self.run_btn.grid(row=2, column=0, sticky="ew", pady=(0, 10))

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
        btns.grid_columnconfigure(1, weight=1)

        self.open_out = ctk.CTkButton(btns, text="出力フォルダ", command=self.open_output, fg_color="#2563eb")
        self.open_out.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.open_out.configure(state="disabled")

        self.open_excel = ctk.CTkButton(btns, text="Excel目次", command=self.open_excel_index, fg_color="#16a34a")
        self.open_excel.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.open_excel.configure(state="disabled")

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

        # 警告ポップアップ
        ans = messagebox.askyesno("確認", f"出力フォルダ ({os.path.basename(outdir)}) の内容を再構築（上書き）します。\nよろしいですか？\n※元のファイルは消えません。")
        if not ans:
            return

        os.makedirs(outdir, exist_ok=True)
        self.set_busy(True)
        self.open_out.configure(state="disabled")
        self.open_excel.configure(state="disabled")
        self.progress.set(0)
        self.status.configure(text="開始準備中…", text_color="gray")
        self.append_log("=== 処理開始 ===")

        t = threading.Thread(target=self._worker, args=(indir, outdir, self.use_ocr.get()), daemon=True)
        t.start()

    def _worker(self, indir: str, outdir: str, do_ocr: bool):
        try:
            def cb(curr: int, total: int, fn: str, status_msg: str = ""):
                msg = f"[{curr}/{total}] {fn} {status_msg}"
                self.after(0, lambda: self._progress(curr, total, msg))

            cfg = dict(core.DEFAULTS)
            cfg["use_ocr"] = do_ocr
            total, needs = core.process_folder(indir, outdir, cfg, cb)
            msg = f"完了しました。総数: {total} / 要確認: {needs}\nNotebookLM用の結合データを作成しました。"
            self.after(0, lambda: self._done(msg, False, outdir))
        except Exception as e:
            msg = f"エラー: {type(e).__name__}: {e}"
            self.after(0, lambda: self._done(msg, True, outdir))

    def _progress(self, curr: int, total: int, msg: str):
        if total > 0:
            self.progress.set(curr / total)
        self.status.configure(text=f"処理中… {msg}", text_color="gray")
        self.append_log(msg)

    def _done(self, msg: str, is_error: bool, outdir: str):
        self.set_busy(False)
        if is_error:
            self.progress.set(0)
            self.status.configure(text=msg, text_color="#ff5555")
            self.append_log("[ERROR] " + msg)
            messagebox.showerror("処理失敗", msg)
        else:
            self.progress.set(1)
            self.status.configure(text=msg, text_color="#00aa00")
            self.append_log("[DONE] " + msg)
            self.open_out.configure(state="normal")
            self.open_excel.configure(state="normal")
            messagebox.showinfo("処理完了", msg)

    def open_output(self):
        p = self.output_dir.get()
        if p and sys.platform.startswith("win"): os.startfile(p)

    def open_excel_index(self):
        x = os.path.join(self.output_dir.get(), "00_統合目次.xlsx")
        if os.path.exists(x) and sys.platform.startswith("win"): os.startfile(x)

if __name__ == "__main__":
    app = App()
    app.mainloop()
