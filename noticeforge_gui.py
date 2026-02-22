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

APP_TITLE = "NoticeForge v3.1（目次Excelつき・初心者向け）"

HELP_RIGHT = """ここだけ読めばOK（小学生向け）

【A】新しい通知を入れる場所
→ あなたが選んだ「入力フォルダ」の中。

【B】入れたら勝手に動く？
→ 基本は動きません。
→ もう一度「処理開始」を押すだけ。

【C】概要（目次）はどこ？
→ 出力フォルダの
   00_統合目次.xlsx（Excel）
   00_統合目次.md（メモ帳でも見れる）
   にあります。

【D】次に何をする？
→ NotebookLMへ
  1) 00_統合目次.md
  2) docs_txt の中身（txt全部）
"""

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1040x720")
        self.minsize(980, 660)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        self._busy = False

        self._build_ui()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(header, text="NoticeForge", font=ctk.CTkFont(size=26, weight="bold")).pack(pady=(16, 4))
        ctk.CTkLabel(header, text="危険物通知フォルダ → 目次（Excel/Markdown）とテキストを自動作成", text_color="gray").pack(pady=(0, 16))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        step = ctk.CTkFrame(main)
        step.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        step.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(step, text="① 入力フォルダ（通知が入っているフォルダ）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))
        self.in_entry = ctk.CTkEntry(step, textvariable=self.input_dir, placeholder_text="例）C:\\Users\\...\\危険物通知")
        self.in_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.in_btn = ctk.CTkButton(step, text="フォルダを選ぶ…", width=160, command=self.pick_input)
        self.in_btn.grid(row=1, column=2, padx=12, pady=(0, 12))

        ctk.CTkLabel(step, text="② 出力フォルダ（結果の保存先）", font=ctk.CTkFont(size=13, weight="bold")).grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 6))
        self.out_entry = ctk.CTkEntry(step, textvariable=self.output_dir, placeholder_text="通常は自動でOK（必要なら変更）")
        self.out_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        self.out_btn = ctk.CTkButton(step, text="変更…", width=160, command=self.pick_output, fg_color="#6b7280")
        self.out_btn.grid(row=3, column=2, padx=12, pady=(0, 12))

        self.run_btn = ctk.CTkButton(main, text="③ 処理開始", height=48, font=ctk.CTkFont(size=16, weight="bold"), command=self.start)
        self.run_btn.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        st = ctk.CTkFrame(main)
        st.grid(row=2, column=0, sticky="nsew")
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
        side.grid(row=1, column=1, rowspan=2, sticky="nsew", padx=(10, 20), pady=10)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(side, text="使い方（ここに全部書いてあります）", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        self.help = ctk.CTkTextbox(side, width=360)
        self.help.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.help.insert("end", HELP_RIGHT)
        self.help.configure(state="disabled")

        btns = ctk.CTkFrame(side, fg_color="transparent")
        btns.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 12))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        self.open_out = ctk.CTkButton(btns, text="出力フォルダを開く", command=self.open_output, fg_color="#2563eb")
        self.open_out.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.open_out.configure(state="disabled")

        self.open_excel = ctk.CTkButton(btns, text="Excel目次を開く", command=self.open_excel_index, fg_color="#16a34a")
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

    def pick_input(self):
        p = filedialog.askdirectory(title="入力フォルダ（危険物通知フォルダ）を選択")
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

        os.makedirs(outdir, exist_ok=True)
        self.set_busy(True)
        self.open_out.configure(state="disabled")
        self.open_excel.configure(state="disabled")
        self.progress.set(0)
        self.status.configure(text="開始準備中…", text_color="gray")
        self.append_log("=== 開始 ===")

        t = threading.Thread(target=self._worker, args=(indir, outdir), daemon=True)
        t.start()

    def _worker(self, indir: str, outdir: str):
        try:
            def cb(curr: int, total: int, fn: str):
                self.after(0, lambda: self._progress(curr, total, fn))

            cfg = dict(core.DEFAULTS)
            total, needs = core.process_folder(indir, outdir, cfg, cb)
            msg = f"完了しました。総ファイル数: {total} / 要確認: {needs}\n出力先: {outdir}\nExcel目次: 00_統合目次.xlsx"
            self.after(0, lambda: self._done(msg, False, outdir))
        except Exception as e:
            msg = f"エラー: {type(e).__name__}: {e}"
            self.after(0, lambda: self._done(msg, True, outdir))

    def _progress(self, curr: int, total: int, fn: str):
        if total > 0:
            self.progress.set(curr / total)
        self.status.configure(text=f"処理中… [{curr}/{total}] {fn}", text_color="gray")
        self.append_log(f"[{curr}/{total}] {fn}")

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
        if not p:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(p)  # type: ignore
            elif sys.platform == "darwin":
                import subprocess; subprocess.Popen(["open", p])
            else:
                import subprocess; subprocess.Popen(["xdg-open", p])
        except Exception:
            pass

    def open_excel_index(self):
        outdir = self.output_dir.get()
        x = os.path.join(outdir, "00_統合目次.xlsx")
        if not os.path.exists(x):
            messagebox.showwarning("確認", "00_統合目次.xlsx が見つかりません。先に処理開始してください。")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(x)  # type: ignore
            elif sys.platform == "darwin":
                import subprocess; subprocess.Popen(["open", x])
            else:
                import subprocess; subprocess.Popen(["xdg-open", x])
        except Exception:
            messagebox.showwarning("確認", "Excelを開けませんでした。ファイルアプリから開いてください。")

if __name__ == "__main__":
    app = App()
    app.mainloop()
