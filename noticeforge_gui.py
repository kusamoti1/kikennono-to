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

APP_TITLE = "NoticeForge v5.0 ― 通知文書 → NotebookLM 変換ツール"

HELP_TEXT = """\
━━━━━━━━━━━━━━━━━━━━━━
 📖  使い方ガイド（3ステップ）
━━━━━━━━━━━━━━━━━━━━━━

【STEP 1】 📂 フォルダを設定する
─────────────────────────
通知文書（PDF・Word・Excel等）
が入ったフォルダを選択します。
出力フォルダは自動で設定されます。

【STEP 2】 ▶ まず通常処理
─────────────────────────
「処理開始」を押します。
OCRなしで高速に処理します。
終了後に結果が表示されます。

【STEP 3】 🔍 必要なら OCR 再処理
─────────────────────────
⚠️ 要確認ファイルが多い場合は
「OCRで再処理」ボタンを押します。
画像PDFも全ページ読み取ります。
※処理時間は長くなります。

━━━━━━━━━━━━━━━━━━━━━━
 📁  出力ファイルの使い方
━━━━━━━━━━━━━━━━━━━━━━

【人が確認する用】
  📊 00_統合目次.xlsx
    → 全ファイル一覧
    　 ✅ 緑＝正常
    　 ⚠️ 黄＝要確認
    　 ❌ 赤＝エラー

【NotebookLMに入れる用】
  📄 00_統合目次.md
  📄 NotebookLM用_統合データ_*.txt
    → この2種類をNotebookLMに
      アップロードしてください

━━━━━━━━━━━━━━━━━━━━━━
 ⚠️  注意事項
━━━━━━━━━━━━━━━━━━━━━━
・処理前に 00_統合目次.xlsx を
  閉じてください（上書きエラー防止）
・元のファイルは変更されません
・OCRには Tesseract-OCR が必要です
"""


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1140x820")
        self.minsize(1000, 740)

        self.input_dir  = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.use_ocr    = tk.BooleanVar(value=False)

        self._busy = False
        self._build_ui()

    # ─────────────────────────────────────────────────────────────
    #  UI 構築
    # ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(1, weight=1)

        # ── ヘッダーバー ──────────────────────────────────────────
        hdr = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1a3a5c", "#0d2137"))
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(
            hdr, text="📋  NoticeForge v5.0",
            font=ctk.CTkFont(size=24, weight="bold"), text_color="white",
        ).pack(side="left", padx=20, pady=12)
        ctk.CTkLabel(
            hdr, text="危険物通知文書 → NotebookLM 自動変換ツール",
            font=ctk.CTkFont(size=13), text_color="#90caf9",
        ).pack(side="left", pady=12)

        # ── 左メインエリア ─────────────────────────────────────────
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=(16, 8), pady=12)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)   # ログ行が伸びる

        # ── STEP 1: フォルダ設定 ─────────────────────────────────
        s1_inner = self._section(main, "STEP 1  📂  フォルダを設定する", row=0)
        s1_inner.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(s1_inner, text="入力フォルダ", width=90, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(12, 6), pady=(12, 4))
        self.in_entry = ctk.CTkEntry(
            s1_inner, textvariable=self.input_dir,
            placeholder_text="通知文書が入ったフォルダを選択...")
        self.in_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=(12, 4))
        self.in_btn = ctk.CTkButton(
            s1_inner, text="📂 選択", width=110, command=self.pick_input)
        self.in_btn.grid(row=0, column=2, padx=(0, 12), pady=(12, 4))

        ctk.CTkLabel(s1_inner, text="出力フォルダ", width=90, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(12, 6), pady=(4, 12))
        self.out_entry = ctk.CTkEntry(
            s1_inner, textvariable=self.output_dir,
            placeholder_text="処理結果の保存先（自動設定されます）")
        self.out_entry.grid(row=1, column=1, sticky="ew", padx=(0, 6), pady=(4, 12))
        self.out_btn = ctk.CTkButton(
            s1_inner, text="✏️ 変更", width=110, fg_color="#6b7280",
            command=self.pick_output)
        self.out_btn.grid(row=1, column=2, padx=(0, 12), pady=(4, 12))

        # ── STEP 2: オプション ────────────────────────────────────
        s2_inner = self._section(main, "STEP 2  ⚙️  処理オプション", row=1)
        self.chk_ocr = ctk.CTkCheckBox(
            s2_inner,
            text="🔍 OCR（画像PDF対応）を有効にする  "
                 "※処理時間が増えますが、読み取り精度が向上します",
            variable=self.use_ocr,
            font=ctk.CTkFont(size=13),
        )
        self.chk_ocr.pack(padx=14, pady=14, anchor="w")

        # ── STEP 3: 処理開始ボタン ────────────────────────────────
        s3_inner = self._section(main, "STEP 3  ▶  処理を開始する", row=2)
        self.run_btn = ctk.CTkButton(
            s3_inner,
            text="▶  処理開始（出力先フォルダの内容は上書きされます）",
            height=50, font=ctk.CTkFont(size=16, weight="bold"),
            command=self.start,
        )
        self.run_btn.pack(fill="x", padx=12, pady=12)

        # ── 処理状況 ──────────────────────────────────────────────
        s4_inner = self._section(main, "処理状況", row=3)

        self.progress = ctk.CTkProgressBar(s4_inner, height=16)
        self.progress.pack(fill="x", padx=12, pady=(12, 4))
        self.progress.set(0)

        self.status_lbl = ctk.CTkLabel(
            s4_inner,
            text="準備完了。フォルダを選んで「処理開始」を押してください。",
            text_color="gray", anchor="w",
        )
        self.status_lbl.pack(fill="x", padx=12, pady=(0, 6))

        # 統計表示（処理後）
        stats_row = ctk.CTkFrame(s4_inner, fg_color="transparent")
        stats_row.pack(fill="x", padx=12, pady=(0, 4))
        self.stats_ok_lbl  = ctk.CTkLabel(
            stats_row, text="", font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#22c55e")
        self.stats_ok_lbl.pack(side="left", padx=(0, 20))
        self.stats_rev_lbl = ctk.CTkLabel(
            stats_row, text="", font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#f59e0b")
        self.stats_rev_lbl.pack(side="left")

        # アクションボタン行
        act_row = ctk.CTkFrame(s4_inner, fg_color="transparent")
        act_row.pack(fill="x", padx=12, pady=(4, 12))
        self.open_out_btn = ctk.CTkButton(
            act_row, text="📂 出力フォルダを開く",
            state="disabled", command=self.open_output, fg_color="#2563eb")
        self.open_out_btn.pack(side="left", padx=(0, 8))
        self.open_excel_btn = ctk.CTkButton(
            act_row, text="📊 Excel目次を開く",
            state="disabled", command=self.open_excel_index, fg_color="#16a34a")
        self.open_excel_btn.pack(side="left", padx=(0, 8))
        self.ocr_retry_btn = ctk.CTkButton(
            act_row, text="🔍 OCRで再処理（STEP 3）",
            state="disabled", command=self.retry_with_ocr, fg_color="#9333ea")
        self.ocr_retry_btn.pack(side="left")

        # ── ログ ──────────────────────────────────────────────────
        s5_outer = ctk.CTkFrame(main)
        s5_outer.grid(row=4, column=0, sticky="nsew", pady=(0, 0))
        s5_outer.grid_columnconfigure(0, weight=1)
        s5_outer.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(
            s5_outer, text="  📄  処理ログ",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w", fg_color=("#dbeafe", "#1e3a5f"), corner_radius=4, height=30,
        ).grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))
        self.log = ctk.CTkTextbox(
            s5_outer, font=ctk.CTkFont(family="Courier", size=12))
        self.log.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self.log.insert("end", "ログがここに出ます。\n")

        # ── 右サイドバー（使い方） ─────────────────────────────────
        side = ctk.CTkFrame(self, width=310)
        side.grid(row=1, column=1, sticky="nsew", padx=(0, 16), pady=12)
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            side, text="📖  使い方ガイド",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        help_box = ctk.CTkTextbox(side, font=ctk.CTkFont(size=12))
        help_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        help_box.insert("end", HELP_TEXT)
        help_box.configure(state="disabled")

    def _section(self, parent, title: str, row: int) -> ctk.CTkFrame:
        """タイトル付きセクションを作成。コンテンツフレームを返す。"""
        outer = ctk.CTkFrame(parent)
        outer.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        outer.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            outer, text=f"  {title}",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w", fg_color=("#dbeafe", "#1e3a5f"),
            corner_radius=4, height=30,
        ).grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))

        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
        inner.grid_columnconfigure(0, weight=1)
        return inner

    # ─────────────────────────────────────────────────────────────
    #  ログ
    # ─────────────────────────────────────────────────────────────
    def append_log(self, s: str, level: str = "info"):
        prefix = {"ok": "✅", "warn": "⚠️", "error": "❌"}.get(level, "  ")
        self.log.insert("end", f"{prefix} {s}\n")
        self.log.see("end")

    # ─────────────────────────────────────────────────────────────
    #  UI 状態制御
    # ─────────────────────────────────────────────────────────────
    def set_busy(self, busy: bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for w in (self.run_btn, self.in_btn, self.out_btn,
                  self.in_entry, self.out_entry, self.chk_ocr):
            w.configure(state=state)

    # ─────────────────────────────────────────────────────────────
    #  フォルダ選択
    # ─────────────────────────────────────────────────────────────
    def pick_input(self):
        p = filedialog.askdirectory(title="入力フォルダを選択")
        if p:
            self.input_dir.set(p)
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(p, "出力"))

    def pick_output(self):
        p = filedialog.askdirectory(title="出力フォルダを選択")
        if p:
            self.output_dir.set(p)

    # ─────────────────────────────────────────────────────────────
    #  処理開始
    # ─────────────────────────────────────────────────────────────
    def start(self):
        indir  = self.input_dir.get()
        outdir = self.output_dir.get()
        if not indir or not os.path.isdir(indir):
            messagebox.showwarning("確認", "入力フォルダが選択されていません。")
            return
        if not outdir:
            messagebox.showwarning("確認", "出力フォルダが選択されていません。")
            return
        ans = messagebox.askyesno(
            "確認",
            f"出力フォルダ「{os.path.basename(outdir)}」の内容を上書きして処理を開始します。\n"
            f"よろしいですか？\n\n（元の通知ファイルは変更されません）"
        )
        if not ans:
            return
        self._run(indir, outdir, self.use_ocr.get())

    def retry_with_ocr(self):
        """OCR有効で再処理"""
        indir  = self.input_dir.get()
        outdir = self.output_dir.get()
        if not indir or not outdir:
            return
        ans = messagebox.askyesno(
            "🔍 OCRで再処理",
            "OCRを有効にして再処理します。\n"
            "画像PDFも全ページ読み取るため、時間がかかります。\n\n"
            "よろしいですか？"
        )
        if not ans:
            return
        self.use_ocr.set(True)
        self._run(indir, outdir, True)

    def _run(self, indir: str, outdir: str, do_ocr: bool):
        os.makedirs(outdir, exist_ok=True)
        self.set_busy(True)
        self.open_out_btn.configure(state="disabled")
        self.open_excel_btn.configure(state="disabled")
        self.ocr_retry_btn.configure(state="disabled")
        self.stats_ok_lbl.configure(text="")
        self.stats_rev_lbl.configure(text="")
        self.progress.set(0)
        self.status_lbl.configure(
            text="開始準備中…", text_color="gray")
        ocr_label = "（OCR有効）" if do_ocr else ""
        self.append_log(f"=== 処理開始 {ocr_label} ===")
        t = threading.Thread(
            target=self._worker, args=(indir, outdir, do_ocr), daemon=True)
        t.start()

    # ─────────────────────────────────────────────────────────────
    #  バックグラウンドワーカー
    # ─────────────────────────────────────────────────────────────
    def _worker(self, indir: str, outdir: str, do_ocr: bool):
        try:
            def cb(curr: int, total: int, fn: str, status_msg: str = ""):
                msg = f"[{curr}/{total}] {fn} {status_msg}"
                self.after(0, lambda: self._progress(curr, total, msg))

            cfg = dict(core.DEFAULTS)
            cfg["use_ocr"] = do_ocr
            total, needs = core.process_folder(indir, outdir, cfg, cb)
            self.after(0, lambda: self._done(total, needs, outdir, False))
        except PermissionError as pe:
            msg = str(pe)
            self.after(0, lambda: self._done(0, 0, outdir, True, msg))
        except Exception as e:
            msg = f"致命的なエラー: {type(e).__name__}: {e}"
            self.after(0, lambda: self._done(0, 0, outdir, True, msg))

    def _progress(self, curr: int, total: int, msg: str):
        if total > 0:
            self.progress.set(curr / total)
        self.status_lbl.configure(text=f"処理中… {msg}", text_color="gray")
        self.append_log(msg)

    def _done(self, total: int, needs: int, outdir: str,
              is_error: bool, error_msg: str = ""):
        self.set_busy(False)
        if is_error:
            self.progress.set(0)
            self.status_lbl.configure(text=error_msg, text_color="#ef4444")
            self.append_log(error_msg, level="error")
            messagebox.showerror("処理失敗", error_msg)
        else:
            ok = total - needs
            self.progress.set(1)
            self.status_lbl.configure(
                text=f"✅ 処理完了  総数: {total}件", text_color="#22c55e")
            self.stats_ok_lbl.configure(text=f"✅ 正常: {ok}件")
            self.stats_rev_lbl.configure(text=f"⚠️ 要確認: {needs}件")
            self.append_log(
                f"完了 — 総数: {total} / ✅ 正常: {ok} / ⚠️ 要確認: {needs}",
                level="ok")
            self.open_out_btn.configure(state="normal")
            self.open_excel_btn.configure(state="normal")

            # 要確認ファイルがあってOCRを未使用のとき → 再処理ボタンを表示
            if needs > 0 and not self.use_ocr.get():
                self.ocr_retry_btn.configure(state="normal")
                self.append_log(
                    f"{needs}件が要確認。「OCRで再処理」ボタンで読み取り精度を上げられます。",
                    level="warn")

            msg = (f"処理完了\n\n"
                   f"✅ 正常: {ok}件\n"
                   f"⚠️ 要確認: {needs}件\n\n"
                   f"出力フォルダに結果が保存されました。\n"
                   + (f"\n⚠️ 要確認ファイルがあります。\n"
                      f"「OCRで再処理」ボタンで精度を上げられます。"
                      if needs > 0 and not self.use_ocr.get() else ""))
            messagebox.showinfo("処理完了", msg)

    # ─────────────────────────────────────────────────────────────
    #  ファイルを開く
    # ─────────────────────────────────────────────────────────────
    def open_output(self):
        p = self.output_dir.get()
        if p and sys.platform.startswith("win"):
            os.startfile(p)

    def open_excel_index(self):
        x = os.path.join(self.output_dir.get(), "00_統合目次.xlsx")
        if os.path.exists(x) and sys.platform.startswith("win"):
            os.startfile(x)


if __name__ == "__main__":
    app = App()
    app.mainloop()
