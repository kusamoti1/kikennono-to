# -*- coding: utf-8 -*-
"""
NoticeForge Core Logic v5.0 (Ultimate: DocuWorks/Excel-MD/LongPath/Binder)
"""
from __future__ import annotations
import os, sys, re, json, time, hashlib, csv, subprocess
from dataclasses import dataclass, asdict
from typing import Dict, List, Tuple, Optional, Callable

# Tesseractの設定 (Windowsの一般的なパス)
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import pytesseract
    from PIL import Image
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    TESSERACT_AVAILABLE = True
except Exception:
    TESSERACT_AVAILABLE = False

try:
    from docx import Document
except Exception:
    Document = None

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
except Exception:
    openpyxl = None

try:
    import xlrd
except Exception:
    xlrd = None

DEFAULTS: Dict[str, object] = {
    "min_chars_mainbody": 400, # 基準を少し甘くして抽出漏れを防止
    "max_depth": 30,
    "summary_chars": 900,
    "main_attach_split_keywords": [r"^\s*別添", r"^\s*別紙", r"^\s*【別添】", r"^\s*【別紙】", r"^\s*【参考】", r"^\s*記\s*$"],
    "bind_bytes_limit": 15 * 1024 * 1024,
    "use_ocr": False,
}

FACILITY_TAGS: Dict[str, List[str]] = {
    "製造所": [r"製造所"],
    "屋外タンク貯蔵所": [r"屋外タンク貯蔵所", r"浮屋根", r"固定屋根", r"アニュラ", r"タンク底", r"泡放射", r"防油堤"],
    "屋内貯蔵所": [r"屋内貯蔵所"],
    "地下タンク貯蔵所": [r"地下タンク貯蔵所", r"FRPタンク", r"漏えい検知"],
    "簡易タンク貯蔵所": [r"簡易タンク貯蔵所"],
    "移動タンク貯蔵所": [r"移動タンク貯蔵所", r"タンクローリー"],
    "給油取扱所": [r"給油取扱所", r"計量機", r"ノズル", r"\bSS\b", r"サービスステーション"],
    "販売取扱所": [r"販売取扱所"],
    "移送取扱所": [r"移送取扱所", r"荷卸し", r"荷積み"],
    "一般取扱所": [r"一般取扱所", r"塗装", r"洗浄", r"混合", r"充填", r"乾燥"],
    "共通": [r"危険物", r"消防法", r"政令", r"規則", r"運用", r"取扱い", r"質疑", r"Q&A", r"解釈"],
}

WORK_TAGS: Dict[str, List[str]] = {
    "申請・届出": [r"許可", r"届出", r"申請", r"変更", r"仮使用", r"完成検査", r"予防規程", r"承認", r"届書", r"様式"],
    "技術基準・設備": [r"技術基準", r"基準", r"構造", r"設備", r"配管", r"タンク", r"保有空地", r"耐震", r"腐食", r"漏えい検知"],
    "運用解釈・Q&A": [r"取扱い", r"運用", r"解釈", r"質疑", r"問", r"答", r"Q&A", r"照会", r"回答"],
    "事故・漏えい・火災": [r"事故", r"漏えい", r"流出", r"火災", r"爆発", r"災害", r"原因", r"再発防止"],
    "消火・防災": [r"泡", r"消火", r"固定消火", r"警報", r"緊急遮断", r"避難", r"防災", r"消火設備"],
    "立入検査・指導": [r"立入", r"検査", r"指導", r"是正", r"改善", r"確認", r"点検", r"報告"],
    "教育・体制": [r"保安監督", r"危険物保安監督者", r"保安統括", r"教育", r"訓練", r"体制", r"責任者"],
}

@dataclass
class Record:
    relpath: str
    ext: str
    size: int
    mtime: float
    sha1: str
    method: str
    pages: Optional[int]
    text_chars: int
    needs_review: bool
    reason: str
    title_guess: str
    date_guess: str
    issuer_guess: str
    summary: str
    tags_facility: List[str]
    tags_work: List[str]
    tag_evidence: Dict[str, List[str]]
    out_txt: str
    full_text_for_bind: str = ""

def get_safe_path(path: str) -> str:
    """Windowsの260文字制限(MAX_PATH)を突破するための安全なパス変換"""
    abs_path = os.path.abspath(path)
    if sys.platform.startswith("win") and not abs_path.startswith("\\\\?\\"):
        return "\\\\?\\" + abs_path
    return abs_path

def extract_pdf(path: str, use_ocr: bool) -> Tuple[str, Optional[int], str]:
    if not fitz: return "", None, "pymupdf_missing"
    text_parts = []
    method = "pdf_text"
    try:
        doc = fitz.open(get_safe_path(path))
        pages = doc.page_count

        # まず全ページのテキストを収集
        raw_texts = [doc.load_page(i).get_text("text") or "" for i in range(pages)]
        total_chars = sum(len(t.strip()) for t in raw_texts)

        # 画像PDF判定：1ページ平均30字未満 → OCR有効時は全ページOCRする
        is_image_pdf = use_ocr and TESSERACT_AVAILABLE and total_chars < max(30 * pages, 100)

        for i in range(pages):
            page_text = raw_texts[i]
            do_ocr_this_page = (
                use_ocr and TESSERACT_AVAILABLE
                and (is_image_pdf or len(page_text.strip()) < 50)
            )
            if do_ocr_this_page:
                try:
                    pix = doc.load_page(i).get_pixmap(dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_text = pytesseract.image_to_string(img, lang="jpn")
                    ocr_text = re.sub(r'([ぁ-んァ-ン一-龥])\s+([ぁ-んァ-ン一-龥])', r'\1\2', ocr_text)
                    # 画像PDF：OCRテキストのみ / 部分欠損：テキスト＋OCRを結合
                    page_text = ocr_text if is_image_pdf else (page_text + "\n" + ocr_text)
                    method = "pdf_ocr" if is_image_pdf else "pdf_ocr_partial"
                except Exception:
                    pass
            text_parts.append(page_text)

        doc.close()
        return "\n".join(text_parts), pages, method
    except Exception as e:
        return "", None, f"pdf_err:{e.__class__.__name__}"

def extract_docx(path: str) -> Tuple[str, str]:
    if not Document: return "", "docx_missing"
    try:
        doc = Document(get_safe_path(path))
        return "\n".join([p.text for p in doc.paragraphs if p.text]), "docx_text"
    except Exception as e:
        return "", f"docx_err:{e.__class__.__name__}"

def extract_excel(path: str) -> Tuple[str, str]:
    """新旧エクセルを読み込み、AIが理解しやすいMarkdown表形式に整形する"""
    out = []
    ext = os.path.splitext(path)[1].lower()
    safe_p = get_safe_path(path)
    try:
        if ext in (".xlsx", ".xlsm") and openpyxl:
            wb = openpyxl.load_workbook(safe_p, data_only=True, read_only=True)
            for ws in wb.worksheets[:10]:
                out.append(f"## Sheet: {ws.title}")
                for row in ws.iter_rows(max_row=400, max_col=40, values_only=True):
                    if any(row):
                        out.append("| " + " | ".join([str(c).strip().replace("\n", " ") if c is not None else "" for c in row]) + " |")
                out.append("")
            wb.close()
            return "\n".join(out), "xlsx_md"
        elif ext == ".xls" and xlrd:
            wb = xlrd.open_workbook(safe_p)
            for sheet_idx in range(min(10, wb.nsheets)):
                ws = wb.sheet_by_index(sheet_idx)
                out.append(f"## Sheet: {ws.name}")
                for row_idx in range(min(400, ws.nrows)):
                    row = ws.row_values(row_idx)
                    if any(row):
                        out.append("| " + " | ".join([str(c).strip().replace("\n", " ") if c else "" for c in row]) + " |")
                out.append("")
            return "\n".join(out), "xls_md"
        else:
            return "", "excel_lib_missing"
    except Exception as e:
        return "", f"excel_err:{e.__class__.__name__}"

def extract_xdw(path: str) -> Tuple[str, str]:
    """DocuWorksから直接テキストを抽出（xdw2textコマンドを使用）"""
    safe_p = get_safe_path(path)
    try:
        result = subprocess.run(["xdw2text", safe_p], capture_output=True, text=True, encoding="cp932", errors="ignore")
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout, "xdw_text"
        return "", "xdw_empty_or_protected"
    except FileNotFoundError:
        return "", "xdw2text_missing (要xdw2text.exe導入)"
    except Exception as e:
        return "", f"xdw_err:{e.__class__.__name__}"

def split_main_attach(text: str, kws: List[str]) -> Tuple[str, str]:
    lines = text.splitlines()
    cut_idx = -1
    for i, line in enumerate(lines):
        for k in kws:
            if re.match(k, line):
                cut_idx = i
                break
        if cut_idx != -1: break

    if cut_idx > 5:
        main_text = "\n".join(lines[:cut_idx])
        attach_text = "\n".join(lines[cut_idx:])
        return main_text.strip(), attach_text.strip()
    return text.strip(), ""

def convert_japanese_year(text: str) -> str:
    def replacer(match):
        era = match.group(1)
        year_str = match.group(2)
        year = 1 if year_str == "元" else int(year_str)
        if era == "令和": west_year = 2018 + year
        elif era == "平成": west_year = 1988 + year
        elif era == "昭和": west_year = 1925 + year
        else: return match.group(0)
        return f"{match.group(0)}（{west_year}年）"
    return re.sub(r"(令和|平成|昭和)\s*([0-9元]+)\s*年", replacer, text)

def guess_title(text: str, fallback: str) -> str:
    for l in text.splitlines()[:50]:
        s = l.strip()
        if 6 <= len(s) <= 120 and not re.match(r"^[\d\-\s\(\)]+$", s): return s
    return fallback

def guess_date(text: str) -> str:
    m = re.search(r"(令和|平成|昭和)\s*[0-9元]+\s*年\s*\d+\s*月\s*\d+\s*日(（\d{4}年）)?", text)
    if m: return m.group(0)
    m2 = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", text)
    return m2.group(0) if m2 else ""

def guess_issuer(text: str) -> str:
    for cand in ["消防庁", "総務省消防庁", "消防局", "危険物保安室", "予防課"]:
        if cand in text: return cand
    return ""

def tag_text(text: str) -> Tuple[List[str], List[str], Dict[str, List[str]]]:
    ev: Dict[str, List[str]] = {}; fac: List[str] = []; work: List[str] = []
    target = text[:8000]
    for t, ps in FACILITY_TAGS.items():
        if hits := [p for p in ps if re.search(p, target)]:
            fac.append(t); ev[t] = hits[:3]
    for t, ps in WORK_TAGS.items():
        if hits := [p for p in ps if re.search(p, target)]:
            work.append(t); ev[t] = hits[:3]
    if not fac and re.search(r"危険物|消防法", target): fac.append("共通")
    return fac, work, ev

def make_summary(main_text: str, n: int) -> str:
    s = re.sub(r"\s+", " ", main_text.strip())
    return s[:n] + ("…" if len(s) > n else "")

_ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def _xls_safe(s) -> str:
    """Excelに書き込めない制御文字を除去する"""
    if not isinstance(s, str):
        return s
    return _ILLEGAL_CHARS_RE.sub("", s)

def write_excel_index(outdir: str, records: List[Record]):
    if not openpyxl: return

    HDR_FILL  = PatternFill("solid", fgColor="1F4E79")   # 濃紺：ヘッダー
    OK_FILL   = PatternFill("solid", fgColor="E8F5E9")   # 薄緑：正常
    REV_FILL  = PatternFill("solid", fgColor="FFF9C4")   # 薄黄：要確認
    ERR_FILL  = PatternFill("solid", fgColor="FFEBEE")   # 薄赤：エラー
    HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
    BODY_FONT = Font(size=10)
    WRAP      = Alignment(wrap_text=True, vertical="top")
    HDR_ALIGN = Alignment(wrap_text=True, vertical="center", horizontal="center")

    def _style_header(ws, col_widths):
        for i, (cell, w) in enumerate(zip(ws[1], col_widths), 1):
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.alignment = HDR_ALIGN
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"

    def _style_row(ws, row_num: int, fill):
        for cell in ws[row_num]:
            cell.fill = fill
            cell.font = BODY_FONT
            cell.alignment = WRAP

    wb = openpyxl.Workbook()

    # ─── Sheet 1: 全体目次（人が確認用）───────────────────────────────
    ws1 = wb.active
    ws1.title = "① 全体目次"
    ws1.append(["No", "タイトル（推定）", "日付（推定）", "発出者", "施設タグ", "業務タグ", "状態", "理由・メモ", "概要（先頭900字）", "元ファイルパス"])
    _style_header(ws1, [5, 38, 14, 14, 20, 24, 10, 22, 55, 42])

    for idx, r in enumerate(records, 1):
        if "err:" in r.method or r.method in ("error", "unhandled"):
            state, fill = "❌エラー", ERR_FILL
        elif r.needs_review:
            state, fill = "⚠️要確認", REV_FILL
        else:
            state, fill = "✅OK", OK_FILL
        ws1.append([
            idx,
            _xls_safe(r.title_guess), _xls_safe(r.date_guess), _xls_safe(r.issuer_guess),
            " / ".join(r.tags_facility), " / ".join(r.tags_work),
            state, _xls_safe(r.reason), _xls_safe(r.summary), _xls_safe(r.relpath),
        ])
        _style_row(ws1, idx + 1, fill)
        ws1.row_dimensions[idx + 1].height = 55

    # ─── Sheet 2: 要確認リスト ──────────────────────────────────────
    ws2 = wb.create_sheet("② 要確認リスト")
    rev_list = [r for r in records if r.needs_review]
    ws2.append(["No", "タイトル（推定）", "理由・メモ", "抽出方式", "文字数", "元ファイルパス"])
    _style_header(ws2, [5, 40, 30, 18, 8, 45])
    if rev_list:
        for idx, r in enumerate(rev_list, 1):
            ws2.append([idx, _xls_safe(r.title_guess), _xls_safe(r.reason), r.method, r.text_chars, _xls_safe(r.relpath)])
            _style_row(ws2, idx + 1, REV_FILL)
    else:
        ws2.append(["", "✅ 要確認ファイルはありません", "", "", "", ""])

    excel_path = os.path.join(outdir, "00_統合目次.xlsx")
    try:
        wb.save(excel_path)
    except PermissionError:
        raise PermissionError("00_統合目次.xlsx が他のアプリで開かれています。閉じてからやり直してください。")

def write_md_indices(outdir: str, records: List[Record]):
    ok_recs  = [r for r in records if not r.needs_review]
    rev_recs = [r for r in records if r.needs_review]

    with open(os.path.join(outdir, "00_統合目次.md"), "w", encoding="utf-8") as f:
        f.write("# 通知文書 統合目次\n\n")
        f.write(f"> 処理ファイル総数: **{len(records)}件** | "
                f"✅ 正常: **{len(ok_recs)}件** | "
                f"⚠️ 要確認: **{len(rev_recs)}件**\n\n")
        f.write("---\n\n")

        # ── 正常ファイル一覧（表形式）────────────────────────────────
        f.write("## ✅ 正常処理ファイル一覧\n\n")
        f.write("| No | タイトル（推定） | 日付（推定） | 発出者 | 施設タグ | 業務タグ |\n")
        f.write("|---|---|---|---|---|---|\n")
        for i, r in enumerate(ok_recs, 1):
            title = r.title_guess.replace("|", "｜").replace("\n", " ")[:60]
            f.write(f"| {i} | {title} | {r.date_guess} | {r.issuer_guess} "
                    f"| {' / '.join(r.tags_facility)} | {' / '.join(r.tags_work)} |\n")

        # ── 要確認ファイル一覧 ────────────────────────────────────
        if rev_recs:
            f.write("\n---\n\n## ⚠️ 要確認ファイル一覧\n\n")
            f.write("| No | ファイル名 | 理由 | 抽出方式 |\n")
            f.write("|---|---|---|---|\n")
            for i, r in enumerate(rev_recs, 1):
                fn = os.path.basename(r.relpath).replace("|", "｜")
                f.write(f"| {i} | {fn} | {r.reason} | {r.method} |\n")

        # ── 概要一覧（NotebookLM読み込み用）────────────────────────
        f.write("\n---\n\n## 各ファイル概要（NotebookLM用）\n\n")
        for r in records:
            title = r.title_guess.replace("\n", " ")
            f.write(f"### {title}\n\n")
            f.write(f"- **日付**: {r.date_guess}\n")
            f.write(f"- **発出者**: {r.issuer_guess}\n")
            f.write(f"- **施設タグ**: {' / '.join(r.tags_facility)}\n")
            f.write(f"- **業務タグ**: {' / '.join(r.tags_work)}\n")
            f.write(f"- **元ファイル**: `{r.relpath}`\n\n")
            f.write(f"> {r.summary}\n\n")

def write_binded_texts(outdir: str, records: List[Record], limit_bytes: int):
    chunk_idx = 1
    current_size = 0
    current_lines = []
    
    def flush():
        nonlocal chunk_idx, current_size, current_lines
        if not current_lines: return
        with open(os.path.join(outdir, f"NotebookLM用_統合データ_{chunk_idx:02d}.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(current_lines))
        chunk_idx += 1
        current_size = 0
        current_lines = []

    for r in records:
        if not r.full_text_for_bind.strip(): continue
        block = f"\n\n{'='*60}\n【DOCUMENT START】\n元ファイル: {r.relpath}\n抽出方式: {r.method}\n{'-'*60}\n{r.full_text_for_bind}\n{'='*60}\n\n"
        b_len = len(block.encode("utf-8"))
        if current_size + b_len > limit_bytes and current_size > 0: flush()
        current_lines.append(block)
        current_size += b_len
    flush()

def process_folder(indir: str, outdir: str, cfg: Dict[str, object], progress_callback: Optional[Callable[[int, int, str, str], None]] = None) -> Tuple[int, int]:
    os.makedirs(outdir, exist_ok=True)
    max_depth = int(cfg.get("max_depth", 30))
    split_kws = list(cfg.get("main_attach_split_keywords", []))
    min_chars = int(cfg.get("min_chars_mainbody", 800))
    use_ocr = bool(cfg.get("use_ocr", False))
    limit_bytes = int(cfg.get("bind_bytes_limit", 15000000))

    targets = [os.path.join(root, fn) for root, _, files in os.walk(indir) if os.path.relpath(root, indir).count(os.sep) < max_depth for fn in files]
    total_files = len(targets)
    records: List[Record] = []

    for i, path in enumerate(targets):
        rel = os.path.relpath(path, indir)
        ext = os.path.splitext(path)[1].lower()
        if progress_callback: progress_callback(i + 1, total_files, rel, "(抽出中...)")

        text, method, reason, pages = "", "unhandled", "", None
        
        try:
            if ext == ".pdf":
                if use_ocr and progress_callback: progress_callback(i + 1, total_files, rel, "(OCR処理中...時間がかかります)")
                text, pages, method = extract_pdf(path, use_ocr)
            elif ext == ".docx":
                text, method = extract_docx(path)
            elif ext in (".xlsx", ".xlsm", ".xls"):
                text, method = extract_excel(path)
            elif ext in (".xdw", ".xbd"):
                text, method = extract_xdw(path)
        except Exception as e:
            method, reason = "error", f"抽出エラー: {e.__class__.__name__}"

        text = convert_japanese_year(text)
        main, attach = split_main_attach(text, split_kws)
        title = guess_title(main or text, os.path.basename(path))
        date_guess = guess_date(text)
        issuer_guess = guess_issuer(text)
        fac, work, ev = tag_text(main or text)

        needs_rev = False
        if method in ("unhandled", "error") or "missing" in method or len(main or text) < min_chars:
            needs_rev = True
            reason = reason or "本文が短すぎる、または画像ファイル"

        summary = make_summary(main or text, int(cfg.get("summary_chars", 900)))
        payload = f"タイトル(推定): {title}\n日付(推定): {date_guess}\n発出者(推定): {issuer_guess}\n\n# 本文\n{main.strip()}"
        if attach.strip(): payload += f"\n\n# 添付資料\n{attach.strip()}"

        records.append(Record(relpath=rel, ext=ext, size=os.path.getsize(get_safe_path(path)), mtime=os.path.getmtime(get_safe_path(path)), sha1="", method=method, pages=pages, text_chars=len(text), needs_review=needs_rev, reason=reason, title_guess=title, date_guess=date_guess, issuer_guess=issuer_guess, summary=summary, tags_facility=fac, tags_work=work, tag_evidence=ev, out_txt="", full_text_for_bind=payload))

    write_excel_index(outdir, records)
    write_md_indices(outdir, records)
    write_binded_texts(outdir, records, limit_bytes)
    
    return len(records), len([r for r in records if r.needs_review])
