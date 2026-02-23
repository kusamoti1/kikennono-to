# -*- coding: utf-8 -*-
"""
NoticeForge Core Logic v4.0 (NotebookLM Binder & OCR Integrated)
"""
from __future__ import annotations
import os, re, json, time, hashlib, csv
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
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
except Exception:
    openpyxl = None

DEFAULTS: Dict[str, object] = {
    "min_chars_mainbody": 800,
    "max_depth": 30,
    "summary_chars": 900,
    # 厳格な分割キーワード (行頭付近にある場合のみ切断)
    "main_attach_split_keywords": [r"^\s*別添", r"^\s*別紙", r"^\s*【別添】", r"^\s*【別紙】", r"^\s*【参考】", r"^\s*記\s*$"],
    "bind_bytes_limit": 15 * 1024 * 1024, # 約15MBごとに分割（NotebookLM用）
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
    full_text_for_bind: str = "" # バインド用のフルテキスト保持

def sha1_file(path: str) -> str:
    h = hashlib.sha1()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(2 * 1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return "000000"

def extract_pdf(path: str, use_ocr: bool) -> Tuple[str, Optional[int], str]:
    if not fitz: return "", None, "pymupdf_missing"
    text_parts = []
    method = "pdf_text"
    try:
        doc = fitz.open(path)
        pages = doc.page_count
        for i in range(pages):
            page = doc.load_page(i)
            page_text = page.get_text("text") or ""
            
            # テキストが極端に少ない場合、画像PDFとみなしてOCRを実行
            if use_ocr and len(page_text.strip()) < 50 and TESSERACT_AVAILABLE:
                try:
                    pix = page.get_pixmap(dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_text = pytesseract.image_to_string(img, lang="jpn")
                    # OCRのノイズ掃除（過剰なスペース除去）
                    ocr_text = re.sub(r'([ぁ-んァ-ン一-龥])\s+([ぁ-んァ-ン一-龥])', r'\1\2', ocr_text)
                    page_text += "\n" + ocr_text
                    method = "pdf_ocr"
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
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs if p.text]), "docx_text"
    except Exception as e:
        return "", f"docx_err:{e.__class__.__name__}"

def split_main_attach(text: str, kws: List[str]) -> Tuple[str, str]:
    # 校正者指摘：行頭付近のキーワードのみで切断する
    lines = text.splitlines()
    cut_idx = -1
    for i, line in enumerate(lines):
        for k in kws:
            if re.match(k, line):
                cut_idx = i
                break
        if cut_idx != -1: break

    if cut_idx > 5: # あまりに早すぎる切断は無視
        main_text = "\n".join(lines[:cut_idx])
        attach_text = "\n".join(lines[cut_idx:])
        return main_text.strip(), attach_text.strip()
    return text.strip(), ""

def convert_japanese_year(text: str) -> str:
    # NotebookLM最適化：和暦を西暦に翻訳補完
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
        if 6 <= len(s) <= 120 and not re.match(r"^[\d\-\s\(\)]+$", s):
            return s
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
    if not fac and re.search(r"危険物|消防法", target):
        fac.append("共通")
    return fac, work, ev

def make_summary(main_text: str, n: int) -> str:
    s = re.sub(r"\s+", " ", main_text.strip())
    return s[:n] + ("…" if len(s) > n else "")

def write_excel_index(outdir: str, records: List[Record]):
    if not openpyxl: return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Index"
    headers = ["タイトル(推定)", "日付(推定)", "発出者(推定)", "施設タグ", "業務タグ", "needs_review", "理由", "概要(先頭)", "元ファイル"]
    ws.append(headers)
    for r in records:
        ws.append([r.title_guess, r.date_guess, r.issuer_guess, " / ".join(r.tags_facility), " / ".join(r.tags_work), "TRUE" if r.needs_review else "FALSE", r.reason, r.summary, r.relpath])
    wb.save(os.path.join(outdir, "00_統合目次.xlsx"))

def write_md_indices(outdir: str, records: List[Record]):
    with open(os.path.join(outdir, "00_統合目次.md"), "w", encoding="utf-8") as f:
        f.write("# 統合目次（概要付き）\n\n")
        for r in records:
            f.write(f"- **{r.title_guess}**\n  - 日付: {r.date_guess} / 発出: {r.issuer_guess}\n  - タグ: [{'/'.join(r.tags_facility)}] [{'/'.join(r.tags_work)}]\n  - 概要: {r.summary}\n  - 元: `{r.relpath}`\n\n")

def write_binded_texts(outdir: str, records: List[Record], limit_bytes: int):
    # NotebookLM用の巨大結合ファイルを作成
    chunk_idx = 1
    current_size = 0
    current_lines = []
    
    def flush():
        nonlocal chunk_idx, current_size, current_lines
        if not current_lines: return
        path = os.path.join(outdir, f"NotebookLM用_統合データ_{chunk_idx:02d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(current_lines))
        chunk_idx += 1
        current_size = 0
        current_lines = []

    for r in records:
        if not r.full_text_for_bind.strip(): continue
        # 明確な区切り線
        block = f"\n\n{'='*60}\n【DOCUMENT START】\n元ファイル: {r.relpath}\n抽出方式: {r.method}\n{'-'*60}\n{r.full_text_for_bind}\n{'='*60}\n\n"
        b_len = len(block.encode("utf-8"))
        if current_size + b_len > limit_bytes and current_size > 0:
            flush()
        current_lines.append(block)
        current_size += b_len
    flush()

def process_folder(indir: str, outdir: str, cfg: Dict[str, object], progress_callback: Optional[Callable[[int, int, str, str], None]] = None) -> Tuple[int, int]:
    # 安全のため毎回作り直す仕様
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
            # NotebookLMに不要なファイルは一旦スキップするか抽出機能を残す
        except Exception as e:
            method, reason = "error", f"抽出エラー: {e.__class__.__name__}"

        text = convert_japanese_year(text)
        main, attach = split_main_attach(text, split_kws)
        title = guess_title(main or text, os.path.basename(path))
        date_guess = guess_date(text)
        issuer_guess = guess_issuer(text)
        fac, work, ev = tag_text(main or text)

        needs_rev = False
        if method in ("unhandled", "error") or len(main or text) < min_chars:
            needs_rev = True
            reason = reason or "本文が短すぎる、または画像PDF（要OCR）"

        summary = make_summary(main or text, int(cfg.get("summary_chars", 900)))
        
        # NotebookLM用合体テキスト作成のための文字列
        payload = f"タイトル(推定): {title}\n日付(推定): {date_guess}\n発出者(推定): {issuer_guess}\n\n# 本文\n{main.strip()}"
        if attach.strip(): payload += f"\n\n# 添付資料\n{attach.strip()}"

        records.append(Record(relpath=rel, ext=ext, size=os.path.getsize(path), mtime=os.path.getmtime(path), sha1="", method=method, pages=pages, text_chars=len(text), needs_review=needs_rev, reason=reason, title_guess=title, date_guess=date_guess, issuer_guess=issuer_guess, summary=summary, tags_facility=fac, tags_work=work, tag_evidence=ev, out_txt="", full_text_for_bind=payload))

    # 結果の出力
    write_excel_index(outdir, records)
    write_md_indices(outdir, records)
    write_binded_texts(outdir, records, limit_bytes)
    
    needs = [r for r in records if r.needs_review]
    return len(records), len(needs)
