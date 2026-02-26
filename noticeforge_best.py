# -*- coding: utf-8 -*-
"""
NoticeForge Core Logic v5.0 (Ultimate: DocuWorks/Excel-MD/LongPath/Binder)
"""
from __future__ import annotations
import os, sys, re, json, time, hashlib, csv, subprocess, html as _html
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

try:
    import xlrd
except Exception:
    xlrd = None

try:
    import xdwlib
    XDWLIB_AVAILABLE = True
except Exception:
    XDWLIB_AVAILABLE = False

# xdw2text.exe の候補パス（DocuWorksの一般的なインストール先を網羅）
XDW2TEXT_CANDIDATES = [
    "xdw2text",  # PATH上にある場合
    r"C:\Program Files\Fuji Xerox\DocuWorks\xdw2text.exe",
    r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\xdw2text.exe",
    r"C:\Program Files\FUJIFILM\DocuWorks\xdw2text.exe",
    r"C:\Program Files (x86)\FUJIFILM\DocuWorks\xdw2text.exe",
    r"C:\Program Files\DocuWorks\xdw2text.exe",
    r"C:\Program Files (x86)\DocuWorks\xdw2text.exe",
]

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
        for i in range(pages):
            page = doc.load_page(i)
            page_text = page.get_text("text") or ""
            if use_ocr and len(page_text.strip()) < 50 and TESSERACT_AVAILABLE:
                try:
                    pix = page.get_pixmap(dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_text = pytesseract.image_to_string(img, lang="jpn")
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
        doc = Document(get_safe_path(path))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                if any(cells):
                    parts.append("| " + " | ".join(cells) + " |")
        return "\n".join(parts), "docx_text"
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
    """DocuWorksから直接テキストを抽出（xdwlib優先、次にxdw2text複数パス試行）"""
    safe_p = get_safe_path(path)

    # 方法1: xdwlib（Python製DocuWorksバインディング）を優先的に試す
    if XDWLIB_AVAILABLE:
        try:
            doc = xdwlib.xdwopen(path)
            texts = []
            for pg in range(doc.pages):
                page = doc[pg]
                texts.append(page.text)
            doc.close()
            result = "\n".join(texts)
            if result.strip():
                return result, "xdw_xdwlib"
        except Exception:
            pass  # 失敗したらxdw2textにフォールバック

    # 方法2: xdw2text.exe を複数の候補パスで試す
    for cmd in XDW2TEXT_CANDIDATES:
        try:
            result = subprocess.run(
                [cmd, safe_p],
                capture_output=True, text=True,
                encoding="cp932", errors="ignore",
                timeout=30
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout, "xdw_text"
        except FileNotFoundError:
            continue  # このパスにはexeがないので次を試す
        except Exception:
            continue

    return "", "xdw2text_missing (要xdw2text.exe導入: DocuWorksインストールフォルダ内)"

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
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Index"
    ws.append(["タイトル(推定)", "日付(推定)", "発出者(推定)", "施設タグ", "業務タグ", "needs_review", "理由", "概要(先頭)", "元ファイル"])
    for r in records:
        ws.append([_xls_safe(r.title_guess), _xls_safe(r.date_guess), _xls_safe(r.issuer_guess), " / ".join(r.tags_facility), " / ".join(r.tags_work), "TRUE" if r.needs_review else "FALSE", _xls_safe(r.reason), _xls_safe(r.summary), _xls_safe(r.relpath)])
    
    excel_path = os.path.join(outdir, "00_統合目次.xlsx")
    try:
        wb.save(excel_path)
    except PermissionError:
        raise PermissionError("00_統合目次.xlsx が他のアプリで開かれています。閉じてからやり直してください。")

def write_md_indices(outdir: str, records: List[Record]):
    with open(os.path.join(outdir, "00_統合目次.md"), "w", encoding="utf-8") as f:
        f.write("# 統合目次（概要付き）\n\n")
        for r in records:
            f.write(f"- **{r.title_guess}**\n  - 日付: {r.date_guess} / 発出: {r.issuer_guess}\n  - タグ: [{'/'.join(r.tags_facility)}] [{'/'.join(r.tags_work)}]\n  - 概要: {r.summary}\n  - 元: `{r.relpath}`\n\n")

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

def compute_sha1(path: str) -> str:
    """ファイルのSHA1ハッシュを計算して重複ファイル検出に使う"""
    h = hashlib.sha1()
    try:
        with open(get_safe_path(path), "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""

def extract_txt(path: str) -> Tuple[str, str]:
    """プレーンテキストファイルを読み込む（文字コードを自動判定）"""
    for enc in ("utf-8-sig", "cp932", "utf-8", "latin-1"):
        try:
            with open(get_safe_path(path), "r", encoding=enc, errors="ignore") as f:
                return f.read(), "txt_read"
        except Exception:
            continue
    return "", "txt_err"

def extract_csv(path: str) -> Tuple[str, str]:
    """CSVファイルをMarkdown表形式に整形する"""
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            with open(get_safe_path(path), "r", encoding=enc, newline="", errors="ignore") as f:
                rows = list(csv.reader(f))
            if not rows:
                return "", "csv_empty"
            out = []
            for row in rows[:400]:
                if any(c.strip() for c in row):
                    out.append("| " + " | ".join([c.strip().replace("\n", " ") for c in row]) + " |")
            return "\n".join(out), "csv_md"
        except Exception:
            continue
    return "", "csv_err"

def write_html_report(outdir: str, records: List[Record]):
    """人間が見やすいHTMLレポートを生成する（ブラウザで開くだけでOK）"""
    def esc(s: object) -> str:
        return _html.escape(str(s) if s is not None else "")

    total = len(records)
    ok_count = sum(1 for r in records if not r.needs_review)
    needs_rev_count = total - ok_count

    # 抽出方式ごとの集計
    method_counts: Dict[str, int] = {}
    for r in records:
        method_counts[r.method] = method_counts.get(r.method, 0) + 1
    method_rows = "".join(
        f"<tr><td>{esc(m)}</td><td style='text-align:right'>{c}</td></tr>"
        for m, c in sorted(method_counts.items(), key=lambda x: -x[1])
    )

    # 施設タグ・業務タグ用のバッジ色マップ
    FAC_COLOR  = "#2563eb"
    WORK_COLOR = "#16a34a"

    def make_badge(text: str, color: str) -> str:
        return f'<span class="badge" style="background:{color}">{esc(text)}</span>'

    # ファイルカード生成
    cards_html = []
    for r in records:
        card_cls  = "card-review" if r.needs_review else "card-ok"
        rev_badge = '<span class="rev-badge">⚠ 要確認</span>' if r.needs_review else \
                    '<span class="ok-badge">✓ 正常</span>'
        fac_badges  = "".join(make_badge(t, FAC_COLOR)  for t in r.tags_facility)
        work_badges = "".join(make_badge(t, WORK_COLOR) for t in r.tags_work)
        tags_html   = (fac_badges + work_badges) or '<span style="color:#94a3b8;font-size:12px">タグなし</span>'

        date_str   = esc(r.date_guess)   or "日付不明"
        issuer_str = esc(r.issuer_guess) or "発出者不明"
        pages_str  = f"/{r.pages}p" if r.pages else ""
        method_str = esc(r.method)
        size_kb    = f"{r.size // 1024:,} KB" if r.size >= 1024 else f"{r.size} B"

        reason_html = (
            f'<div class="reason-box">⚠ {esc(r.reason)}</div>'
            if r.reason else ""
        )

        # data-search に検索対象テキストを全部まとめる（小文字化はJS側で行う）
        search_data = " ".join([
            r.title_guess, r.summary, r.relpath,
            r.date_guess, r.issuer_guess,
            " ".join(r.tags_facility), " ".join(r.tags_work),
            r.reason, r.method,
        ]).replace('"', '')

        cards_html.append(f"""
<div class="card {card_cls}" data-search="{esc(search_data.lower())}">
  <div class="card-header">
    <div class="card-title">{esc(r.title_guess)}</div>
    {rev_badge}
  </div>
  <div class="meta">
    <span>📅 {date_str}</span>
    <span>🏢 {issuer_str}</span>
    <span>📄 {esc(r.ext.upper().lstrip('.'))}{pages_str} · {size_kb}</span>
    <span class="method-tag">抽出: {method_str}</span>
  </div>
  <div class="tags">{tags_html}</div>
  <div class="summary">{esc(r.summary) or '<i style="color:#94a3b8">本文を抽出できませんでした</i>'}</div>
  <div class="filepath">📁 {esc(r.relpath)}</div>
  {reason_html}
</div>""")

    html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>NoticeForge 処理レポート</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Meiryo UI','Yu Gothic UI','Hiragino Sans',sans-serif;background:#f1f5f9;color:#1e293b;font-size:14px}}
/* ─── ヘッダー ─── */
.header{{background:linear-gradient(135deg,#1e40af,#2563eb);color:white;padding:24px 32px;display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:8px}}
.header h1{{font-size:22px;font-weight:bold}}
.header .sub{{opacity:.75;font-size:13px;margin-top:4px}}
/* ─── 統計バー ─── */
.stats-bar{{background:white;border-bottom:1px solid #e2e8f0;padding:16px 32px;display:flex;gap:12px;flex-wrap:wrap;align-items:center}}
.stat-box{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px 20px;text-align:center;min-width:100px}}
.stat-box .num{{font-size:26px;font-weight:bold;color:#1e40af}}
.stat-box .lbl{{font-size:11px;color:#64748b;margin-top:2px}}
.stat-box.warn .num{{color:#dc2626}}
.stat-box.good .num{{color:#16a34a}}
.method-table{{margin-left:auto;font-size:12px;border-collapse:collapse}}
.method-table td{{padding:2px 8px;border-bottom:1px solid #f1f5f9}}
.method-table tr:last-child td{{border-bottom:none}}
/* ─── カード一覧 ─── */
.container{{max-width:1080px;margin:24px auto;padding:0 16px}}
/* ─── 検索バー ─── */
.search-bar{{background:white;padding:12px 32px;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100;box-shadow:0 2px 6px rgba(0,0,0,.06)}}
.search-input{{flex:1;max-width:680px;padding:10px 16px 10px 42px;border:2px solid #e2e8f0;border-radius:8px;font-size:14px;font-family:inherit;outline:none;transition:border-color .2s;background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='18' height='18' fill='none' stroke='%2394a3b8' stroke-width='2' viewBox='0 0 24 24'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 12px center}}
.search-input:focus{{border-color:#2563eb}}
.search-hint{{font-size:12px;color:#94a3b8}}
.search-count{{font-size:13px;color:#64748b;font-weight:bold;white-space:nowrap}}
.no-results{{text-align:center;padding:64px 16px;color:#94a3b8;font-size:15px;display:none}}
.card{{background:white;border-radius:10px;padding:18px 22px;margin-bottom:14px;border-left:5px solid #94a3b8;box-shadow:0 1px 4px rgba(0,0,0,.07);transition:box-shadow .2s}}
.card:hover{{box-shadow:0 3px 10px rgba(0,0,0,.12)}}
.card-ok{{border-left-color:#16a34a}}
.card-review{{border-left-color:#dc2626}}
.card-header{{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px}}
.card-title{{font-size:15px;font-weight:bold;color:#0f172a;line-height:1.5;flex:1}}
.ok-badge{{background:#dcfce7;color:#16a34a;border:1px solid #86efac;border-radius:6px;padding:2px 10px;font-size:12px;font-weight:bold;white-space:nowrap}}
.rev-badge{{background:#fee2e2;color:#dc2626;border:1px solid #fca5a5;border-radius:6px;padding:2px 10px;font-size:12px;font-weight:bold;white-space:nowrap}}
.meta{{display:flex;gap:14px;flex-wrap:wrap;color:#64748b;font-size:12px;margin-bottom:10px}}
.method-tag{{color:#94a3b8;font-size:11px}}
.tags{{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}}
.badge{{color:white;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:500}}
.summary{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:10px 14px;font-size:13px;line-height:1.75;color:#334155;max-height:150px;overflow-y:auto;margin-bottom:10px;white-space:pre-wrap}}
.filepath{{font-size:11px;color:#94a3b8;font-family:'Consolas','Courier New',monospace;word-break:break-all}}
.reason-box{{margin-top:8px;font-size:12px;color:#92400e;background:#fffbeb;border:1px solid #fde68a;border-radius:5px;padding:6px 12px}}
/* ─── フッター ─── */
.footer{{text-align:center;color:#94a3b8;font-size:11px;padding:24px;margin-top:8px}}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>NoticeForge 処理レポート</h1>
    <div class="sub">生成日時: {time.strftime('%Y年%m月%d日 %H:%M:%S')}</div>
  </div>
</div>
<div class="stats-bar">
  <div class="stat-box"><div class="num">{total}</div><div class="lbl">総ファイル数</div></div>
  <div class="stat-box good"><div class="num">{ok_count}</div><div class="lbl">正常抽出</div></div>
  <div class="stat-box warn"><div class="num">{needs_rev_count}</div><div class="lbl">要確認</div></div>
  <table class="method-table">
    <tr><td colspan="2" style="font-weight:bold;padding-bottom:4px">抽出方式別</td></tr>
    {method_rows}
  </table>
</div>
<div class="search-bar">
  <input class="search-input" id="searchInput" type="text"
    placeholder="キーワードで絞り込む（タイトル・発出者・ファイル名など。NotebookLMの引用文をそのまま貼り付けてもOK）"
    oninput="filterCards()">
  <span class="search-hint">→ 元ファイルを素早く探せます</span>
  <span class="search-count" id="searchCount"></span>
</div>
<div class="container">
{''.join(cards_html)}
  <div class="no-results" id="noResults">該当するファイルが見つかりませんでした。別のキーワードを試してください。</div>
</div>
<div class="footer">NoticeForge &mdash; NotebookLM 連携ツール</div>
<script>
function filterCards() {{
  var q = document.getElementById('searchInput').value.toLowerCase();
  var cards = document.querySelectorAll('.card');
  var shown = 0;
  cards.forEach(function(card) {{
    var text = card.getAttribute('data-search');
    var match = !q || text.includes(q);
    card.style.display = match ? '' : 'none';
    if (match) shown++;
  }});
  var countEl = document.getElementById('searchCount');
  var noRes   = document.getElementById('noResults');
  countEl.textContent = q ? (shown + ' 件 / ' + cards.length + ' 件中') : (cards.length + ' 件');
  noRes.style.display = (q && shown === 0) ? 'block' : 'none';
}}
window.addEventListener('load', function() {{
  document.getElementById('searchCount').textContent = document.querySelectorAll('.card').length + ' 件';
}});
</script>
</body>
</html>"""

    with open(os.path.join(outdir, "00_人間用レポート.html"), "w", encoding="utf-8") as f:
        f.write(html_content)


def process_folder(indir: str, outdir: str, cfg: Dict[str, object], progress_callback: Optional[Callable[[int, int, str, str], None]] = None) -> Tuple[int, int, str]:
    os.makedirs(outdir, exist_ok=True)

    # ① 前回の生成ファイルを削除（古いデータがNotebookLMに混入しないように）
    for fname in os.listdir(outdir):
        if fname.startswith("NotebookLM用_統合データ_") and fname.endswith(".txt"):
            try: os.remove(os.path.join(outdir, fname))
            except Exception: pass
    for fname in ("00_統合目次.md", "00_統合目次.xlsx", "00_人間用レポート.html", "00_処理ログ.txt"):
        p = os.path.join(outdir, fname)
        if os.path.exists(p):
            try: os.remove(p)
            except Exception: pass

    max_depth = int(cfg.get("max_depth", 30))
    split_kws = list(cfg.get("main_attach_split_keywords", []))
    min_chars = int(cfg.get("min_chars_mainbody", 800))
    use_ocr = bool(cfg.get("use_ocr", False))
    limit_bytes = int(cfg.get("bind_bytes_limit", 15000000))

    # システムファイル・一時ファイルを除外するフィルター
    SKIP_FILENAMES = frozenset({"thumbs.db", "desktop.ini", ".ds_store"})
    SKIP_EXTENSIONS = frozenset({".db", ".tmp", ".bak", ".lnk", ".ini", ".cache"})

    targets = [
        os.path.join(root, fn)
        for root, _, files in os.walk(indir)
        if os.path.relpath(root, indir).count(os.sep) < max_depth
        for fn in files
        if fn.lower() not in SKIP_FILENAMES
        and os.path.splitext(fn)[1].lower() not in SKIP_EXTENSIONS
        and not fn.startswith("~$")
    ]
    total_files = len(targets)
    records: List[Record] = []

    # ④ SHA1 重複検出用
    seen_sha1: set = set()
    skipped_dup = 0

    # ⑥ 処理ログ
    log_lines: List[str] = [
        "=== NoticeForge 処理ログ ===",
        f"処理日時: {time.strftime('%Y年%m月%d日 %H:%M:%S')}",
        f"入力フォルダ: {indir}",
        f"出力フォルダ: {outdir}",
        "",
        "--- 各ファイルの処理結果 ---",
    ]

    for i, path in enumerate(targets):
        rel = os.path.relpath(path, indir)
        ext = os.path.splitext(path)[1].lower()
        if progress_callback: progress_callback(i + 1, total_files, rel, "(抽出中...)")

        # ④ SHA1 重複チェック
        sha1 = compute_sha1(path)
        if sha1 and sha1 in seen_sha1:
            if progress_callback: progress_callback(i + 1, total_files, rel, "(重複ファイル・スキップ)")
            log_lines.append(f"[重複スキップ] {rel}")
            skipped_dup += 1
            continue
        if sha1:
            seen_sha1.add(sha1)

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
            elif ext == ".txt":                          # ③ .txt 対応
                text, method = extract_txt(path)
            elif ext == ".csv":                          # ③ .csv 対応
                text, method = extract_csv(path)
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

        log_lines.append(f"[{method}] {rel}")
        if reason:
            log_lines.append(f"  → {reason}")

        records.append(Record(relpath=rel, ext=ext, size=os.path.getsize(get_safe_path(path)), mtime=os.path.getmtime(get_safe_path(path)), sha1=sha1, method=method, pages=pages, text_chars=len(text), needs_review=needs_rev, reason=reason, title_guess=title, date_guess=date_guess, issuer_guess=issuer_guess, summary=summary, tags_facility=fac, tags_work=work, tag_evidence=ev, out_txt="", full_text_for_bind=payload))

    write_excel_index(outdir, records)
    write_md_indices(outdir, records)
    write_binded_texts(outdir, records, limit_bytes)
    write_html_report(outdir, records)

    # ⑥ サマリーを集計してログファイルに保存
    needs_rev_count = len([r for r in records if r.needs_review])
    review_breakdown: Dict[str, int] = {}
    for r in records:
        if r.needs_review:
            key = r.method if ("missing" in r.method or r.method in ("unhandled", "error")) else "本文が短すぎる"
            review_breakdown[key] = review_breakdown.get(key, 0) + 1

    log_lines += [
        "",
        "--- サマリー ---",
        f"総処理数: {len(records)} 件",
        f"正常抽出: {len(records) - needs_rev_count} 件",
        f"要確認: {needs_rev_count} 件",
    ]
    for k, v in sorted(review_breakdown.items(), key=lambda x: -x[1]):
        log_lines.append(f"  ・{k}: {v} 件")
    if skipped_dup:
        log_lines.append(f"重複スキップ: {skipped_dup} 件")

    with open(os.path.join(outdir, "00_処理ログ.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))

    # ⑥ GUI に渡す内訳文字列
    breakdown_str = "　".join(f"{k}: {v}件" for k, v in sorted(review_breakdown.items(), key=lambda x: -x[1]))
    return len(records), needs_rev_count, breakdown_str
