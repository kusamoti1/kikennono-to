# -*- coding: utf-8 -*-
"""
NoticeForge Core Logic v3.1 (GUI Optimized + Excel Index + LONG SUMMARY)
"""
from __future__ import annotations
import os, re, json, time, hashlib, csv
from dataclasses import dataclass, asdict
from typing import Dict, List, Tuple, Optional, Callable

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

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
    "main_attach_split_keywords": ["別添", "別紙", "参考", "添付", "（写）", "(写)", "【別添】", "【別紙】", "【参考】"],
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

def sha1_file(path: str) -> str:
    h = hashlib.sha1()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(2 * 1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return "000000"

def safe_mkdir(p: str):
    os.makedirs(p, exist_ok=True)

def safe_filename(base: str) -> str:
    s = re.sub(r"[^0-9A-Za-zぁ-んァ-ン一-龥\-\_]+", "_", base)[:110]
    return s or "doc"

def sniff_type(path: str) -> str:
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        if head.startswith(b"%PDF"):
            return "pdf"
        if head.startswith(b"PK\x03\x04"):
            return "zip"
    except Exception:
        pass
    return "unknown"

def extract_pdf(path: str) -> Tuple[str, Optional[int], str]:
    if not fitz:
        return "", None, "pymupdf_missing"
    try:
        doc = fitz.open(path)
        pages = doc.page_count
        parts = []
        for i in range(pages):
            parts.append(doc.load_page(i).get_text("text") or "")
        doc.close()
        return "\n".join(parts), pages, "pdf_text"
    except Exception as e:
        return "", None, f"pdf_err:{e.__class__.__name__}"

def extract_docx(path: str) -> Tuple[str, str]:
    if not Document:
        return "", "docx_missing"
    try:
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs if p.text]), "docx_text"
    except Exception as e:
        return "", f"docx_err:{e.__class__.__name__}"

def extract_xlsx(path: str) -> Tuple[str, str]:
    if not openpyxl:
        return "", "openpyxl_missing"
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        out = []
        for ws in wb.worksheets[:10]:
            out.append(f"## Sheet: {ws.title}")
            for row in ws.iter_rows(max_row=400, max_col=40, values_only=True):
                if any(row):
                    out.append(" | ".join(["" if c is None else str(c).strip() for c in row]))
            out.append("")
        wb.close()
        return "\n".join(out), "xlsx"
    except Exception as e:
        return "", f"xlsx_err:{e.__class__.__name__}"

def extract_text_file(path: str) -> Tuple[str, str]:
    for enc in ("utf-8", "utf-8-sig", "cp932", "shift_jis"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), f"text_{enc}"
        except Exception:
            continue
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(), "text_lossy"
    except Exception:
        return "", "text_read_fail"

def split_main_attach(text: str, kws: List[str]) -> Tuple[str, str]:
    idxs = []
    for k in kws:
        m = re.search(re.escape(k), text)
        if m:
            idxs.append(m.start())
    cut = min(idxs) if idxs else -1
    if cut > 200:
        return text[:cut].strip(), text[cut:].strip()
    return text.strip(), ""

def guess_title(text: str, fallback: str) -> str:
    for l in text.splitlines()[:50]:
        s = l.strip()
        if 6 <= len(s) <= 120 and not re.match(r"^[\d\-\s\(\)]+$", s):
            return s
    return fallback

def guess_date(text: str) -> str:
    m = re.search(r"(令和|平成)\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日", text)
    if m:
        return m.group(0)
    m2 = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", text)
    return m2.group(0) if m2 else ""

def guess_issuer(text: str) -> str:
    for cand in ["消防庁", "総務省消防庁", "消防局", "危険物保安室", "予防課"]:
        if cand in text:
            return cand
    return ""

def tag_text(text: str) -> Tuple[List[str], List[str], Dict[str, List[str]]]:
    ev: Dict[str, List[str]] = {}
    fac: List[str] = []
    work: List[str] = []
    target = text[:8000]
    for t, ps in FACILITY_TAGS.items():
        hits = [p for p in ps if re.search(p, target)]
        if hits:
            fac.append(t); ev[t] = hits[:3]
    for t, ps in WORK_TAGS.items():
        hits = [p for p in ps if re.search(p, target)]
        if hits:
            work.append(t); ev[t] = hits[:3]
    if not fac and re.search(r"危険物|消防法", target):
        fac.append("共通")
    return fac, work, ev

def make_summary(main_text: str, n: int) -> str:
    s = re.sub(r"\s+", " ", main_text.strip())
    return s[:n] + ("…" if len(s) > n else "")

def write_excel_index(outdir: str, records: List[Record]):
    if not openpyxl:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Index"

    headers = [
        "タイトル(推定)", "日付(推定)", "発出者(推定)",
        "施設タグ", "業務タグ", "needs_review", "理由", "概要(先頭)",
        "元ファイル(relpath)", "生成テキスト(out_txt)"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for r in records:
        ws.append([
            r.title_guess, r.date_guess, r.issuer_guess,
            " / ".join(r.tags_facility) if r.tags_facility else "",
            " / ".join(r.tags_work) if r.tags_work else "",
            "TRUE" if r.needs_review else "FALSE",
            r.reason,
            r.summary,
            r.relpath,
            r.out_txt,
        ])

    widths = [40, 16, 14, 24, 24, 12, 26, 80, 50, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    ws2 = wb.create_sheet("NeedsReview")
    ws2.append(headers)
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    for r in records:
        if r.needs_review:
            ws2.append([
                r.title_guess, r.date_guess, r.issuer_guess,
                " / ".join(r.tags_facility) if r.tags_facility else "",
                " / ".join(r.tags_work) if r.tags_work else "",
                "TRUE", r.reason, r.summary, r.relpath, r.out_txt
            ])
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(os.path.join(outdir, "00_統合目次.xlsx"))

def write_md_indices(outdir: str, records: List[Record]):
    def md_line(r: Record) -> str:
        return (
            f"- **{r.title_guess}**\n"
            f"  - 日付(推定): {r.date_guess} / 発出者(推定): {r.issuer_guess}\n"
            f"  - タグ: 施設=[{' / '.join(r.tags_facility)}] / 業務=[{' / '.join(r.tags_work)}]\n"
            f"  - needs_review: {'True' if r.needs_review else 'False'} / 理由: {r.reason}\n"
            f"  - 概要: {r.summary}\n"
            f"  - 元: `{r.relpath}`\n"
            f"  - テキスト: `{r.out_txt}`\n\n"
        )
    with open(os.path.join(outdir, "00_統合目次.md"), "w", encoding="utf-8") as f:
        f.write("# 統合目次（概要付き）\n\n")
        for r in records:
            f.write(md_line(r))

def process_folder(indir: str, outdir: str, cfg: Dict[str, object], progress_callback: Optional[Callable[[int, int, str], None]] = None) -> Tuple[int, int]:
    safe_mkdir(outdir)
    safe_mkdir(os.path.join(outdir, "docs_txt"))

    max_depth = int(cfg.get("max_depth", 30))
    split_kws = list(cfg.get("main_attach_split_keywords", DEFAULTS["main_attach_split_keywords"]))  # type: ignore
    min_chars = int(cfg.get("min_chars_mainbody", 800))
    summary_chars = int(cfg.get("summary_chars", 900))

    targets: List[str] = []
    for root, _, files in os.walk(indir):
        if os.path.relpath(root, indir).count(os.sep) >= max_depth:
            continue
        for fn in files:
            targets.append(os.path.join(root, fn))

    total_files = len(targets)
    records: List[Record] = []

    for i, path in enumerate(targets):
        rel = os.path.relpath(path, indir)
        if progress_callback:
            progress_callback(i + 1, total_files, rel)

        try:
            st = os.stat(path)
        except Exception:
            continue

        ext = os.path.splitext(path)[1].lower()
        sniff = sniff_type(path)
        sha1 = sha1_file(path)

        text = ""
        pages: Optional[int] = None
        method = "unhandled"
        reason = ""

        try:
            if ext == ".pdf" or sniff == "pdf":
                text, pages, method = extract_pdf(path)
            elif ext == ".docx":
                text, method = extract_docx(path)
            elif ext in (".xlsx", ".xlsm"):
                text, method = extract_xlsx(path)
            elif ext in (".txt", ".md", ".csv"):
                text, method = extract_text_file(path)
            elif ext in (".xdw", ".xbd"):
                method = "xdw_skipped"
                reason = "DocuWorksは原本確認が必要（本文抽出は未対応）"
            else:
                reason = f"未対応拡張子: {ext}"
        except Exception as e:
            method = "error"
            reason = f"抽出エラー: {e.__class__.__name__}"

        if not text and not reason and method != "xdw_skipped":
            reason = f"本文が空（抽出方式={method}）"

        main, attach = split_main_attach(text, split_kws)
        title = guess_title(main or text, os.path.basename(path))
        date_guess = guess_date(text)
        issuer_guess = guess_issuer(text)
        fac, work, ev = tag_text(main or text)

        needs_rev = False
        if method in ("unhandled", "error", "xdw_skipped") or reason.startswith("未対応拡張子"):
            needs_rev = True
        if not needs_rev and len(main or text) < min_chars:
            needs_rev = True
            reason = reason or "本文が短い（要確認）"

        summary = make_summary(main or text, summary_chars)

        stem = safe_filename(os.path.splitext(rel)[0])
        out_txt_rel = f"docs_txt/{stem}_{sha1[:6]}.txt"
        out_txt_abs = os.path.join(outdir, out_txt_rel)

        payload = []
        payload.append("# 文書メタデータ")
        payload.append(f"- 元ファイル: {rel}")
        payload.append(f"- タイトル(推定): {title}")
        if date_guess:
            payload.append(f"- 日付(推定): {date_guess}")
        if issuer_guess:
            payload.append(f"- 発出者(推定): {issuer_guess}")
        payload.append(f"- 抽出方式: {method}")
        payload.append(f"- needs_review: {'True' if needs_rev else 'False'}")
        if reason:
            payload.append(f"- 理由: {reason}")
        payload.append(f"- 施設タグ: {' / '.join(fac) if fac else '（なし）'}")
        payload.append(f"- 業務タグ: {' / '.join(work) if work else '（なし）'}")
        payload.append("")
        payload.append("# 概要（先頭）")
        payload.append(summary if summary else "(概要なし)")
        payload.append("")
        payload.append("# 本文")
        payload.append(main.strip() if main.strip() else "(本文抽出なし)")
        if attach.strip():
            payload.append("")
            payload.append("# 別添・別紙・参考")
            payload.append(attach.strip())

        try:
            with open(out_txt_abs, "w", encoding="utf-8") as f:
                f.write("\n".join(payload))
        except Exception as e:
            needs_rev = True
            reason = f"書き込み失敗: {e.__class__.__name__}"

        records.append(Record(
            relpath=rel, ext=ext, size=int(st.st_size), mtime=float(st.st_mtime),
            sha1=sha1, method=method, pages=pages,
            text_chars=len(text), needs_review=needs_rev, reason=reason,
            title_guess=title, date_guess=date_guess, issuer_guess=issuer_guess,
            summary=summary, tags_facility=fac, tags_work=work, tag_evidence=ev,
            out_txt=out_txt_rel
        ))

        time.sleep(0.001)

    write_excel_index(outdir, records)
    write_md_indices(outdir, records)

    needs = [r for r in records if r.needs_review]
    if needs:
        with open(os.path.join(outdir, "needs_review.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["relpath", "method", "reason", "title_guess", "out_txt"])
            for r in needs:
                w.writerow([r.relpath, r.method, r.reason, r.title_guess, r.out_txt])

    with open(os.path.join(outdir, "ledger.json"), "w", encoding="utf-8") as f:
        json.dump([asdict(r) for r in records], f, ensure_ascii=False, indent=2)

    return len(records), len(needs)
