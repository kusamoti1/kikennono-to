#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
egov_downloader.py - 外部APIを使わず、ユーザー選択資料を収集するツール

このスクリプトは、e-Gov など外部サイトへのアクセスを行いません。
取り込む資料は、ユーザーが選んだファイルだけです。

かんたん操作（おすすめ）:
  1) フォルダを丸ごと候補化（CSV自動作成）
     python egov_downloader.py --from-dir "C:\\資料フォルダ"

  2) 作られた CSV で「取り込む(1/0)」を調整して再実行
     python egov_downloader.py --apply-csv 危険物法令/取り込み候補_記入用.csv

最短で実行したい場合:
  python egov_downloader.py --from-dir "C:\\資料フォルダ" --auto-collect
"""
from __future__ import annotations

import argparse
import csv
import shutil
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

GUIDE_FILE = "00_手動取り込みガイド.txt"
CSV_FILE = "取り込み候補_記入用.csv"
PROFILE_FILE = "00_NotebookLM_指標メモ.txt"
DEST_DIR = "法令"

SUPPORTED_EXTS: Sequence[str] = (
    ".pdf", ".txt", ".doc", ".docx", ".xls", ".xlsx", ".xdw", ".xbd",
)


def _write_guide(out_dir: Path) -> Path:
    guide = out_dir / GUIDE_FILE
    text = """【このフォルダの使い方】

このツールは外部ダウンロードをしません。
取り込むものは、あなたが決めます。

おすすめ手順（かんたん）:
1. まず次を実行（候補CSVを自動作成）
   python egov_downloader.py --from-dir "あなたの資料フォルダ"
2. 「取り込み候補_記入用.csv」を開く
3. 取り込む列を 1（使わない行は 0）にする
4. 次を実行
   python egov_downloader.py --apply-csv 危険物法令/取り込み候補_記入用.csv

補足:
- もっと簡単にするなら、--auto-collect を付けるとCSV生成後すぐに収集します
- 収集されたファイルだけが NotebookLM の基準データになります
"""
    guide.write_text(text, encoding="utf-8")
    return guide


def _write_csv_template(out_dir: Path) -> Path:
    csv_path = out_dir / CSV_FILE
    if csv_path.exists():
        return csv_path
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["取り込む(1/0)", "ファイルパス", "文書タイプ(法令/通知/マニュアル)", "優先度(高/中/低)", "メモ"])
        w.writerow(["1", r"C:\資料\消防法施行令.pdf", "法令", "高", "例"])
        w.writerow(["0", "", "通知", "中", "使わない行は0のまま"])
    return csv_path


def prepare_template(out_dir: Path) -> Tuple[Path, Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / DEST_DIR
    dest.mkdir(parents=True, exist_ok=True)
    guide = _write_guide(out_dir)
    csv_path = _write_csv_template(out_dir)
    return guide, csv_path, dest


def _is_enabled(v: str) -> bool:
    return str(v).strip() in {"1", "true", "True", "TRUE", "yes", "YES", "y", "Y"}


def _detect_doc_type(path: Path) -> str:
    s = str(path).lower()
    if any(k in s for k in ("法令", "施行令", "施行規則", "政令", "規則", "法律")):
        return "法令"
    if any(k in s for k in ("通知", "通達", "事務連絡")):
        return "通知"
    return "マニュアル"


def _iter_files(base_dir: Path, recursive: bool) -> Iterable[Path]:
    if recursive:
        for p in base_dir.rglob("*"):
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                yield p
    else:
        for p in base_dir.glob("*"):
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                yield p


def seed_csv_from_dir(source_dir: Path, csv_path: Path, recursive: bool = True) -> int:
    files = sorted(_iter_files(source_dir, recursive=recursive))
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["取り込む(1/0)", "ファイルパス", "文書タイプ(法令/通知/マニュアル)", "優先度(高/中/低)", "メモ"])
        for p in files:
            w.writerow(["1", str(p), _detect_doc_type(p), "中", "自動候補"])
    return len(files)


def _write_notebooklm_profile(out_dir: Path, rows: List[dict]) -> Path:
    profile = out_dir / PROFILE_FILE
    selected = [r for r in rows if _is_enabled(r.get("取り込む(1/0)", "0"))]
    type_counts = {"法令": 0, "通知": 0, "マニュアル": 0}
    for r in selected:
        dtype = (r.get("文書タイプ(法令/通知/マニュアル)") or "").strip()
        if dtype in type_counts:
            type_counts[dtype] += 1

    lines = [
        "【NotebookLM 指標メモ】",
        "このメモは、取り込んだ資料の意図をNotebookLMに伝えるために使います。",
        "",
        "推奨プロンプト（そのまま貼り付け可）:",
        "- このプロジェクトでは、添付資料のみを根拠として回答してください。",
        "- 優先度が高の資料を優先し、次に中、低の順で判断してください。",
        "- 文章に矛盾がある場合は、原本PDFと法令ファイルを最優先してください。",
        "",
        f"収集件数: {len(selected)}件（法令: {type_counts['法令']} / 通知: {type_counts['通知']} / マニュアル: {type_counts['マニュアル']}）",
    ]
    profile.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return profile


def apply_csv(csv_path: Path, out_dir: Path, source_base: Path | None = None) -> List[str]:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSVが見つかりません: {csv_path}")

    _, _, dest_dir = prepare_template(out_dir)

    copied: List[str] = []
    skipped: List[str] = []
    rows: List[dict] = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for idx, row in enumerate(reader, start=1):
            rows.append(row)
            use = row.get("取り込む(1/0)", "0")
            raw_path = (row.get("ファイルパス") or "").strip().strip('"')
            if not _is_enabled(use) or not raw_path:
                continue

            src = Path(raw_path)
            if not src.is_absolute() and source_base:
                src = source_base / src

            if not src.exists() or not src.is_file():
                skipped.append(f"{idx}: 見つからないためスキップ -> {src}")
                continue

            dst = dest_dir / f"{idx:03d}_{src.name}"
            shutil.copy2(src, dst)
            copied.append(str(dst))

    profile = _write_notebooklm_profile(out_dir, rows)

    print("\n=== 取り込み結果 ===")
    print(f"コピー完了: {len(copied)}件")
    for p in copied:
        print(f"  + {p}")

    if skipped:
        print(f"スキップ: {len(skipped)}件")
        for s in skipped:
            print(f"  - {s}")

    print(f"指標メモを作成: {profile}")
    return copied


def main() -> None:
    parser = argparse.ArgumentParser(
        description="外部ダウンロードなしで、ユーザー選択資料を収集します。",
    )
    parser.add_argument("--out", "-o", default="危険物法令", help="出力フォルダ（デフォルト: 危険物法令/）")
    parser.add_argument("--apply-csv", help="記入済みCSVを指定すると、取り込む=1のファイルだけを収集します。")
    parser.add_argument("--source-base", help="CSVの相対パスを解決する基準フォルダ（省略可）。")
    parser.add_argument("--from-dir", help="このフォルダを走査してCSV候補を自動作成します。")
    parser.add_argument("--no-recursive", action="store_true", help="--from-dir 時にサブフォルダを走査しません。")
    parser.add_argument("--auto-collect", action="store_true", help="--from-dir 後に自動で収集まで実行します。")
    args = parser.parse_args()

    out_dir = Path(args.out)
    guide, csv_path, dest = prepare_template(out_dir)

    print(f"雛形を作成しました: {out_dir.resolve()}")
    print(f"- ガイド: {guide.name}")
    print(f"- 記入用CSV: {csv_path.name}")
    print(f"- 収集先フォルダ: {dest.name}")

    if args.from_dir:
        source_dir = Path(args.from_dir)
        if not source_dir.exists() or not source_dir.is_dir():
            raise NotADirectoryError(f"--from-dir がフォルダではありません: {source_dir}")
        count = seed_csv_from_dir(source_dir, csv_path, recursive=not args.no_recursive)
        print(f"候補CSVを更新: {csv_path}（{count}件）")

    if args.apply_csv:
        apply_csv(
            csv_path=Path(args.apply_csv),
            out_dir=out_dir,
            source_base=Path(args.source_base) if args.source_base else None,
        )
    elif args.auto_collect and args.from_dir:
        apply_csv(csv_path=csv_path, out_dir=out_dir)


if __name__ == "__main__":
    main()
