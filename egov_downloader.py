#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
egov_downloader.py - 外部APIを使わず、手動取り込み用の雛形を作るツール

このスクリプトは、e-Gov など外部サイトへのアクセスを行いません。
取り込む法令・通知・資料は、ユーザーが自分で決めて管理します。

使い方（かんたん）:
  1) 雛形を作る
     python egov_downloader.py

  2) 作られた CSV に、取り込みたいファイルパスを記入

  3) 記入した CSV を使って、選んだファイルだけを収集
     python egov_downloader.py --apply-csv 危険物法令/取り込み候補_記入用.csv
"""
from __future__ import annotations

import argparse
import csv
import shutil
from pathlib import Path
from typing import List, Tuple

GUIDE_FILE = "00_手動取り込みガイド.txt"
CSV_FILE = "取り込み候補_記入用.csv"
DEST_DIR = "法令"


def _write_guide(out_dir: Path) -> Path:
    guide = out_dir / GUIDE_FILE
    text = """【このフォルダの使い方】

このツールは外部ダウンロードをしません。
取り込むものは、あなたが決めます。

手順:
1. 「取り込み候補_記入用.csv」を開く
2. 取り込みたいファイルのパスを入力する
3. 取り込む列に 1 を入れる
4. 次のコマンドを実行する
   python egov_downloader.py --apply-csv 取り込み候補_記入用.csv

実行後:
- 選んだファイルだけが「法令」フォルダにコピーされます
- その「法令」フォルダを NoticeForge の入力フォルダ内に置けば、
  NotebookLM 用の指標（目次や要約）は、その内容を基準に作られます
"""
    guide.write_text(text, encoding="utf-8")
    return guide


def _write_csv_template(out_dir: Path) -> Path:
    csv_path = out_dir / CSV_FILE
    if csv_path.exists():
        return csv_path
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["取り込む(1/0)", "ファイルパス", "メモ"])
        w.writerow(["1", r"C:\資料\消防法施行令.pdf", "例"])
        w.writerow(["0", "", "使わない行は0のまま"])
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


def apply_csv(csv_path: Path, out_dir: Path, source_base: Path | None = None) -> List[str]:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSVが見つかりません: {csv_path}")

    _, _, dest_dir = prepare_template(out_dir)

    copied: List[str] = []
    skipped: List[str] = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for idx, row in enumerate(reader, start=1):
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

            dst = dest_dir / f"{idx:02d}_{src.name}"
            shutil.copy2(src, dst)
            copied.append(str(dst))

    print("\n=== 取り込み結果 ===")
    print(f"コピー完了: {len(copied)}件")
    for p in copied:
        print(f"  + {p}")

    if skipped:
        print(f"スキップ: {len(skipped)}件")
        for s in skipped:
            print(f"  - {s}")

    return copied


def main() -> None:
    parser = argparse.ArgumentParser(
        description="外部ダウンロードなしで、手動取り込み用フォルダを作成します。",
    )
    parser.add_argument(
        "--out", "-o",
        default="危険物法令",
        help="出力フォルダ（デフォルト: 危険物法令/）",
    )
    parser.add_argument(
        "--apply-csv",
        help="記入済みCSVを指定すると、取り込む=1のファイルだけを収集します。",
    )
    parser.add_argument(
        "--source-base",
        help="CSVの相対パスを解決する基準フォルダ（省略可）。",
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

    if args.apply_csv:
        apply_csv(
            csv_path=Path(args.apply_csv),
            out_dir=out_dir,
            source_base=Path(args.source_base) if args.source_base else None,
        )


if __name__ == "__main__":
    main()
