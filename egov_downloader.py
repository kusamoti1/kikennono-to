#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
egov_downloader.py - 外部APIを使わず、手動取り込みを行うツール

このスクリプトは、e-Gov など外部サイトへのアクセスを行いません。
取り込む法令・通知・資料は、ユーザーが自分で決めて管理します。

使い方（かんたん）:
  A) 選択式（CSV）
     python egov_downloader.py
     python egov_downloader.py --apply-csv 危険物法令/取り込み候補_記入用.csv

  B) まるごと上書き（初心者向け）
     python egov_downloader.py --import-all-dir "C:\資料フォルダ"

     ※ 既存の「危険物法令/法令」内は全削除してから、
        指定フォルダ内のファイルを再コピーします。
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

# NoticeForgeで扱うことが多い形式を対象
SUPPORTED_EXTS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt", ".csv", ".xdw", ".xbd"
}


def _write_guide(out_dir: Path) -> Path:
    guide = out_dir / GUIDE_FILE
    text = """【このフォルダの使い方】

このツールは外部ダウンロードをしません。
取り込むものは、あなたが決めます。

手順A（CSVで選ぶ）:
1. 「取り込み候補_記入用.csv」を開く
2. 取り込みたいファイルのパスを入力する
3. 取り込む列に 1 を入れる
4. 次のコマンドを実行する
   python egov_downloader.py --apply-csv 取り込み候補_記入用.csv

手順B（おすすめ：まるごと上書き）:
1. 取り込みたい資料を1つのフォルダに集める
2. 次のコマンドを実行する
   python egov_downloader.py --import-all-dir "C:\資料フォルダ"

実行後:
- 「危険物法令/法令」フォルダが作られます
- 手順Bでは、法令フォルダ内をいったん空にしてから再作成します
- この法令フォルダを NoticeForge の入力フォルダに置くと、
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


def _clear_directory(path: Path) -> None:
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
        return
    for child in path.iterdir():
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
        else:
            try:
                child.unlink()
            except FileNotFoundError:
                pass


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

            dst = dest_dir / f"{idx:03d}_{src.name}"
            shutil.copy2(src, dst)
            copied.append(str(dst))

    print("\n=== 取り込み結果（CSV選択）===")
    print(f"コピー完了: {len(copied)}件")
    for p in copied:
        print(f"  + {p}")

    if skipped:
        print(f"スキップ: {len(skipped)}件")
        for s in skipped:
            print(f"  - {s}")

    return copied


def import_all_overwrite(import_dir: Path, out_dir: Path) -> List[str]:
    """指定フォルダ内の対応ファイルを、法令フォルダへ全上書きコピーする。"""
    if not import_dir.exists() or not import_dir.is_dir():
        raise FileNotFoundError(f"取り込み元フォルダが見つかりません: {import_dir}")

    _, _, dest_dir = prepare_template(out_dir)
    _clear_directory(dest_dir)

    files = [
        p for p in import_dir.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
    ]
    # 「最新が最後」になるよう、更新日時が古い順で並べる
    files.sort(key=lambda p: p.stat().st_mtime)

    copied: List[str] = []
    used_names: set[str] = set()

    for idx, src in enumerate(files, start=1):
        rel = src.relative_to(import_dir)
        # フォルダ構造を潰して重複回避しやすい名前にする
        base_name = str(rel).replace("/", "_").replace("\\", "_")
        stem = f"{idx:03d}_{base_name}"
        candidate = stem
        c = 1
        while candidate.lower() in used_names:
            p = Path(stem)
            candidate = f"{p.stem}_{c}{p.suffix}"
            c += 1
        used_names.add(candidate.lower())

        dst = dest_dir / candidate
        shutil.copy2(src, dst)
        copied.append(str(dst))

    print("\n=== 取り込み結果（まるごと上書き）===")
    print(f"取り込み元: {import_dir}")
    print(f"コピー完了: {len(copied)}件")
    print("※ 既存の法令フォルダ内は上書き再作成しました")
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
    parser.add_argument(
        "--import-all-dir",
        help="指定フォルダ内の対応ファイルを全上書きで取り込みます（初心者向け）。",
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

    if args.import_all_dir:
        import_all_overwrite(Path(args.import_all_dir), out_dir)
        return

    if args.apply_csv:
        apply_csv(
            csv_path=Path(args.apply_csv),
            out_dir=out_dir,
            source_base=Path(args.source_base) if args.source_base else None,
        )


if __name__ == "__main__":
    main()
