#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
egov_downloader.py - e-Gov法令APIv2から消防法危険物関連法令をダウンロード

取得対象:
  1. 消防法 第三章（危険物）
  2. 消防法施行令（全文）
  3. 危険物の規制に関する政令（全文）
  4. 危険物の規制に関する規則（全文）
  5. 消防法施行規則（全文）

使い方:
  python egov_downloader.py                 # 出力先: 危険物法令/
  python egov_downloader.py --out 別フォルダ
  python egov_downloader.py --list-only      # ダウンロードせず一覧だけ表示

必要パッケージ:
  pip install requests
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import requests
except ImportError:
    print("ERROR: requests がインストールされていません。")
    print("  pip install requests")
    sys.exit(1)

# ─────────────────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────────────────
API_BASE  = "https://laws.e-gov.go.jp/api/2"
DELAY_SEC = 1.2   # APIレート制限対応（連続呼び出し間隔）
TIMEOUT   = 30    # HTTPタイムアウト（秒）

# 法令タイプコード（e-Gov API v2 の law_type パラメータ）
LAW_TYPE_ACT   = 2   # 法律
LAW_TYPE_CAB   = 3   # 政令（Cabinet Order）
LAW_TYPE_ORD   = 5   # 府省令（Ministerial Ordinance）

# ─────────────────────────────────────────────────────────
# ダウンロード対象定義
# exact_name: 検索結果から選ぶ際に完全一致を優先する名前
# chapter   : 特定章のみ抽出したい場合に指定（Noneで全文）
# ─────────────────────────────────────────────────────────
TARGETS: List[Dict[str, Any]] = [
    {
        "label":      "消防法（第三章 危険物）",
        "search":     "消防法",
        "law_type":   LAW_TYPE_ACT,
        "exact_name": "消防法",
        "chapter":    "第三章",
        "filename":   "01_消防法_第三章（危険物）.txt",
    },
    {
        "label":      "消防法施行令",
        "search":     "消防法施行令",
        "law_type":   LAW_TYPE_CAB,
        "exact_name": "消防法施行令",
        "chapter":    None,
        "filename":   "02_消防法施行令.txt",
    },
    {
        "label":      "危険物の規制に関する政令",
        "search":     "危険物の規制に関する政令",
        "law_type":   LAW_TYPE_CAB,
        "exact_name": "危険物の規制に関する政令",
        "chapter":    None,
        "filename":   "03_危険物の規制に関する政令.txt",
    },
    {
        "label":      "危険物の規制に関する規則",
        "search":     "危険物の規制に関する規則",
        "law_type":   LAW_TYPE_ORD,
        "exact_name": "危険物の規制に関する規則",
        "chapter":    None,
        "filename":   "04_危険物の規制に関する規則.txt",
    },
    {
        "label":      "消防法施行規則",
        "search":     "消防法施行規則",
        "law_type":   LAW_TYPE_ORD,
        "exact_name": "消防法施行規則",
        "chapter":    None,
        "filename":   "05_消防法施行規則.txt",
    },
]


# ─────────────────────────────────────────────────────────
# API クライアント
# ─────────────────────────────────────────────────────────
def _get(session: requests.Session, url: str, params: Optional[Dict] = None) -> Dict:
    """GETリクエストを送信してJSONを返す。失敗時は例外を発生させる。"""
    resp = session.get(url, params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def search_law(session: requests.Session, keyword: str, law_type: int) -> List[Dict]:
    """キーワードと法令タイプで法令を検索し、候補リストを返す。"""
    url = f"{API_BASE}/laws"
    params = {
        "keyword":   keyword,
        "law_type":  law_type,
        "offset":    0,
        "limit":     20,
    }
    data = _get(session, url, params)
    # APIレスポンス: {"laws": [...], "total_count": N}
    return data.get("laws", [])


def pick_law(candidates: List[Dict], exact_name: str) -> Optional[Dict]:
    """候補から exact_name に最も近い法令を返す。"""
    # 完全一致を優先
    for c in candidates:
        if c.get("law_name") == exact_name:
            return c
    # 前方一致
    for c in candidates:
        if c.get("law_name", "").startswith(exact_name):
            return c
    # 部分一致
    for c in candidates:
        if exact_name in c.get("law_name", ""):
            return c
    # マッチなし → 先頭を返す
    return candidates[0] if candidates else None


def fetch_law_text(session: requests.Session, law_id: str) -> str:
    """法令IDを指定して条文テキストを取得する。"""
    url = f"{API_BASE}/law_data/{law_id}"
    params = {"response_format": "json"}
    data = _get(session, url, params)
    # API v2レスポンス: {"law_data": {"law": {...}}}
    law_node = (
        data.get("law_data", {}).get("law")
        or data.get("law_data", {})
        or data
    )
    return _node_to_text(law_node)


# ─────────────────────────────────────────────────────────
# JSON ノード → テキスト変換
# ─────────────────────────────────────────────────────────
def _node_to_text(node: Any, depth: int = 0) -> str:
    """e-Gov API v2 の JSON ノードを再帰的に平文テキストに変換する。

    ノード構造（例）:
      {"tag": "Chapter", "attr": {"Num": "第三章"}, "children": [...]}
      {"tag": "Sentence", "children": ["テキスト文字列"]}
      "直接の文字列"
    """
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return "".join(_node_to_text(n, depth) for n in node)
    if not isinstance(node, dict):
        return str(node)

    tag      = node.get("tag", "")
    children = node.get("children", [])
    attr     = node.get("attr", {})

    # ── 構造タグに応じて見出しを付与 ──
    indent = "　" * depth

    if tag in ("Law",):
        # 法令ルートノード: タイトルは LawTitle 子要素にある
        return _node_to_text(children, depth)

    if tag == "LawBody":
        return _node_to_text(children, depth)

    if tag == "LawTitle":
        title_text = _node_to_text(children, depth)
        return f"{'=' * 60}\n{title_text}\n{'=' * 60}\n\n"

    if tag == "MainProvision":
        return _node_to_text(children, depth)

    # 章 (Chapter / Part)
    if tag in ("Part", "Chapter"):
        num   = attr.get("Num", "")
        title = _find_title_text(children)
        heading = f"\n{'─' * 50}\n{num}　{title}\n{'─' * 50}\n"
        return heading + _node_to_text(children, depth + 1)

    if tag in ("PartTitle", "ChapterTitle"):
        return ""   # 章見出しは上で処理済み

    # 節 (Section)
    if tag == "Section":
        num   = attr.get("Num", "")
        title = _find_title_text(children)
        return f"\n【{num}　{title}】\n" + _node_to_text(children, depth + 1)

    if tag == "SectionTitle":
        return ""

    # 款 (Subsection)
    if tag == "Subsection":
        num   = attr.get("Num", "")
        title = _find_title_text(children)
        return f"\n〔{num}　{title}〕\n" + _node_to_text(children, depth + 1)

    if tag == "SubsectionTitle":
        return ""

    # 条 (Article)
    if tag == "Article":
        num = attr.get("Num", "")
        return f"\n第{_arabic_to_kanji_article(num)}条\n" + _node_to_text(children, depth + 1)

    if tag == "ArticleTitle":
        title_text = _node_to_text(children, depth).strip()
        return f"（{title_text}）\n" if title_text else ""

    # 項 (Paragraph)
    if tag == "Paragraph":
        num  = attr.get("Num", "")
        text = _node_to_text(children, depth + 1).strip()
        prefix = f"{num}　" if num and num != "1" else ""
        return f"{indent}{prefix}{text}\n"

    if tag == "ParagraphNum":
        return ""   # 項番号は Paragraph 側で処理

    # 号 (Item)
    if tag == "Item":
        num  = attr.get("Num", "")
        text = _node_to_text(children, depth + 1).strip()
        return f"{indent}  {num}　{text}\n"

    if tag == "ItemTitle":
        return _node_to_text(children, depth).strip() + "　"

    # 文 (Sentence)
    if tag == "Sentence":
        return _node_to_text(children, depth)

    # 表 (Table)
    if tag == "Table":
        return "\n[表省略]\n"

    # 図 (Fig)
    if tag == "Fig":
        return "\n[図省略]\n"

    # 付則 (Suppl)
    if tag == "SupplProvision":
        return f"\n\n【附　則】\n" + _node_to_text(children, depth)

    # 別表 (AppdxTable)
    if tag in ("AppdxTable", "Appdx", "AppdxStyle"):
        num = attr.get("Num", "")
        return f"\n\n【別表{num}】\n" + _node_to_text(children, depth)

    # その他: 子要素を再帰処理
    return _node_to_text(children, depth)


def _find_title_text(children: Any) -> str:
    """子ノードリストから *Title タグの文字列を探す。"""
    if not isinstance(children, list):
        return ""
    for child in children:
        if isinstance(child, dict) and child.get("tag", "").endswith("Title"):
            return _node_to_text(child.get("children", "")).strip()
    return ""


def _arabic_to_kanji_article(num_str: str) -> str:
    """アラビア数字の条番号を漢数字に変換する（簡易版）。
    例: '3' → '三'、'14' → '十四'"""
    try:
        n = int(num_str)
    except (ValueError, TypeError):
        return num_str
    digits = {1:"一",2:"二",3:"三",4:"四",5:"五",
              6:"六",7:"七",8:"八",9:"九"}
    if 1 <= n <= 9:
        return digits[n]
    if 10 <= n <= 99:
        tens  = n // 10
        units = n % 10
        s = ("" if tens == 1 else digits.get(tens, str(tens))) + "十"
        s += digits.get(units, "") if units else ""
        return s
    if 100 <= n <= 999:
        hundreds = n // 100
        rest     = n % 100
        s = ("" if hundreds == 1 else digits.get(hundreds, str(hundreds))) + "百"
        if rest >= 10:
            tens  = rest // 10
            units = rest % 10
            s += ("" if tens == 1 else digits.get(tens, str(tens))) + "十"
            s += digits.get(units, "") if units else ""
        elif rest > 0:
            s += digits.get(rest, str(rest))
        return s
    return num_str


# ─────────────────────────────────────────────────────────
# 章フィルタ
# ─────────────────────────────────────────────────────────
def extract_chapter(full_text: str, chapter_label: str) -> str:
    """平文テキストから特定章を抽出する。

    章見出し行から次の章見出し行（またはEOF）までを返す。
    """
    # 章見出し行のパターン（`─` 区切り行 + 章名 + `─` 区切り行）
    sep = "─" * 50
    pattern = re.compile(
        rf"({re.escape(sep)}\n(?:[^\n]*\n)*?(?:[^\n]*{re.escape(chapter_label)}[^\n]*\n)(?:[^\n]*\n)*?{re.escape(sep)}\n"
        rf".*?)(?={re.escape(sep)}|\Z)",
        re.DOTALL,
    )
    m = pattern.search(full_text)
    if m:
        return m.group(1).strip()

    # フォールバック: 章番号の行から次の章区切りまで
    lines = full_text.splitlines(keepends=True)
    start = None
    result: List[str] = []
    for i, line in enumerate(lines):
        if chapter_label in line and sep in (lines[i-1] if i > 0 else ""):
            start = i - 1   # ─ 区切り行から開始
        if start is not None:
            result.append(line)
            # 次の章が始まったら終了
            if i > start + 3 and line.strip() == sep and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                # 次行が章名らしければ終了
                if re.match(r"第[一二三四五六七八九十百]+章", next_line):
                    break
    if result:
        return "".join(result).strip()

    # どのフォールバックもマッチしなければ全文を返す
    return full_text


# ─────────────────────────────────────────────────────────
# メイン処理
# ─────────────────────────────────────────────────────────
def run(out_dir: Path, list_only: bool = False) -> None:
    session = requests.Session()
    session.headers.update({
        "Accept":     "application/json",
        "User-Agent": "egov-downloader/1.0 (危険物法令収集スクリプト)",
    })

    if not list_only:
        out_dir.mkdir(parents=True, exist_ok=True)
        print(f"出力先: {out_dir.resolve()}\n")

    for target in TARGETS:
        label     = target["label"]
        search    = target["search"]
        law_type  = target["law_type"]
        exact     = target["exact_name"]
        chapter   = target["chapter"]
        filename  = target["filename"]

        print(f"[検索] {label} ...", end=" ", flush=True)

        try:
            candidates = search_law(session, search, law_type)
        except requests.RequestException as e:
            print(f"\n  ERROR (検索): {e}")
            continue

        time.sleep(DELAY_SEC)

        if not candidates:
            print(f"\n  WARNING: 検索結果なし（keyword={search!r}, law_type={law_type}）")
            continue

        law = pick_law(candidates, exact)
        if law is None:
            print(f"\n  WARNING: 候補から選択できませんでした")
            continue

        law_id   = law.get("law_id", "")
        law_name = law.get("law_name", "")
        law_num  = law.get("law_num", "")
        print(f"→ {law_name}（{law_num}）[{law_id}]")

        if list_only:
            continue

        # 条文取得
        print(f"  [取得] 条文ダウンロード中 ...", end=" ", flush=True)
        try:
            full_text = fetch_law_text(session, law_id)
        except requests.RequestException as e:
            print(f"\n  ERROR (条文取得): {e}")
            continue

        time.sleep(DELAY_SEC)

        # 章フィルタ
        if chapter:
            print(f"→ {chapter} を抽出中 ...", end=" ", flush=True)
            text = extract_chapter(full_text, chapter)
            if text == full_text:
                print(f"※ {chapter} が見つからず全文を保存")
            else:
                print("OK")
        else:
            text = full_text
            print("OK（全文）")

        # 保存
        out_path = out_dir / filename
        header = (
            f"法令名: {law_name}\n"
            f"法令番号: {law_num}\n"
            f"法令ID: {law_id}\n"
            f"出典: e-Gov法令APIv2 (https://laws.e-gov.go.jp/api/2)\n"
            f"{'=' * 60}\n\n"
        )
        out_path.write_text(header + text, encoding="utf-8")
        size_kb = out_path.stat().st_size / 1024
        print(f"  [保存] {out_path.name}  ({size_kb:.1f} KB)")

    if not list_only:
        print(f"\n完了: {out_dir.resolve()}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="e-Gov法令APIv2から消防法危険物関連法令をダウンロードする",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--out", "-o",
        default="危険物法令",
        help="出力フォルダ（デフォルト: 危険物法令/）",
    )
    parser.add_argument(
        "--list-only", "-l",
        action="store_true",
        help="ダウンロードせずに法令IDの一覧だけ表示する",
    )
    args = parser.parse_args()

    run(Path(args.out), list_only=args.list_only)


if __name__ == "__main__":
    main()
