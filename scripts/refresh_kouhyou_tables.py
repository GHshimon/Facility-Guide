# -*- coding: utf-8 -*-
"""
既存 kaigo_kouhyou_cache の JSON について、公表の基本(kihon)・特色(feature)のみ再取得して tables を更新。

パーサー修正後に医療ケア関連の加算・喀痰吸引登録を反映するために使用。
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from kaigo_kouhyou_client import JigyosyoMatch, KaigoKouhyouClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
CACHE_DIR = ROOT / "data/公的データ/kaigo_kouhyou_cache"


def match_from_record(rec: dict) -> JigyosyoMatch | None:
    if rec.get("error") or not rec.get("jigyosyo_cd"):
        return None
    return JigyosyoMatch(
        pref_cd=str(rec.get("pref_cd", "46")),
        jigyosyo_cd=str(rec["jigyosyo_cd"]),
        jigyosyo_sub_cd=str(rec.get("jigyosyo_sub_cd", "00")),
        service_cd=str(rec["service_cd"]),
        version_cd=str(rec.get("version_cd", "")),
        name=str(rec.get("name", "")),
    )


def refresh_record(client: KaigoKouhyouClient, rec: dict, pages: tuple[str, ...]) -> dict:
    match = match_from_record(rec)
    if not match:
        return rec
    html_pages = client.fetch_detail_pages(match, only=pages)
    tables = rec.setdefault("tables", {})
    for page in pages:
        tables[page] = client.parse_html_tables(html_pages.get(page, ""))
    rec["tables_refreshed_at"] = datetime.now(timezone.utc).isoformat()
    return rec


def main() -> int:
    ap = argparse.ArgumentParser(description="公表キャッシュの tables を部分再取得")
    ap.add_argument("--pref", default="46")
    ap.add_argument("--delay", type=float, default=1.2)
    ap.add_argument("--limit", type=int)
    ap.add_argument(
        "--pages",
        default="kihon,feature",
        help="再取得するページ（kihon,feature,kani）",
    )
    args = ap.parse_args()
    pages = tuple(p.strip() for p in args.pages.split(",") if p.strip())

    client = KaigoKouhyouClient(pref_cd=args.pref, delay_sec=args.delay)
    paths = sorted(CACHE_DIR.glob("*.json"))
    if args.limit:
        paths = paths[: args.limit]

    ok = err = 0
    for i, path in enumerate(paths, 1):
        with open(path, encoding="utf-8") as f:
            rec = json.load(f)
        if rec.get("error"):
            continue
        num = rec.get("office_number", path.stem)
        print(f"[{i}/{len(paths)}] refresh {num} ({pages})")
        try:
            rec = refresh_record(client, rec, pages)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(rec, f, ensure_ascii=False, indent=2)
            ok += 1
        except Exception as e:
            print(f"  error: {e}")
            err += 1

    print(f"完了: 更新 {ok} / 失敗 {err}")
    return 0 if err == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
