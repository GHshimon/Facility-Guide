# -*- coding: utf-8 -*-
"""
介護サービス情報公表システムから事業所番号で詳細情報を取得する。

使用例:
  python scripts/fetch_kaigo_kouhyou.py --office-number 4670108317
  python scripts/fetch_kaigo_kouhyou.py --from-csv --limit 5
  python scripts/fetch_kaigo_kouhyou.py --from-xlsx --infer-care-levels

取得データは data/公的データ/kaigo_kouhyou_cache/ に JSON 保存。
--apply-xlsx で 施設データ.xlsx の要介護列・検索情報を更新可能（--infer-care-levels 必須）。
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

# 同ディレクトリのモジュール
sys.path.insert(0, str(Path(__file__).resolve().parent))
from kaigo_kouhyou_client import CSV_BASENAME_TO_SERVICE_CD, KaigoKouhyouClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
MHLW_DIR = ROOT / "data/公的データ/厚生労働省_介護サービス情報公表オープンデータ_2025年12月末"
CACHE_DIR = ROOT / "data/公的データ/kaigo_kouhyou_cache"
XLSX_PATH = ROOT / "施設データ.xlsx"

CARE_COLS = {
    "1": "要介護1",
    "2": "要介護2",
    "3": "要介護3",
    "4": "要介護4",
    "5": "要介護5",
    "j1": "自立1",
    "j2": "自立2",
}


def load_offices_from_csv(pref: str = "鹿児島県", limit: int | None = None) -> list[dict]:
    rows: list[dict] = []
    for basename, service_cd in CSV_BASENAME_TO_SERVICE_CD.items():
        path = MHLW_DIR / basename
        if not path.exists():
            continue
        with open(path, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                if row.get("都道府県名") != pref:
                    continue
                num = str(row.get("事業所番号", "")).strip()
                if not num:
                    continue
                rows.append(
                    {
                        "office_number": num,
                        "name": row.get("事業所名", ""),
                        "service_cd": service_cd,
                        "csv_file": basename,
                    }
                )
    # 同一事業所番号は service_cd ごとに別レコード（老健+特養など）
    if limit:
        rows = rows[:limit]
    return rows


def extract_office_number_from_search_info(text: str) -> str | None:
    m = re.search(r"事業所番号[:：]\s*(\d{10})", text or "")
    return m.group(1) if m else None


def load_offices_from_xlsx(limit: int | None = None) -> list[dict]:
    import openpyxl

    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb["施設一覧"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    rows: list[dict] = []
    for r in range(4, ws.max_row + 1):
        info = ws.cell(r, col["検索情報"]).value if col.get("検索情報") else ""
        num = extract_office_number_from_search_info(str(info or ""))
        if not num:
            continue
        rows.append(
            {
                "office_number": num,
                "name": ws.cell(r, col["施設名"]).value,
                "excel_row": r,
                "service_cd": None,
            }
        )
    wb.close()
    if limit:
        rows = rows[:limit]
    return rows


def save_cache(record: dict) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    num = record.get("office_number", "unknown")
    svc = record.get("service_cd") or ("err" if record.get("error") else "x")
    path = CACHE_DIR / f"{num}_{svc}.json"
    record["fetched_at"] = datetime.now(timezone.utc).isoformat()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
    return path


def infer_care_marks(residents: dict) -> dict[str, str]:
    """入所者数>0 の要介護度のみ ○（それ以外は空欄）。"""
    marks: dict[str, str] = {}
    for key, col in CARE_COLS.items():
        n = residents.get(key)
        if n is not None and n > 0:
            marks[col] = "○"
    return marks


def apply_to_xlsx(records: list[dict], infer: bool) -> None:
    import openpyxl
    from openpyxl import load_workbook

    wb = load_workbook(XLSX_PATH)
    ws = wb["施設一覧"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    by_num: dict[str, dict] = {}
    for rec in records:
        if rec.get("error"):
            continue
        key = (rec["office_number"], rec.get("service_cd"))
        by_num[key] = rec

    updated = 0
    for r in range(4, ws.max_row + 1):
        info = str(ws.cell(r, col["検索情報"]).value or "")
        num = extract_office_number_from_search_info(info)
        if not num:
            continue
        rec = None
        for (on, _), item in by_num.items():
            if on == num:
                rec = item
                break
        if not rec:
            continue

        note = f"公表取得:{rec.get('kohyobi', '')}"
        if note not in info:
            ws.cell(r, col["検索情報"]).value = (info + " / " + note).strip(" /")

        if infer and rec.get("care_level_residents"):
            marks = infer_care_marks(rec["care_level_residents"])
            for col_name, mark in marks.items():
                if col_name in col:
                    ws.cell(r, col[col_name]).value = mark
            updated += 1

    wb.save(XLSX_PATH)
    print(f"Excel更新: {updated} 行（要介護○は入所者実績>0のみ）→ {XLSX_PATH}")


def main() -> int:
    ap = argparse.ArgumentParser(description="介護サービス情報公表システムから事業所詳細を取得")
    ap.add_argument("--pref", default="46", help="都道府県コード（鹿児島=46）")
    ap.add_argument("--delay", type=float, default=1.5, help="リクエスト間隔（秒）")
    ap.add_argument("--office-number", help="事業所番号（10桁）")
    ap.add_argument("--service-cd", help="サービス種類コード（510=特養 等）")
    ap.add_argument("--from-csv", action="store_true", help="厚労省CSVの鹿児島県分を対象")
    ap.add_argument("--from-xlsx", action="store_true", help="施設データ.xlsx の検索情報から事業所番号を読む")
    ap.add_argument("--limit", type=int, help="処理件数上限")
    ap.add_argument("--skip-cached", action="store_true", help="既存JSONがある場合はスキップ")
    ap.add_argument("--apply-xlsx", action="store_true", help="施設データ.xlsx を更新")
    ap.add_argument(
        "--infer-care-levels",
        action="store_true",
        help="入所者実績>0 の要介護度を Excel で ○ にする（--apply-xlsx と併用）",
    )
    ap.add_argument(
        "--apply-cache-only",
        action="store_true",
        help="kaigo_kouhyou_cache の JSON のみで Excel 更新（Web取得しない）",
    )
    args = ap.parse_args()

    if args.apply_cache_only:
        if not args.apply_xlsx:
            ap.error("--apply-cache-only は --apply-xlsx と併用してください")
        results = []
        for path in sorted(CACHE_DIR.glob("*.json")):
            with open(path, encoding="utf-8") as f:
                results.append(json.load(f))
        print(f"キャッシュ読込: {len(results)} 件 → Excel 反映")
        apply_to_xlsx(results, infer=args.infer_care_levels)
        return 0

    if not args.office_number and not args.from_csv and not args.from_xlsx:
        ap.error("--office-number / --from-csv / --from-xlsx のいずれかを指定してください")

    targets: list[dict] = []
    if args.office_number:
        targets.append(
            {
                "office_number": args.office_number,
                "service_cd": args.service_cd,
                "name": "",
            }
        )
    elif args.from_csv:
        targets = load_offices_from_csv(limit=args.limit)
    elif args.from_xlsx:
        targets = load_offices_from_xlsx(limit=args.limit)

    client = KaigoKouhyouClient(pref_cd=args.pref, delay_sec=args.delay)
    results: list[dict] = []
    ok = err = skip = 0

    for i, t in enumerate(targets, 1):
        num = t["office_number"]
        svc = t.get("service_cd") or args.service_cd
        cache_glob = list(CACHE_DIR.glob(f"{num}_*.json")) if CACHE_DIR.exists() else []
        if args.skip_cached and cache_glob:
            with open(cache_glob[0], encoding="utf-8") as f:
                results.append(json.load(f))
            skip += 1
            print(f"[{i}/{len(targets)}] skip {num}")
            continue

        print(f"[{i}/{len(targets)}] fetch {num} ({t.get('name', '')[:20]}) service={svc or 'auto'}")
        try:
            rec = client.fetch_office(num, service_cd=svc)
        except Exception as e:
            rec = {"office_number": num, "error": str(e), "service_cd": svc}
            err += 1
        else:
            if rec.get("error"):
                err += 1
            else:
                ok += 1

        save_cache(rec)
        results.append(rec)
        residents = rec.get("care_level_residents") or {}
        if residents:
            print("  入所者実績:", {k: v for k, v in residents.items() if v})

    print(f"\n完了: 成功 {ok} / 失敗 {err} / スキップ {skip} / 合計 {len(targets)}")
    print(f"キャッシュ: {CACHE_DIR}")

    if args.apply_xlsx:
        apply_to_xlsx(results, infer=args.infer_care_levels)

    return 0 if err == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
