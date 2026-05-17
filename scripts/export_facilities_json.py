# -*- coding: utf-8 -*-
"""施設データ.xlsx → data/facilities.json（index.html と同じ列マッピング）"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "施設データ.xlsx"
OUT = ROOT / "data" / "facilities.json"

TYPE_CODE_MAP = {
    "特別養護老人ホーム": "tokuyou",
    "介護老人保健施設": "roken",
    "介護医療院": "iryoin",
    "サービス付き高齢者向け住宅": "sahoju",
    "特定施設入居者生活介護": "tokuteishisetsu",
    "グループホーム": "group",
    "その他": "other",
}


def _cell(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def row_to_facility(r: list) -> dict | None:
    if not r or not r[0] or not r[1]:
        return None
    care_level_keys: list[str] = []
    if _cell(r[3]) == "○":
        care_level_keys.append("j1")
    if _cell(r[4]) == "○":
        care_level_keys.append("j2")
    for i in range(5):
        if _cell(r[5 + i]) == "○":
            care_level_keys.append(str(i + 1))

    room_types: list[str] = []
    if _cell(r[10]) == "○":
        room_types.append("個室")
    if _cell(r[11]) == "○":
        room_types.append("多床室")

    stage_keys = ["1", "2", "3-1", "3-2", "4"]
    limit_stages = [stage_keys[i] for i in range(5) if _cell(r[20 + i]) == "○"]

    strengths = [s.strip() for s in re.split(r"[,、・]", _cell(r[34])) if s.strip()]

    def med(i: int) -> str:
        v = _cell(r[i]) if i < len(r) else ""
        return v if v else "不明"

    ftype = _cell(r[2])
    return {
        "id": int(float(r[0])) if str(r[0]).replace(".", "", 1).isdigit() else r[0],
        "name": _cell(r[1]),
        "type": ftype,
        "type_code": TYPE_CODE_MAP.get(ftype, "other"),
        "area": _cell(r[25]),
        "address": _cell(r[26]),
        "tel": _cell(r[27]),
        "fax": _cell(r[28]),
        "care_level_keys": care_level_keys,
        "room_types": room_types,
        "medical_care": {
            "insulin": med(12),
            "gastrostomy": med(13),
            "tube_feeding": med(14),
            "suction": med(15),
            "balloon": med(16),
            "pressure_sore": med(17),
            "dialysis": med(18),
        },
        "economics": {
            "seikohogo": _cell(r[19]) == "○",
            "limit_stages": limit_stages,
            "monthly_fee_note": _cell(r[35]),
        },
        "insider_info": {
            "flexibility": _cell(r[31]) or "不明",
            "vacancy_trend": _cell(r[32]) or "不明",
            "meeting_frequency": _cell(r[33]) or "不明",
            "strengths": strengths,
            "staff_name": _cell(r[29]),
            "staff_note": _cell(r[30]),
            "notes": _cell(r[36]),
        },
    }


def main() -> int:
    import openpyxl

    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["施設一覧"]
    facilities: list[dict] = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        r = list(row)
        f = row_to_facility(r)
        if f:
            facilities.append(f)
    wb.close()

    payload = {
        "exported_at": __import__("datetime").datetime.now().isoformat(timespec="seconds"),
        "source": str(XLSX.name),
        "count": len(facilities),
        "facilities": facilities,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"Wrote {len(facilities)} facilities → {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
