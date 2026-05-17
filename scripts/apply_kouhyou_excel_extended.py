# -*- coding: utf-8 -*-
"""
介護サービス情報公表キャッシュ (kaigo_kouhyou_cache) から
施設データ.xlsx の未充足列を補充する。

※「Web検索」に相当するオンライン公表データ。LIFULL等の民間ポータルは別途。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
CACHE_DIR = ROOT / "data/公的データ/kaigo_kouhyou_cache"
XLSX_PATH = ROOT / "施設データ.xlsx"

# Excel列 ← 公表キャッシュから埋められる項目（整理表）
COLUMN_SOURCES = """
| Excel列 | 公表での主な出典 | 補充ルール |
|---------|------------------|------------|
| 個室/多床室 | 簡易・基本の居室・療養室 | 室数>0なら○ |
| 月額費用目安 | 食費・居住費の算定方法 | 公表テキストを要約（種別相場は上書き可） |
| 空床傾向 | 空き人数・更新日 | 空き>0→あり、0→なし、不明→不明 |
| 施設の強み | サービスの特色 | 先頭200字程度 |
| 担当者名 | 施設長・管理者 | 氏名のみ |
| 生保・負担段階 | 食費/居住費の段階説明 | 記載のある段階を○（空欄のみ） |
| 医療ケア各列 | 特色・所定疾患等の記述 | キーワード→○/要相談（空欄のみ） |
| 備考（生情報） | 待機者数・入所日数・URL | 末尾に追記 |
| 検索情報 | 公表URL・HP | 末尾に追記 |
| 相談員系 | — | 公表では取得不可（手入力） |
"""

ROOM_PRIVATE_KEYS = ("個室", "居室の状況 / 個室", "療養室の状況 / 個室")
ROOM_SHARED_KEYS = (
    "2人部屋",
    "3人部屋",
    "4人部屋",
    "5人部屋以上",
    "多床室",
    "療養室の数",
)

MEDICAL_KEYWORDS = {
    "インスリン": (r"インスリン|糖尿病", "要相談"),
    "胃瘻": (r"胃瘻|胃ろう|胃ロウ", "要相談"),
    "経管栄養": (r"経管栄養|経鼻経管|チューブ", "要相談"),
    "喀痰吸引": (r"喀痰|吸引", "要相談"),
    "バルーン": (r"バルーン|留置カテーテル", "要相談"),
    "褥瘡処置": (r"褥瘡|床ずれ", "要相談"),
    "透析": (r"透析", "要相談"),
}

STAGE_PATTERNS = {
    "第1段階": r"第[１1]段階",
    "第2段階": r"第[２2]段階",
    "第3-1段階": r"第[３3]段階[①1]|3段階[①1]",
    "第3-2段階": r"第[３3]段階[②2]|3段階[②2]",
    "第4段階": r"第[４4]段階",
}


def extract_office_number(text: str) -> str | None:
    m = re.search(r"事業所番号[:：]\s*(\d{10})", text or "")
    return m.group(1) if m else None


def _tables(rec: dict) -> dict[str, str]:
    t = rec.get("tables") or {}
    merged: dict[str, str] = {}
    for part in ("kani", "kihon", "feature", "overview"):
        merged.update(t.get(part) or {})
    return merged


def _blob(rec: dict) -> str:
    return json.dumps(_tables(rec), ensure_ascii=False)


def _text_blob(tables: dict[str, str]) -> str:
    """列名を含めず値のみ連結（キーワード誤検知防止）。"""
    return "\n".join(str(v) for v in tables.values() if v)


def _room_count(text: str) -> int:
    if not text:
        return 0
    m = re.search(r"(\d+)\s*室", text)
    if m:
        return int(m.group(1))
    m = re.search(r"療養室の数['\"]?\s*:\s*['\"]?(\d+)", text)
    if m:
        return int(m.group(1))
    parts = re.findall(r"(\d+)\s*/", text)
    if parts and "療養室" in text:
        return sum(int(x) for x in parts[:4])
    return 0


def infer_rooms(tables: dict[str, str]) -> tuple[str | None, str | None]:
    priv = 0
    shared = 0
    for k, v in tables.items():
        if any(pk in k for pk in ROOM_PRIVATE_KEYS) or k == "個室":
            priv += _room_count(v)
        if any(sk in k for sk in ROOM_SHARED_KEYS) or re.match(r"^[2-5]人", k):
            shared += _room_count(v)
    if "居室の数" in tables:
        nums = re.findall(r"(\d+)", tables["居室の数"])
        if len(nums) >= 2:
            priv = int(nums[0])
            shared = sum(int(n) for n in nums[1:5])
    return (
        "○" if priv > 0 else None,
        "○" if shared > 0 else None,
    )


def infer_monthly_fee(tables: dict[str, str]) -> str | None:
    parts = []
    for key in ("食費とその算定方法", "居住費とその算定方法"):
        v = tables.get(key, "").strip()
        if v:
            parts.append(f"{key[:2]}:{v[:120]}{'…' if len(v) > 120 else ''}")
    if not parts:
        return None
    return " / ".join(parts) + "（出典:介護サービス情報公表）"


def infer_vacancy(rec: dict) -> str | None:
    n = rec.get("free_num")
    if n is None:
        upd = rec.get("free_num_update")
        if upd and upd not in ("-", "－", "|", "ー"):
            return "不明"
        return None
    if n > 0:
        return "あり"
    return "なし"


def infer_strength(tables: dict[str, str]) -> str | None:
    t = (tables.get("サービスの特色") or tables.get("（その内容）") or "").strip()
    if not t:
        return None
    t = re.sub(r"\s+", " ", t)
    return (t[:200] + "…") if len(t) > 200 else t


def infer_manager(tables: dict[str, str]) -> str | None:
    v = tables.get("施設の管理者の氏名及び職名 / 氏名", "").strip()
    if v:
        return v
    return tables.get("法人等の代表者の<br />氏名及び職名 / 氏名", "").strip() or None


def infer_economics(blob: str) -> dict[str, str]:
    marks: dict[str, str] = {}
    if re.search(r"生活保護.*受給|生保.*受入|生活保護者", blob):
        marks["生保受入"] = "○"
    for col, pat in STAGE_PATTERNS.items():
        if re.search(pat, blob):
            marks[col] = "○"
    return marks


def _merge_marks(*parts: dict[str, str]) -> dict[str, str]:
    rank = {"○": 3, "要相談": 2, "×": 1}
    out: dict[str, str] = {}
    for part in parts:
        for col, val in part.items():
            if not val:
                continue
            if col not in out or rank.get(val, 0) > rank.get(out[col], 0):
                out[col] = val
    return out


def _table_ari(tables: dict[str, str], *needles: str) -> bool:
    for key, val in tables.items():
        if val == "あり" and any(n in key for n in needles):
            return True
    return False


def infer_medical(blob: str) -> dict[str, str]:
    marks: dict[str, str] = {}
    for col, (pat, default) in MEDICAL_KEYWORDS.items():
        if re.search(pat, blob, re.I):
            if re.search(r"対応可能|対応も可能|ケアも行|実施", blob):
                marks[col] = "○"
            else:
                marks[col] = default
    return marks


def infer_medical_from_kouhyou(tables: dict[str, str], rec: dict) -> dict[str, str]:
    """公表の加算・登録・施設種別から医療ケア列を推定（空欄のみ補充用）。"""
    marks: dict[str, str] = {}

    if tables.get(
        "社会福祉士及び介護福祉士法第４８条の３に規定する登録喀痰吸引等事業者"
    ) == "あり" or _table_ari(tables, "登録喀痰吸引"):
        marks["喀痰吸引"] = "○"

    if _table_ari(tables, "経口移行加算", "経口維持加算"):
        marks["経管栄養"] = "要相談"

    if _table_ari(tables, "看護体制加算"):
        marks.setdefault("インスリン", "要相談")
        marks.setdefault("褥瘡処置", "要相談")

    if _table_ari(tables, "療養食加算"):
        marks.setdefault("インスリン", "要相談")

    shotei = str(tables.get("所定疾患施設療養費の算定回数（前年度）", ""))
    m = re.search(r"(\d+)\s*回", shotei)
    if m and int(m.group(1)) > 0 and rec.get("service_cd") == "520":
        for col in ("胃瘻", "経管栄養", "喀痰吸引", "褥瘡処置", "インスリン"):
            marks.setdefault(col, "要相談")

    if rec.get("service_cd") == "550":
        for col in MEDICAL_KEYWORDS:
            marks.setdefault(col, "要相談")

    return marks


def infer_care_marks(residents: dict) -> dict[str, str]:
    col_map = {
        "1": "要介護1",
        "2": "要介護2",
        "3": "要介護3",
        "4": "要介護4",
        "5": "要介護5",
    }
    marks = {}
    for k, col in col_map.items():
        n = residents.get(k)
        if n is not None and n > 0:
            marks[col] = "○"
    return marks


def build_supplement(rec: dict) -> dict[str, str | None]:
    if rec.get("error"):
        return {}
    tables = _tables(rec)
    blob = _blob(rec)
    text_blob = _text_blob(tables)
    priv, shared = infer_rooms(tables)
    extra_note = []
    if tables.get("待機者数"):
        extra_note.append(f"待機者数:{tables['待機者数']}")
    if tables.get("入所者の平均的な入所日数") or tables.get(
        "入所者の平均的な入所日数<br>　※＜＞内の数値は都道府県平均"
    ):
        d = tables.get("入所者の平均的な入所日数") or tables.get(
            "入所者の平均的な入所日数<br>　※＜＞内の数値は都道府県平均"
        )
        extra_note.append(f"平均入所日数:{d}")
    hp = tables.get("ホームページ", "")
    if "http" in hp:
        url = [p.strip() for p in hp.split("/") if p.strip().startswith("http")]
        if url:
            extra_note.append(f"HP:{url[0]}")

    urls = rec.get("urls") or {}
    if urls.get("kani"):
        extra_note.append(f"公表:{urls['kani']}")

    out: dict[str, str | None] = {
        "個室": priv,
        "多床室": shared,
        "月額費用目安": infer_monthly_fee(tables),
        "空床傾向": infer_vacancy(rec),
        "施設の強み": infer_strength(tables),
        "担当者名": infer_manager(tables),
        "備考（生情報）": " / ".join(extra_note) if extra_note else None,
    }
    out.update(infer_economics(blob))
    out.update(
        _merge_marks(
            infer_medical(text_blob),
            infer_medical_from_kouhyou(tables, rec),
        )
    )
    out.update(infer_care_marks(rec.get("care_level_residents") or {}))
    return out


def _append_cell(existing: str | None, addition: str) -> str:
    ex = str(existing or "").strip()
    if not addition:
        return ex
    if addition in ex:
        return ex
    return (ex + " / " + addition).strip(" / ") if ex else addition


def apply(overwrite_fee: bool = False, overwrite_ox: bool = False) -> None:
    from openpyxl import load_workbook

    records: list[dict] = []
    for path in sorted(CACHE_DIR.glob("*.json")):
        with open(path, encoding="utf-8") as f:
            records.append(json.load(f))

    by_num: dict[str, dict] = {}
    for rec in records:
        if rec.get("error"):
            continue
        num = rec.get("office_number")
        if not num:
            continue
        if num not in by_num:
            by_num[num] = rec

    wb = load_workbook(XLSX_PATH)
    ws = wb["施設一覧"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    col = {h: i + 1 for i, h in enumerate(headers) if h}

    stats = {k: 0 for k in ["rows", "個室", "多床室", "月額", "空床", "強み", "担当", "経済", "医療", "備考", "介護"]}

    for r in range(4, ws.max_row + 1):
        info = str(ws.cell(r, col["検索情報"]).value or "")
        num = extract_office_number(info)
        if not num or num not in by_num:
            continue
        sup = build_supplement(by_num[num])
        if not sup:
            continue
        stats["rows"] += 1

        ox_cols = [
            "個室",
            "多床室",
            "生保受入",
            "第1段階",
            "第2段階",
            "第3-1段階",
            "第3-2段階",
            "第4段階",
            "インスリン",
            "胃瘻",
            "経管栄養",
            "喀痰吸引",
            "バルーン",
            "褥瘡処置",
            "透析",
            "要介護1",
            "要介護2",
            "要介護3",
            "要介護4",
            "要介護5",
        ]
        for c in ox_cols:
            val = sup.get(c)
            if not val or c not in col:
                continue
            cur = ws.cell(r, col[c]).value
            if overwrite_ox or not cur:
                ws.cell(r, col[c]).value = val
                if c.startswith("要介護"):
                    stats["介護"] += 1
                elif c in ("個室", "多床室"):
                    stats[c] += 1
                elif c in MEDICAL_KEYWORDS or c == "褥瘡処置":
                    stats["医療"] += 1
                else:
                    stats["経済"] += 1

        fee = sup.get("月額費用目安")
        if fee and "月額費用目安" in col:
            cur = ws.cell(r, col["月額費用目安"]).value
            generic = cur and "LIFULL" in str(cur)
            if overwrite_fee or not cur or generic:
                ws.cell(r, col["月額費用目安"]).value = fee
                stats["月額"] += 1

        for c, key in [("空床傾向", "空床"), ("施設の強み", "強み"), ("担当者名", "担当")]:
            val = sup.get(c)
            if val and c in col:
                cur = ws.cell(r, col[c]).value
                if not cur:
                    ws.cell(r, col[c]).value = val
                    stats[key] += 1

        if sup.get("備考（生情報）") and "備考（生情報）" in col:
            ws.cell(r, col["備考（生情報）"]).value = _append_cell(
                ws.cell(r, col["備考（生情報）"]).value, sup["備考（生情報）"]
            )
            stats["備考"] += 1

        if "検索情報" in col:
            kani_url = (by_num[num].get("urls") or {}).get("kani", "")
            if kani_url:
                ws.cell(r, col["検索情報"]).value = _append_cell(
                    ws.cell(r, col["検索情報"]).value, f"公表詳細:{kani_url}"
                )

    # 入力ガイドにオンライン補充マップを追記
    if "オンライン補充マップ" in wb.sheetnames:
        wg = wb["オンライン補充マップ"]
    else:
        wg = wb.create_sheet("オンライン補充マップ")
        wg["A1"] = "列名"
        wg["B1"] = "補充元"
        wg["C1"] = "備考"
        row = 2
        for line in COLUMN_SOURCES.strip().splitlines():
            if line.startswith("|") and not line.startswith("|-"):
                cells = [c.strip() for c in line.split("|")[1:-1]]
                if cells and cells[0] != "Excel列":
                    wg.cell(row, 1, cells[0])
                    wg.cell(row, 2, cells[1] if len(cells) > 1 else "")
                    wg.cell(row, 3, cells[2] if len(cells) > 2 else "")
                    row += 1

    out_path = XLSX_PATH
    try:
        wb.save(XLSX_PATH)
    except PermissionError:
        out_path = XLSX_PATH.with_name("施設データ_公表補充済み.xlsx")
        wb.save(out_path)
        print(f"※ {XLSX_PATH.name} が開かれているため {out_path.name} に保存しました。")
    print(f"更新行: {stats['rows']}")
    print(
        f"  個室{stats['個室']} 多床室{stats['多床室']} 月額{stats['月額']} "
        f"空床{stats['空床']} 強み{stats['強み']} 担当{stats['担当']}"
    )
    print(f"  経済・居室系{stats['経済']} 医療キーワード{stats['医療']} 要介護{stats['介護']} 備考追記{stats['備考']}")
    print(f"保存: {out_path}")
    print("シート「オンライン補充マップ」に列ごとの出典を追加しました。")


def main() -> int:
    ap = argparse.ArgumentParser(description="公表キャッシュから Excel 列を補充")
    ap.add_argument("--overwrite-fee", action="store_true", help="月額費用目安を公表で上書き")
    ap.add_argument("--overwrite-ox", action="store_true", help="○×列を上書き")
    args = ap.parse_args()
    apply(overwrite_fee=args.overwrite_fee, overwrite_ox=args.overwrite_ox)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
