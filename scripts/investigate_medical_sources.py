# -*- coding: utf-8 -*-
"""医療ケア情報の取得状況を調査（公表キャッシュ・サンプルHTML）"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
CACHE_DIR = ROOT / "data/公的データ/kaigo_kouhyou_cache"
SAMPLE_HTML = ROOT / "scripts/_kihon_sample.html"

sys.path.insert(0, str(ROOT / "scripts"))
from kaigo_kouhyou_client import KaigoKouhyouClient  # noqa: E402

MEDICAL_COLS = [
    "インスリン",
    "胃瘻",
    "経管栄養",
    "喀痰吸引",
    "バルーン",
    "褥瘡処置",
    "透析",
]

ADDON_KEYS = (
    "登録喀痰吸引等事業者",
    "経口移行加算",
    "経口維持加算",
    "看護体制加算",
    "療養食加算",
    "看取り介護加算",
)


def scan_cache() -> dict:
    n = 0
    feature_nonempty = 0
    shotei_positive = 0
    service_550 = 0
    keyword_hits: dict[str, int] = {c: 0 for c in MEDICAL_COLS}
    addon_in_cache = 0

    for path in sorted(CACHE_DIR.glob("*.json")):
        with open(path, encoding="utf-8") as f:
            rec = json.load(f)
        if rec.get("error"):
            continue
        n += 1
        if rec.get("service_cd") == "550":
            service_550 += 1
        tables = rec.get("tables") or {}
        if tables.get("feature"):
            feature_nonempty += 1
        merged = {}
        for part in ("kani", "kihon", "feature", "overview"):
            merged.update(tables.get(part) or {})
        blob = json.dumps(merged, ensure_ascii=False)
        for col in MEDICAL_COLS:
            if col in ("インスリン",):
                if re.search(r"インスリン|糖尿病", blob):
                    keyword_hits[col] += 1
            elif col == "胃瘻" and re.search(r"胃瘻|胃ろう", blob):
                keyword_hits[col] += 1
            elif col == "経管栄養" and re.search(r"経管|経鼻", blob):
                keyword_hits[col] += 1
            elif col == "喀痰吸引" and re.search(r"喀痰|吸引", blob):
                keyword_hits[col] += 1
            elif col == "バルーン" and re.search(r"バルーン|留置カテーテル", blob):
                keyword_hits[col] += 1
            elif col == "褥瘡処置" and re.search(r"褥瘡|床ずれ", blob):
                keyword_hits[col] += 1
            elif col == "透析" and "透析" in blob:
                keyword_hits[col] += 1
        for k in ADDON_KEYS:
            if k in blob:
                addon_in_cache += 1
                break
        v = merged.get("所定疾患施設療養費の算定回数（前年度）", "")
        m = re.search(r"(\d+)\s*回", str(v))
        if m and int(m.group(1)) > 0:
            shotei_positive += 1

    return {
        "facilities": n,
        "feature_nonempty": feature_nonempty,
        "service_550": service_550,
        "shotei_positive": shotei_positive,
        "addon_in_cache": addon_in_cache,
        "keyword_hits": keyword_hits,
    }


def test_parser_on_sample() -> dict:
    client = KaigoKouhyouClient()
    html = SAMPLE_HTML.read_text(encoding="utf-8", errors="replace")
    tables = client.parse_html_tables(html)
    katakansen = tables.get(
        "社会福祉士及び介護福祉士法第４８条の３に規定する登録喀痰吸引等事業者"
    )
    addons_ari = sum(1 for k, v in tables.items() if "加算" in k and v == "あり")
    addons_total = sum(1 for k in tables if "加算" in k)
    return {
        "table_keys": len(tables),
        "katakansen": katakansen,
        "addons_ari": addons_ari,
        "addons_total": addons_total,
    }


def main() -> None:
    print("=== 公表キャッシュ現状（パーサー修正前のJSON） ===")
    c = scan_cache()
    print(f"施設数: {c['facilities']}")
    print(f"feature 非空: {c['feature_nonempty']}")
    print(f"介護医療院(550): {c['service_550']}")
    print(f"所定疾患療養費>0回（老健等）: {c['shotei_positive']}")
    print(f"加算項目がキャッシュに存在: {c['addon_in_cache']} 件")
    print("キーワードのみヒット（テキスト記述）:")
    for col, cnt in c["keyword_hits"].items():
        print(f"  {col}: {cnt}")

    print("\n=== サンプルHTML + 修正後パーサー ===")
    if SAMPLE_HTML.is_file():
        p = test_parser_on_sample()
        print(f"抽出キー数: {p['table_keys']}")
        print(f"登録喀痰吸引等事業者: {p['katakansen']!r}")
        print(f"加算項目（あり）: {p['addons_ari']} / {p['addons_total']}")
    else:
        print("(scripts/_kihon_sample.html なし)")


if __name__ == "__main__":
    main()
