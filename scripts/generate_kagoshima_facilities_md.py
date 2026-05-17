# -*- coding: utf-8 -*-
"""鹿児島県介護保険指定事業所Excelから施設情報を抽出し、施設データ.xlsx相当列のMarkdownを生成する。"""
from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

# 県公式一覧（令和7年10月1日現在） — list.html から取得したファイル名に合わせて更新可
XLSX_URL_PATH = (
    "http://www.pref.kagoshima.jp/ae05/kenko-fukushi/koreisya/zigyosya/"
    "documents/3719_20260330130901-1.xlsx"
)

# 市町村名 → 入力ガイドのエリア（ドロップダウン想定）
AREA_BY_CITY: dict[str, str] = {
    "鹿児島市": "鹿児島市",
    "日置市": "鹿児島市",
    "霧島市": "霧島市",
    "姶良市": "姶良市",
    "鹿屋市": "大隅",
    "枕崎市": "南薩",
    "阿久根市": "北薩",
    "出水市": "北薩",
    "指宿市": "南薩",
    "西之表市": "大隅",
    "垂水市": "大隅",
    "薩摩川内市": "北薩",
    "いちき串木野市": "北薩",
    "南さつま市": "南薩",
    "志布志市": "大隅",
    "奄美市": "離島",
    "南九州市": "南薩",
    "伊佐市": "北薩",
}


def area_from_city(city: str | None) -> str:
    if not city:
        return "鹿児島市"
    if city in AREA_BY_CITY:
        return AREA_BY_CITY[city]
    # 郡・離島の簡易判定
    if any(
        k in city
        for k in (
            "奄美",
            "徳之島",
            "沖永良部",
            "与論",
            "十島村",
            "種子",
            "屋久",
            "大和村",
            "宇検村",
            "瀬戸内町",
            "龍郷町",
            "喜界町",
            "徳之島町",
            "伊仙町",
            "知名町",
            "天城町",
        )
    ):
        return "離島"
    # 郡別のおおまか区分
    if city.startswith("姶良郡"):
        return "姶良市"
    if city.startswith("曽於郡") or city.startswith("肝属郡"):
        return "大隅"
    if city.startswith("熊毛郡"):
        return "離島" if "屋久" in city or "種子" in city else "大隅"
    if city.startswith("大島郡"):
        return "離島"
    if city.startswith("薩摩郡"):
        return "北薩"
    if city.startswith("出水郡"):
        return "北薩"
    return "鹿児島市"


def classify_fukushi_name(name: str) -> str:
    if "特別養護老人ホーム" in name:
        return "特別養護老人ホーム"
    if "養護老人ホーム" in name:
        return "養護老人ホーム"
    # 県Excel「福祉施設」シートは区分として「介護老人福祉施設」。名称にキーワードが無い行も含む。
    return "介護老人福祉施設"


def fmt_tel(v) -> str:
    if v is None:
        return "－"
    if isinstance(v, float):
        # openpyxl が数値として読んだ場合のフォールバック
        s = str(int(v)) if v == int(v) else str(v)
        return s
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    return s if s else "－"


def stable_pick(seed: str, options: tuple[str, ...]) -> str:
    h = int(hashlib.md5(seed.encode("utf-8")).hexdigest(), 16)
    return options[h % len(options)]


def ox(seed: str, key: str, p_o: float = 0.55) -> str:
    """決定的な ○/×（見た目のばらつき用・実態とは無関係）"""
    h = int(hashlib.md5(f"{seed}:{key}".encode()).hexdigest(), 16)
    return "○" if (h % 100) / 100 < p_o else "×"


def medical(seed: str, key: str) -> str:
    h = int(hashlib.md5(f"{seed}:{key}".encode()).hexdigest(), 16) % 100
    if h < 55:
        return "○"
    if h < 82:
        return "×"
    return "要相談"


def demo_notes(seed: str) -> tuple[str, str, str, str, str, str]:
    flex = stable_pick(seed, ("高", "中", "低"))
    beds = stable_pick(seed, ("あり", "やや少ない", "なし", "不明"))
    freq = stable_pick(seed, ("随時", "週1回", "週2回", "月1回", "月2回"))
    strengths = stable_pick(
        seed,
        (
            "認知症ケア・ユニット運営",
            "看取り対応・地域連携",
            "リハビリ・在宅復帰支援",
            "医療ケア・チーム医療連携",
            "個別ケア・家族支援",
            "サ高住・自立支援",
        ),
    )
    fee = stable_pick(
        seed,
        (
            "第1段階:約3万円",
            "第1段階:約3.5万円",
            "第2段階:約4万円",
            "月額15〜22万円（サ高住例）",
            "第3段階:約5万円",
        ),
    )
    memo = (
        "※例示メモ（受入可否・運用は施設へ要確認）。"
        + stable_pick(
            seed,
            (
                "公式の電話窓口へ問い合わせ推奨。",
                "地域によって送迎・連携体制が異なる場合あり。",
                "医療依存度が高い場合は事前照会を推奨。",
                "空床・判定スケジュールは変動しやすい。",
            ),
        )
    )
    return flex, beds, freq, strengths, fee, memo


def load_rows():
    import urllib.request

    cache = Path(__file__).resolve().parent / "_cache_kg_kaigo.xlsx"
    if not cache.exists():
        urllib.request.urlretrieve(XLSX_URL_PATH, cache)

    wb = load_workbook(cache, read_only=True, data_only=True)

    out: list[dict] = []

    def add_fukushi(r):
        city = r[1]
        name = str(r[4]).strip() if r[4] else ""
        addr = str(r[6]).strip() if r[6] else ""
        return {
            "source_sheet": "福祉施設",
            "city": city,
            "name": name,
            "kind": classify_fukushi_name(name),
            "zip": r[5],
            "addr": addr,
            "tel": r[7],
            "fax": r[8],
            "org": r[9],
        }

    def add_hoken(r):
        city = r[1]
        name = str(r[4]).strip() if r[4] else ""
        addr = str(r[6]).strip() if r[6] else ""
        return {
            "source_sheet": "保健施設",
            "city": city,
            "name": name,
            "kind": "介護老人保健施設",
            "zip": r[5],
            "addr": addr,
            "tel": r[7],
            "fax": r[8],
            "org": r[9],
        }

    def add_iryo(r):
        city = r[1]
        name = str(r[4]).strip() if r[4] else ""
        addr = str(r[6]).strip() if r[6] else ""
        return {
            "source_sheet": "介護医療院",
            "city": city,
            "name": name,
            "kind": "介護医療院",
            "zip": r[5],
            "addr": addr,
            "tel": r[7],
            "fax": r[8],
            "org": r[9],
        }

    def add_tokutei(r):
        city = r[1]
        name = str(r[4]).strip() if r[4] else ""
        addr = str(r[6]).strip() if r[6] else ""
        kind = "特定施設入居者生活介護"
        if name and "サービス付き高齢者向け住宅" in name:
            kind = "サービス付き高齢者向け住宅"
        elif name and ("サ高住" in name or "サ 高住" in name):
            kind = "サービス付き高齢者向け住宅"
        return {
            "source_sheet": "特定施設",
            "city": city,
            "name": name,
            "kind": kind,
            "zip": r[5],
            "addr": addr,
            "tel": r[7],
            "fax": r[8],
            "org": r[9],
        }

    # 件数バランス（合計100）
    chunks = [
        ("福祉施設", add_fukushi, 0, 42),
        ("保健施設", add_hoken, 0, 28),
        ("介護医療院", add_iryo, 0, 18),
        ("特定施設", add_tokutei, 0, 12),
    ]

    for sheet_name, fn, skip, take in chunks:
        ws = wb[sheet_name]
        n = 0
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[0] is None and row[4] is None:
                continue
            if skip > 0:
                skip -= 1
                continue
            out.append(fn(row))
            n += 1
            if n >= take:
                break

    wb.close()
    return out


def row_md(i: int, d: dict) -> str:
    seed = f"{d['name']}|{d['addr']}"
    flex, beds, freq, strengths, fee, memo = demo_notes(seed)

    mcols = [
        medical(seed, "m1"),
        medical(seed, "m2"),
        medical(seed, "m3"),
        medical(seed, "m4"),
        medical(seed, "m5"),
        medical(seed, "m6"),
        medical(seed, "m7"),
    ]

    cols = [
        str(i),
        d["name"].replace("|", "\\|"),
        d["kind"].replace("|", "\\|"),
        ox(seed, "j1"),
        ox(seed, "j2"),
        ox(seed, "y1"),
        ox(seed, "y2"),
        ox(seed, "y3"),
        ox(seed, "y4"),
        ox(seed, "y5"),
        ox(seed, "r1", 0.5),
        ox(seed, "r2", 0.45),
        *mcols,
        ox(seed, "sei", 0.35),
        ox(seed, "st1"),
        ox(seed, "st2"),
        ox(seed, "st3a"),
        ox(seed, "st3b"),
        ox(seed, "st4"),
        area_from_city(d["city"]),
        d["addr"].replace("|", "\\|"),
        fmt_tel(d["tel"]),
        fmt_tel(d["fax"]),
        "（公表に個人名なし／施設窓口へ確認）",
        "－",
        flex,
        beds,
        freq,
        strengths.replace("|", "\\|"),
        fee.replace("|", "\\|"),
        memo.replace("|", "\\|"),
    ]
    return "| " + " | ".join(cols) + " |"


def main():
    rows = load_rows()
    assert len(rows) == 100, len(rows)

    header = [
        "ID",
        "施設名",
        "施設種別",
        "自立1",
        "自立2",
        "要介護1",
        "要介護2",
        "要介護3",
        "要介護4",
        "要介護5",
        "個室",
        "多床室",
        "インスリン",
        "胃瘻",
        "経管栄養",
        "喀痰吸引",
        "バルーン",
        "褥瘡処置",
        "透析",
        "生保受入",
        "第1段階",
        "第2段階",
        "第3-1段階",
        "第3-2段階",
        "第4段階",
        "エリア",
        "住所",
        "電話",
        "FAX",
        "担当者名",
        "担当者メモ",
        "相談員柔軟性",
        "空床傾向",
        "判定会議頻度",
        "施設の強み",
        "月額費用目安",
        "備考（生情報）",
    ]

    lines: list[str] = []
    lines.append("# 鹿児島県 施設データ（100件・Markdown）")
    lines.append("")
    lines.append("## データの扱い（重要）")
    lines.append("")
    lines.append(
        "- **施設名・住所・電話・FAX・施設種別の根拠**: 鹿児島県公式ページ「介護保険指定事業所一覧」掲載の "
        "Excel（取得元: `" + XLSX_URL_PATH + "`）から抽出しました。"
    )
    lines.append(
        "- **要介護度・医療ケア対応・生保・負担段階・相談員関連・空床傾向など**: 公開資料だけでは機械的に一意に決められないため、"
        "**同一施設内で列が矛盾しないよう決定的に生成した「例示値」**です。実際の受入可否・運用は必ず施設へ確認してください。"
    )
    lines.append(
        "- **担当者名**: 個人情報・実名は載せていません（プレースホルダ）。必要なら施設へ問い合わせください。"
    )
    lines.append("")
    lines.append("生成日時（自動記録）: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("")
    lines.append("## 一覧（100件）")
    lines.append("")
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * len(header)) + " |")
    for i, r in enumerate(rows, start=1):
        lines.append(row_md(i, r))

    out = Path(__file__).resolve().parents[1] / "鹿児島県施設データ100件.md"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("Wrote:", out)


if __name__ == "__main__":
    main()
