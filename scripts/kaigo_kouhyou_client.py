# -*- coding: utf-8 -*-
"""
介護サービス情報公表システム (kaigokensaku.mhlw.go.jp) クライアント

事業所番号で検索し、詳細ページ（概要・簡易・基本・特色）を取得して解析する。
利用規約・robots.txt を確認のうえ、適切な間隔でアクセスすること。
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from html.parser import HTMLParser
from http.cookiejar import CookieJar
from typing import Any

BASE_URL = "https://www.kaigokensaku.mhlw.go.jp"
USER_AGENT = "SoudaninApp/1.0 (+local research; contact project maintainer)"

# 厚労省オープンデータ CSV ファイル名 → ServiceCd
CSV_BASENAME_TO_SERVICE_CD: dict[str, str] = {
    "510_介護老人福祉施設.csv": "510",
    "520_介護老人保健施設.csv": "520",
    "540_地域密着型介護老人福祉施設入所者生活介護.csv": "540",
    "550_介護医療院.csv": "550",
    "331_特定施設入居者生活介護_有料老人ホーム.csv": "331",
    "332_特定施設入居者生活介護_軽費老人ホーム.csv": "332",
    "334_特定施設入居者生活介護_サービス付き高齢者向け住宅.csv": "334",
}

CARE_LEVEL_ABBR_MAP = {
    "要介護１": "1",
    "要介護２": "2",
    "要介護３": "3",
    "要介護４": "4",
    "要介護５": "5",
    "要支援１": "s1",
    "要支援２": "s2",
    "自立": "j",
    "自立１": "j1",
    "自立２": "j2",
}


@dataclass
class JigyosyoMatch:
    pref_cd: str
    jigyosyo_cd: str
    jigyosyo_sub_cd: str
    service_cd: str
    version_cd: str
    name: str
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def jigyosyo_cd_full(self) -> str:
        return f"{self.jigyosyo_cd}-{self.jigyosyo_sub_cd}"


class _TableParser(HTMLParser):
    """th[abbr] と同じ行の td テキストを収集する。"""

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[tuple[str, str]] = []
        self._in_tr = False
        self._in_th = False
        self._in_td = False
        self._cur_abbr = ""
        self._buf: list[str] = []
        self._th_parts: list[str] = []
        self._td_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        ad = {k: (v or "") for k, v in attrs}
        if tag == "tr":
            self._in_tr = True
            self._th_parts = []
            self._td_parts = []
        elif self._in_tr and tag == "th":
            self._in_th = True
            self._cur_abbr = ad.get("abbr", "")
            self._buf = []
        elif self._in_tr and tag == "td":
            self._in_td = True
            self._buf = []
        elif self._in_tr and tag == "img" and (self._in_td or self._in_th):
            alt = ad.get("alt", "").strip()
            if alt in ("あり", "なし"):
                self._buf.append(alt)

    def handle_endtag(self, tag: str) -> None:
        if tag == "th" and self._in_th:
            self._in_th = False
            if self._cur_abbr:
                self._th_parts.append(self._cur_abbr)
            self._cur_abbr = ""
        elif tag == "td" and self._in_td:
            self._in_td = False
            self._td_parts.append(self._normalize("".join(self._buf)))
        elif tag == "tr" and self._in_tr:
            self._in_tr = False
            self._flush_row()

    def handle_data(self, data: str) -> None:
        if self._in_th or self._in_td:
            self._buf.append(data)

    def _flush_row(self) -> None:
        if not self._td_parts:
            return
        if len(self._th_parts) == 1 and len(self._td_parts) == 1:
            self.rows.append((self._th_parts[0], self._td_parts[0]))
            return
        # 要介護度別入所者数: th=要介護N, td=人数
        if len(self._th_parts) == 1 and self._th_parts[0] in CARE_LEVEL_ABBR_MAP:
            self.rows.append((self._th_parts[0], self._td_parts[0]))
            return
        # 複合行はラベル: 値 を連結
        if self._th_parts and self._td_parts:
            label = " / ".join(self._th_parts)
            val = " / ".join(self._td_parts)
            self.rows.append((label, val))

    @staticmethod
    def _normalize(text: str) -> str:
        text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
        return text


class KaigoKouhyouClient:
    def __init__(self, pref_cd: str = "46", delay_sec: float = 1.5) -> None:
        self.pref_cd = pref_cd
        self.delay_sec = delay_sec
        self._cj = CookieJar()
        self._opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self._cj))
        self._opener.addheaders = [("User-Agent", USER_AGENT)]
        self._session_ready = False

    def _init_session(self) -> None:
        """検索API利用前に一覧ページへアクセスし Cookie を初期化する。"""
        if self._session_ready:
            return
        base = f"{BASE_URL}/{self.pref_cd}/index.php"
        self._request(f"{base}?action_kouhyou_pref_search_list_list=true")
        self._sleep()
        self._request(f"{base}?action_kouhyou_pref_topjigyosyo_index=true")
        self._sleep()
        self._session_ready = True

    def _sleep(self) -> None:
        if self.delay_sec > 0:
            time.sleep(self.delay_sec)

    def _request(self, url: str, data: dict[str, str] | None = None) -> str:
        if data is not None:
            body = urllib.parse.urlencode(data).encode("utf-8")
            req = urllib.request.Request(
                url,
                data=body,
                headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"},
            )
        else:
            req = urllib.request.Request(url)
        try:
            with self._opener.open(req, timeout=60) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace") if e.fp else ""
            raise RuntimeError(f"HTTP {e.code} for {url}: {body[:200]}") from e

    def search_by_office_number(self, office_number: str) -> list[JigyosyoMatch]:
        """事業所番号でキーワード検索し、該当サービス一覧を返す。"""
        self._init_session()
        search_url = f"{BASE_URL}/{self.pref_cd}/index.php?action_kouhyou_pref_topjigyosyo_index=true"
        self._request(
            search_url,
            {
                "method": "search",
                "action_kouhyou_pref_topjigyosyo_index": "true",
                "PrefCd": self.pref_cd,
                "FromPage": "kaigoTopPage",
                "SearchConditions": "",
                "LatLng": "",
                "SearchKeyword": office_number.strip(),
                "KeywordConjunction": "0",
            },
        )
        self._sleep()

        list_url = f"{BASE_URL}/{self.pref_cd}/index.php?action_kouhyou_pref_search_list_list=true"
        qs = urllib.parse.urlencode(
            {
                "action_kouhyou_pref_search_search": "true",
                "method": "search",
                "p_count": 50,
                "p_offset": 0,
                "p_sort_name": "JigyosyoCd",
                "p_order": 0,
            }
        )
        raw = self._request(f"{list_url}&{qs}")
        self._sleep()
        payload = json.loads(raw)
        if payload.get("status") != "success":
            raise RuntimeError(f"検索API失敗: {payload.get('data', payload)}")

        matches: list[JigyosyoMatch] = []
        for item in payload.get("data") or []:
            if str(item.get("JigyosyoCd", "")).strip() != str(office_number).strip():
                continue
            matches.append(
                JigyosyoMatch(
                    pref_cd=str(item.get("PrefCd", self.pref_cd)),
                    jigyosyo_cd=str(item["JigyosyoCd"]),
                    jigyosyo_sub_cd=str(item.get("JigyosyoSubCd", "00")),
                    service_cd=str(item["ServiceCd"]),
                    version_cd=str(item["VersionCd"]),
                    name=str(item.get("JigyosyoName", "")),
                    raw=item,
                )
            )
        return matches

    def pick_match(
        self, matches: list[JigyosyoMatch], service_cd: str | None = None
    ) -> JigyosyoMatch | None:
        if not matches:
            return None
        if service_cd:
            for m in matches:
                if m.service_cd == service_cd:
                    return m
        # 施設系を優先（510, 520, 540, 550, 331-334）
        facility_codes = ("510", "520", "530", "540", "550", "331", "332", "334", "335", "336", "337")
        for code in facility_codes:
            for m in matches:
                if m.service_cd == code:
                    return m
        return matches[0]

    def detail_url(self, action: str, match: JigyosyoMatch, extra: dict[str, str] | None = None) -> str:
        params: dict[str, str] = {
            "JigyosyoCd": match.jigyosyo_cd_full,
            "ServiceCd": match.service_cd,
        }
        if extra:
            params.update(extra)
        q = urllib.parse.urlencode(params)
        return f"{BASE_URL}/{self.pref_cd}/index.php?{action}&{q}"

    def _detail_page_urls(self, match: JigyosyoMatch) -> dict[str, str]:
        ver = match.version_cd
        return {
            "overview": self.detail_url("action_kouhyou_detail_overview_index=true", match),
            "kani": self.detail_url(f"action_kouhyou_detail_{ver}_kani=true", match, {"Type": "1"}),
            "kihon": self.detail_url(f"action_kouhyou_detail_{ver}_kihon=true", match),
            "feature": self.detail_url("action_kouhyou_detail_feature_index=true", match),
        }

    def fetch_detail_pages(
        self, match: JigyosyoMatch, only: tuple[str, ...] | None = None
    ) -> dict[str, str]:
        """概要・簡易・基本・特色の HTML を取得。only 指定時はそのページのみ。"""
        pages = self._detail_page_urls(match)
        keys = only if only else tuple(pages.keys())
        html: dict[str, str] = {}
        for key in keys:
            url = pages.get(key)
            if not url:
                continue
            html[key] = self._request(url)
            self._sleep()
        return html

    def parse_html_tables(self, html: str) -> dict[str, str]:
        parser = _TableParser()
        parser.feed(html)
        return {k: v for k, v in parser.rows if k and v}

    def parse_care_level_residents(self, tables: dict[str, str]) -> dict[str, int | None]:
        """要介護度別入所者数などから人数を抽出（合計行は除外）。"""
        result: dict[str, int | None] = {
            "j1": None,
            "j2": None,
            "s1": None,
            "s2": None,
            "1": None,
            "2": None,
            "3": None,
            "4": None,
            "5": None,
        }

        for label, value in tables.items():
            if "要介護度別入所者数" in label or label == "入所者の人数":
                continue
            care_label = label.split(" / ")[-1].strip()
            if care_label in CARE_LEVEL_ABBR_MAP:
                key = CARE_LEVEL_ABBR_MAP[care_label]
                if key in result:
                    result[key] = _parse_person_count(value)

        # 入所者の人数テーブル（年齢別×要介護）の合計列を拾う
        if all(result[str(i)] is None for i in range(1, 6)):
            totals = _parse_nyukyo_matrix_totals(tables)
            for k, v in totals.items():
                if result.get(k) is None:
                    result[k] = v
        return result

    def build_record(
        self,
        office_number: str,
        match: JigyosyoMatch,
        html_pages: dict[str, str],
    ) -> dict[str, Any]:
        tables_kani = self.parse_html_tables(html_pages.get("kani", ""))
        tables_kihon = self.parse_html_tables(html_pages.get("kihon", ""))
        tables_feature = self.parse_html_tables(html_pages.get("feature", ""))
        tables_overview = self.parse_html_tables(html_pages.get("overview", ""))

        merged = {**tables_kihon, **tables_kani, **tables_feature, **tables_overview}
        residents = self.parse_care_level_residents({**tables_kani, **tables_kihon})

        free_num = _extract_free_num(html_pages.get("feature", ""), match.raw.get("FreeNum"))
        vacancy_text = _first_matching(merged, ("空き人数", "受け入れ可能人数", "空き数"))

        return {
            "office_number": office_number,
            "pref_cd": match.pref_cd,
            "jigyosyo_cd": match.jigyosyo_cd,
            "jigyosyo_sub_cd": match.jigyosyo_sub_cd,
            "service_cd": match.service_cd,
            "version_cd": match.version_cd,
            "name": match.name,
            "kohyobi": match.raw.get("Kohyobi"),
            "capacity": _parse_person_count(str(match.raw.get("CapacityNum") or ""))
            or _parse_person_count(merged.get("入所定員", "")),
            "total_users": _parse_person_count(str(match.raw.get("TotalUserNum") or "")),
            "free_num": free_num,
            "free_num_update": match.raw.get("FreeNumUpdateDate") or vacancy_text,
            "care_level_residents": residents,
            "tables": {
                "kani": tables_kani,
                "kihon": tables_kihon,
                "feature": tables_feature,
                "overview": tables_overview,
            },
            "urls": {
                "overview": self.detail_url("action_kouhyou_detail_overview_index=true", match),
                "kani": self.detail_url(
                    f"action_kouhyou_detail_{match.version_cd}_kani=true", match, {"Type": "1"}
                ),
                "kihon": self.detail_url(
                    f"action_kouhyou_detail_{match.version_cd}_kihon=true", match
                ),
                "feature": self.detail_url("action_kouhyou_detail_feature_index=true", match),
            },
            "source": "介護サービス情報公表システム https://www.kaigokensaku.mhlw.go.jp/",
            "note": (
                "care_level_residents は入所者実績（受入可能範囲の公式宣言ではない）。"
                "Excelの○×へ反映する場合は --infer-care-levels を明示すること。"
            ),
        }

    def fetch_office(
        self, office_number: str, service_cd: str | None = None
    ) -> dict[str, Any]:
        matches = self.search_by_office_number(office_number)
        match = self.pick_match(matches, service_cd)
        if not match:
            return {
                "office_number": office_number,
                "error": "not_found",
                "matches_count": 0,
            }
        html_pages = self.fetch_detail_pages(match)
        record = self.build_record(office_number, match, html_pages)
        record["all_services"] = [
            {
                "service_cd": m.service_cd,
                "version_cd": m.version_cd,
                "name": m.name,
                "short_name": m.raw.get("ShortName"),
            }
            for m in matches
        ]
        return record


def _parse_person_count(text: str) -> int | None:
    if not text or text in ("-", "―", "－", "|", "ー"):
        return None
    m = re.search(r"(\d+)\s*人", text)
    if m:
        return int(m.group(1))
    m = re.search(r"^(\d+)$", text.strip())
    return int(m.group(1)) if m else None


def _extract_free_num(feature_html: str, raw_free: Any) -> int | None:
    n = _parse_person_count(str(raw_free or ""))
    if n is not None:
        return n
    if not feature_html:
        return None
    m = re.search(r"空き数[^0-9]*(\d+)\s*人", feature_html)
    if m:
        return int(m.group(1))
    m = re.search(r"現在の空き数\s*(\d+)\s*人", feature_html)
    return int(m.group(1)) if m else None


def _first_matching(tables: dict[str, str], keys: tuple[str, ...]) -> str | None:
    for k in keys:
        if k in tables and tables[k]:
            return tables[k]
    return None


def _parse_nyukyo_matrix_totals(tables: dict[str, str]) -> dict[str, int | None]:
    """基本ページの年齢別入所者表から要介護1-5の合計らしき値を推定（簡易）。"""
    # 実装は kani の縦並び表を優先するため、ここでは未使用に近い
    return {}
