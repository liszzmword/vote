"""
정책·공약마당(policy.nec.go.kr) 후보자공약 크롤러
- 대상: 제9회 전국동시지방선거 (sgId=20260603, menuId=CNDDT25)
- 출력: pledges.json {huboId: {giho, party, name, sgg, subSgName, pdfUrl, fileTypeName, pdfPrvwYn, fileDispYn}}
- 재실행 안전: 기존 항목 갱신 (huboId 기준 dedupe)

NEC API 호출 흐름:
  1) initUCACommiment.do?menuId=CNDDT25  → 세션 부트스트랩
  2) initUCACommimentRegion.do (sgId, subSgId)  → 시도(wiwid) 목록
  3) initUCACommimentGu.do (sgId, subSgId, hRegionId)  → 시군구(guid) 목록
  4) initUCACommimentSgg.do (sgId, subSgId, hRegionId, hGuId)  → 선거구(sggid) 목록
  5) initUCACommimentList.do (전부) → 후보자 목록 (huboid, fileinfo 포함)

선거 종류별 drill 깊이:
  - 시도단위 (3:시도지사, 8:광역비례, 11:교육감) → Region 만
  - 구시군단위 (4:구시군장, 9:기초비례, 2:국회의원) → Region+Gu
  - 시도의원 (5) → Region+Gu+Sgg
  - 구시군의원 (6) → Region+Gu+Sgg
"""

from __future__ import annotations

import argparse
import json
import logging
import time
from pathlib import Path

import requests

BASE = "https://policy.nec.go.kr"
SG_ID = "20260603"
MENU_ID = "CNDDT25"
CDN_BASE = "https://cdn.nec.go.kr/policy_pdf/"
REFERER = f"{BASE}/plc/commiment/initUCACommiment.do?menuId={MENU_ID}"

# (subSgId, sgTypecode, drill_depth, name)
# NEC fnSearch 분기에 따라:
#   depth=1: Region 만 (시도지사 3, 광역비례 8, 교육감 11)
#   depth=3: Region → Gu → Sgg → List (그 외 모두)
ELECTIONS = [
    ("320260603", "3", 1, "시도지사선거"),
    ("420260603", "4", 3, "구시군의장선거"),
    ("520260603", "5", 3, "시도의회의원선거"),
    ("620260603", "6", 3, "구시군의회의원선거"),
    ("820260603", "8", 1, "광역의원비례대표선거"),
    ("920260603", "9", 3, "기초의원비례대표선거"),
    ("1120260603", "11", 1, "교육감선거"),
    ("220260603", "2", 3, "국회의원선거"),
]

DELAY = 0.08
TIMEOUT = 30

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("pledges")


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Referer": REFERER,
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "ko-KR,ko;q=0.9",
    })
    # bootstrap
    s.get(f"{BASE}/plc/commiment/initUCACommiment.do", params={"menuId": MENU_ID}, timeout=TIMEOUT)
    return s


def _post_json(s: requests.Session, path: str, data: dict) -> dict | None:
    """NEC가 application/json이 아닌 text/plain으로도 JSON을 돌려주는 경우가 있어
    Content-Type 무시하고 본문으로 판별. 에러 페이지는 HTML(<!DOCTYPE...)."""
    for attempt in range(3):
        try:
            r = s.post(f"{BASE}{path}", data=data, timeout=TIMEOUT)
            body = r.text.lstrip()
            if r.status_code == 200 and body.startswith(("{", "[")):
                return r.json()
            if r.status_code == 200 and body.startswith("<"):
                log.warning("error page from %s (len=%d)", path, len(r.text))
                return None
            log.warning("unexpected %s status=%s len=%d", path, r.status_code, len(r.text))
            return None
        except (requests.RequestException, ValueError) as e:
            log.warning("retry %d %s: %s", attempt + 1, path, e)
            time.sleep(0.5 * (attempt + 1))
    return None


def get_regions(s: requests.Session, sub_sg_id: str) -> list[dict]:
    j = _post_json(s, "/plc/commiment/initUCACommimentRegion.do",
                   {"sgId": SG_ID, "subSgId": sub_sg_id})
    if not j:
        return []
    # First entry can be a stub - filter
    return [r for r in j.get("regionlist", []) if r.get("wiwid")]


def get_gus(s: requests.Session, sub_sg_id: str, region_id: str) -> list[dict]:
    """region_id = wiwsidocode (시도 코드)"""
    j = _post_json(s, "/plc/commiment/initUCACommimentGu.do",
                   {"sgId": SG_ID, "subSgId": sub_sg_id,
                    "wiwsidocode": region_id, "sortYn": "Y"})
    if not j:
        return []
    return [r for r in j.get("gulist", []) if r.get("wiwid")]


def get_sggs(s: requests.Session, sub_sg_id: str, region_id: str, gu_id: str) -> list[dict]:
    """region_id=시도, gu_id=구시군(wiwid)"""
    j = _post_json(s, "/plc/commiment/initUCACommimentSgg.do",
                   {"sgId": SG_ID, "subSgId": sub_sg_id,
                    "wiwsidocode": region_id, "wiwid": gu_id, "sortYn": "Y"})
    if not j:
        return []
    return [r for r in j.get("sgglist", []) if r.get("sggid")]


def get_list(s: requests.Session, sub_sg_id: str, sg_typecode: str,
             region_id: str, gu_id: str = "", sgg_id: str = "",
             page_index: int = 1) -> dict | None:
    """List 호출 — fnSearch 동일 파라미터 셋."""
    data = {
        "sgId": SG_ID, "subSgId": sub_sg_id, "menuId": MENU_ID,
        "sgTypecode": sg_typecode,
        "hSgId": sub_sg_id,
        "hRegionId": region_id,
        "hGuId": gu_id,
        "hSggId": sgg_id,
        "pageIndex": str(page_index),
        "allSearchWord": "",
        "searchWord": "",
        "sortYn": "Y",
        "elecEndYn": "N",
        "phGuId": "",
    }
    return _post_json(s, "/plc/commiment/initUCACommimentList.do", data)


def parse_files_from_fileinfo(fileinfo: str | None) -> dict[str, str]:
    """fileinfo 예:
        '선거공보||20260603/PDF/PBINFO/1100/003_100157144_20260520_1.pdf'
        '선거공보||A.pdf||||1||HEIGHT||Y||00||01,선거공약서||||||0||HEIGHT||Y||||00,5대공약||B.pdf||..'
    파일종류 → CDN 경로 dict 반환. 경로가 비어있으면 미제출.
    """
    if not fileinfo:
        return {}
    out: dict[str, str] = {}
    for chunk in fileinfo.split(","):
        parts = chunk.split("||")
        if len(parts) < 2:
            continue
        ftype = parts[0].strip()
        fpath = parts[1].strip()
        if ftype and fpath:
            out[ftype] = fpath
    return out


def _paginate_list(s, sub_sg_id, sg_typecode, region_id, gu_id="", sgg_id=""):
    out = []
    page = 1
    while True:
        j = get_list(s, sub_sg_id, sg_typecode, region_id, gu_id=gu_id, sgg_id=sgg_id, page_index=page)
        if not j: break
        rows = j.get("list", [])
        if not rows: break
        out.extend(rows)
        total = int(j.get("totalCnt", 0))
        if page * 10 >= total: break
        page += 1
        time.sleep(DELAY)
    return out


def candidates_for_election(s: requests.Session, sub_sg_id: str, sg_typecode: str, depth: int) -> list[dict]:
    """depth=1: Region만, depth=3: Region→Gu→Sgg→List."""
    out: list[dict] = []
    regions = get_regions(s, sub_sg_id)
    log.info("  regions=%d", len(regions))
    for ri, region in enumerate(regions, 1):
        region_id = region["wiwid"]
        region_name = region.get("wiwname", "")
        time.sleep(DELAY)

        if depth == 1:
            rows = _paginate_list(s, sub_sg_id, sg_typecode, region_id)
            out.extend(rows)
            log.info("  [%s/%s] %s cumulative=%d", ri, len(regions), region_name, len(out))
            continue

        # depth == 3: 모든 Region → Gu → Sgg
        gus = get_gus(s, sub_sg_id, region_id)
        for gu in gus:
            gu_id = gu.get("wiwid")
            if not gu_id: continue
            time.sleep(DELAY)
            sggs = get_sggs(s, sub_sg_id, region_id, gu_id)
            for sgg in sggs:
                sgg_id = sgg["sggid"]
                rows = _paginate_list(s, sub_sg_id, sg_typecode, region_id, gu_id=gu_id, sgg_id=sgg_id)
                out.extend(rows)
                time.sleep(DELAY)
        log.info("  [%s/%s] %s cumulative=%d (gus=%d)", ri, len(regions), region_name, len(out), len(gus))
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="pledges.json")
    ap.add_argument("--only", help="election code (e.g. 3) to crawl just one")
    args = ap.parse_args()

    out_path = Path(args.out)
    db: dict[str, dict] = {}
    if out_path.exists():
        db = json.loads(out_path.read_text(encoding="utf-8"))
        log.info("loaded existing pledges.json (%d entries)", len(db))

    s = make_session()

    for sub_sg_id, type_code, depth, name in ELECTIONS:
        if args.only and args.only != type_code:
            continue
        log.info("=== %s (subSgId=%s, depth=%d) ===", name, sub_sg_id, depth)
        rows = candidates_for_election(s, sub_sg_id, type_code, depth)

        added = 0
        for row in rows:
            hubo_id = str(row.get("huboid", "") or "").strip()
            if not hubo_id:
                continue
            files = parse_files_from_fileinfo(row.get("fileinfo"))
            file_urls = {k: CDN_BASE + v for k, v in files.items()}
            db[hubo_id] = {
                "huboId": hubo_id,
                "subSgName": (row.get("subSgName") or name).strip(),
                "sggname": (row.get("sggname") or "").strip(),
                "wiwname": (row.get("wiwname") or "").strip(),
                "party": (row.get("jdname") or "").strip(),
                "giho": str(row.get("hbjgiho") or "").strip(),
                "name": (row.get("hbjname") or "").strip(),
                "files": file_urls,   # {선거공보: url, 선거공약서: url, 5대공약: url, ...}
                "fileDispYn": row.get("fileDispYn"),
                "openStatusCode": row.get("openStatusCode"),
            }
            added += 1
        log.info("  → %s rows merged (total db=%d)", added, len(db))

        # checkpoint after each election
        out_path.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")

    log.info("done. total %d entries → %s", len(db), out_path)


if __name__ == "__main__":
    main()
