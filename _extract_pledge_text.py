"""
공약 PDF → 텍스트 추출 (Gemini Flash Lite multimodal OCR)
- 입력: pledges.json
- 출력: pledges_text.json  {huboId: extracted_text}
- 재실행 안전: 이미 추출된 huboId는 건너뜀
- 비용 (대략): 7,826 후보 × 평균 5p PDF, gemini-2.5-flash-lite 기준 ~$10-20

사용 전 설치: pip install google-genai
환경변수 또는 --key 옵션으로 GEMINI_API_KEY 지정.

⚠️ API 키 노출 주의: 키를 셸 히스토리에 남기지 않으려면
   export GEMINI_API_KEY=...  후에 옵션 없이 실행하는 게 안전.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("extract_text")

API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

PROMPT = """이 PDF는 한국 선거 후보자가 중앙선관위에 제출한 공약 자료(선거공보 또는 5대공약)입니다.
다음 규칙으로 본문 텍스트를 추출해 주세요:
1. 페이지 순서대로 본문 텍스트만 출력. 표지/페이지 번호 같은 무의미한 텍스트는 생략.
2. 공약 항목(분야, 제목, 세부내용)이 명확하면 그 구조를 유지하며 markdown 으로 정리.
3. 정치 슬로건/구호도 그대로 포함.
4. 표가 있으면 행마다 "- 항목: 값" 형식으로.
5. 추측하거나 요약하지 말고, 문서에 적힌 내용 그대로 옮길 것.
6. 한국어로 출력."""


def extract_one(api_key: str, model: str, hubo_id: str, pdf_url: str, timeout: int = 120) -> str | None:
    try:
        r = requests.get(pdf_url, timeout=timeout)
        r.raise_for_status()
        pdf_bytes = r.content
    except Exception as e:
        log.warning("%s PDF fetch fail: %s", hubo_id, e)
        return None

    import base64
    body = {
        "contents": [{
            "role": "user",
            "parts": [
                {"inline_data": {"mime_type": "application/pdf",
                                 "data": base64.b64encode(pdf_bytes).decode()}},
                {"text": PROMPT},
            ],
        }],
        "generationConfig": {"temperature": 0.0, "maxOutputTokens": 8192},
    }
    url = API_URL.format(model=model) + f"?key={api_key}"
    try:
        r = requests.post(url, json=body, timeout=timeout * 2)
        if not r.ok:
            log.warning("%s Gemini %s: %s", hubo_id, r.status_code, r.text[:200])
            return None
        j = r.json()
        parts = j.get("candidates", [{}])[0].get("content", {}).get("parts", [])
        text = "".join(p.get("text", "") for p in parts).strip()
        return text or None
    except Exception as e:
        log.warning("%s Gemini call fail: %s", hubo_id, e)
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pledges", default="pledges.json")
    ap.add_argument("--out", default="pledges_text.json")
    ap.add_argument("--key", default=os.environ.get("GEMINI_API_KEY", ""),
                    help="Gemini API key (또는 GEMINI_API_KEY env)")
    ap.add_argument("--model", default="gemini-2.5-flash-lite")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--limit", type=int, default=0, help="처음 N명만")
    ap.add_argument("--prefer", default="선거공보",
                    help="여러 파일이 있을 때 우선 사용할 종류 (선거공보 | 5대공약)")
    args = ap.parse_args()

    if not args.key:
        log.error("API key가 없어. --key 또는 GEMINI_API_KEY env 설정 필요.")
        sys.exit(1)

    pledges = json.loads(Path(args.pledges).read_text(encoding="utf-8"))
    out_path = Path(args.out)
    out: dict[str, str] = {}
    if out_path.exists():
        out = json.loads(out_path.read_text(encoding="utf-8"))
        log.info("loaded existing %s (%d entries)", out_path, len(out))

    # 작업 큐: 텍스트가 없는 후보자
    tasks = []
    for hubo_id, p in pledges.items():
        if hubo_id in out:
            continue
        files = p.get("files") or {}
        if not files:
            continue
        url = files.get(args.prefer) or next(iter(files.values()))
        tasks.append((hubo_id, url))

    if args.limit:
        tasks = tasks[: args.limit]
    log.info("to extract: %d", len(tasks))

    def worker(item):
        hubo_id, url = item
        text = extract_one(args.key, args.model, hubo_id, url)
        return hubo_id, text

    done = 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(worker, t) for t in tasks]
        for fut in as_completed(futs):
            hubo_id, text = fut.result()
            if text:
                out[hubo_id] = text
            done += 1
            if done % 20 == 0:
                out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
                log.info("checkpoint %d/%d (extracted=%d)", done, len(tasks), len(out))

    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("done. extracted=%d → %s", len(out), out_path)


if __name__ == "__main__":
    main()
