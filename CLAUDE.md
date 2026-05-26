# 2026 지방선거 후보자 분석 프로젝트

중앙선거관리위원회 선거통계시스템(`info.nec.go.kr`)에서 **제9회 전국동시지방선거** (2026.06.03) 후보자 데이터를 크롤링하고, 사진 DB와 단일 HTML 대시보드를 만드는 프로젝트.

- 선거 ID: `0020260603`
- 후보자 총 7,826명 (8종 선거 합계)
- 사진 보유: 7,313장 (NEC에 등록된 후보)
- 모든 산출물은 오프라인 동작 가능 (Chart.js만 CDN)

## 파일 구조

| 파일 | 종류 | 역할 |
|---|---|---|
| `crawler.py` | 크롤러 | NEC 후보자 목록 → `nec_2026_지방선거_후보자.xlsx` 생성 |
| `_fetch_photo_db.py` | 크롤러 | NEC 상세페이지에서 **정확한** photo URL 추출 → `photos/`에 썸네일 다운로드 + `.photo_db.json` |
| `_fetch_pledges.py` | 크롤러 | NEC 정책·공약마당(`policy.nec.go.kr`)에서 후보자 공약 PDF 메타데이터 → `pledges.json` |
| `_rebuild_photo_urls.py` | 보조 | 시군구 코드 매핑으로 사진 URL **추측** (fallback용. 정확하진 않음) |
| `_check_photos.py` | 진단 | (참조용) Content-Type 으로 사진 유무 검증 |
| `build_dashboard.py` | 빌더 | xlsx + photo_db → `dashboard.html` 생성. 모달에 "공약·AI Q&A 보기" CTA |
| `build_candidates.py` | 빌더 | xlsx + photo_db + pledges → `candidates/{huboId}.html` 7,826개 + 공유 `style.css`, `qa.js` |
| `nec_2026_지방선거_후보자.xlsx` | 데이터 | 7,826행, 22컬럼 (전체 + 8개 선거 시트) |
| `.photo_db.json` | 캐시 | `{huboId: {url, thumb_url, local, has_photo}}` — 1순위 |
| `.photo_urls.json` | 캐시 | huboId → 재구성된 URL (fallback) |
| `pledges.json` | 캐시 | `{huboId: {party, name, sgg, files: {선거공보: pdfUrl, 5대공약: pdfUrl, ...}}}` |
| `photos/{huboId}.jpg` | 자산 | 썸네일 7,313장 (~68MB) |
| `dashboard.html` | 산출물 | 단일 HTML, 데이터·UI 임베드 |
| `candidates/{huboId}.html` | 산출물 | 후보자별 페이지 (기본정보·재산·공약 PDF iframe·Gemini Q&A) |
| `candidates/style.css`·`qa.js` | 공유 | 모든 후보자 페이지가 참조 |

## 워크플로우

```powershell
# 한글 출력을 위해 항상 먼저
$env:PYTHONIOENCODING='utf-8'

# 1) 후보자 목록 갱신 (사퇴/등록 변동 시) — 약 10-15분
python crawler.py

# 2) 사진 갱신 — 재실행 안전 (캐시에 없거나 미확인 항목만 처리)
python _fetch_photo_db.py

# 3) 공약 PDF 메타데이터 갱신 — ~5-10분
python _fetch_pledges.py

# 4) 빌드 — 두 빌더는 독립적이므로 순서 상관없음
python build_dashboard.py             # dashboard.html
python build_candidates.py            # candidates/{huboId}.html × 7,826

# (선택) Gemini API 키를 페이지에 임베드해서 빌드
python build_candidates.py --gemini-key "AIzaSy..."

# 5) 배포
$dst = "$env:USERPROFILE\Downloads\2026_지방선거_대시보드"
New-Item -ItemType Directory -Force -Path $dst,"$dst\photos","$dst\candidates" | Out-Null
Copy-Item dashboard.html "$dst\dashboard.html" -Force
Copy-Item photos\*.jpg "$dst\photos\" -Force
Copy-Item candidates\* "$dst\candidates\" -Force
```

## 주요 설계 결정

### 사진 URL은 추측하지 말 것
NEC 사진 디렉토리는 선거 종류에 따라 다른 코드 체계를 씀:
- **시도 단위 선거** (시도지사 3, 광역비례 8, 교육감 11): `Gsg{시도cityCode}` (예: 서울 1100)
- **구시군 단위 선거** (구시군장 4, 시도의원 5, 구시군의원 6, 기초비례 9, 국회 2): `Gsg{시군구코드}` (예: 종로구 1101, 부산 동래구 2608)

시군구 코드는 `sggCityCode[:4]` 같은 단순 절단으로 추측하면 일부 시·군에서 어긋남. **항상 NEC 상세페이지 (`candidate_detail_info.xhtml?huboId=...`)에서 정규식 `/photo_20260603/[^"'?]+\.JPG`로 추출**.

### 통합특별시 (2900)
2026 선거부터 광주광역시 + 전라남도가 "전남광주통합특별시"로 통합. cityCode 2900으로 처리.
- 시도단위 선거(3, 5, 8, 11): `selectbox_cityCodeDIBySgJson.json` 사용 → 16개 시도
- 구시군단위 선거(2, 4, 6, 9): `selectbox_cityCodeBySgJson.json` 사용 → 17개 시도 (분리)

### 사진 누락 = NEC HTML fallback 응답
NEC는 사진 없으면 **200 OK + HTML 안내페이지** (~3.3KB) 반환. 404 아님. 감지 방법:
- 서버측: `Content-Type` 헤더 + `Content-Length < 5KB` 체크 (`_fetch_photo_db.py:download_thumb`)
- 클라이언트측 (`safeImg` JS 헬퍼): `<img onload>`에서 `naturalWidth < 10`이면 강제 `error` 이벤트 dispatch → fallback HTML로 교체

### 정당명 정규화
비례대표는 후순위 번호가 붙어옴 (예: `더불어민주당(1)`, `국민의힘(2)`). 모든 차트/집계 전에 `re.sub(r'\(\d+\)\s*$', '', p)`로 제거 (`build_dashboard.py:normalize_party`).

### 차트 필터 규칙
대시보드에서 필터(정당 탭 / 선거종류 / 지역 / 선거구) 적용 동작:
- **정당별 비교 차트 3개** (정당별 전과율, 정당별 평균 자산, 정당별 평균 나이) → **항상 `DATA` 전체 기준** (필터 무관). 정당 간 절대 비교를 유지하려고 의도적으로 고정. h3 옆에 보라 "전체 기준 · 필터 무관" 라벨.
- **그 외 모든 차트와 KPI** → 필터 반영. 초록 "필터 반영" 라벨.

### 색상 규약
- **정당이 등장하는 차트** → 정당 고유 색 (`PARTY_COLORS` 객체)
- **그 외 모든 차트** → primary `#5645d4` 단일색
- **자산 히스토그램** → 음수(빚) 구간만 `#dc2626` 빨강, 나머지 primary
- **병역 도넛** → 보라 톤 + 미필/면제만 빨강

## 디자인 시스템

Notion 디자인 시스템 (Inter 폰트, 8px 버튼 라운드, 12px 카드 라운드, primary `#5645d4`). 색상 토큰은 `build_dashboard.py`의 `:root` CSS 변수 참조.

## 핵심 NEC 엔드포인트

| 메서드 | 경로 | 용도 |
|---|---|---|
| GET | `/main/showDocument.xhtml` | 세션 부트스트랩 (JSESSIONID) |
| POST | `/electioninfo/electionInfo_report.xhtml` | 후보자 목록 (HTML 테이블) |
| GET | `/electioninfo/candidate_detail_info.xhtml?electionId=&huboId=` | 후보자 상세 (정확한 photo `<img src>` 포함) |
| POST | `/bizcommon/selectbox/selectbox_cityCodeBySgJson.json` | 시도 목록 (일반) |
| POST | `/bizcommon/selectbox/selectbox_cityCodeDIBySgJson.json` | 시도 목록 (통합권역) |
| POST | `/bizcommon/selectbox/selectbox_townCodeBySgJson.json` | 시도→시군구 |
| POST | `/bizcommon/selectbox/selectbox_townCodeByCityIntgSgJson.json` | 시도→시군구 (통합권역) |
| POST | `/bizcommon/selectbox/selectbox_getSggCityCodeJson.json` | 시도→sggCityCode |
| POST | `/bizcommon/selectbox/selectbox_getSggTownCodeJson.json` | 시군구→sggTownCode (선거구) |

응답 JSON 형식: `{"jsonResult": {"body": [{"CODE": "...", "NAME": "..."}, ...]}}`

## 선거 종류 코드

| code | 이름 | drill-down | cityCode flavor |
|---|---|---|---|
| 3 | 시도지사선거 | single (cityCode만) | DI (통합) |
| 4 | 구시군의장선거 | sgg (+sggCityCode) | 일반 |
| 5 | 시도의회의원선거 | drill (+townCode+sggTownCode) | DI (통합) |
| 6 | 구시군의회의원선거 | drill (+townCode+sggTownCode) | 일반 |
| 8 | 광역의원비례대표선거 | single | DI |
| 9 | 기초의원비례대표선거 | sgg | 일반 |
| 11 | 교육감선거 | single (9개 시도만) | DI |
| 2 | 국회의원선거(보궐) | sgg | 일반 |

## 주의/제약

- **저작권/이용제한** — 후보자 정보공개자료는 공직선거법 제49조 제12항에 따라 선거일까지만 공개. 목적외 사용 금지.
- **NEC 부하** — 사진/데이터 크롤링은 동시 6 worker + 0.05~0.15s sleep 유지. 한꺼번에 두드리지 말 것.
- **한글 출력** — Windows PowerShell에서 Python 실행 시 `$env:PYTHONIOENCODING='utf-8'` 먼저 설정.
- **CSV로 변환 시** — UTF-8 BOM (`utf-8-sig`)으로 저장해야 한국 Windows Excel에서 안 깨짐.

## 환경

- Python 3.13.2 (Windows)
- 패키지: `requests`, `beautifulsoup4`, `lxml`, `openpyxl`
- Chart.js 4.4.1 (CDN, dashboard 전용)
- Inter 폰트 (Google Fonts CDN)

## 공약·AI Q&A 시스템 (2026-05 추가)

### 공약 데이터 출처
공약은 `info.nec.go.kr`이 아니라 별도 사이트 **`policy.nec.go.kr` (NEC 정책·공약마당)**. 메뉴 ID `CNDDT25`가 9회 지방선거.

크롤 cascade (`_fetch_pledges.py`):
1. `initUCACommiment.do?menuId=CNDDT25` — 세션 부트스트랩
2. `initUCACommimentRegion.do` (`sgId`, `subSgId`) — 시도 목록
3. `initUCACommimentGu.do` (`sgId`, `subSgId`, **`wiwsidocode`**, `sortYn`) — 구시군 목록 ← param 이름 함정
4. `initUCACommimentSgg.do` (+ **`wiwid`**) — 선거구 목록
5. `initUCACommimentList.do` (sgTypecode + hRegionId/hGuId/hSggId) — 후보자 목록

선거 종류별 drill 깊이:
- depth=1 (시도지사 3, 광역비례 8, 교육감 11): Region 만
- depth=3 (그 외 모두): Region → Gu → Sgg → List

응답 row의 `fileinfo`는 `"종류||경로||종류||경로,..."` 형식. 여러 파일(선거공보 / 선거공약서 / 5대공약)이 같이 옴. PDF는 `https://cdn.nec.go.kr/policy_pdf/{path}`에서 CORS `*`로 직접 접근 가능 → 후보자 페이지에서 iframe 임베드.

### 상세 스캔(재산·병역·전과·학력·공직선거경력)은 hotlink 차단됨
`info.nec.go.kr/electioninfo/candidate_detail_scanSearchJson.json` 으로 PDF 경로는 받을 수 있지만 `/unielec_pdf_file/...` 직접 GET은 referer/세션 검사로 막힘. 우리 페이지에선 **NEC 상세보기 외부 링크** 버튼으로만 처리.

### Gemini Q&A
모델: `gemini-2.5-flash-lite` (가장 저렴, multimodal). 키는 다음 우선순위:
1. `window.GEMINI_API_KEY` (빌드 시 `--gemini-key` 로 임베드)
2. localStorage `nec_gemini_api_key` (페이지 상단 input에서 사용자가 입력)

전송 컨텍스트:
- 사전 추출된 텍스트(`window.PLEDGE_TEXT`)가 있으면 텍스트만 전송 (저렴)
- 없으면 클라이언트가 PDF를 fetch → base64 inline_data 로 한 번 캐시한 뒤 질문마다 재사용 (`cachedPdfParts`)

system_instruction에 후보자명/정당/선거 정보 박아두고 "공약 문서에 근거하지 않은 추측 금지" 원칙 명시. 추가 토큰 제어: `temperature: 0.2`, `maxOutputTokens: 1024`.

**API 키 노출 주의** — `--gemini-key`로 빌드하면 키가 모든 HTML에 평문으로 들어가 누구나 추출 가능. 데모/내부용으로만 쓰고, 배포 후엔 [aistudio.google.com/apikey](https://aistudio.google.com/apikey) 에서 키 재발급(rotate) 권장. 운영 배포가 필요해지면 Vercel/Cloudflare Workers의 작은 프록시로 옮길 것.

## Vercel 배포 메모

- 사진(~68MB) + dashboard.html(~4.7MB) + candidates/(~50MB, 7,826개) + 공유 css/js → **총 ~125MB, 파일 ~15,140개**
- 무료 Hobby 한도(100MB / 15,000개)를 **둘 다 약간 넘김**. 옵션:
  1. **Pro 플랜** (가장 단순)
  2. **per-candidate HTML 폐기 → 단일 `candidate.html` + 해시 라우팅** (e.g. `candidate.html#100157144`). 데이터는 `candidate_data.json` 하나로 통합. → 파일 8천대로 감소
  3. **photos를 Vercel Blob/외부 CDN으로 분리** → 파일·용량 둘 다 해소
- Git push 말고 **CLI 직접 배포 권장**: 폴더에서 `vercel --prod`
- `dashboard.html` 의 사진 경로(`photos/{id}.jpg`)와 candidate 페이지 내부 사진 경로(`../photos/{id}.jpg`)는 둘 다 상대경로
