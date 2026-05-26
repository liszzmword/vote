"""
가상 여론조사 결과 → virtual_poll.html 빌더

입력:
  - virtual_poll_raw.jsonl   (race 별 페르소나 응답)
  - pledges.json             (후보 메타)
  - .photo_db.json           (후보 사진)

처리:
  - race 별 집계: 후보별 vote 수, confidence 분포, abstain/none 비율, key_issue Top, comments sample
  - 결과 → virtual_poll.json (요약 데이터)
  - 결과 → virtual_poll.html (Chart.js 임베드)
"""

from __future__ import annotations

import html
import json
import logging
import random
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("build_poll")

ROOT = Path(__file__).parent
RAW = ROOT / "virtual_poll_raw.jsonl"
PHOTO_DB = ROOT / ".photo_db.json"
AGG_OUT = ROOT / "virtual_poll.json"
HTML_OUT = ROOT / "virtual_poll.html"

PARTY_COLORS = {
    "더불어민주당": "#004EA2", "국민의힘": "#E61E2B", "조국혁신당": "#0073CF",
    "개혁신당": "#FF7920", "진보당": "#D6001C", "정의당": "#FFCC00",
    "기본소득당": "#00B5A8", "새미래민주당": "#1AA75E", "자유와혁신": "#8A2BE2",
    "노동당": "#EB121A", "자유통일당": "#A6324D", "무소속": "#9ca3af",
}


def party_color(p): return PARTY_COLORS.get(p, "#7c3aed")


def aggregate_race(race: dict) -> dict:
    cands = race["candidates"]
    letter_to_cand = {c["letter"]: c for c in cands if c.get("letter")}
    n_total = len(race["responses"])

    tally = {l: {"strong": 0, "moderate": 0, "reluctant": 0, "total": 0} for l in letter_to_cand}
    tally["none"] = {"total": 0}
    tally["abstain"] = {"total": 0}
    key_issues = Counter()
    likeability = {l: [] for l in letter_to_cand}
    comments = []  # (age, sex, comment, vote)

    for r in race["responses"]:
        resp = r.get("response") or {}
        vote = resp.get("vote") or ""
        conf = resp.get("confidence") or "moderate"
        if conf not in ("strong", "moderate", "reluctant"):
            conf = "moderate"
        if vote in letter_to_cand:
            tally[vote][conf] = tally[vote].get(conf, 0) + 1
            tally[vote]["total"] += 1
        elif vote == "none":
            tally["none"]["total"] += 1
        elif vote == "abstain":
            tally["abstain"]["total"] += 1
        else:
            # malformed — count as abstain
            tally["abstain"]["total"] += 1

        ki = (resp.get("key_issue") or "").strip()
        if ki: key_issues[ki] += 1

        ratings = resp.get("ratings") or {}
        for l, rt in ratings.items():
            if l in likeability and isinstance(rt, dict):
                lk = rt.get("likeability")
                if isinstance(lk, (int, float)):
                    likeability[l].append(float(lk))

        cm = (resp.get("comment") or "").strip()
        meta = r.get("persona_meta") or {}
        if cm:
            comments.append({
                "age": meta.get("age"),
                "sex": meta.get("sex"),
                "occupation": meta.get("occupation"),
                "district": meta.get("district"),
                "vote": vote,
                "text": cm[:240],
            })

    # percentage
    pct_base = max(n_total, 1)
    for k, v in tally.items():
        v["pct"] = round(100 * v["total"] / pct_base, 1)
    # average likeability
    for l in likeability:
        arr = likeability[l]
        tally[l]["avg_likeability"] = round(sum(arr)/len(arr), 2) if arr else None

    # diverse comments (sample 5, prefer different age buckets)
    rng = random.Random(42)
    by_age = defaultdict(list)
    for c in comments:
        try:
            a = int(c["age"] or 0)
            bucket = "20s" if a<30 else "30s" if a<40 else "40s" if a<50 else "50s" if a<60 else "60+"
        except: bucket = "??"
        by_age[bucket].append(c)
    sample_comments = []
    for bucket in ["20s","30s","40s","50s","60+","??"]:
        arr = by_age.get(bucket, [])
        if arr: sample_comments.append(rng.choice(arr))
        if len(sample_comments) >= 5: break
    if len(sample_comments) < 5:
        sample_comments += rng.sample(comments, min(5 - len(sample_comments), len(comments)))

    return {
        "race_key": race["race_key"],
        "sheet": race["sheet"],
        "sido": race["sido"],
        "region": race["region"],
        "level": race["level"],
        "n_voters": n_total,
        "n_sampled": race.get("n_personas_sampled", n_total),
        "candidates": cands,
        "tally": tally,
        "key_issues_top": key_issues.most_common(8),
        "comments": sample_comments,
    }


def aggregate_all():
    if not RAW.exists():
        log.error("%s 없음 — 먼저 _run_virtual_vote.py 실행", RAW)
        return None
    races = []
    with RAW.open(encoding="utf-8") as f:
        for line in f:
            try: races.append(json.loads(line))
            except json.JSONDecodeError: continue
    log.info("raw races: %d", len(races))
    return [aggregate_race(r) for r in races]


# ---------- HTML ---------- #

CSS = r"""
:root {
  --primary:#5645d4; --primary-pressed:#4534b3; --on-primary:#fff;
  --brand-navy:#0a1530; --link-blue:#0075de;
  --tint-mint:#d9f3e1; --tint-rose:#fde0ec; --tint-yellow:#fef7d6; --tint-sky:#dcecfa;
  --canvas:#fff; --surface:#f6f5f4; --surface-soft:#fafaf9;
  --hairline:#e5e3df; --hairline-soft:#ede9e4;
  --ink:#1a1a1a; --charcoal:#37352f; --slate:#5d5b54; --steel:#787671; --stone:#a4a097; --muted:#bbb8b1;
  --r-sm:6px; --r-md:8px; --r-lg:12px; --r-xl:16px; --r-full:9999px;
  --shadow-card: rgba(15,15,15,0.06) 0px 4px 12px 0px;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  font-family:'Inter',-apple-system,BlinkMacSystemFont,'Apple SD Gothic Neo','맑은 고딕',sans-serif;
  background:var(--surface); color:var(--ink); font-size:15px; line-height:1.55;
}
a{color:var(--link-blue);text-decoration:none}
.container{max-width:1200px;margin:0 auto;padding:24px 20px 80px}
.topbar{display:flex;align-items:center;gap:10px;padding:8px 0 16px;font-size:13px;color:var(--steel)}
.topbar a{color:var(--steel)}

.hero{
  background:var(--brand-navy);color:#fff;
  padding:48px 20px 36px;text-align:center;
  border-radius:0 0 var(--r-xl) var(--r-xl);
  margin-bottom:24px;
}
.hero h1{font-size:32px;margin:0 0 10px;letter-spacing:-0.02em;font-weight:700}
.hero .sub{color:#bbb;font-size:14px;margin-bottom:16px}
.hero .meta{display:flex;flex-wrap:wrap;justify-content:center;gap:14px;font-size:13px;color:#cfc7e8}
.hero .meta .pill{
  background:rgba(255,255,255,0.08);padding:5px 12px;border-radius:var(--r-full);
  border:1px solid rgba(255,255,255,0.12);
}
.disclaimer{
  background:#fff7ed;border:1px solid #fcd9b3;color:#7c2d12;
  padding:12px 16px;border-radius:var(--r-md);font-size:12.5px;line-height:1.55;margin-bottom:24px;
}

.tabs{display:flex;gap:0;border-bottom:1px solid var(--hairline);margin-bottom:20px}
.tab{
  padding:12px 18px;cursor:pointer;font-size:14px;font-weight:600;color:var(--steel);
  border-bottom:2px solid transparent;background:none;border:none;font-family:inherit;
}
.tab:hover{color:var(--charcoal)}
.tab.active{color:var(--primary);border-bottom-color:var(--primary)}

.controls{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;align-items:center}
.controls select{
  padding:8px 12px;border:1px solid var(--hairline);border-radius:var(--r-md);
  background:#fff;font-family:inherit;font-size:13.5px;color:var(--charcoal);min-width:180px;
}
.controls select:focus{outline:none;border-color:var(--primary)}
.controls .race-info{color:var(--steel);font-size:13px;margin-left:auto}

.race-card{
  background:var(--canvas);border:1px solid var(--hairline);border-radius:var(--r-lg);
  box-shadow:var(--shadow-card);padding:24px 28px;margin-bottom:20px;
}
.race-head{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-end;gap:12px;margin-bottom:18px}
.race-title{font-size:20px;font-weight:700;color:var(--charcoal);letter-spacing:-0.01em}
.race-sub{color:var(--steel);font-size:13px;margin-top:4px}
.race-stat{text-align:right}
.race-stat .n{font-size:24px;font-weight:700;color:var(--primary)}
.race-stat .lbl{font-size:11px;color:var(--steel);text-transform:uppercase;letter-spacing:.5px}

.cards-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:24px}
.cand-card{
  background:var(--surface-soft);border:1px solid var(--hairline);border-radius:var(--r-md);
  padding:14px;display:flex;gap:12px;align-items:center;
}
.cand-photo{width:50px;height:66px;border-radius:var(--r-sm);object-fit:cover;background:var(--surface);border:1px solid var(--hairline)}
.cand-photo.missing{display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:18px;font-weight:bold;background:var(--surface)}
.cand-meta .letter{font-size:11px;color:var(--steel);font-weight:700;letter-spacing:.5px;margin-bottom:2px}
.cand-meta .name{font-size:14px;font-weight:600;color:var(--charcoal)}
.cand-meta .party{font-size:11px;color:var(--slate);margin-top:2px}
.cand-meta .pct{font-size:18px;font-weight:700;margin-top:4px}

.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px}
@media (max-width:760px){.chart-grid{grid-template-columns:1fr}}
.chart-card{
  background:var(--surface-soft);border:1px solid var(--hairline);border-radius:var(--r-md);padding:16px;
}
.chart-card h4{margin:0 0 12px;font-size:13px;color:var(--charcoal);font-weight:600;display:flex;align-items:center;gap:8px}
.chart-card h4 .badge{font-size:10px;color:var(--steel);background:var(--canvas);padding:2px 7px;border-radius:var(--r-full);border:1px solid var(--hairline);font-weight:500}
.chart-wrap{height:280px;position:relative}

.comments{display:flex;flex-direction:column;gap:10px;margin-top:8px}
.comment{
  background:var(--surface-soft);border-left:3px solid var(--primary);
  padding:12px 16px;border-radius:0 var(--r-md) var(--r-md) 0;
}
.comment .quote{font-size:14px;color:var(--charcoal);font-style:italic;line-height:1.6;margin-bottom:6px}
.comment .who{font-size:11.5px;color:var(--steel)}

.empty-state{padding:60px 20px;text-align:center;color:var(--steel);font-size:14px}
"""


def render_html(aggregated: list[dict]) -> str:
    # photo DB
    photo_db = {}
    if PHOTO_DB.exists():
        photo_db = json.loads(PHOTO_DB.read_text(encoding="utf-8"))

    # group races by sheet
    by_sheet = defaultdict(list)
    for r in aggregated:
        by_sheet[r["sheet"]].append(r)
    for sheet in by_sheet:
        by_sheet[sheet].sort(key=lambda r: (r["sido"], r["region"]))

    # Attach photo info to candidates
    for r in aggregated:
        for c in r["candidates"]:
            hid = c["huboId"]
            ph = photo_db.get(hid) or {}
            c["photo"] = f"photos/{hid}.jpg" if ph.get("has_photo") else None

    total_voters = sum(r["n_voters"] for r in aggregated)
    total_races = len(aggregated)
    sheets = sorted(by_sheet.keys())
    gen_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    data_js = json.dumps(by_sheet, ensure_ascii=False)
    party_colors_js = json.dumps(PARTY_COLORS, ensure_ascii=False)

    return f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>가상 여론조사 — 2026 지방선거</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>{CSS}</style>
</head>
<body>

<div class="container">
  <div class="topbar"><a href="dashboard.html">← 대시보드로</a></div>
</div>

<header class="hero">
  <h1>🗳️ 가상 여론조사 결과</h1>
  <div class="sub">NVIDIA Nemotron 한국 페르소나로 시뮬레이션한 2026 지방선거 가상 투표</div>
  <div class="meta">
    <div class="pill">총 {total_races}개 선거</div>
    <div class="pill">{total_voters:,}명 페르소나 응답</div>
    <div class="pill">모델: gemini-2.5-flash-lite</div>
    <div class="pill">생성: {gen_at}</div>
  </div>
</header>

<div class="container">
  <div class="disclaimer">
    ⚠️ <strong>면책</strong> · 이 페이지는 NVIDIA Nemotron 한국 페르소나(100만)에서 추출한 가상의 유권자가 후보 공약 텍스트를 읽고 답한 시뮬레이션 결과입니다.
    실제 여론조사·예측이 아니며, 페르소나는 후보 이름·정당을 모른 채(알파벳 익명) 공약 내용만으로 판단했습니다.
    AI 모델 자체의 학습 편향이 결과에 반영될 수 있으니 참고용으로만 봐주세요.
  </div>

  <div class="tabs" id="tabs">
    {''.join(f'<button class="tab{" active" if i==0 else ""}" data-sheet="{html.escape(s)}">{html.escape(s)}</button>' for i, s in enumerate(sheets))}
  </div>

  <div class="controls">
    <select id="raceSelect"></select>
    <div class="race-info" id="raceInfo"></div>
  </div>

  <div id="raceContent">
    <div class="empty-state">선거와 지역을 선택하세요.</div>
  </div>
</div>

<script>
const DATA = {data_js};
const PARTY_COLORS = {party_colors_js};
const partyColor = p => PARTY_COLORS[p] || "#7c3aed";
const confLabel = {{strong:"확실", moderate:"약간", reluctant:"마지못해"}};

const tabsEl = document.getElementById('tabs');
const raceSelect = document.getElementById('raceSelect');
const raceContent = document.getElementById('raceContent');
const raceInfo = document.getElementById('raceInfo');

let currentSheet = Object.keys(DATA)[0];
let currentChart = null, currentConfChart = null, currentIssueChart = null;

function rebuildRaceSelect() {{
  const races = DATA[currentSheet] || [];
  raceSelect.innerHTML = races.map((r, i) =>
    `<option value="${{i}}">${{r.sido}}${{r.level==='sgg' && r.region !== r.sido ? ' · ' + r.region : ''}}</option>`
  ).join('');
  if (races.length) {{
    raceSelect.value = '0';
    renderRace(races[0]);
  }} else {{
    raceContent.innerHTML = '<div class="empty-state">이 선거의 결과가 없습니다.</div>';
  }}
}}

function pctRow(label, pct, color) {{
  return `<div style="display:flex;align-items:center;gap:8px;font-size:13px;margin:4px 0">
    <span style="width:90px;color:var(--slate)">${{label}}</span>
    <div style="flex:1;height:14px;background:var(--hairline-soft);border-radius:7px;overflow:hidden">
      <div style="height:100%;width:${{Math.max(pct,0)}}%;background:${{color}}"></div>
    </div>
    <span style="width:54px;text-align:right;font-variant-numeric:tabular-nums">${{pct.toFixed(1)}}%</span>
  </div>`;
}}

function renderRace(r) {{
  // Race header
  const validLetters = Object.keys(r.tally).filter(k => k !== 'none' && k !== 'abstain');
  const letterToCand = {{}};
  for (const c of r.candidates) if (c.letter) letterToCand[c.letter] = c;

  raceInfo.textContent = `유효 후보 ${{validLetters.length}}명 · 페르소나 ${{r.n_voters}}명 응답 (샘플 ${{r.n_sampled}})`;

  // Sorted by total votes desc
  const sortedLetters = [...validLetters].sort((a, b) => r.tally[b].total - r.tally[a].total);

  const candCards = sortedLetters.map(l => {{
    const c = letterToCand[l] || {{}};
    const t = r.tally[l] || {{total:0, pct:0}};
    const ph = c.photo
      ? `<img class="cand-photo" src="${{c.photo}}" onerror="this.onerror=null;this.classList.add('missing');this.removeAttribute('src');this.textContent=' ';">`
      : `<div class="cand-photo missing">?</div>`;
    return `<div class="cand-card">
      ${{ph}}
      <div class="cand-meta">
        <div class="letter">기호 ${{c.giho||'-'}} · ${{l}}</div>
        <div class="name">${{c.name||'(이름 없음)'}}</div>
        <div class="party" style="color:${{partyColor(c.party)}}">${{c.party||''}}</div>
        <div class="pct" style="color:${{partyColor(c.party)}}">${{t.pct.toFixed(1)}}%</div>
      </div>
    </div>`;
  }}).join('');

  const noneT = r.tally.none || {{total:0,pct:0}};
  const absT = r.tally.abstain || {{total:0,pct:0}};

  const pctSection = `
    ${{sortedLetters.map(l => {{
      const c = letterToCand[l] || {{}};
      const t = r.tally[l] || {{total:0,pct:0}};
      return pctRow(`기호${{c.giho||'-'}} ${{c.name||''}}`, t.pct, partyColor(c.party));
    }}).join('')}}
    ${{pctRow('지지없음', noneT.pct, '#bbb8b1')}}
    ${{pctRow('기권', absT.pct, '#787671')}}
  `;

  // Comments
  const commentsHtml = (r.comments || []).map(c => {{
    const lc = letterToCand[c.vote];
    const voteTag = lc ? `→ 기호 ${{lc.giho}} ${{lc.name}} 지지` : (c.vote === 'none' ? '→ 지지없음' : c.vote === 'abstain' ? '→ 기권' : '');
    return `<div class="comment">
      <div class="quote">"${{c.text}}"</div>
      <div class="who">${{c.age||'?'}}세 ${{c.sex||''}} · ${{c.occupation||''}} · ${{c.district||''}} ${{voteTag}}</div>
    </div>`;
  }}).join('');

  raceContent.innerHTML = `
    <div class="race-card">
      <div class="race-head">
        <div>
          <div class="race-title">${{r.sido}}${{r.level==='sgg' && r.region !== r.sido ? ' · ' + r.region : ''}}</div>
          <div class="race-sub">${{r.sheet}} · 페르소나 ${{r.n_voters}}명 응답</div>
        </div>
        <div class="race-stat">
          <div class="n">${{((noneT.pct + absT.pct).toFixed(1))}}%</div>
          <div class="lbl">기권·지지없음</div>
        </div>
      </div>
      <div class="cards-row">${{candCards}}</div>
      <div style="margin:12px 0 22px">${{pctSection}}</div>

      <div class="chart-grid">
        <div class="chart-card">
          <h4>득표율 (응답 기준)</h4>
          <div class="chart-wrap"><canvas id="chBar"></canvas></div>
        </div>
        <div class="chart-card">
          <h4>지지 강도 분포</h4>
          <div class="chart-wrap"><canvas id="chConf"></canvas></div>
        </div>
      </div>

      <div class="chart-card" style="margin-bottom:20px">
        <h4>결정에 영향 준 분야 Top</h4>
        <div class="chart-wrap" style="height:200px"><canvas id="chIssues"></canvas></div>
      </div>

      <h4 style="margin:24px 0 12px;color:var(--charcoal);font-size:14px">페르소나 코멘트 샘플</h4>
      <div class="comments">${{commentsHtml || '<div class="empty-state">코멘트 없음</div>'}}</div>
    </div>
  `;

  // Charts
  if (currentChart) currentChart.destroy();
  if (currentConfChart) currentConfChart.destroy();
  if (currentIssueChart) currentIssueChart.destroy();

  const ctx1 = document.getElementById('chBar').getContext('2d');
  currentChart = new Chart(ctx1, {{
    type:'bar',
    data:{{
      labels: sortedLetters.map(l => {{
        const c = letterToCand[l];
        return `기호 ${{c.giho}} ${{c.name}}`;
      }}),
      datasets:[{{
        label:'득표율',
        data: sortedLetters.map(l => r.tally[l].pct),
        backgroundColor: sortedLetters.map(l => partyColor((letterToCand[l]||{{}}).party)),
        borderRadius: 6,
      }}]
    }},
    options:{{
      indexAxis:'y',
      maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}, tooltip:{{callbacks:{{label:ctx => ctx.parsed.x.toFixed(1)+'%'}}}}}},
      scales:{{x:{{ticks:{{callback: v => v+'%'}}, suggestedMax: 100}}}}
    }}
  }});

  const ctx2 = document.getElementById('chConf').getContext('2d');
  currentConfChart = new Chart(ctx2, {{
    type:'bar',
    data:{{
      labels: sortedLetters.map(l => {{ const c=letterToCand[l]; return `${{c.name}}`; }}),
      datasets: ['strong','moderate','reluctant'].map((k,i) => ({{
        label: confLabel[k],
        data: sortedLetters.map(l => r.tally[l][k]||0),
        backgroundColor: ['#5645d4','#9087e3','#cfcaf0'][i],
      }}))
    }},
    options:{{
      maintainAspectRatio:false,
      scales:{{x:{{stacked:true}}, y:{{stacked:true, ticks:{{precision:0}}}}}},
      plugins:{{legend:{{position:'top',labels:{{boxWidth:12,font:{{size:11}}}}}}}}
    }}
  }});

  // Issues bar
  const issues = (r.key_issues_top || []).slice(0, 8);
  const ctx3 = document.getElementById('chIssues').getContext('2d');
  currentIssueChart = new Chart(ctx3, {{
    type:'bar',
    data:{{
      labels: issues.map(x => x[0]),
      datasets:[{{label:'언급 수', data: issues.map(x => x[1]), backgroundColor:'#5645d4', borderRadius:6}}]
    }},
    options:{{maintainAspectRatio:false, indexAxis:'y',
      plugins:{{legend:{{display:false}}}}, scales:{{x:{{ticks:{{precision:0}}}}}}}}
  }});
}}

tabsEl.querySelectorAll('.tab').forEach(t => {{
  t.addEventListener('click', () => {{
    tabsEl.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    t.classList.add('active');
    currentSheet = t.dataset.sheet;
    rebuildRaceSelect();
  }});
}});
raceSelect.addEventListener('change', () => {{
  const races = DATA[currentSheet] || [];
  renderRace(races[parseInt(raceSelect.value, 10)]);
}});

rebuildRaceSelect();
</script>
</body>
</html>
"""


def main():
    aggregated = aggregate_all()
    if aggregated is None:
        return
    log.info("races aggregated: %d", len(aggregated))

    AGG_OUT.write_text(json.dumps(aggregated, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("wrote %s (%.1f KB)", AGG_OUT, AGG_OUT.stat().st_size/1024)

    html_str = render_html(aggregated)
    HTML_OUT.write_text(html_str, encoding="utf-8")
    log.info("wrote %s (%.1f KB)", HTML_OUT, HTML_OUT.stat().st_size/1024)


if __name__ == "__main__":
    main()
