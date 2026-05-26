// Vercel Serverless Function — 후보자 공약 Q&A
//   POST /api/qa
//   body: { question, candidate: {name, party, giho, subSgName, sgg}, pledgeText?, pdfUrls?: [{type,url}] }
//   resp: { answer } | { error }
//
// 키 관리: GEMINI_API_KEY 환경 변수 (Vercel Dashboard → Settings → Environment Variables)
// 모델: gemini-2.5-flash-lite
//
// 흐름:
//   - pledgeText 가 있으면 텍스트로 컨텍스트 구성 (가장 저렴/빠름)
//   - 없으면 서버에서 pdfUrls 의 PDF를 직접 fetch → inline_data 로 Gemini 에 전달
//     (Vercel 함수가 PDF를 가져와서 모델로 보내므로 클라이언트 요청 크기 제한 무관)

export const config = { maxDuration: 60 };

const MODEL = 'gemini-2.5-flash-lite';
const ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent`;
const PDF_SIZE_LIMIT = 18 * 1024 * 1024; // 18MB — inline_data 한도 고려

function buildSystemPrompt(c = {}, mode = 'single') {
  let focusLine = '';
  let modeRule = '';
  if (mode === 'compare') {
    focusLine = '작업 모드: 여러 후보의 공약을 비교 분석.';
    modeRule = `
[비교 모드 추가 규칙]
- 첨부된 PDF 각각이 어느 후보의 것인지 (위 PDF: 후보명(정당)) 표시가 옆에 붙어 있으니 그걸 기준으로 후보별 입장을 정리하세요.
- 분야(예: 주거, 교통, 일자리, 복지, 교육, 환경 등)별로 후보별 입장을 짧게 비교하세요.
- 한 분야에서 한 후보만 언급했고 다른 후보는 안 다뤘다면 "(다른 후보는 해당 분야 미언급)" 처럼 명시.
- "더 낫다/잘했다/올바르다" 같은 평가는 절대 금지. 사실만 제시.`;
  } else if (c && c.name) {
    focusLine = `대상 후보자: ${c.name} (${c.party || ''}, 기호 ${c.giho || '?'}번, ${c.subSgName || ''} ${c.sgg || ''})`;
  }

  return `당신은 2026년 6월 3일 제9회 전국동시지방선거 관련 질문에만 답변하는 AI 어시스턴트입니다.
${focusLine}

[답변 가능 범위]
- 후보자의 공약, 정책, 비전
- 후보자의 기본 정보(정당, 선거구, 기호 등 공개된 정보)
- 2026 지방선거 제도, 일정, 절차에 관한 일반 정보
- 여러 후보 공약의 분야별 비교 (사실 기반, 평가 금지)

[답변 금지 범위]
- 선거와 무관한 일반 질문 (날씨, 음식, 코딩, 일상 잡담 등)
- 다른 선거(대선/총선 등) 또는 과거 선거 관련 질문
- 특정 후보를 비방하거나 정치적으로 평가하는 질문
- "어느 후보가 더 낫다/적합하다" 같은 우열 평가
- 자료에 근거하지 않은 추측

[선거 무관 질문에 대한 응답 규칙]
질문이 위의 "답변 가능 범위"에 해당하지 않으면, 컨텍스트 자료를 무시하고 정확히 아래 문장만 답변하세요:
"저는 2026 지방선거 후보자와 공약에 관한 질문만 답변할 수 있어요. 선거 관련해서 다른 궁금한 점 있으세요?"

[답변 형식]
- 한국어, 간결하게(단일 후보 3~6문장 / 비교는 분야별로 1~2문장씩, 전체 12문장 이내).
- 마크다운 사용 금지: **굵게**, *기울임*, # 헤더, > 인용, --- 구분선, [링크](url) 등 일체 금지.
- 항목 나열이 꼭 필요하면 "1) ... 2) ... 3) ..." 또는 빈 줄로 구분된 짧은 문장.
- 자료에 없는 내용은 추측하지 말고 "제출된 공약 자료에서 확인되지 않습니다"라고 답변.${modeRule}`;
}

function cleanAnswer(text) {
  if (!text) return '';
  // remove markdown bold/italic markers, but keep the inner text
  let t = text.replace(/\*\*(.+?)\*\*/g, '$1');
  t = t.replace(/(^|[^*])\*(?!\s)([^*\n]+?)\*/g, '$1$2'); // *italic*
  t = t.replace(/^#{1,6}\s+/gm, ''); // headers
  t = t.replace(/^>\s+/gm, '');       // blockquote
  t = t.replace(/^\s*[-+]\s+/gm, '· '); // bullet → dot
  // also strip any straggling **
  t = t.replace(/\*\*/g, '');
  return t.trim();
}

async function fetchPdfPart(url, type) {
  const r = await fetch(url, { redirect: 'follow' });
  if (!r.ok) return null;
  const buf = Buffer.from(await r.arrayBuffer());
  if (buf.length === 0 || buf.length > PDF_SIZE_LIMIT) return null;
  return {
    label: type,
    part: { inline_data: { mime_type: 'application/pdf', data: buf.toString('base64') } },
  };
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'POST only' });
  }

  const key = process.env.GEMINI_API_KEY;
  if (!key) {
    return res.status(500).json({
      error: 'GEMINI_API_KEY 환경 변수가 Vercel에 설정되지 않았습니다. Settings → Environment Variables 에서 추가하세요.',
    });
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  const { question = '', candidate = {}, pledgeText = '', pdfUrls = [], mode = 'single' } = body || {};
  if (!question || typeof question !== 'string') {
    return res.status(400).json({ error: 'question 필드가 필요합니다.' });
  }

  // ---------- build context parts ----------
  const parts = [];
  let usedSource = 'none';

  if (pledgeText && pledgeText.length > 80) {
    parts.push({ text: '아래는 후보자가 NEC에 제출한 공약 자료에서 추출한 텍스트입니다.\n\n' + pledgeText });
    usedSource = 'text';
  } else if (Array.isArray(pdfUrls) && pdfUrls.length > 0) {
    const results = await Promise.all(
      pdfUrls.slice(0, 4).map((p) => fetchPdfPart(p.url, p.type).catch(() => null))
    );
    for (const r of results) {
      if (!r) continue;
      parts.push(r.part);
      parts.push({ text: `(위 PDF: ${r.label})` });
    }
    if (parts.length > 0) usedSource = 'pdf';
  }

  if (parts.length === 0) {
    return res.status(200).json({
      answer: '이 후보자는 NEC에 제출된 공약 자료가 없어 답변할 수 없습니다.',
      source: 'empty',
    });
  }

  parts.push({ text: '질문: ' + question });

  const reqBody = {
    system_instruction: { parts: [{ text: buildSystemPrompt(candidate, mode) }] },
    contents: [{ role: 'user', parts }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: mode === 'compare' ? 2048 : 1024,
    },
  };

  try {
    const r = await fetch(`${ENDPOINT}?key=${encodeURIComponent(key)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(reqBody),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      return res.status(r.status).json({
        error: j?.error?.message || ('Gemini ' + r.status),
      });
    }
    const raw = j?.candidates?.[0]?.content?.parts?.map((p) => p.text || '').join('') || '';
    return res.status(200).json({ answer: cleanAnswer(raw), source: usedSource });
  } catch (e) {
    return res.status(500).json({ error: e?.message || String(e) });
  }
}
