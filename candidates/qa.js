// 후보자 페이지 클라이언트 로직
//   1) PDF 미리보기: NEC CDN 은 application/x-pdf 로 보내서 iframe 직접 src 가 안 먹는다.
//      → fetch → Blob(application/pdf) → URL.createObjectURL → iframe.src
//   2) Q&A: POST /api/qa (Vercel 서버리스). 키는 서버의 GEMINI_API_KEY 에서 읽음.
// 데이터:  window.CANDIDATE = {huboId, name, party, ...}
//          window.PLEDGE_TEXT = "..."  (선택: 사전 추출 텍스트)
//          window.PLEDGE_PDFS = [{type, url}, ...]
(function(){
  const $ = sel => document.querySelector(sel);
  const $$ = sel => document.querySelectorAll(sel);

  // ----- PDF preview (blob URL) -----
  const blobCache = new Map();          // url → blob URL
  async function loadPdfIntoFrame(url){
    const frame = $('#pdfFrame');
    const loading = $('#pdfLoading');
    if (!frame) return;
    if (loading){ loading.classList.remove('hidden'); loading.textContent = 'PDF 불러오는 중…'; }
    try{
      if (!blobCache.has(url)){
        const r = await fetch(url);
        if (!r.ok) throw new Error('HTTP '+r.status);
        const buf = await r.arrayBuffer();
        const blob = new Blob([buf], { type: 'application/pdf' });
        blobCache.set(url, URL.createObjectURL(blob));
      }
      frame.src = blobCache.get(url);
      if (loading){
        setTimeout(() => loading.classList.add('hidden'), 200);
      }
    }catch(e){
      if (loading){ loading.textContent = 'PDF 로딩 실패: '+(e.message||e); }
    }
  }

  function setActiveTab(tab){
    $$('.pdf-tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    const url = tab.dataset.pdf;
    if (url) loadPdfIntoFrame(url);
    // download button updates to currently active PDF
    const dl = $('#pdfDownloadActive');
    if (dl){
      dl.href = url;
      dl.download = (window.CANDIDATE?.name || 'pledge') + '_' + tab.textContent.trim() + '.pdf';
      dl.querySelector('.label').textContent = tab.textContent.trim() + ' PDF 다운로드';
    }
  }

  // ----- Q&A (uses /api/qa) -----
  function addMsg(role, text, extraClass=''){
    const div = document.createElement('div');
    div.className = 'qa-msg ' + role + (extraClass ? ' '+extraClass : '');
    div.textContent = text;
    const msgs = $('#qaMsgs');
    msgs.appendChild(div);
    msgs.scrollTop = msgs.scrollHeight;
    return div;
  }

  async function ask(question){
    addMsg('user', question);
    const loading = addMsg('bot', '공약 자료를 읽고 답변 작성 중…', 'loading');
    $('#qaInput').value = '';
    $('#qaSend').disabled = true;

    try{
      const r = await fetch('/api/qa', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({
          question,
          candidate: window.CANDIDATE,
          pledgeText: window.PLEDGE_TEXT || '',
          pdfUrls: window.PLEDGE_PDFS || []
        })
      });
      const j = await r.json().catch(() => ({error:'응답 파싱 실패'}));
      loading.remove();
      if (!r.ok){
        addMsg('bot', j.error || ('서버 오류 '+r.status), 'error');
        return;
      }
      addMsg('bot', (j.answer || '').trim() || '(빈 응답)');
    } catch(e){
      loading.remove();
      addMsg('bot', '요청 실패: '+(e.message||e), 'error');
    } finally {
      $('#qaSend').disabled = false;
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    // PDF: first tab 자동 로드
    const first = $('.pdf-tab.active') || $('.pdf-tab');
    if (first) setActiveTab(first);
    $$('.pdf-tab').forEach(tab => tab.addEventListener('click', () => setActiveTab(tab)));

    // Q&A 폼
    const form = $('#qaForm');
    if (form){
      form.addEventListener('submit', e => {
        e.preventDefault();
        const q = $('#qaInput').value.trim();
        if (q) ask(q);
      });
      $$('.qa-chip').forEach(chip => {
        chip.addEventListener('click', () => ask(chip.textContent));
      });
    }
  });
})();