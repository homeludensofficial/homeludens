(function(){
/* ---------- 목록 정의 (엑셀 Data 시트) ---------- */
const OPT = {
  severity:['S (긴급)','A (높음)','B (보통)'],
  claimType:['감성형','현장성_경미','현장성_누수','현장성_화재','현장성_담배','현장성_기타','외부_건물설비','외부_관리사무소','외부_기타'],
  channel:['전화','방문','카카오톡','온라인','기타'],
  reportType:['최초 보고','중간 보고','종결 보고'],
  process:['철거','설비','목공','전기','타일','도장','도배','가구','마감','기타'],
  causeType:['시공팀','협력업체','자재','설계/견적','영업 설명 부족','고객 요인','기타'],
  action:['재시공','부분 보수','자재 교체','금액 보상','설명 및 협의','기타'],
  costBearer:['자사','협력업체','고객','분담'],
  satisfaction:['매우 만족','만족','보통','불만'],
  closed:['종결','추가 조치 필요'],
  shareTargets:['영업팀','시공팀','협력업체','전사']
};
const ROLES = ['팀원 (작성)','팀장','본부장','부대표','대표'];
const STD = {
  S:{grade:'중대', steps:5},
  A:{grade:'보통', steps:4},
  B:{grade:'경미', steps:3}
};
const COSTS = [['rework','재시공 / 보수 비용'],['material','자재비'],['comp','보상 (감액 / 환불)'],['etc','기타']];

/* ---------- 보고서 항목 (엑셀 클레임보고서 시트) ---------- */
const SECTIONS = [
  {no:'', title:'기본 정보', fields:[
    ['docNo','문서번호','text',{ph:'비워 두면 자동 부여 (CLM-2026-001)'}],
    ['createdAt','작성일자','date'],
    ['author','작성자','text',{ph:'(팀 소속) / (성명)'}],
    ['reportType','보고 구분','select'],
    ['customer','고객명','text'],
    ['address','현장 주소','text'],
    ['contractNo','계약번호','text'],
    ['contractAmount','계약금액 (원)','money'],
    ['periodStart','공사 시작일','date'],
    ['periodEnd','공사 종료일','date'],
    ['receivedAt','클레임 접수일','date'],
    ['channel','접수 경로','multi'],
    ['claimType','클레임 유형','multi',{wide:true}],
    ['severity','중요도','select',{wide:true}]
  ]},
  {no:'1', title:'사고 개요 및 상세 경위', guide:'사실 중심의 발생 상황을 육하원칙에 따라 기재해 주세요.', fields:[
    ['occurredAt','발생일시','datetime'],
    ['process','발생 공정','multi',{wide:true}],
    ['contractor','시공 업체','vendors',{wide:true}],
    ['summary','주요 사고내용','textarea',{wide:true}],
    ['schedFinish','기존일정 (마감)','date'],
    ['schedMoveIn','기존일정 (입주)','date'],
    ['customerRequest','고객 요청사항','textarea',{ph:'예: 데드라인 요구 / 보상 요구 / 태도 기타 등',wide:true}],
    ['firstAware','최초 인지 경로','textarea',{wide:true}],
    ['before','작업 전 상황','textarea',{wide:true}],
    ['progress','작업 진행 과정','textarea',{wide:true}],
    ['after','이후 대응 과정','textarea',{wide:true}]
  ]},
  {no:'2', title:'사고 원인 분석', guide:'직접적 원인과 더불어 시스템적 근본 원인을 분석해 주세요.', fields:[
    ['causeType','원인 구분','multi',{wide:true}],
    ['responsibility','책임 소재','text',{wide:true}],
    ['cause','사고 원인','textarea',{wide:true}]
  ]},
  {no:'3', title:'피해 내용', guide:'파손, 재시공 범위, 민원 등', fields:[
    ['damage','고객 피해내용','textarea',{wide:true}]
  ]},
  {no:'4', title:'대응 조치 및 결과', guide:'클레임 해결을 위해 취한 조치와 현재 상태를 기재해 주세요.', fields:[
    ['action','처리 방안','multi',{wide:true}],
    ['actionOwner','조치 담당','text'],
    ['dueDate','완료 예정일','date'],
    ['done','대응 완료 사항','textarea',{wide:true}],
    ['planned','조치 예정사항','textarea',{wide:true}],
    ['negotiation','고객 협의 결과','textarea',{wide:true}]
  ]},
  {no:'5', title:'비용 및 손실 산정', guide:'발생 비용을 항목별로 기재해 주세요. 합계는 자동 계산됩니다. (예: 욕실 방수 재시공 1,500,000)', costs:true, fields:[
    ['costBearer','비용 부담 주체','multi'],
    ['shareRatio','분담 비율','text',{ph:'자사 OO% / 협력업체 OO% / 고객 OO%'}]
  ]},
  {no:'6', title:'재발 방지 대책', guide:'공유 대상은 해당하는 곳을 모두 고르세요.', fields:[
    ['prevention','재발 방지 방안','textarea',{wide:true}],
    ['training','교육 / 프로세스 개선 사항','textarea',{wide:true}],
    ['shareTargets','공유 대상','multi',{wide:true}]
  ]},
  {no:'7', title:'처리 결과 (종결 보고)', guide:'종결 보고 시 작성해 주세요.', fields:[
    ['finalDate','최종 처리일','date'],
    ['satisfaction','고객 만족도','select'],
    ['result','처리 결과','textarea',{wide:true}],
    ['closed','종결 여부','select',{wide:true}]
  ]},
  {no:'', title:'결재 의견 (반려 사유)', fields:[
    ['approvalComment','결재 의견','textarea',{wide:true}]
  ]}
];

/* ---------- 상태 ---------- */
const S = {
  view:'list', id:null, claims:[], all:[], examples:[], loaded:false, cfg:{minor:'',major:''},
  f:{status:'전체', sev:'전체', type:'전체', q:'', mine:false},
  me:{role:null, isAdmin:false, name:''}, members:[], staff:false,
  session:null, authReady:false, email:null, canWrite:true, offline:false,
  photoUrls:{}, confirmDel:false, busy:false
};
const app = document.getElementById('app');

/* ---------- 유틸 ---------- */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const won = n => (n===''||n==null||isNaN(n)) ? '' : Number(n).toLocaleString('ko-KR');
const parseNum = v => { const n = Number(String(v??'').replace(/[^\d.-]/g,'')); return isFinite(n)&&String(v??'').trim()!=='' ? n : ''; };
const todayStr = () => { const d=new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); };
const sevKey = s => (s||'').charAt(0);
const costTotal = c => COSTS.reduce((a,[k]) => a + (Number(c?.costs?.[k]?.amt)||0), 0);
// 여러 개 고를 수 있는 항목: 예전 보고서는 문자열 하나, 새 보고서는 배열
const toArr = v => Array.isArray(v) ? v.filter(Boolean) : (v ? [v] : []);
const joinV = v => toArr(v).join(', ');
// 시공 업체 여러 곳: 새 보고서는 contractors 배열, 예전 보고서는 '업체 / 시공자 / 전화' 글자 하나
function vendorsOf(c){
  if(Array.isArray(c.contractors) && c.contractors.length) return c.contractors;
  if(!c.contractor) return [];
  return String(c.contractor).split(/\s*;\s*/).filter(Boolean).map(t=>{ const [name='',worker='',phone=''] = t.split(/\s*\/\s*/); return {name,worker,phone}; });
}
const vendorLine = v => [v.name,v.worker,v.phone].map(x=>String(x||'').trim()).filter(Boolean).join(' / ');
const statusOf = c => c.closed==='종결' ? '종결' : (toArr(c.action).length ? '처리중' : '접수');
function elapsed(c){
  if(!c.receivedAt) return '';
  const start = new Date(c.receivedAt+'T00:00:00');
  const end = (c.closed==='종결' && c.finalDate) ? new Date(c.finalDate+'T00:00:00') : new Date(todayStr()+'T00:00:00');
  const d = Math.round((end-start)/86400000);
  return isNaN(d) ? '' : d;
}
function fmtDate(s){ if(!s) return ''; return String(s).replace('T',' '); }
function toast(msg){ const t=document.getElementById('toast'); t.textContent=msg; t.hidden=false; clearTimeout(toast._t); toast._t=setTimeout(()=>t.hidden=true,2600); }
function stdText(k){
  const mi = S.cfg.minor!=='' && S.cfg.minor!=null ? S.cfg.minor : 'OO';
  const ma = S.cfg.major!=='' && S.cfg.major!=null ? S.cfg.major : 'OO';
  if(k==='S') return `보수 비용 ${ma}만 원 이상, 재시공, 법적 분쟁 우려`;
  if(k==='A') return `보수 비용 ${mi}만 원 이상 ${ma}만 원 미만`;
  return `보수 비용 ${mi}만 원 미만, 단순 응대·설명`;
}
const routeOf = k => STD[k] ? ROLES.slice(0,STD[k].steps).map(r=>r.replace(' (작성)','')).join(' → ') : '';
function suggestSev(totalWon){
  const mi = Number(S.cfg.minor), ma = Number(S.cfg.major);
  if(!totalWon || !S.cfg.minor || !S.cfg.major) return '';
  const man = totalWon/10000;
  return man>=ma ? 'S (긴급)' : man>=mi ? 'A (높음)' : 'B (보통)';
}
function nextDocNo(){
  const y = new Date().getFullYear();
  const re = new RegExp('^CLM-'+y+'-(\\d+)$');
  const max = S.claims.reduce((m,c)=>{ const r=re.exec(c.docNo||''); return r?Math.max(m,+r[1]):m; },0);
  return `CLM-${y}-${String(max+1).padStart(3,'0')}`;
}

/* ---------- 화면 전환 ---------- */
function go(view,id){ S.view=view; S.id=id||null; S.confirmDel=false; render(); window.scrollTo(0,0); }
document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>go(b.dataset.view)));
document.getElementById('btn-new').addEventListener('click',()=>go('edit'));

function render(){
  document.getElementById('nav').hidden = !S.session;
  document.getElementById('who').hidden = !S.session;
  if(!S.session){ renderLogin(); return; }
  if(S.loaded && !S.staff){ document.getElementById('nav').hidden = true; document.getElementById('btn-new').hidden = true; renderBlocked(); return; }
  document.querySelectorAll('.tab').forEach(b=>b.setAttribute('aria-current', String(b.dataset.view===S.view || (b.dataset.view==='list' && (S.view==='doc'||S.view==='edit')))));
  document.getElementById('btn-new').hidden = !S.canWrite || !S.session;
  if(S.view==='list') renderList();
  else if(S.view==='doc') renderDoc();
  else if(S.view==='edit') renderEdit();
  else if(S.view==='std') renderStd();
}

/* ---------- 관리대장 ---------- */
function roleOfTitle(t){
  t = String(t||'').trim();
  if(t.startsWith('부대표')) return '부대표';
  if(t.startsWith('대표')) return '대표';
  if(t.startsWith('본부장')) return '본부장';
  if(t.startsWith('팀장')) return '팀장';
  return null;
}
function memberName(email){ const m = S.members.find(x=>x.email===String(email||'').toLowerCase()); return m?.name || String(email||'').split('@')[0]; }
function needsMe(c){
  const r = S.me.role; if(!r) return false;
  const steps = STD[sevKey(c.severity)]?.steps || 0;
  return ROLES.indexOf(r) > 0 && ROLES.indexOf(r) < steps && !c.approvals?.[r]?.date;
}
function filtered(){
  const q = S.f.q.trim().toLowerCase();
  return S.claims.filter(c=>{
    if(S.f.mine && !needsMe(c)) return false;
    if(S.f.status!=='전체' && statusOf(c)!==S.f.status) return false;
    if(S.f.sev!=='전체' && sevKey(c.severity)!==S.f.sev) return false;
    if(S.f.type!=='전체' && !toArr(c.claimType).includes(S.f.type)) return false;
    if(q){
      const hay = [c.docNo,c.customer,c.address,c.author,c.summary,c.contractor,c.contractNo].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  }).sort((a,b)=> String(b.receivedAt||b.createdAt||'').localeCompare(String(a.receivedAt||a.createdAt||'')) || String(b.docNo||'').localeCompare(String(a.docNo||'')));
}
function segHTML(name,vals,cur){
  return `<div class="seg" role="group" aria-label="${esc(name)}">`+vals.map(v=>`<button type="button" data-seg="${esc(name)}" data-v="${esc(v)}" aria-pressed="${v===cur}">${esc(v)}</button>`).join('')+`</div>`;
}
function renderList(){
  const rows = filtered();
  const open = rows.filter(c=>statusOf(c)!=='종결').length;
  const sumContract = rows.reduce((a,c)=>a+(Number(c.contractAmount)||0),0);
  const sumCost = rows.reduce((a,c)=>a+costTotal(c),0);
  let body;
  if(!S.loaded){
    body = `<div class="empty"><h3>관리대장을 불러오는 중입니다</h3><p>잠시만 기다려 주세요.</p></div>`;
  } else if(S.claims.length===0){
    body = `<div class="empty"><h3>아직 등록된 클레임이 없습니다</h3><p>보고서를 작성하면 이 관리대장에 1건당 1행으로 자동 기록됩니다.${S.examples.length?' 위의 작성 예시를 참고하세요.':''}</p>${S.canWrite?'<button class="btn primary" data-act="new">첫 클레임 보고서 작성</button>':''}</div>`;
  } else if(rows.length===0){
    body = `<div class="empty"><h3>조건에 맞는 클레임이 없습니다</h3><p>필터를 바꾸거나 검색어를 지워 보세요.</p><button class="btn" data-act="reset">필터 초기화</button></div>`;
  } else {
    body = `<div class="tbl-wrap"><table class="ledger"><thead><tr>
      <th>문서번호</th><th>고객 / 현장</th><th>접수일</th><th>클레임 유형</th><th>중요도</th><th>처리 방안</th><th class="num">비용 합계(원)</th><th>상태</th><th class="num">경과일</th>
    </tr></thead><tbody>`+rows.map(c=>{
      const k=sevKey(c.severity), st=statusOf(c), el=elapsed(c);
      return `<tr data-open="${esc(c._id)}" tabindex="0">
        <td class="mono">${esc(c.docNo||'—')}</td>
        <td class="cust"><b>${esc(c.customer||'(고객명 미입력)')}</b><span>${esc(c.address||'')}</span></td>
        <td data-c="hide" class="num">${esc(c.receivedAt||'')}</td>
        <td data-c="hide">${esc(joinV(c.claimType))}</td>
        <td>${k?`<span class="pill sev-${k}">${esc(c.severity)}</span>`:''}</td>
        <td data-c="hide">${esc(joinV(c.action))}</td>
        <td class="num" data-c="hide">${costTotal(c)?won(costTotal(c)):''}</td>
        <td><span class="pill st-${st}">${st}</span></td>
        <td class="num ${st!=='종결'&&el>=14?'over':''}">${el===''?'':el+'일'}</td>
      </tr>`;}).join('')+`</tbody></table></div>`;
  }
  app.innerHTML = `
    <h1 class="page">클레임 관리대장</h1>
    <p class="sub">행을 누르면 보고서가 열립니다. 위 합계는 걸러 낸 결과만 계산합니다.${S.offline?' · 저장소에 연결되지 않아 이 화면에서는 저장되지 않습니다.':''}</p>
    <div class="kpis">
      <div class="kpi"><small>표시 건수</small><strong>${rows.length}<em>건</em></strong></div>
      <div class="kpi"><small>처리 중 (미종결)</small><strong>${open}<em>건</em></strong></div>
      <div class="kpi"><small>계약금액 합계</small><strong>${won(sumContract)}<em>원</em></strong></div>
      <div class="kpi"><small>비용 합계</small><strong>${won(sumCost)}<em>원</em></strong></div>
    </div>
    ${S.examples.length?`<div class="ex-strip"><div class="ex-top"><span class="ex-label">작성 예시</span><span class="ex-hint">처음 쓰는 분은 예시를 먼저 열어 보세요. 합계와 대장에는 포함되지 않습니다.</span></div><div class="ex-cards">${S.examples.map(e=>`<button class="ex-card" data-open="${esc(e._id)}"><b>${esc(joinV(e.claimType)||'예시')} · ${esc(joinV(e.process))}</b><span>${esc((e.summary||'').slice(0,48))}${(e.summary||'').length>48?'…':''}</span></button>`).join('')}</div></div>`:''}
    <div class="filters">
      ${S.me.role?`<button type="button" class="btn ${S.f.mine?'primary':''}" data-act="mine" aria-pressed="${S.f.mine}">내 결재 대기 ${S.claims.filter(needsMe).length}건</button>`:''}
      <span><span class="flabel">처리 상태</span>${segHTML('status',['전체','접수','처리중','종결'],S.f.status)}</span>
      <span><span class="flabel">중요도</span>${segHTML('sev',['전체','S','A','B'],S.f.sev)}</span>
      <span><span class="flabel">유형</span><select class="sm" id="f-type"><option>전체</option>${OPT.claimType.map(t=>`<option ${t===S.f.type?'selected':''}>${esc(t)}</option>`).join('')}</select></span>
      <input class="search" id="f-q" type="search" placeholder="문서번호, 고객명, 주소, 작성자 검색" value="${esc(S.f.q)}">
      ${S.claims.length?'<button class="btn ghost" data-act="csv">CSV 내보내기</button>':''}
    </div>
    ${body}`;
  app.querySelectorAll('[data-seg]').forEach(b=>b.addEventListener('click',()=>{ S.f[b.dataset.seg]=b.dataset.v; renderList(); }));
  app.querySelector('#f-type').addEventListener('change',e=>{ S.f.type=e.target.value; renderList(); });
  const q = app.querySelector('#f-q');
  q.addEventListener('input',e=>{ S.f.q=e.target.value; clearTimeout(q._t); q._t=setTimeout(()=>{ renderList(); const n=app.querySelector('#f-q'); n.focus(); n.setSelectionRange(n.value.length,n.value.length); },250); });
  app.querySelectorAll('[data-open]').forEach(tr=>{
    tr.addEventListener('click',()=>go('doc',tr.dataset.open));
    tr.addEventListener('keydown',e=>{ if(e.key==='Enter') go('doc',tr.dataset.open); });
  });
  app.querySelector('[data-act="new"]')?.addEventListener('click',()=>go('edit'));
  app.querySelector('[data-act="reset"]')?.addEventListener('click',()=>{ S.f={status:'전체',sev:'전체',type:'전체',q:'',mine:false}; renderList(); });
  app.querySelector('[data-act="mine"]')?.addEventListener('click',()=>{ S.f.mine=!S.f.mine; renderList(); });
  app.querySelector('[data-act="csv"]')?.addEventListener('click',()=>exportCSV(rows));
}

function exportCSV(rows){
  const head = ['No','문서번호','작성일자','작성자','보고 구분','고객명','현장 주소','계약번호','계약금액(원)','공사 시작일','공사 종료일','접수일','접수 경로','클레임 유형','중요도','결재 범위','발생일시','발생 공정','시공 업체','주요 사고내용','원인 구분','책임 소재','처리 방안','조치 담당','완료 예정일','비용 합계(원)','비용 부담 주체','최종 처리일','고객 만족도','종결 여부','처리 상태','경과일'];
  const q = v => { const s=String(v??''); return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s; };
  const lines = [head.join(',')].concat(rows.map((c,i)=>[i+1,c.docNo,c.createdAt,c.author,c.reportType,c.customer,c.address,c.contractNo,c.contractAmount,c.periodStart,c.periodEnd,c.receivedAt,joinV(c.channel),joinV(c.claimType),c.severity,routeOf(sevKey(c.severity)),fmtDate(c.occurredAt),joinV(c.process),c.contractor,c.summary,joinV(c.causeType),c.responsibility,joinV(c.action),c.actionOwner,c.dueDate,costTotal(c)||'',joinV(c.costBearer),c.finalDate,c.satisfaction,c.closed,statusOf(c),elapsed(c)].map(q).join(',')));
  const blob = new Blob(['\ufeff'+lines.join('\r\n')], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `클레임_관리대장_${todayStr()}.csv`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
}

/* ---------- 보고서 보기 ---------- */
function valHTML(c,key,type){
  let v = c[key];
  if(type==='money') v = v!==''&&v!=null ? won(v)+' 원' : '';
  else if(type==='datetime') v = fmtDate(v);
  else if(type==='multi') v = joinV(v);
  else if(type==='vendors') v = vendorsOf(c).map(vendorLine).filter(Boolean).join('\n');
  return v ? `<dd>${esc(v)}</dd>` : `<dd class="blank">—</dd>`;
}
async function renderDoc(){
  const c = S.all.find(x=>x._id===S.id);
  if(!c){ app.innerHTML = `<div class="empty" style="margin-top:28px"><h3>보고서를 찾을 수 없습니다</h3><p>삭제되었거나 아직 불러오는 중입니다.</p><button class="btn" data-act="back">관리대장으로</button></div>`; app.querySelector('[data-act="back"]').onclick=()=>go('list'); return; }
  const canW = S.canWrite && !c.example;
  const k = sevKey(c.severity), st = statusOf(c), steps = STD[k]?.steps || 0;
  const ap = c.approvals || {};
  const apCells = ROLES.map((r,i)=>{
    if(i===0) return `<td><div class="stamp-btn" style="cursor:default"><span class="stamp">작성</span><span class="stamp-date">${esc(c.createdAt||'')}</span></div></td>`;
    if(!steps || i>=steps) return `<td class="off"><span>${steps?'해당 없음':'중요도 선택'}</span></td>`;
    const a = ap[r];
    const mine = canW && (S.me.role===r || S.me.isAdmin);
    if(a?.date){
      const nm = a.by ? memberName(a.by) : '';
      const inner = `<span class="stamp">${esc(r.slice(0,2))}</span><span class="stamp-date">${esc(a.date.slice(5).replace('-','/'))}${nm?' '+esc(nm):''}</span>`;
      return mine ? `<td><button class="stamp-btn" data-stamp="${esc(r)}" title="누르면 결재를 취소합니다">${inner}</button></td>` : `<td><div class="stamp-btn" style="cursor:default">${inner}</div></td>`;
    }
    return mine ? `<td><button class="stamp-btn mine" data-stamp="${esc(r)}">결재하기</button></td>` : `<td><div class="stamp-btn" style="cursor:default">결재 대기</div></td>`;
  }).join('');
  const secs = SECTIONS.map(sec=>{
    const kv = sec.fields.map(([key,label,type,o])=>`<div class="${o?.wide||type==='textarea'?'wide':''}"><dt>${esc(label)}</dt>${valHTML(c,key,type)}</div>`).join('');
    let extra = '';
    if(sec.costs){
      extra = `<table class="cost"><thead><tr><th>항목</th><th class="num">금액 (원)</th><th>비고</th></tr></thead><tbody>`+
        COSTS.map(([ck,cl])=>`<tr><td>${cl}</td><td class="num">${won(c.costs?.[ck]?.amt)||'—'}</td><td>${esc(c.costs?.[ck]?.note||'')}</td></tr>`).join('')+
        `<tr class="total"><td>합계</td><td class="num">${won(costTotal(c))}</td><td></td></tr></tbody></table><div style="height:10px"></div>`;
    }
    return `<section class="sec"><h2>${sec.no?`<i>${sec.no}.</i>`:''}${esc(sec.title)}</h2>${extra}<dl class="kv">${kv}</dl></section>`;
  }).join('');
  const photos = (c.photos||[]);
  app.innerHTML = `
    <div class="back"><button class="btn ghost" data-act="back">← 관리대장</button></div>
    ${c.example?`<div class="ex-note"><b>작성 예시</b><span>실제 클레임이 아니라 작성 방법을 보여 주는 참고용 보고서입니다. 항목마다 어느 정도로 쓰면 되는지 보고 <button class="linkish" data-act="new-from-ex">새 보고서 작성</button>을 시작하세요. 이 예시는 수정·결재·삭제되지 않습니다.</span></div>`:''}
    <article class="doc">
      <div class="doc-head">
        <div>
          <h1 class="doc-title">클레임 보고서</h1>
          <div class="doc-meta">
            ${c.example?'<span class="pill ex">작성 예시</span>':''}<span class="mono">${esc(c.docNo||'')}</span>
            ${k?`<span class="pill sev-${k}">${esc(c.severity)} · ${STD[k].grade}</span>`:''}
            <span class="pill st-${st}">${st}</span>
            ${c.reportType?`<span style="color:var(--ink-3);font-size:12px">${esc(c.reportType)}</span>`:''}
          </div>
        </div>
        <div style="overflow-x:auto;max-width:100%">
          <table class="approve"><tr>${ROLES.map(r=>`<th>${esc(r)}</th>`).join('')}</tr><tr>${apCells}</tr></table>
          <div class="approve-route">${steps?'결재 범위: '+esc(routeOf(k)):'중요도를 선택하면 결재 범위가 정해집니다'}</div>
        </div>
      </div>
      ${secs}
      <section class="sec"><h2><i>8.</i>사고 현장 증빙</h2><p class="guide">사진을 누르면 원본이 새 탭에서 열립니다.</p>
        <div class="photos">
          ${photos.map(p=>`<div class="photo"><a href="${esc(S.photoUrls[p]||'#')}" target="_blank" rel="noopener"><img data-photo="${esc(p)}" ${S.photoUrls[p]?`src="${esc(S.photoUrls[p])}"`:''} alt="현장 사진" loading="lazy"></a>${canW?`<button data-rmphoto="${esc(p)}">빼기</button>`:''}</div>`).join('')}
          ${canW?`<label class="dropzone" for="photo-in">${S.busy?'올리는 중…':'+ 사진 추가<br>(여러 장 선택 가능)'}</label><input id="photo-in" type="file" accept="image/*" multiple hidden>`:''}
          ${!photos.length && !(canW)?'<p class="guide">첨부된 사진이 없습니다.</p>':''}
        </div>
      </section>
      <p class="sign">위와 같이 사고 발생에 대한 경위 및 대책을 보고합니다.<br><span style="font-weight:400;font-size:13px">${esc(c.createdAt||'')} · 보고자 ${esc(c.author||'')}</span></p>
    </article>
    <div class="doc-actions">
      ${canW?'<button class="btn primary" data-act="edit">보고서 수정</button>':''}
      <span style="flex:1"></span>
      ${canW?(S.confirmDel?'<span class="confirm">이 보고서를 삭제할까요? <button class="btn danger" data-act="del-yes">삭제</button><button class="btn" data-act="del-no">취소</button></span>':'<button class="btn ghost" data-act="del">삭제</button>'):''}
    </div>`;
  loadPhotoUrls(photos);
  app.querySelector('[data-act="back"]').onclick=()=>go('list');
  app.querySelector('[data-act="new-from-ex"]')?.addEventListener('click',()=>go('edit'));
  app.querySelector('[data-act="edit"]')?.addEventListener('click',()=>go('edit',c._id));
  app.querySelector('[data-act="del"]')?.addEventListener('click',()=>{ S.confirmDel=true; renderDoc(); });
  app.querySelector('[data-act="del-no"]')?.addEventListener('click',()=>{ S.confirmDel=false; renderDoc(); });
  app.querySelector('[data-act="del-yes"]')?.addEventListener('click',async()=>{
    try{ await api.remove(c._id); toast('삭제했습니다.'); go('list'); }catch(e){ toast('삭제하지 못했습니다. 권한을 확인해 주세요.'); }
  });
  app.querySelectorAll('[data-stamp]').forEach(b=>b.addEventListener('click',async()=>{
    const r = b.dataset.stamp; const next = Object.assign({}, c.approvals||{});
    if(next[r]?.date) delete next[r]; else next[r] = {date:todayStr(), by:S.email||null};
    await patch(c._id,{approvals:next}, next[r]?`${r} 결재를 기록했습니다.`:`${r} 결재를 취소했습니다.`);
  }));
  app.querySelectorAll('[data-rmphoto]').forEach(b=>b.addEventListener('click',async e=>{
    e.preventDefault();
    const p = b.dataset.rmphoto;
    await patch(c._id,{photos:(c.photos||[]).filter(x=>x!==p)},'사진을 뺐습니다.');
    api.removePhotos([p]).catch(()=>{});
  }));
  app.querySelector('#photo-in')?.addEventListener('change',async e=>{
    const files=[...e.target.files]; if(!files.length) return;
    S.busy=true; renderDoc();
    const added=[];
    for(const f of files){
      try{ added.push(await api.uploadPhoto(c._id, f)); }
      catch(err){ toast(err?.tooLarge?'20MB보다 큰 사진은 올릴 수 없습니다.':'사진 한 장을 올리지 못했습니다.'); }
    }
    S.busy=false;
    const cur = S.claims.find(x=>x._id===c._id);
    if(added.length) await patch(c._id,{photos:[...(cur?.photos||[]),...added]},`사진 ${added.length}장을 추가했습니다.`);
    else renderDoc();
  });
}
async function patch(id,data,msg){
  try{ await api.update(id, Object.assign({updatedAt:new Date().toISOString()},data)); toast(msg); }
  catch(e){ toast(e?.code==='42501' && e?.message ? e.message : '저장하지 못했습니다. 권한을 확인해 주세요.'); }
}

/* ---------- 작성 / 수정 ---------- */
function vendorRowHTML(v={}){
  return `<div class="vendor-row">
    <input type="text" data-v="name" value="${esc(v.name||'')}" placeholder="업체명" aria-label="업체명">
    <input type="text" data-v="worker" value="${esc(v.worker||'')}" placeholder="시공자" aria-label="시공자">
    <input type="text" data-v="phone" value="${esc(v.phone||'')}" placeholder="전화번호" inputmode="tel" aria-label="전화번호">
    <button type="button" class="btn ghost" data-act="vendor-del" aria-label="이 업체 빼기">×</button>
  </div>`;
}
function fieldHTML(c,[key,label,type,o]){
  const v = c[key] ?? '';
  const id = 'f_'+key, wide = (o?.wide||type==='textarea')?' wide':'';
  const ph = o?.ph ? ` placeholder="${esc(o.ph)}"` : '';
  let input;
  if(type==='select'){
    const opts = OPT[key];
    if(opts.length<=6){
      input = `<div class="chips" role="radiogroup" aria-label="${esc(label)}">`+opts.map((op,i)=>`<label class="chip"><input type="radio" name="${id}" id="${id}_${i}" value="${esc(op)}" ${op===v?'checked':''}><span>${esc(op)}</span></label>`).join('')+`</div>`;
    } else {
      input = `<select id="${id}"><option value="">선택하세요</option>${opts.map(op=>`<option ${op===v?'selected':''}>${esc(op)}</option>`).join('')}</select>`;
    }
    if(key==='severity') input += `<div class="auto" id="sev-auto"></div>`;
    return `<div class="fld${wide}"><span class="lbl">${esc(label)}</span>${input}</div>`;
  }
  if(type==='multi'){
    const arr = toArr(v);
    input = `<div class="chips">`+OPT[key].map((op,i)=>`<label class="chip"><input type="checkbox" name="${id}" id="${id}_${i}" value="${esc(op)}" ${arr.includes(op)?'checked':''}><span>${esc(op)}</span></label>`).join('')+`</div>`;
    return `<div class="fld${wide}"><span class="lbl">${esc(label)} <span class="guide" style="font-weight:400">· 여러 개 선택 가능</span></span>${input}</div>`;
  }
  if(type==='vendors'){
    const list = vendorsOf(c);
    return `<div class="fld${wide}"><span class="lbl">${esc(label)} <span class="guide" style="font-weight:400">· 여러 곳이면 + 업체 추가</span></span>
      <div class="vendors" id="${id}">${(list.length?list:[{}]).map(vendorRowHTML).join('')}</div>
      <div><button type="button" class="btn ghost" data-act="vendor-add">+ 업체 추가</button></div></div>`;
  }
  if(type==='textarea') input = `<textarea id="${id}"${ph}>${esc(v)}</textarea>`;
  else if(type==='date') input = `<input type="date" id="${id}" value="${esc(v)}">`;
  else if(type==='datetime') input = `<input type="datetime-local" id="${id}" value="${esc(v)}">`;
  else if(type==='money') input = `<input type="text" inputmode="numeric" id="${id}" value="${esc(won(v))}" data-money style="text-align:right">`;
  else input = `<input type="text" id="${id}" value="${esc(v)}"${ph}>`;
  return `<div class="fld${wide}"><label for="${id}">${esc(label)}</label>${input}</div>`;
}
function renderEdit(){
  const editing = S.id ? S.claims.find(x=>x._id===S.id) : null;
  const c = editing || {createdAt:todayStr(), receivedAt:todayStr(), reportType:'최초 보고'};
  app.innerHTML = `
    <div class="back"><button class="btn ghost" data-act="cancel">← ${editing?'보고서로':'관리대장으로'}</button></div>
    <h1 class="page">${editing?'보고서 수정':'새 클레임 보고서'}</h1>
    <p class="sub">${editing?esc(c.docNo||''):'필요한 칸만 채워도 저장할 수 있습니다. 사진과 결재는 저장 후 보고서 화면에서 처리합니다.'}${!editing&&S.examples.length?' <button type="button" class="linkish" data-act="see-ex">작성 예시 보기</button>':''}</p>
    <form id="frm" novalidate>
      ${SECTIONS.map(sec=>`<section class="form-sec"><h2>${sec.no?`<i>${sec.no}.</i>`:''}${esc(sec.title)}</h2>${sec.guide?`<p class="guide">${esc(sec.guide)}</p>`:''}
        ${sec.costs?`<table class="cost-in"><thead><tr><th style="width:34%">항목</th><th style="width:28%">금액 (원)</th><th>비고</th></tr></thead><tbody>`+
          COSTS.map(([ck,cl])=>`<tr><td>${cl}</td><td><input class="money" inputmode="numeric" id="c_${ck}_amt" data-money value="${esc(won(c.costs?.[ck]?.amt))}" aria-label="${cl} 금액"></td><td><input id="c_${ck}_note" value="${esc(c.costs?.[ck]?.note||'')}" aria-label="${cl} 비고"></td></tr>`).join('')+
          `<tr class="total"><td>합계</td><td class="num" id="cost-total">0</td><td id="cost-sug" style="font-weight:400;color:var(--ink-3);font-size:12px"></td></tr></tbody></table>`:''}
        <div class="grid">${sec.fields.map(f=>fieldHTML(c,f)).join('')}</div></section>`).join('')}
      <div class="save-bar">
        <button type="button" class="btn" data-act="cancel">취소</button>
        <button type="submit" class="btn primary" id="btn-save">${editing?'변경 내용 저장':'보고서 저장'}</button>
      </div>
    </form>`;
  const frm = app.querySelector('#frm');
  app.querySelectorAll('[data-act="cancel"]').forEach(b=>b.onclick=()=>editing?go('doc',editing._id):go('list'));
  app.querySelector('[data-act="see-ex"]')?.addEventListener('click',()=>go('doc',S.examples[0]._id));
  frm.querySelectorAll('[data-money]').forEach(inp=>inp.addEventListener('input',()=>{
    const n = parseNum(inp.value); const pos = inp.value.length - inp.selectionStart;
    inp.value = n===''?'':won(n); const p = Math.max(0,inp.value.length-pos); inp.setSelectionRange(p,p); updateAuto();
  }));
  frm.addEventListener('change',updateAuto);
  frm.addEventListener('click',e=>{
    const add = e.target.closest('[data-act="vendor-add"]');
    if(add){ const box = frm.querySelector('#f_contractor'); box.insertAdjacentHTML('beforeend', vendorRowHTML()); box.lastElementChild.querySelector('input').focus(); return; }
    const del = e.target.closest('[data-act="vendor-del"]');
    if(del){ const box = frm.querySelector('#f_contractor'); const row = del.closest('.vendor-row');
      if(box.children.length>1) row.remove(); else row.querySelectorAll('input').forEach(i=>i.value=''); }
  });
  function updateAuto(){
    const tot = COSTS.reduce((a,[ck])=>a+(Number(parseNum(frm.querySelector('#c_'+ck+'_amt').value))||0),0);
    frm.querySelector('#cost-total').textContent = won(tot);
    const sug = suggestSev(tot);
    frm.querySelector('#cost-sug').textContent = sug ? `금액 기준 권장 중요도: ${sug}` : '';
    const sev = frm.querySelector('input[name="f_severity"]:checked')?.value || '';
    const k = sevKey(sev);
    frm.querySelector('#sev-auto').innerHTML = k ? `결재 범위 (자동): <b>${esc(routeOf(k))}</b><br><span style="color:var(--ink-3)">${esc(STD[k].grade)} · ${esc(stdText(k))}</span>` : '중요도를 선택하면 결재 범위가 자동으로 표시됩니다.';
  }
  updateAuto();
  frm.addEventListener('submit',async e=>{
    e.preventDefault();
    if(S.offline){ toast('데이터베이스에 연결되지 않아 저장할 수 없습니다.'); return; }
    const data = {};
    SECTIONS.forEach(sec=>sec.fields.forEach(([key,,type])=>{
      if(type==='select') data[key] = (frm.querySelector(`[name="f_${key}"]:checked`)?.value) ?? (frm.querySelector('#f_'+key)?.value||'');
      else if(type==='multi') data[key] = [...frm.querySelectorAll(`[name="f_${key}"]:checked`)].map(x=>x.value);
      else if(type==='money') data[key] = parseNum(frm.querySelector('#f_'+key).value);
      else if(type==='vendors'){
        data.contractors = [...frm.querySelectorAll('#f_'+key+' .vendor-row')].map(r=>({
          name:r.querySelector('[data-v="name"]').value.trim(),
          worker:r.querySelector('[data-v="worker"]').value.trim(),
          phone:r.querySelector('[data-v="phone"]').value.trim()
        })).filter(v=>v.name||v.worker||v.phone);
        data[key] = data.contractors.map(vendorLine).join('; ');   // 검색·CSV·예전 화면용 글자
      }
      else data[key] = frm.querySelector('#f_'+key).value.trim();
    }));
    data.costs = {};
    COSTS.forEach(([ck])=>{ data.costs[ck] = {amt:parseNum(frm.querySelector('#c_'+ck+'_amt').value), note:frm.querySelector('#c_'+ck+'_note').value.trim()}; });
    data.costTotal = COSTS.reduce((a,[ck])=>a+(Number(data.costs[ck].amt)||0),0);
    if(!data.docNo) data.docNo = editing?.docNo || nextDocNo();
    data.updatedAt = new Date().toISOString();
    const btn = frm.querySelector('#btn-save'); btn.disabled = true; btn.textContent='저장 중…';
    try{
      if(editing){ await api.update(editing._id, data); toast('저장했습니다.'); go('doc',editing._id); }
      else { data.photos=[]; data.approvals={}; data.createdBy=S.email||null; const newId = await api.insert(data); toast(`${data.docNo} 보고서를 등록했습니다.`); go('doc',newId); }
    }catch(err){
      btn.disabled=false; btn.textContent= editing?'변경 내용 저장':'보고서 저장';
      toast('저장하지 못했습니다. 인터넷 연결을 확인하고 다시 시도해 주세요.');
    }
  });
}

/* ---------- 결재 기준 ---------- */
function renderStd(){
  app.innerHTML = `
    <h1 class="page">결재 기준 (중요도별)</h1>
    <p class="sub">금액 기준을 입력하면 아래 표와 보고서의 권장 중요도에 바로 반영됩니다. 비어 있으면 OO로 표시됩니다.</p>
    <section class="form-sec">
      <h2>금액 기준 입력</h2>
      <div class="set-grid" style="margin-top:12px">
        <div class="fld"><label for="cfg-minor">경미 / 보통 경계 (만 원)</label><input type="number" id="cfg-minor" min="0" value="${esc(S.cfg.minor??'')}" ${S.me.isAdmin?'':'disabled'}></div>
        <div class="fld"><label for="cfg-major">보통 / 중대 경계 (만 원)</label><input type="number" id="cfg-major" min="0" value="${esc(S.cfg.major??'')}" ${S.me.isAdmin?'':'disabled'}></div>
      </div>
      ${S.me.isAdmin?'<div style="margin-top:14px"><button class="btn primary" id="cfg-save">기준 저장</button></div>':'<p class="guide" style="margin-top:12px">금액 기준은 관리자만 바꿀 수 있습니다.</p>'}
    </section>
    <section class="form-sec">
      <h2>기준표</h2>
      <div style="overflow-x:auto"><table class="std"><thead><tr><th>중요도</th><th>클레임 등급</th><th>기준</th><th>결재 범위</th><th class="num">결재 단계</th></tr></thead><tbody>
        ${OPT.severity.map(s=>{const k=sevKey(s);return `<tr><td><span class="pill sev-${k}">${esc(s)}</span></td><td>${STD[k].grade}</td><td>${esc(stdText(k))}</td><td>${esc(routeOf(k))}</td><td class="num">${STD[k].steps}명</td></tr>`;}).join('')}
      </tbody></table></div>
      <p class="guide" style="margin-top:12px">보고서에서 중요도를 고르면 결재 범위가 자동으로 정해지고, 결재가 필요 없는 칸은 빗금으로 바뀝니다. 결재자는 보고서 화면의 결재란을 눌러 결재합니다.</p>
    </section>`;
  app.querySelector('#cfg-save')?.addEventListener('click',async()=>{
    const mi = app.querySelector('#cfg-minor').value, ma = app.querySelector('#cfg-major').value;
    if(mi!=='' && ma!=='' && Number(mi)>=Number(ma)){ toast('보통 / 중대 경계는 경미 / 보통 경계보다 커야 합니다.'); return; }
    try{ await api.saveCfg({minor: mi===''?'':Number(mi), major: ma===''?'':Number(ma)}); toast('결재 기준을 저장했습니다.'); }
    catch(e){ toast('저장하지 못했습니다. 권한을 확인해 주세요.'); }
  });
}

function renderBlocked(){
  app.innerHTML = `<div class="login"><h1 class="page">접속할 수 없는 계정입니다</h1>
    <p class="sub">${esc(S.email||'')} 은(는) 업무지원포털 직원 명단에 없거나 재직 중이 아닙니다. 포털에 등록된 이메일로 다시 로그인하거나 관리자에게 문의하세요.</p>
    <button class="btn" id="blk-out">다른 이메일로 로그인</button></div>`;
  app.querySelector('#blk-out').onclick = ()=>sb.auth.signOut();
}

/* ---------- 로그인 ---------- */
function renderLogin(){
  if(!S.authReady){ app.innerHTML = `<div class="empty" style="margin-top:28px"><h3>불러오는 중입니다</h3></div>`; return; }
  const dom = DOMAIN ? '@'+DOMAIN : '';
  app.innerHTML = `
    <div class="login">
      <h1 class="page">로그인</h1>
      <p class="sub">업무지원포털에 등록된 이메일로 6자리 로그인 코드를 받아 접속합니다. 비밀번호는 없습니다.</p>
      <form id="lf" class="form-sec" novalidate>
        <div class="fld"><label for="login-email">이메일 (업무지원포털과 같은 주소)</label><input type="text" id="login-email" inputmode="email" autocomplete="email" placeholder="name@homeludens.co.kr"></div>
        <button type="submit" class="btn primary" id="login-send" style="margin-top:12px">로그인 코드 받기</button>
      </form>
      <form id="cf" class="form-sec" novalidate hidden>
        <p class="guide" id="code-msg"></p>
        <div class="fld"><label for="login-code">메일로 받은 6자리 코드</label><input type="text" id="login-code" inputmode="numeric" autocomplete="one-time-code" maxlength="10"></div>
        <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap"><button type="submit" class="btn primary">확인</button><button type="button" class="btn ghost" id="code-back">이메일 다시 입력</button></div>
      </form>
    </div>`;
  const lf = app.querySelector('#lf'), cf = app.querySelector('#cf');
  let email = '';
  lf.addEventListener('submit', async e=>{
    e.preventDefault();
    email = app.querySelector('#login-email').value.trim().toLowerCase();
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ toast('이메일 주소를 확인해 주세요.'); return; }
    const btn = app.querySelector('#login-send'); btn.disabled = true; btn.textContent = '보내는 중…';
    const {error} = await sb.auth.signInWithOtp({email, options:{shouldCreateUser:true}});
    btn.disabled = false; btn.textContent = '로그인 코드 받기';
    if(error){ toast(error.status===429 ? '잠시 후 다시 시도해 주세요. 코드는 1분에 한 번 보낼 수 있습니다.' : '코드를 보내지 못했습니다. 이메일 주소를 확인해 주세요.'); return; }
    lf.hidden = true; cf.hidden = false;
    app.querySelector('#code-msg').textContent = `${email}로 6자리 코드를 보냈습니다. 메일에 적힌 코드를 입력하세요. (몇 분 안에 안 오면 스팸함도 확인해 주세요)`;
    app.querySelector('#login-code').focus();
  });
  app.querySelector('#code-back').addEventListener('click', ()=>{ cf.hidden = true; lf.hidden = false; });
  cf.addEventListener('submit', async e=>{
    e.preventDefault();
    const token = app.querySelector('#login-code').value.replace(/\D/g,'');
    const {error} = await sb.auth.verifyOtp({email, token, type:'email'});
    if(error) toast('코드가 맞지 않거나 만료되었습니다. 새 코드를 받아 주세요.');
  });
}

/* ---------- Supabase 연결 ---------- */
const CFG = window.CLAIM_APP_CONFIG || {};
const DOMAIN = String(CFG.allowedEmailDomain||'').toLowerCase().replace(/^@/,'');
const BUCKET = 'claim-photos';
const configured = window.supabase && CFG.supabaseUrl && CFG.supabaseAnonKey && !/YOUR-/.test(CFG.supabaseUrl+CFG.supabaseAnonKey);
const sb = configured ? window.supabase.createClient(CFG.supabaseUrl, CFG.supabaseAnonKey) : null;
const plain = c => { const o = Object.assign({}, c); delete o._id; delete o.example; return o; };

const api = {
  async loadAll(){
    const {data, error} = await sb.from('claims').select('id,data,example').order('created_at', {ascending:true});
    if(error) throw error;
    S.all = data.map(r=>Object.assign({}, r.data, {_id:r.id, example:r.example}));
    S.claims = S.all.filter(c=>!c.example);
    S.examples = S.all.filter(c=>c.example);
    S.loaded = true;
  },
  async loadCfg(){
    const {data, error} = await sb.from('settings').select('value').eq('key','approval').maybeSingle();
    if(error) throw error;
    if(data?.value) S.cfg = {minor:data.value.minor??'', major:data.value.major??''};
  },
  async insert(d){
    const {data, error} = await sb.from('claims').insert({data:d}).select('id').single();
    if(error) throw error;
    await api.loadAll();
    return data.id;
  },
  async update(id, patch){
    // 다른 사람이 방금 결재했을 수 있으니 최신 내용 위에 합칩니다.
    const {data:fresh, error:readErr} = await sb.from('claims').select('data').eq('id', id).single();
    if(readErr) throw readErr;
    const {error} = await sb.from('claims').update({data:Object.assign({}, fresh.data, patch)}).eq('id', id);
    if(error) throw error;
    await api.loadAll();
    if(S.view==='doc' && S.id===id) render();
  },
  async remove(id){
    const cur = S.all.find(x=>x._id===id);
    const {error} = await sb.from('claims').delete().eq('id', id);
    if(error) throw error;
    if(cur?.photos?.length) await api.removePhotos(cur.photos).catch(()=>{});
    await api.loadAll();
  },
  async loadMembers(){
    const [emp, adm] = await Promise.all([
      sb.from('employees').select('email,name,dept,title,status,synced_at'),
      sb.from('app_admins').select('email')
    ]);
    if(emp.error) throw emp.error;
    if(adm.error) throw adm.error;
    const admins = new Set(adm.data.map(a=>a.email));
    S.members = emp.data.map(e=>Object.assign({}, e, {role:roleOfTitle(e.title), is_admin:admins.has(e.email)}));
    const m = S.members.find(x=>x.email===S.email);
    // 관리자(app_admins)는 직원 시트에 없어도 들어온다 — 회사 대표 계정(official@) 같은 경우
    S.staff = (!!m && m.status==='재직') || admins.has(S.email);
    S.me = {role:S.staff?m.role:null, isAdmin:S.staff&&m.is_admin, name:m?.name||''};
    document.getElementById('who-email').textContent = `${S.me.name||S.email}${S.me.role?' · '+S.me.role:''}${S.me.isAdmin?' · 관리자':''}`;
  },
  async saveCfg(v){
    const {error} = await sb.from('settings').upsert({key:'approval', value:v});
    if(error) throw error;
    S.cfg = {minor:v.minor, major:v.major};
  },
  async uploadPhoto(claimId, file){
    if(file.size > 20*1024*1024) throw {tooLarge:true};
    const ext = ((file.name||'').split('.').pop()||'jpg').toLowerCase().replace(/[^a-z0-9]/g,'') || 'jpg';
    const path = `${claimId}/${Date.now()}-${Math.random().toString(36).slice(2,8)}.${ext}`;
    const {error} = await sb.storage.from(BUCKET).upload(path, file, {contentType:file.type||'image/jpeg'});
    if(error) throw error;
    return path;
  },
  async removePhotos(paths){
    const {error} = await sb.storage.from(BUCKET).remove(paths);
    if(error) throw error;
  }
};

async function loadPhotoUrls(paths){
  const need = paths.filter(p=>!S.photoUrls[p]);
  if(need.length){
    const {data} = await sb.storage.from(BUCKET).createSignedUrls(need, 60*60*6);
    (data||[]).forEach(r=>{ if(r.signedUrl && r.path) S.photoUrls[r.path] = r.signedUrl; });
  }
  app.querySelectorAll('img[data-photo]').forEach(img=>{
    const u = S.photoUrls[img.dataset.photo];
    if(u && img.getAttribute('src')!==u){ img.src = u; img.closest('a').href = u; }
  });
}

function subscribe(){
  if(S.chan) return;
  const refresh = ()=>{ clearTimeout(S._rt); S._rt = setTimeout(async()=>{
    try{ await Promise.all([api.loadAll(), api.loadCfg(), api.loadMembers()]); }catch(_){ return; }
    if(S.view==='list'||S.view==='doc'||(S.view==='std'&&!app.contains(document.activeElement))) render();
  }, 300); };
  S.chan = sb.channel('claims-live')
    .on('postgres_changes', {event:'*', schema:'public', table:'claims'}, refresh)
    .on('postgres_changes', {event:'*', schema:'public', table:'settings'}, refresh)
    .on('postgres_changes', {event:'*', schema:'public', table:'employees'}, refresh)
    .subscribe();
}

async function onSession(session){
  const sameUser = S.session && session && S.session.user.id===session.user.id;
  S.session = session; S.authReady = true;
  S.email = session?.user?.email?.toLowerCase() || null;
  document.getElementById('who-email').textContent = S.email || '';
  if(!session){
    if(S.chan){ sb.removeChannel(S.chan); S.chan = null; }
    S.all = []; S.claims = []; S.examples = []; S.members = []; S.staff = false; S.me = {role:null, isAdmin:false, name:''}; S.loaded = false; S.view = 'list';
    render(); return;
  }
  if(sameUser && S.loaded) return;
  render();
  try{ await Promise.all([api.loadAll(), api.loadCfg(), api.loadMembers()]); S.offline = false; }
  catch(e){ S.loaded = true; S.offline = true; toast('데이터를 불러오지 못했습니다. 새로고침해 주세요.'); }
  subscribe();
  render();
}

document.getElementById('btn-logout').addEventListener('click', ()=>sb.auth.signOut());

if(!sb){
  document.getElementById('nav').hidden = true;
  app.innerHTML = `<div class="empty" style="margin-top:28px"><h3>아직 설정되지 않았습니다</h3><p>config.js에 Supabase 주소와 anon 키를 넣어 주세요. 방법은 README.md에 있습니다.</p></div>`;
} else {
  sb.auth.onAuthStateChange((event, session)=>{
    if(event==='INITIAL_SESSION' || event==='SIGNED_IN' || event==='SIGNED_OUT') setTimeout(()=>onSession(session), 0);
  });
  render();
}
})();
