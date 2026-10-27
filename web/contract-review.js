"use strict";
// Additive integration: retain existing import, source UI, chat and Word export handlers.
const contractUI = {busy:false, mutating:false, stop:false, filter:'all', issue:null, candidate:null};
let activeProjectRequests=0;
const originalContractApi=api;
api=async function(path,options={}){
  const tracked=options.method==='POST'&&!path.startsWith('/api/contract-review/');
  if(tracked)activeProjectRequests++;
  try{
    const result=await originalContractApi(path,options);
    if(path==='/api/import'&&result.id){const url=new URL(location.href);url.searchParams.set('project',result.id);history.replaceState(null,'',url)}
    return result;
  }finally{if(tracked)activeProjectRequests--}
};
const chatPanel = document.querySelector('.question-panel');
chatPanel.classList.remove('claims-panel');
const reviewPanel = document.createElement('aside');
reviewPanel.className = 'contract-panel';
reviewPanel.innerHTML = `<div class="panel-head"><div><small>合同审核</small><strong>全文结论与修改意见</strong></div></div>
 <div id="review-overview" class="review-overview"></div>
 <div class="review-filters" role="group" aria-label="审核意见筛选"><button data-review-filter="all">全部</button><button data-review-filter="text">文字</button><button data-review-filter="math">计算</button><button data-review-filter="logic">逻辑</button><button data-review-filter="risk" hidden>其他</button></div>
 <div id="review-issues" class="review-issues"></div><div class="review-foot">点击意见定位原文；修改后请重新审核。</div>`;
chatPanel.replaceWith(reviewPanel);
const taskbar = document.createElement('section');
taskbar.id = 'review-taskbar'; taskbar.className = 'review-taskbar hidden';
taskbar.innerHTML = `<b>当前文件审核</b><label>审核模型 <select id="review-provider"><option value="api">API 模型 · 现有配置</option><option value="local">本地小模型 · 现有配置</option></select></label><span id="review-scope" aria-live="polite"></span><label>补查 <select id="review-extra-checks" aria-label="补查次数"><option value="0">不补查</option><option value="1">再查1次</option><option value="2">再查2次（更慢）</option></select></label><button id="review-start" class="button primary">审核全文</button><button id="review-resume" class="button small" hidden>继续审核</button><button id="review-stop" class="button small" hidden>停止审核</button><span id="review-progress"></span>`;
$('workspace').before(taskbar);
const conversation = document.createElement('section');
conversation.id = 'contract-chat'; conversation.className = 'contract-chat hidden';
conversation.innerHTML = `<button id="contract-chat-toggle" aria-expanded="true"><b>MCP / API 对话</b><span>当前段落 · 查法条、追问与核对</span><em>收起 ⌄</em></button>`;
conversation.append(chatPanel); $('app').append(conversation);
const tray = document.createElement('div'); tray.className = 'review-selection';
tray.innerHTML = `<div id="review-selection-label">点击左侧审核意见可定位原文</div><button id="review-ignore">忽略此项</button><button id="review-apply">查看修改建议</button>`;
document.querySelector('.artifact-panel').append(tray);
const changeDialog = document.createElement('dialog'); changeDialog.id='review-change-dialog';
changeDialog.innerHTML = `<div class="dialog-card"><header><h2>确认修改建议</h2><button id="review-change-close" class="icon-button">×</button></header><label>原文</label><pre id="review-change-before"></pre><label>建议替换为</label><pre id="review-change-after"></pre><p class="hint">确认后更新工作区文字，原件不变；旧依据转入历史，需重新溯源和审核。最后使用“生成修订副本”另存为带修订痕迹的 .docx。</p><footer><button id="review-change-confirm" class="button primary">确认采用此修改</button></footer></div>`;
document.body.append(changeDialog);
const capacitySettings=document.createElement('section');
capacitySettings.innerHTML='<h3>全文审核上下文容量</h3><p class="hint">填写模型当前实际设置的容量，不是宣传的最大容量。仅用于全文容量估算，不改变模型设置；本地留0尝试读取已加载实例，API需填写。</p><label>本地模型当前上下文（token）<input id="review-local-context" type="number" min="0" max="2097152" value="0"></label><label>API模型上下文（token）<input id="review-api-context" type="number" min="0" max="2097152" value="0"></label>';
$('settings-dialog').querySelector('footer').before(capacitySettings);
const originalCapacityLoad=loadSettings;
loadSettings=async function(){await originalCapacityLoad();const c=await api('/api/contract-review/settings');$('review-local-context').value=c.localReviewContextTokens||0;$('review-api-context').value=c.apiReviewContextTokens||0};
const originalCapacitySave=$('save-settings').onclick;
$('save-settings').onclick=async()=>{try{await api('/api/contract-review/settings',{method:'POST',body:JSON.stringify({localReviewContextTokens:Number($('review-local-context').value),apiReviewContextTokens:Number($('review-api-context').value)})});await originalCapacitySave()}catch(error){toast(error.message)}};
document.title='审核溯源台';
document.querySelector('.brand strong').textContent='审核溯源台';
document.querySelector('.brand-mark').innerHTML='<img src="/audit-mark.svg" alt="" aria-hidden="true">';
document.querySelector('.brand small').textContent='逐项审核 · 原文核对';
document.querySelector('.welcome .eyebrow').textContent='单文件审核工具';
document.querySelector('.welcome h1').innerHTML='审核一个文件，<br>逐项核对修改。';
document.querySelector('.welcome-copy>p').textContent='导入一份合同或文档，本地模型核对文字、称呼、计算和前后逻辑；API 模型还会提出法律疑点。左侧查看意见，中间核对修改，右侧查来源；Word 可另存带修订痕迹的原格式副本。';
$('import-button').textContent='导入文件';$('back-button').textContent='更换文件';
document.querySelector('.welcome-visual').innerHTML=`<div class="welcome-console-head"><b>审核 / 原文 / 来源</b><span>功能示意</span></div><div class="welcome-sample"><small>当前文件 · 合同价款</small><p>本合同总价为 <mark>100,000 元</mark>。<br>分期支付：30,000 元、50,000 元、10,000 元。</p></div><div class="welcome-result"><small>计算核对</small><b>分期合计与总价相差 10,000 元</b><p>原句、算式和修改建议放在一起，逐项由你确认。</p></div><div class="welcome-console-foot"><span>保留原件</span><span>修改另存修订副本 ↗</span></div>`;

function contractReview(){
  const review=state.project?.contractReview;
  if(review && Object.keys(review.textSnapshot||{}).some(id=>claimById(id)?.text!==review.textSnapshot[id])) review.status='stale';
  if(review && Object.keys(review.textSnapshot||{}).length!==(state.project?.claims||[]).length) review.status='stale';
  return review;
}
function selectedReviewIssue(){return contractReview()?.issues.find(i=>i.id===contractUI.issue)}
function reviewVariantsHtml(issue){return issue.variants?.length?`<details class="review-quote-alignment"><summary>同一原句的其他意见（${issue.variants.length}）</summary>${issue.variants.map(v=>`<p><b>${v.sourcePass?'补查'+v.sourcePass:'首轮'}</b>：${reviewReferenceHtml(v.reason)}<br>${reviewReferenceHtml(v.suggestion)}</p>`).join('')}</details>`:''}
function reviewPassDone(review,index,pass){return Boolean(review.passCompleted?.full?.includes(pass))}
const categoryLabel={text:'文字',math:'计算',logic:'逻辑',risk:'其他'};
function categoryName(category,review){return category==='risk'&&review?.instructionProfile==='legal-v1'?'法律':categoryLabel[category]||'其他'}
function compactReviewOverview(review,issues,stale){
  const legacy=review&&review.plan!=='full-document-v1';
  let status=!review?'导入完成 · 可全文审核':stale?'原文已修改 · 请重审':legacy?'旧审核记录 · 请重新全文审核':review.status==='complete'?'审核完成':contractUI.busy?(review.activePass?'全文补查'+review.activePass:'全文首审中'):'全文结果尚未完成';
  if(contractUI.error)status='启动或审核未完成';
  const info=[];
  if(contractUI.error)info.push(`<p class="review-error">${esc(contractUI.error)}</p>`);
  if(review){
    info.push(`<p>审核范围：${review.instructionProfile==='legal-v1'?'文字校对与法律疑点':review.instructionProfile==='proofread-v1'?'文字校对与一致性':'旧记录，请以当时审核范围为准'}。</p>`);
    info.push(`<p>全文一次输入；${Object.keys(review.textSnapshot||{}).length} 段定位编号。${review.extraChecks?'全文补查 '+review.extraChecks+' 次。':'未启用补查。'}</p>`);
    if(review.contextTokens)info.push(`<p>上下文 ${esc(review.contextTokens)} token；全文与输出预留估算约 ${esc(review.estimatedTokens)}。${esc(review.contextSource||'')}。估算以服务端实际校验为准。</p>`);
    info.push(...Object.values(review.failed||{}).map(v=>`<p class="review-error">${esc(v)}</p>`));
    info.push(...(review.warnings||[]).map(v=>`<p>${esc(v)}</p>`));
    if(review.unlocated?.length)info.push(`<p>${review.unlocated.length} 条意见未定位，已在下方单列。</p>`);
  }else info.push('<p>全文提供给所选模型。API模式会向配置地址发送原文。</p>');
  if(review&&!stale&&!contractUI.busy)info.push('<button id="review-relocate" class="button small">重新匹配已有意见</button>');
  info.push('<p>意见需人工核对。Word原件不变，修改另存修订副本。</p>');
  return `<div class="review-status-line"><span class="${contractUI.error||Object.keys(review?.failed||{}).length?"is-error":""}">${esc(status)}</span><details class="review-details"><summary>详情</summary><div>${info.join('')}</div></details></div><div class="review-stats"><span>${issues.filter(i=>i.severity==='high').length} 重点核对</span><span>${issues.length} 条意见</span><span>${issues.filter(i=>i.decision!=='pending').length} 已处理</span></div>`;
}
// Model-written IDs are navigation hints, not verified quotations.
function reviewReferenceHtml(value){
  return String(value||'').split(/(\bC\d{1,6}\b)/gi).map(part=>{
    if(!/^C\d{1,6}$/i.test(part))return esc(part);
    const id='C'+String(Number(part.slice(1))).padStart(3,'0');
    return claimById(id)?`<button type="button" class="review-id-link" data-review-reference="${esc(id)}" title="查看模型提及的位置，需核对原文">${esc(part)}</button>`:esc(part);
  }).join('');
}
function reviewIssueHtml(i){
  const status=({partial:'部分定位 · 其余待核',approximate:'候选位置 · 待核对'})[i.locationStatus]||'已定位原文';
  return `<article class="review-issue ${i.id===contractUI.issue?'active':''} ${i.targets.some(t=>t.claimId===state.selected)?'corresponding':''}" data-review-issue="${esc(i.id)}"><div class="review-issue-select"><div class="review-issue-meta"><span>${categoryName(i.category,contractReview())}</span><span>${esc(status)}</span><span>${i.sourcePass?'补查'+i.sourcePass:'首轮'}</span><span>${({pending:'待确认',ignore:'已忽略',applied:'已采纳'})[i.decision]||esc(i.decision)}</span></div><h3>${reviewReferenceHtml(i.title)}</h3><p>${reviewReferenceHtml(i.reason)}</p><div class="review-suggestion">${reviewReferenceHtml(i.suggestion)}</div></div>${i.targets.map((t,ai)=>`<button class="review-anchor-link" data-review-anchor="${ai}" data-issue="${esc(i.id)}">查看此处原文 →</button><blockquote>${esc(t.quote)}</blockquote><small>${esc(t.location)} · ${esc(t.claimId)}${t.alignment?.startsWith('global-')?' · 全文匹配已纠正编号':''}</small>${t.modelQuote?`<details class="review-quote-alignment"><summary>已核对排版差异 · 查看模型引文</summary><p>${esc(t.modelQuote)}</p></details>`:''}`).join('')}${(i.unresolvedTargets||[]).map((u,ti)=>`<details class="review-quote-alignment"><summary>待核对引文${u.candidates?.length?' · '+u.candidates.length+' 个候选':''}</summary><blockquote>模型引文：${esc(u.modelQuote)}</blockquote><p>${esc(u.reason)}</p>${(u.candidates||[]).map((c,ci)=>`<button class="review-candidate-link" data-review-candidate="${ci}" data-target="${ti}" data-issue="${esc(i.id)}">查看 ${esc(c.location||c.claimId)}${c.alignment==='ambiguous'?' · 多处同文':' · 近似 '+Math.round(c.score*100)+'%'}${c.numberDifference?' · 数字不同':''}</button><blockquote>${esc(c.quote)}</blockquote>`).join('')}</details>`).join('')}${reviewVariantsHtml(i)}<footer><button type="button" class="review-issue-open">查看意见对应原文 →</button></footer></article>`;
}
function renderContractReview(){
  const review=contractReview();
  const visible=Boolean(state.project); $('review-taskbar').classList.toggle('hidden',!visible); conversation.classList.toggle('hidden',!visible);
  if(!visible)return;
  const stale=review?.status==='stale';
  $('review-start').disabled=contractUI.busy;
  $('review-provider').disabled=contractUI.busy; $('review-extra-checks').disabled=contractUI.busy;
  $('review-resume').hidden=!review||stale||review.status==='complete'||contractUI.busy;
  $('review-stop').hidden=!contractUI.busy;
  $('review-progress').textContent=review?`${review.model} · ${stale?'原文已修改':review.status==='complete'?'审核完成':contractUI.busy?(review.activePass?'全文补查'+review.activePass:'全文首审'):'待继续全文审核'}`:'全文一次送审 · 编号只用于定位';
  $('review-scope').textContent=$('review-provider').value==='api'?'下次审核：校对 + 法律':'下次审核：文档校对';
  const issues=review?.issues||[];
  $('review-overview').innerHTML=compactReviewOverview(review,issues,stale);
  const otherVisible=review?.instructionProfile==='legal-v1'||[...issues,...(review?.unlocated||[])].some(i=>i.category==='risk');
  document.querySelector('[data-review-filter="risk"]').textContent=categoryName('risk',review);
  document.querySelector('[data-review-filter="risk"]').hidden=!otherVisible;if(!otherVisible&&contractUI.filter==='risk')contractUI.filter='all';
  const filtered=issues.filter(i=>contractUI.filter==='all'||i.category===contractUI.filter);
  $('review-issues').innerHTML=filtered.length?filtered.map(reviewIssueHtml).join(''):'<div class="review-empty">此分类暂无意见。</div>';
  const unlocated=(review?.unlocated||[]).filter(i=>contractUI.filter==='all'||i.category===contractUI.filter);
  if(unlocated.length)$('review-issues').insertAdjacentHTML('beforeend',`<div class="review-unlocated-heading">未定位到原句 · 仅供人工核对</div>`+unlocated.map(i=>`<article class="review-issue review-unlocated"><div class="review-issue-meta"><span>待人工核对</span><span>${categoryName(i.category,review)}</span></div><h3>${reviewReferenceHtml(i.title)}</h3><p>${reviewReferenceHtml(i.reason)}</p><div class="review-suggestion">${reviewReferenceHtml(i.suggestion)}</div>${i.quotes.map(q=>`<blockquote>模型所引文字（未验证）：${esc(q)}</blockquote>`).join('')}<small>相关范围：${esc(i.range||'未定位')}</small><footer>无法自动定位或采纳，请对照当前文件核查。</footer></article>`).join(''));
  document.querySelectorAll('[data-review-issue]').forEach(node=>node.onclick=event=>{if(!event.target.closest('details,[data-review-candidate],[data-review-anchor],[data-review-reference]'))selectContractIssue(node.dataset.reviewIssue)});
  document.querySelectorAll('[data-review-filter]').forEach(node=>node.classList.toggle('active',contractUI.filter===node.dataset.reviewFilter));
  document.querySelectorAll('[data-review-reference]').forEach(node=>node.onclick=event=>{
    event.stopPropagation();const card=node.closest('[data-review-issue]');const id=node.dataset.reviewReference;
    if(!claimById(id))return toast('该编号在当前文档中不存在。');
    contractUI.issue=card?.dataset.reviewIssue||null;contractUI.candidate={claimId:id,reference:true};
    setSelected(id);renderContractReview();highlightReviewTargets();
  });
  document.querySelectorAll('[data-review-anchor]').forEach(node=>node.onclick=()=>{const i=issues.find(i=>i.id===node.dataset.issue);contractUI.issue=i.id;contractUI.candidate=null;setSelected(i.targets[Number(node.dataset.reviewAnchor)].claimId);renderContractReview();highlightReviewTargets()});
  document.querySelectorAll('[data-review-candidate]').forEach(node=>node.onclick=()=>{const i=issues.find(i=>i.id===node.dataset.issue);const c=i.unresolvedTargets[Number(node.dataset.target)].candidates[Number(node.dataset.reviewCandidate)];contractUI.issue=i.id;contractUI.candidate=c;setSelected(c.claimId);renderContractReview();highlightReviewTargets()});
  const relocate=$('review-relocate');if(relocate)relocate.onclick=async()=>{if(contractUI.busy||contractUI.mutating||activeProjectRequests)return;contractUI.mutating=true;try{state.project=await reviewRequest('relocate');contractUI.issue=null;contractUI.candidate=null;render();toast('已重新匹配，未调用模型。')}catch(e){toast(e.message)}finally{contractUI.mutating=false}};
  const issue=selectedReviewIssue();
  $('review-selection-label').innerHTML=issue?`<b>${esc(issue.title)}</b><span>${contractUI.candidate?.reference?'模型提及的编号位置，请核对原文':contractUI.candidate?'候选原文，尚未确认':issue.targets.map(t=>esc(t.location)).join(' ↔ ')}</span>`:contractUI.candidate?.reference?'<b>模型提及的编号位置</b><span>请核对原文</span>':'点击左侧意见定位原文；也可选择正文查看对应意见。';
  $('review-ignore').disabled=!issue||stale||contractUI.busy||issue.decision==='applied';
  $('review-ignore').textContent=issue?.decision==='ignore'?'撤销忽略':'忽略此项';
  $('review-apply').disabled=!issue||stale||contractUI.busy||state.project.artifact.kind!=='word'||issue.decision==='applied'||!issue.targets.length||Boolean(contractUI.candidate);
  $('review-apply').textContent=(!issue?.targets.length||contractUI.candidate)?'需核对定位':issue?.replacement!=null?'查看修改建议':'人工编辑条款';
}
function highlightReviewTargets(){
  const review=contractReview(); if(!review||review.status==='stale')return;
  const issue=selectedReviewIssue();
  document.querySelectorAll('[data-artifact-claim]').forEach(node=>{
    const cid=node.dataset.artifactClaim, text=claimById(cid)?.text;
    let targets=issue?.targets.filter(t=>t.claimId===cid)||[];
    const candidate=contractUI.candidate;node.classList.toggle('review-candidate',Boolean(candidate?.claimId===cid));if(candidate?.claimId===cid&&!candidate.reference)targets=[candidate];
    node.classList.toggle('review-linked',Boolean(targets.length));
    if(text && node.classList.contains('artifact-sentence'))node.innerHTML=esc(text);
    if(text && targets.length && node.classList.contains('artifact-sentence')){
      const spans=targets.map(t=>({start:t.start??text.indexOf(t.quote),end:t.end??(text.indexOf(t.quote)+t.quote.length)})).sort((a,b)=>a.start-b.start);
      let result='',cursor=0;
      for(const span of spans){if(span.start<cursor)continue;result+=esc(text.slice(cursor,span.start))+'<mark>'+esc(text.slice(span.start,span.end))+'</mark>';cursor=span.end}
      node.innerHTML=result+esc(text.slice(cursor));
    }
  });
}
function selectContractIssue(id){
  contractUI.issue=id;contractUI.candidate=null;
  const issue=selectedReviewIssue();if(!issue)return;
  const anchor=issue.targets[0];if(anchor)setSelected(anchor.claimId);else{const candidate=issue.unresolvedTargets?.flatMap(t=>t.candidates||[])[0];if(candidate){contractUI.candidate=candidate;setSelected(candidate.claimId)}}
  renderContractReview(); highlightReviewTargets();
  const card=document.querySelector(`[data-review-issue="${CSS.escape(id)}"]`);
  if(card){const list=$('review-issues');list.scrollTop+=card.getBoundingClientRect().top-list.getBoundingClientRect().top}
}
const originalRenderClaims=renderClaims;
renderClaims=function(){originalRenderClaims();renderContractReview()};
const originalRenderArtifact=renderArtifact;
renderArtifact=function(){originalRenderArtifact();highlightReviewTargets()};
// Keep the existing local-material check available, folded beside the sources,
// so it no longer separates contract clauses that need to be compared together.
const originalInlineReviewHtml=inlineReviewHtml;
inlineReviewHtml=function(){return ''};
const originalContractSources=renderSources;
renderSources=function(){
  originalContractSources();
  const claim=claimById(state.selected);if(!claim)return;
  const details=document.createElement('details');details.className='contract-local-check';
  details.innerHTML='<summary>本地材料核对提示</summary>'+originalInlineReviewHtml(claim);
  if(state.checkingLocal||claim.localCheck)details.open=true;
  $('sources').append(details);
};
const originalContractLocalCheck=runLocalConsistencyCheck;
runLocalConsistencyCheck=async function(){const pending=originalContractLocalCheck();renderSources();await pending;renderSources()};
const originalResetWorkspace=resetWorkspace;
resetWorkspace=function(){originalResetWorkspace();contractUI.issue=null;const url=new URL(location.href);url.searchParams.delete('project');history.replaceState(null,'',url);renderContractReview()};

async function reviewRequest(action,extra={}){
  return api('/api/contract-review/'+action,{method:'POST',body:JSON.stringify({project:state.project,...extra})});
}
async function runContractReview(resume=false){
  if(contractUI.busy||contractUI.mutating||!state.project)return;
  if(activeProjectRequests||state.runningAll||state.runningClaim||state.asking||state.structuring||state.exportingRevision||state.editingClaim)return toast('请先完成当前操作或保存编辑。');
  if(!resume && state.project.contractReview && !window.confirm('重新审核将替换当前意见列表和处理状态，已保存的原文修改仍然保留。继续吗？'))return;
  contractUI.busy=true;contractUI.stop=false;document.body.classList.add('review-running');
  let progressTimer,runFinished=false; contractUI.error="";
  try{
    if(!resume){state.project=await reviewRequest('start',{provider:$('review-provider').value,extraChecks:Number($('review-extra-checks').value)});contractUI.issue=null;contractUI.candidate=null}
    render();
    let polling=false,liveSignature='',progressEpoch=0;
    progressTimer=setInterval(async()=>{
      if(polling||!contractUI.busy)return;polling=true;const epoch=progressEpoch,reviewId=state.project.contractReview.id;
      try{
        const p=await api('/api/contract-review/progress?id='+encodeURIComponent(state.project.contractReview.id));
        if(runFinished||!contractUI.busy||epoch!==progressEpoch||reviewId!==state.project.contractReview.id)return;
        if(p.issues&&['reading','generating'].includes(p.stage)){
          const signature=JSON.stringify([p.issues,p.unlocated]);
          if(signature!==liveSignature){state.project.contractReview.issues=p.issues;state.project.contractReview.unlocated=p.unlocated||[];liveSignature=signature;renderContractReview()}
        }
        if(['reading','generating'].includes(p.stage))$('review-progress').textContent=`${state.project.contractReview.activePass?'全文补查'+state.project.contractReview.activePass:'全文首审'} · ${Math.round(p.seconds)} 秒 · 已返回 ${p.chars} 字`;
      }catch{}finally{polling=false}
    },700);
    const extra=state.project.contractReview.extraChecks||0;
    for(let passIndex=0;passIndex<=extra&&!contractUI.stop;passIndex++){
      if(reviewPassDone(state.project.contractReview,'full',passIndex))continue;
      state.project.contractReview.activePass=passIndex;
      $('review-progress').textContent=passIndex?'全文补查'+passIndex+'…':'全文首审…';
      state.project=await reviewRequest('step',{index:'full',passIndex});progressEpoch++;render();
      if(state.project.contractReview.failed.full||!reviewPassDone(state.project.contractReview,'full',passIndex))break;
    }
    toast(contractUI.stop?'已停止审核，现有意见保留，可继续审核当前文件':state.project.contractReview.status==='complete'?'全文审核完成，请逐项核对':'审核部分完成，请查看未完成范围并继续审核');
  }catch(error){contractUI.error=error.message;toast('审核未完成：'+error.message)}
  finally{runFinished=true;clearInterval(progressTimer);contractUI.busy=false;document.body.classList.remove('review-running');$('review-stop').disabled=false;$('review-stop').textContent='停止审核';renderContractReview()}
}
async function decideContractIssue(decision){
  const issue=selectedReviewIssue();if(!issue||contractUI.mutating)return;
  if(activeProjectRequests)return toast('请等待当前操作完成。');
  contractUI.mutating=true;$('review-change-confirm').disabled=true;
  try{state.project=await reviewRequest('decide',{issueId:issue.id,decision});changeDialog.close();render();toast(decision==='apply'?'已保存修改；请重新审核和溯源，最后生成修订副本':'意见状态已更新')}catch(error){toast(error.message)}
  finally{contractUI.mutating=false;$('review-change-confirm').disabled=false}
}
$('review-start').onclick=()=>runContractReview();$('review-resume').onclick=()=>runContractReview(!(contractReview()?.plan!=='full-document-v1'));
$('review-provider').onchange=()=>{$('review-extra-checks').value='0';renderContractReview()};
$('review-stop').onclick=async()=>{contractUI.stop=true;$('review-stop').textContent='正在停止…';$('review-stop').disabled=true;if(state.project?.contractReview){try{await api('/api/contract-review/cancel',{method:'POST',body:JSON.stringify({reviewId:state.project.contractReview.id})})}catch(error){toast(error.message)}}};
document.querySelectorAll('[data-review-filter]').forEach(node=>node.onclick=()=>{contractUI.filter=node.dataset.reviewFilter;renderContractReview()});
$('review-ignore').onclick=()=>decideContractIssue(selectedReviewIssue()?.decision==='ignore'?'pending':'ignore');
$('review-apply').onclick=()=>{
  const issue=selectedReviewIssue();if(!issue||contractUI.candidate||!issue.targets.length)return;
  if(issue.replacement==null){state.editingClaim=(issue.targets.find(t=>t.claimId===state.selected)||issue.targets[0]).claimId;state.selected=state.editingClaim;render();return}
  $('review-change-before').textContent=issue.targets[0].quote;
  $('review-change-after').textContent=issue.replacement||'（删除选中文字）';changeDialog.showModal();
};
$('review-change-close').onclick=()=>changeDialog.close();$('review-change-confirm').onclick=()=>decideContractIssue('apply');
$('contract-chat-toggle').onclick=()=>{
  const collapsed=conversation.classList.toggle('collapsed');document.body.classList.toggle('contract-chat-collapsed',collapsed);
  $('contract-chat-toggle').setAttribute('aria-expanded',String(!collapsed));$('contract-chat-toggle').querySelector('em').textContent=collapsed?'展开 ⌃':'收起 ⌄';
};
// A review response includes project state: lock mutations while it is in flight.
document.addEventListener('click',event=>{
  if(!contractUI.busy&&!contractUI.mutating)return;
  if(event.target.closest('#review-stop'))return;
  if(contractUI.busy&&!contractUI.mutating&&event.target.closest('[data-review-issue],[data-review-filter],[data-artifact-claim],#contract-chat-toggle,summary'))return;
  if(event.target.closest('button,input,select,textarea,summary')){event.preventDefault();event.stopImmediatePropagation();toast('正在审核当前文件，请等待完成或点击停止审核。')}
},true);
document.addEventListener('submit',event=>{if(contractUI.busy){event.preventDefault();event.stopImmediatePropagation()}},true);
const originalSetSelected=setSelected;
setSelected=function(id){
  originalSetSelected(id);
  const issue=selectedReviewIssue();
  if(!issue?.targets.some(t=>t.claimId===id)&&contractUI.candidate?.claimId!==id){
    contractUI.candidate=null;
    contractUI.issue=contractReview()?.issues.find(i=>i.targets.some(t=>t.claimId===id))?.id||null;
    renderContractReview();highlightReviewTargets();
  }
};
// A URL project id allows reopening a saved local session without reimporting it.
const savedProjectId=new URLSearchParams(location.search).get('project');
const originalThemedReviewRender=renderContractReview;
renderContractReview=function(){
  originalThemedReviewRender();
  document.querySelectorAll('.review-stats>span').forEach(node=>{
    const match=node.textContent.match(/^(\d+)\s+(.+)$/);
    if(match)node.innerHTML=`<b>${esc(match[1])}</b><span>${esc(match[2])}</span>`;
  });
};
if(savedProjectId)api('/api/project/'+encodeURIComponent(savedProjectId)).then(project=>{if(project.artifact){state.project=project;state.selected=null;$('review-provider').value=project.contractReview?.provider||'api';$('review-extra-checks').value=project.contractReview?.extraChecks||0;render()}}).catch(error=>toast(error.message));
renderContractReview();
