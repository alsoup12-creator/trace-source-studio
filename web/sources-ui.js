"use strict";

function plainCandidateText(value){
  const stripped=String(value||"").replace(/<[^>]+>/g," ").replace(/[`*_#]+/g,"");
  const decoder=document.createElement("textarea");decoder.innerHTML=stripped;
  return decoder.value.replace(/\s+/g," ").trim();
}

function candidateLocator(source){
  const match=String(source.quote||"").match(/<mark[^>]*>([\s\S]*?)<\/mark>/i);
  return plainCandidateText(match?match[1]:source.quote).slice(0,500);
}

function candidateType(source){
  if(source.reading?.sourceType)return source.reading.sourceType;
  if(source.sourceType)return source.sourceType;
  if(source.path){const ext=source.path.split(".").pop();return ext?ext.toUpperCase()+" 文件":"本地文件"}
  if(String(source.method||"").includes("MCP"))return "MCP资料";
  return source.url?"网页":"候选资料";
}

function candidateHost(source){
  if(!source.url)return "";
  try{return new URL(source.url).hostname.replace(/^www\./,"")}catch{return ""}
}

function isLocalPdf(source){return Boolean(source?.path&&/\.pdf$/i.test(String(source.path)));}

function candidateOrigin(source){
  const method=String(source.method||"");
  const host=candidateHost(source);
  if(source.path&&!source.url)return "本地文件";
  if(method.includes("北大法宝")||host.includes("pkulaw"))return "北大法宝结果";
  if(method.includes("MCP"))return "MCP资料";
  if(method.includes("秘塔")||method.includes("网站")||source.url)return "网页结果";
  return "其他来源";
}

function candidateReviewTier(source){if(source.reviewTier)return source.reviewTier;if(source.path&&!source.url){if(source.relation==="反向候选")return"conflict";return ["逐字命中","规范化命中","关键词组合"].includes(source.matchType)?"primary":"weak"}return"primary"}

function candidateReviewLabel(source){return ({primary:"主要候选",conflict:"可能冲突",weak:"其他线索"})[candidateReviewTier(source)]||"主要候选"}

function sourceFilterLabels(source){return [...new Set([candidateOrigin(source),candidateReviewLabel(source),source.matchType,...sourceViewLabels(source)].filter(Boolean))]}

function sourceViewLabels(source){
  if(String(source.label||"").includes("暂不可用"))return ["检索失败"];
  const labels=(source.reading?.viewpointLabels||[]).map(item=>item.label).filter(Boolean);
  return labels.length?[...new Set(labels)]:["待深读"];
}

function tracePlanHtml(claim){
  const plan=claim.tracePlan;const errors=[claim.localAiError&&`补充检索未完成：${claim.localAiError}`,claim.localRefineError&&`候选精筛未完成：${claim.localRefineError}`].filter(Boolean);const refinement=claim.localRefinement;const localNotice=`${refinement?`<p class="local-refine-note">${esc(refinement.model)} 已精筛 ${refinement.classified}/${refinement.totalLocal} 条本地候选：主要 ${refinement.counts?.primary||0}，冲突 ${refinement.counts?.conflict||0}，其他 ${refinement.counts?.weak||0}${refinement.failedBatches?`；${refinement.failedBatches}批使用规则分层`:""}。</p>`:""}${errors.map(message=>`<p class="plan-error">本地 AI ${esc(message)}；全部规则候选仍已保留。</p>`).join("")}`;if(!plan)return localNotice;
  const queries=[plan.webQuery&&`网站：${plan.webQuery}`,plan.pkulawKeywordQuery&&`法宝关键词：${plan.pkulawKeywordQuery}`,plan.pkulawSemanticQuery&&`法宝语义：${plan.pkulawSemanticQuery}`].filter(Boolean);
  return `<section class="trace-plan"><div class="trace-plan-head"><b>本次审查框架</b><span>${esc(plan.generatedBy||"")}</span></div>${plan.reviewQuestion||plan.coreIssue?`<p><strong>要回答：</strong>${esc(plan.reviewQuestion||plan.coreIssue)}</p>`:""}${plan.viewpointLabels?.length?`<div class="plan-subclaims">${plan.viewpointLabels.map(item=>`<span>${esc(item)}</span>`).join("")}</div>`:""}${plan.subclaims?.length?`<details><summary>查看拆分的目标观点</summary>${plan.subclaims.map(item=>`<p>${esc(item)}</p>`).join("")}</details>`:""}${queries.length?`<details><summary>查看实际检索问题</summary>${queries.map(item=>`<p>${esc(item)}</p>`).join("")}</details>`:""}${plan.error?`<p class="plan-error">${esc(plan.error)}</p>`:""}${localNotice}</section>`;
}

function readingHtml(reading){
  if(!reading)return `<div class="inline-empty">该来源尚未深读。请在右侧点“深读”，核心观点和逐字依据会显示在这里。</div>`;
  const labels=(reading.viewpointLabels||[]).map(item=>`<div class="viewpoint-row"><span>${esc(item.label)}</span><div><p>${esc(item.summary||"")}</p>${item.quote?`<blockquote>${esc(item.quote)}</blockquote>`:""}</div></div>`).join("");
  const quotes=(reading.keyQuotes||[]).map(item=>`<div class="reading-quote"><b>${esc(item.role||"关键原文")}</b><p>${esc(item.quote)}</p></div>`).join("");
  const alignment=(reading.alignment||[]).map(item=>`<div class="alignment-row"><span>${esc(item.status)}</span><div><b>${esc(item.subclaim)}</b><p>${esc(item.explanation)}</p>${item.quote?`<blockquote>${esc(item.quote)}</blockquote>`:""}</div></div>`).join("");
  return `<section class="source-reading"><div class="reading-title"><b>来源核心观点</b><span>${esc(reading.sourceType||"")}</span></div><p class="core-view">${esc(reading.coreView||"未生成概括")}</p>${labels?`<div class="viewpoint-list">${labels}</div>`:""}${quotes?`<details open><summary>关键来源原文</summary>${quotes}</details>`:""}${alignment?`<details><summary>与目标观点逐项对齐</summary>${alignment}</details>`:""}<small>由${esc(reading.generatedBy||"理解层")}概括；分类和概括必须回到逐字原文核对。</small></section>`;
}

function summarySourceIndex(value, allRows){
  const number=Number(value);
  if(!Number.isInteger(number)||number<0||number>=allRows.length)return null;
  return number;
}

function summaryTextHtml(value, allRows){
  const text=String(value||"");let html="",cursor=0;const pattern=/\[\[(\d+)\]\]/g;let match;
  while((match=pattern.exec(text))){html+=esc(text.slice(cursor,match.index));const index=summarySourceIndex(match[1],allRows);const source=index==null?null:allRows[index];const origin=source?candidateOrigin(source).replace("结果",""):"";html+=index==null?esc(match[0]):`<button type="button" class="summary-ref" data-summary-source="${index}" title="${esc(source?.label||"")}">依据${index+1}${origin?` · ${esc(origin)}`:""}</button>`;cursor=match.index+match[0].length}
  return html+esc(text.slice(cursor));
}

function summaryEvidenceHtml(source,index){
  if(!source)return "";
  const reading=source.reading||{};const quote=(reading.keyQuotes||[])[0]?.quote||source.quote||plainCandidateText(source.rawText||source.context||"");
  const preview=isLocalPdf(source)?`<button type="button" class="button" data-preview-pdf="${index}">工作台查看${source.page?`（第${esc(source.page)}页）`:""}</button>`:"";
  const open=source.url||source.path?`<button type="button" class="button quiet" data-summary-open="${index}">${source.url?"打开并定位原文":"用默认程序打开"}</button>`:"";
  return `<div class="summary-evidence"><div class="summary-evidence-head"><b>${esc(source.label||`来源${index+1}`)}</b><span>${esc(candidateType(source))}${source.location?` · ${esc(source.location)}`:""}</span></div>${reading.coreView?`<p class="summary-evidence-view">${esc(reading.coreView)}</p>`:""}<blockquote>${esc(quote||"没有取得可展示原文")}</blockquote><div class="source-actions">${preview}${open}${source.url&&!reading.coreView?`<button type="button" class="button" data-summary-deep-read="${index}">深读此依据</button>`:""}</div></div>`;
}

function summaryRefsFromText(text,allRows){const refs=[];for(const raw of String(text||"").matchAll(/\[\[(\d+)\]\]/g)){const index=summarySourceIndex(raw[1],allRows);if(index!=null&&!refs.includes(index))refs.push(index)}return refs}

function summarySourceButtons(indexes,allRows){return indexes.map(index=>{const source=allRows[index];if(!source)return "";const origin=candidateOrigin(source).replace("结果","");return `<button type="button" class="summary-ref" data-summary-source="${index}" title="${esc(source.label||"")}">依据${index+1}${origin?` · ${esc(origin)}`:""}</button>`}).join("")}

function sourceSummaryHtml(claim, allRows){
  const summary=claim.sourceSummary;
  if(!summary||!summary.status||summary.status==="unavailable")return allRows.length?`<section class="summary-empty"><b>来源观点尚未汇总</b><p>先查看候选依据，再由秘塔把来源归纳成少数几个观点。</p><button class="button" data-summarize-claim ${state.summarizing?"disabled":""}>${state.summarizing?"正在汇总…":"整理来源观点"}</button>${summary?.gaps?.length?`<p class="plan-error">${esc(summary.gaps.join("；"))}</p>`:""}</section>`:"";
  const overallRefs=summaryRefsFromText(summary.overallView,allRows);
  const renderGroup=(group,labelOverride="")=>{const indexes=(group.sourceIndexes||[]).map(index=>summarySourceIndex(index,allRows)).filter(index=>index!=null);const unique=[...new Set(indexes)];const expanded=state.summaryEvidence!=null&&unique.includes(state.summaryEvidence)?summaryEvidenceHtml(allRows[state.summaryEvidence],state.summaryEvidence):"";return `<article class="summary-card"><div class="summary-title"><b>${esc(labelOverride||group.label||"来源判断")}</b><span>${unique.length}项依据</span></div><p>${summaryTextHtml(group.summary||"",allRows)}</p>${unique.length?`<div class="summary-sources">${summarySourceButtons(unique,allRows)}</div>`:""}${expanded}</article>`};
  const groups=(summary.groups||[]).map(group=>renderGroup(group)).join("");
  const fallback=!groups&&summary.overallView?renderGroup({summary:summary.overallView,sourceIndexes:overallRefs},"综合来源判断"):"";
  const plain=summary.status==="text"&&!fallback?`<article class="summary-card"><div class="summary-title"><b>秘塔问答汇总</b><span>自然语言回答</span></div><p>${summaryTextHtml(summary.overallView,allRows)}</p>${overallRefs.length?`<div class="summary-sources">${summarySourceButtons(overallRefs,allRows)}</div>`:""}${state.summaryEvidence!=null&&overallRefs.includes(state.summaryEvidence)?summaryEvidenceHtml(allRows[state.summaryEvidence],state.summaryEvidence):""}</article>`:"";
  return `<section class="source-summary"><div class="summary-head"><b>来源观点总结</b><span>${esc(summary.generatedBy||"")}</span></div>${summary.overallView&&!fallback&&!plain?`<p class="summary-overall">${summaryTextHtml(summary.overallView,allRows)}</p>`:""}${groups||fallback||plain||`<p class="empty-note">暂未形成可核验的来源分组。</p>`}${summary.gaps?.length?`<details><summary>查看未覆盖或存在差异的问题</summary>${summary.gaps.map(item=>`<p>${summaryTextHtml(item,allRows)}</p>`).join("")}</details>`:""}<button class="button quiet small" data-resummarize-claim ${state.summarizing?"disabled":""}>${state.summarizing?"正在汇总…":"重新汇总"}</button></section>`;
}

async function summarizeSelectedClaim(){
  const claim=claimById(state.selected);if(!claim||state.summarizing)return;
  state.summarizing=true;renderArtifact();renderSources();
  try{state.project=await api("/api/summarize-claim",{method:"POST",body:JSON.stringify({project:state.project,claimId:claim.id,prompt:state.project.prompt||"",allowLocalOnlineSummary:state.allowLocalOnlineSummary===true})});toast("来源观点汇总完成")}catch(error){toast(`来源汇总失败：${error.message}`)}finally{state.summarizing=false;renderSources();renderArtifact()}
}

function inlineReviewHtml(claim){
  const check=claim.localCheck;
  const localRows=(claim.candidates||[]).filter(source=>source.path&&!source.url);
  const references=(check?.references||[]).map(index=>{const source=localRows[index];return source?`<li><b>${esc(source.label||"本地文件")}</b> · ${esc(source.location||"")}<blockquote>${esc(String(source.quote||"").slice(0,500))}</blockquote></li>`:""}).join("");
  const result=check?`<div class="local-check-result local-check-${esc(check.status||"unavailable")}"><b>${esc(check.label||"本地材料核对")}</b><p>${esc(check.summary||"")}</p>${references?`<details><summary>查看本次核对的本地原文</summary><ul>${references}</ul></details>`:""}<small>仅比较已检索到的本地文件，不等同于事实、法律或专业结论。</small></div>`:`<div class="inline-empty">${localRows.length?"已找到本地候选，可用本地 AI 进行一致性核对。":"请先以“本地文件夹”溯源当前段落，才可进行本地材料核对。"}</div>`;
  return `<aside class="inline-review local-check" data-inline-review="${esc(claim.id)}"><div class="inline-summary-head"><b>本地材料核对提示</b><span>仅调用本地模型配置</span></div>${result}<button class="button quiet small" data-local-check ${state.checkingLocal||!localRows.length?"disabled":""}>${state.checkingLocal?"正在核对…":"核对本地材料"}</button></aside>`;
}

async function runLocalConsistencyCheck(){
  const claim=claimById(state.selected);if(!claim||state.checkingLocal)return;
  state.checkingLocal=true;renderArtifact();
  try{state.project=await api("/api/check-local-consistency",{method:"POST",body:JSON.stringify({project:state.project,claimId:claim.id})});toast("本地材料核对完成")}catch(error){toast(`本地材料核对失败：${error.message}`)}finally{state.checkingLocal=false;renderArtifact()}
}

renderSources=function(){
  const claim=claimById(state.selected);
  if(!claim){$("source-heading").textContent="选择一个观点";$("source-filters").innerHTML="";$("sources").className="sources empty";$("sources").textContent="从左侧选择观点，这里会显示候选来源索引。";return}
  $("source-heading").textContent=claim.text;
  const allRows=claim.candidates||[];
  const counts={all:allRows.length};
  allRows.forEach(source=>sourceFilterLabels(source).forEach(label=>{counts[label]=(counts[label]||0)+1}));
  const origins=[...new Set(allRows.map(candidateOrigin))];
  const filters=["all",...origins,...Object.keys(counts).filter(key=>key!=="all"&&!origins.includes(key))];
  if(!counts[state.sourceFilter])state.sourceFilter="all";
  $("source-filters").innerHTML=filters.map(value=>`<button class="source-filter ${origins.includes(value)?"origin-filter ":""}${state.sourceFilter===value?"active":""}" data-source-filter="${esc(value)}">${value==="all"?"全部":esc(value)} <span>${counts[value]||0}</span></button>`).join("");
  document.querySelectorAll("[data-source-filter]").forEach(node=>node.onclick=()=>{state.sourceFilter=node.dataset.sourceFilter;renderSources()});
  if(!allRows.length){
    const archived=(claim.evidenceHistory||[]).length;
    const message=claim.status==="待重新溯源"?`文字已经修改。${archived?"修改前的依据已转入审计历史，":""}请对新文字重新溯源。`:claim.status==="已删除待导出"?"这一片段将在修订副本中显示为删除，不需要重新溯源。":claim.status==="未找到候选"?"当前检索范围没有找到候选来源。可以调整提示或检索范围后重试。":"该观点尚未执行溯源。";
    $("sources").className="sources";$("sources").innerHTML=`<div class="empty-note">${message}${archived?`<small class="archived-evidence-note">已保留 ${archived} 个修改前审计快照。</small>`:""}</div>`;renderArtifact();return
  }
  const rows=allRows.map((source,index)=>({source,index})).filter(item=>state.sourceFilter==="all"||sourceFilterLabels(item.source).includes(state.sourceFilter));
  $("sources").className="sources";
  const candidateCard=({source,index})=>{
    const labels=sourceViewLabels(source);
    const previewPdf=isLocalPdf(source)?`<button class="button" data-preview-pdf="${index}">工作台查看${source.page?`（第${esc(source.page)}页）`:""}</button>`:"";
    const openSource=source.url||source.path?`<button class="button quiet" data-open-source="${index}">${source.url?"打开并定位原文":"用默认程序打开"}</button>`:"";
    const badges=[candidateReviewLabel(source),source.aiLabel,candidateOrigin(source),candidateType(source),source.matchType,source.ocrConfidence!=null?`OCR置信度 ${source.ocrConfidence}%`:"",source.timeliness,candidateHost(source),...labels].filter(Boolean);
    const preview=source.reading?.coreView||plainCandidateText(source.quote)||"没有取得可展示原文";
    return `<article class="source-card compact tier-${esc(candidateReviewTier(source))} ${state.selectedSource===index?"active":""}" data-source-index="${index}"><div class="source-title"><b>${esc(source.label||"未命名候选")}</b><div class="source-badges">${badges.map(value=>`<span>${esc(value)}</span>`).join("")}</div></div>${source.aiReason?`<p class="ai-refine-reason"><b>AI精筛：</b>${esc(source.aiReason)}</p>`:""}<p class="source-preview">${esc(preview.slice(0,180))}${preview.length>180?"…":""}</p><div class="source-meta">${esc(source.location)} · ${esc(source.method)} ${source.score==null?"":`<span class="score">匹配 ${esc(source.score)}</span>`}</div>${source.queryUsed?`<details class="query-used"><summary>查看本次检索词</summary><p>${esc(source.queryUsed)}</p></details>`:""}<div class="source-actions">${previewPdf}${openSource}${source.url?`<button class="button" data-deep-read="${index}">${source.reading?"重新深读":"深读并归类"}</button>`:""}</div></article>`;
  };
  const groups=[
    {tier:"primary",title:"主要候选",hint:"优先核对语义直接或部分对应的原文",open:true},
    {tier:"conflict",title:"可能冲突",hint:"数字、主体、否定、条件或结论可能不一致",open:true},
    {tier:"weak",title:"其他线索",hint:"主题相关、疑似无关或尚未精筛，全部保留",open:false},
  ];
  const grouped=groups.map(group=>{const items=rows.filter(item=>candidateReviewTier(item.source)===group.tier);if(!items.length)return"";return `<details class="candidate-group candidate-group-${group.tier}" ${group.open?"open":""}><summary><span><b>${group.title}</b><small>${group.hint}</small></span><em>${items.length}条</em></summary><div class="raw-candidate-list">${items.map(candidateCard).join("")}</div></details>`}).join("");
  $("sources").innerHTML=`${tracePlanHtml(claim)}${grouped||`<div class="empty-note">当前筛选条件下没有候选。</div>`}`;
  document.querySelectorAll("[data-source-index]").forEach(node=>node.onclick=event=>{if(event.target.closest("button"))return;state.selectedSource=Number(node.dataset.sourceIndex);renderSources();renderArtifact()});
  document.querySelectorAll("[data-open-source]").forEach(node=>node.onclick=async()=>{const source=allRows[Number(node.dataset.openSource)];try{const data=await api("/api/open",{method:"POST",body:JSON.stringify({path:source.path,url:source.url,locator:candidateLocator(source),page:source.page})});toast(data.message)}catch(error){toast(error.message)}});
  document.querySelectorAll("[data-deep-read]").forEach(node=>node.onclick=async()=>{const index=Number(node.dataset.deepRead);node.disabled=true;node.textContent="正在深读…";try{state.project=await api("/api/deep-read",{method:"POST",body:JSON.stringify({project:state.project,claimId:claim.id,candidateIndex:index,prompt:state.project.prompt||""})});state.selected=claim.id;state.selectedSource=index;renderSources();renderArtifact();toast("来源观点与逐字依据已显示在正文下方")}catch(error){toast(`深读失败：${error.message}`);node.disabled=false;node.textContent="深读并归类"}});
};

let pdfPreviewUrl="";
let pdfPreviewRequest=0;

function closePdfReader(){
  const dialog=$("pdf-reader-dialog");
  if(dialog.open)dialog.close();
  if(pdfPreviewUrl){URL.revokeObjectURL(pdfPreviewUrl);pdfPreviewUrl="";}
  state.pdfPreview=null;
}

function clampPdfPage(value){
  const view=state.pdfPreview;if(!view)return 1;
  const number=Number.parseInt(value,10);
  return Math.max(1,Math.min(view.pageCount,Number.isFinite(number)?number:view.page));
}

async function renderPdfPage(){
  const view=state.pdfPreview;if(!view)return;
  const image=$("pdf-page-image");const request=++pdfPreviewRequest;
  $("pdf-page-input").value=view.page;$("pdf-page-total").textContent=`/ ${view.pageCount} 页`;
  $("pdf-reader-location").textContent=`${view.location||"来源 PDF"} · 正在载入第 ${view.page} 页`;
  $("pdf-page-prev").disabled=view.page<=1;$("pdf-page-next").disabled=view.page>=view.pageCount;
  image.classList.add("loading");image.removeAttribute("src");
  try{
    const response=await fetch(`/api/pdf-page?path=${encodeURIComponent(view.path)}&page=${view.page}`,{cache:"no-store"});
    if(!response.ok){let message="PDF 页面无法显示";try{message=(await response.json()).error||message}catch{}throw new Error(message)}
    const blob=await response.blob();if(request!==pdfPreviewRequest)return;
    if(pdfPreviewUrl)URL.revokeObjectURL(pdfPreviewUrl);
    pdfPreviewUrl=URL.createObjectURL(blob);image.src=pdfPreviewUrl;
    $("pdf-reader-location").textContent=`${view.location||"来源 PDF"} · 第 ${view.page} / ${view.pageCount} 页`;
  }catch(error){if(request===pdfPreviewRequest){$("pdf-reader-location").textContent="PDF 页面载入失败。可先用默认程序打开原件。";toast(`PDF 阅读失败：${error.message}`)}}finally{if(request===pdfPreviewRequest)image.classList.remove("loading")}
}

async function openPdfReader(source){
  if(!isLocalPdf(source))return;
  try{
    const meta=await api(`/api/pdf-meta?path=${encodeURIComponent(source.path)}`);
    state.pdfPreview={path:source.path,title:source.label||meta.fileName,location:source.location||"",page:1,pageCount:Number(meta.pageCount)||1};
    state.pdfPreview.page=clampPdfPage(source.page||1);
    $("pdf-reader-title").textContent=state.pdfPreview.title;
    const dialog=$("pdf-reader-dialog");if(!dialog.open)dialog.showModal();
    await renderPdfPage();
  }catch(error){toast(`无法打开内置 PDF 阅读区：${error.message}`)}
}

$("pdf-reader-close").onclick=closePdfReader;
$("pdf-reader-dialog").addEventListener("close",()=>{if(pdfPreviewUrl){URL.revokeObjectURL(pdfPreviewUrl);pdfPreviewUrl="";}state.pdfPreview=null;});
$("pdf-page-prev").onclick=()=>{if(!state.pdfPreview)return;state.pdfPreview.page=clampPdfPage(state.pdfPreview.page-1);renderPdfPage();};
$("pdf-page-next").onclick=()=>{if(!state.pdfPreview)return;state.pdfPreview.page=clampPdfPage(state.pdfPreview.page+1);renderPdfPage();};
$("pdf-page-input").onchange=()=>{if(!state.pdfPreview)return;state.pdfPreview.page=clampPdfPage($("pdf-page-input").value);renderPdfPage();};
$("pdf-open-default").onclick=async()=>{const view=state.pdfPreview;if(!view)return;try{const data=await api("/api/open",{method:"POST",body:JSON.stringify({path:view.path,page:view.page})});toast(data.message)}catch(error){toast(error.message)}};

function focusSourceInSidebar(index){const node=document.querySelector(`[data-source-index="${index}"]`);if(!node)return;const container=node.closest(".sources");if(container){const nodeRect=node.getBoundingClientRect();const containerRect=container.getBoundingClientRect();const delta=nodeRect.top-containerRect.top-(containerRect.height-nodeRect.height)/2;container.scrollBy({top:delta,behavior:"smooth"})}else{node.scrollIntoView({block:"center",inline:"nearest",behavior:"smooth"})}node.classList.add("source-focus");setTimeout(()=>node.classList.remove("source-focus"),1800)}

document.addEventListener("click",async event=>{
  const localCheck=event.target.closest("[data-local-check]");
  if(localCheck){event.preventDefault();event.stopPropagation();runLocalConsistencyCheck();return}
  const preview=event.target.closest("[data-preview-pdf]");
  if(preview){event.preventDefault();event.stopPropagation();const claim=claimById(state.selected);const source=claim?.candidates?.[Number(preview.dataset.previewPdf)];if(source)await openPdfReader(source);return}
  const ref=event.target.closest("[data-summary-source]");
  if(ref){event.preventDefault();event.stopPropagation();const index=Number(ref.dataset.summarySource);const claim=claimById(state.selected);const source=claim?.candidates?.[index];state.summaryEvidence=null;document.querySelectorAll(".summary-evidence").forEach(node=>node.remove());state.selectedSource=index;if(source)state.sourceFilter=candidateOrigin(source);renderSources();focusSourceInSidebar(index);return}
  const summarize=event.target.closest("[data-summarize-claim],[data-resummarize-claim]");
  if(summarize){event.preventDefault();event.stopPropagation();summarizeSelectedClaim();return}
  const open=event.target.closest("[data-summary-open]");
  if(open){event.preventDefault();event.stopPropagation();const claim=claimById(state.selected);const source=claim?.candidates?.[Number(open.dataset.summaryOpen)];if(!source)return;try{const data=await api("/api/open",{method:"POST",body:JSON.stringify({path:source.path,url:source.url,locator:candidateLocator(source),page:source.page})});toast(data.message)}catch(error){toast(error.message)}return}
  const deep=event.target.closest("[data-summary-deep-read]");
  if(deep){event.preventDefault();event.stopPropagation();const claim=claimById(state.selected);const index=Number(deep.dataset.summaryDeepRead);if(!claim)return;deep.disabled=true;deep.textContent="正在深读…";try{state.project=await api("/api/deep-read",{method:"POST",body:JSON.stringify({project:state.project,claimId:claim.id,candidateIndex:index,prompt:state.project.prompt||""})});state.summaryEvidence=index;renderArtifact();renderSources();toast("依据已深读并展开")}catch(error){toast(`深读失败：${error.message}`);deep.disabled=false;deep.textContent="深读此依据"}return}
});

const previousBackHandler=$("back-button").onclick;
$("back-button").onclick=()=>{closePdfReader();state.sourceFilter="all";state.selectedSource=0;previousBackHandler()};
