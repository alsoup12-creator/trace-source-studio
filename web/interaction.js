"use strict";

function openTraceDialog(claimId=null){
  state.pendingClaim=claimId;
  state.selected=claimId||state.selected;
  $("run-prompt").value=state.project.prompt||$("run-prompt").value;
  $("source-folder").value=state.project.sourceFolder||$("source-folder").value;
  $("run-dialog-title").textContent=claimId?"溯源当前段落":"开始溯源";
  $("confirm-run").textContent=claimId?"开始溯源当前段落":"开始查找候选来源";
  if(claimId){renderClaims();renderArtifact();renderSources()}
  $("run-dialog").showModal();
}

$("run-button").onclick=()=>openTraceDialog();
$("confirm-run").onclick=()=>state.pendingClaim?runSingleTrace(state.pendingClaim):runTrace();
$("run-dialog").addEventListener("close",()=>{
  if(!state.runningClaim&&!state.runningAll)state.pendingClaim=null;
});
$("back-button").onclick=()=>{state.pendingClaim=null;resetWorkspace()};
