const $ = (selector) => document.querySelector(selector);
let currentTask = null;
let issues = [];
let activeFilter = "all";

function showToast(message){const toast=$("#toast");toast.textContent=message;toast.classList.remove("hidden");setTimeout(()=>toast.classList.add("hidden"),2600)}
function setView(name){for(const id of ["uploadView","loadingView","resultsView"]){$("#"+id).classList.toggle("hidden",id!==name)}}
function escapeHtml(value){return String(value??"").replace(/[&<>'"]/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]))}

document.querySelectorAll('input[type="file"]').forEach(input=>input.addEventListener("change",()=>input.closest(".upload-card").classList.toggle("has-file",input.files.length>0)));

$("#auditForm").addEventListener("submit",async event=>{
  event.preventDefault();
  const subtitleFiles=[...$("#subtitles").files];
  const lexicon=$("#lexicon").files[0];
  if(!subtitleFiles.length||!lexicon){showToast("请先选择字幕文件和审核词表");return}
  setView("loadingView");
  const form=new FormData();
  subtitleFiles.forEach(file=>form.append("subtitles",file));
  form.append("lexicon",lexicon);
  if($("#document").files[0])form.append("document",$("#document").files[0]);
  form.append("video_names",JSON.stringify([...$("#videos").files].map(file=>file.name)));
  try{
    const response=await fetch("/api/check",{method:"POST",body:form});
    if(!response.ok)throw new Error((await response.json()).detail||"检查失败");
    const data=await response.json();currentTask=data.task_id;issues=data.issues;
    $("#issueCount").textContent=issues.length;
    $("#resultSummary").textContent=`已检查 ${data.summary.files} 个字幕文件，使用词表“${data.lexicon_name}”。`;
    renderIssues();setView("resultsView");
  }catch(error){setView("uploadView");showToast(error.message)}
});

function renderIssues(){
  const list=$("#issueList");list.innerHTML="";
  const shown=issues.filter(item=>activeFilter==="all"||item.system_result===activeFilter);
  $("#emptyState").classList.toggle("hidden",shown.length!==0);
  shown.forEach(item=>{
    const statusClass=item.system_result==="需要确认"?"confirm":item.system_result==="已经处理好"?"done":"";
    const card=document.createElement("article");card.className="issue-card";card.dataset.id=item.id;
    card.innerHTML=`<div class="issue-meta"><span class="time">${escapeHtml(item.time)}</span><span class="video">${escapeHtml(item.video)}</span><span class="status ${statusClass}">${escapeHtml(item.system_result)}</span></div><div class="issue-body"><h3>${escapeHtml(item.kind)}</h3><div class="sentence">${escapeHtml(item.sentence)}</div><p class="detail-title">问题在哪里</p><p>${escapeHtml(item.detail)}</p><p class="detail-title">建议怎么处理</p><p>${escapeHtml(item.suggestion)}</p></div><div class="review-box"><label>你的复核结果</label><select><option value="">请选择</option>${["确实有问题","系统提示错了","已经处理好","我不确定"].map(v=>`<option ${item.review_result===v?"selected":""}>${v}</option>`).join("")}</select><textarea placeholder="需要时填写正确句子或原因">${escapeHtml(item.note)}</textarea></div>`;
    card.querySelector("select").addEventListener("change",e=>item.review_result=e.target.value);
    card.querySelector("textarea").addEventListener("input",e=>item.note=e.target.value);
    list.appendChild(card);
  });
}

document.querySelectorAll(".filter").forEach(button=>button.addEventListener("click",()=>{document.querySelectorAll(".filter").forEach(x=>x.classList.remove("active"));button.classList.add("active");activeFilter=button.dataset.filter;renderIssues()}));

async function saveReview(){
  if(!currentTask)return false;
  const response=await fetch(`/api/tasks/${currentTask}/review`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({issues})});
  if(!response.ok){showToast("保存失败，请重试");return false}showToast("复核结果已保存");return true;
}
$("#saveButton").addEventListener("click",saveReview);
$("#downloadButton").addEventListener("click",async()=>{if(await saveReview())location.href=`/api/tasks/${currentTask}/report.xlsx`});
$("#suggestionsButton").addEventListener("click",()=>location.href="/api/improvement-suggestions.xlsx");
$("#backButton").addEventListener("click",async()=>{if(currentTask)await fetch(`/api/tasks/${currentTask}`,{method:"DELETE"});currentTask=null;issues=[];$("#auditForm").reset();document.querySelectorAll(".upload-card").forEach(card=>card.classList.remove("has-file"));setView("uploadView")});
