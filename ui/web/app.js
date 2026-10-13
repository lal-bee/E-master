"use strict";
const roles = {
  prepare: {source: "原始资料 Excel", mapping: "字段映射配置", cleaning: "数据清洗配置", equipment: "类型编码映射配置", template: "标准模板配置", gate: "质量门禁配置"},
  verify: {source: "最初收集的原始源 Excel（回源基准）", system: "业务系统实际导出 Excel", original_config: "原始源接入配置", system_config: "系统导出接入配置", verification: "回源核验配置"}
};
const states = Object.fromEntries(Object.keys(roles).map(k => [k, {files:{}, settings:{sheet:"",header:1}, settingsDirty:false, settingsSync:null, revision:0, result:null, exportable:false, busy:false, status:"尚未运行，请选择文件与 JSON 配置，或加载合成演示。", error:false, filter:"全部", confirmed:false}]));
let active = "prepare";
const labels = {MATCH:"一致",FIELD_CHANGED:"字段变化",MISSING_IN_SYSTEM:"系统缺失",EXTRA_IN_SYSTEM:"系统新增",DUPLICATE_KEY:"重复主键",UNVERIFIABLE:"无法判定",CHANGED:"已确认变化",EQUAL:"一致",PASS:"通过",FAIL:"阻断",BLOCKER:"阻断问题",ERROR:"错误",WARNING:"警告",INFO:"提示",CLEANED:"已清洗",UNCHANGED:"未改变",NO_RULE:"无清洗规则",EMPTY_NORMALIZED:"空值已标准化",INVALID:"无效",UNMATCHED:"未映射",AMBIGUOUS:"映射歧义",CONFLICT:"映射冲突",SKIPPED:"跳过",UNMAPPED:"未映射"};
const esc = s => String(s ?? "").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const value = v => JSON.stringify(v); // None/null、空字符串、0 与 False 明确区分
const pretty = v => esc(JSON.stringify(v,null,2));
const label = v => labels[v] ? `${labels[v]} · ${v}` : v;

async function api(action, body, binary=false) {
  const response = await fetch(`/api/${action}`,{method:"POST",headers:{"Content-Type":"application/json","X-Session-Token":SESSION_TOKEN},body:JSON.stringify({flow:active,...body})});
  if (!response.ok) {
    const error = await response.json();
    if (error.revision !== null && error.revision !== undefined) states[active].revision = error.revision;
    throw new Error(error.error || "处理失败，请重新运行");
  }
  return binary ? response.blob() : response.json();
}
function invalidate(s) {s.result=null;s.exportable=false;s.filter="全部";s.status="输入已更换，旧结果失效，请重新运行";s.error=false;}
function canRun(s) {
  return !s.settingsDirty && Object.keys(roles[active]).every(role=>s.files[role]) &&
    (active==="prepare" ? s.settings.sheet && Number.isInteger(s.settings.header) && s.settings.header>0 : s.confirmed);
}
async function perform(message, work) {
  const s=states[active]; if(s.busy)return;
  s.busy=true;s.status=message;s.error=false;render();
  try {await work(s);} catch(error) {s.status=error.message;s.error=true;}
  finally {s.busy=false;render();}
}
function structure(file) {
  if(!file?.preview?.sheets)return "";
  return `<details open><summary>${esc(file.name)} · Sheet 与结构探查</summary>${file.preview.sheets.map(sheet=>`<p><b>${esc(sheet.sheet_name)}</b>：${sheet.empty?"空 Sheet":`${sheet.data_row_count} 行 / ${sheet.fields.length} 列；探查表头 ${sheet.header_start_row ?? "未识别"}～${sheet.header_end_row ?? "未识别"} 行`}</p>${sheet.fields.map(f=>`<span class="pill">${esc(f.column_letter)} · ${esc(f.name || "空表头")}</span> `).join("")}${sheet.issues.map(i=>`<p class="note">${esc(i.code)} · ${esc(i.message)} (${esc(i.row_number ?? "-")}, ${esc(i.column_letter ?? "-")})</p>`).join("")}`).join("")}<details><summary>完整探查结果（含候选表头和坐标）</summary><pre>${pretty(file.preview)}</pre></details></details>`;
}
function metric(name,v,bad=false) {return `<div class="metric ${bad?"bad":""}">${esc(name)}<strong>${esc(v)}</strong></div>`;}
function table(headers,rows) {
  return `<div class="table-wrap"><table><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join("")}</tr></thead><tbody>${rows.length?rows.map(row=>`<tr>${row.map(cell=>`<td>${cell}</td>`).join("")}</tr>`).join(""):`<tr><td colspan="${headers.length}">当前筛选下无记录</td></tr>`}</tbody></table></div>`;
}
function detail(obj) {return `<details><summary>查看原值、处理值、规则及来源</summary><pre>${pretty(obj)}</pre></details>`;}
function selected(items,s,getStatus) {return items.filter(i=>s.filter==="全部" || getStatus(i)===s.filter);}
function filterControl(s,options) {return `<div class="toolbar"><span class="note">状态筛选；展开明细可查看完整来源。值按 JSON 显示，保留类型。</span><select id="filter" aria-label="结果状态筛选"><option>全部</option>${[...new Set(options)].map(o=>`<option value="${esc(o)}" ${s.filter===o?"selected":""}>${esc(label(o))}</option>`).join("")}</select></div>`;}
function configSummary(s) {
  if(active==="prepare")return `<p class="note">当前处理范围：${esc(s.settings.sheet || "未选择 Sheet")} · 物理表头第 ${esc(s.settings.header)} 行。自动探查是候选结果；执行范围以人工选择为准。</p>`;
  const a=s.files.original_config?.preview,b=s.files.system_config?.preview,v=s.files.verification?.preview;
  return `<p class="note">原始源：${esc(a?.sheet_name ?? "未配置 Sheet")} · 表头 ${esc(a?.header_row_number ?? "未配置")} 行<br>系统导出：${esc(b?.sheet_name ?? "未配置 Sheet")} · 表头 ${esc(b?.header_row_number ?? "未配置")} 行<br>主键：${esc(v?.key_fields?.join?.(" + ") ?? "未配置")} · 比较字段：${esc(v?.comparison_fields?.join?.("、") ?? "未配置")}</p>`;
}
function preparation(s) {
  const r=s.result, g=r.gate;
  const mappings=r.mapping.sheet_results.flatMap(x=>x.entries), cleans=r.cleaning.sheet_results.flatMap(x=>x.rows.flatMap(row=>row.cells)), codes=r.equipment.sheet_results.flatMap(x=>x.audits);
  const options=[...mappings.map(x=>x.status),...cleans.map(x=>x.action),...codes.map(x=>x.status),...g.issues.map(x=>x.severity)];
  return `<div class="metrics">${metric("质量门禁",label(g.status),!g.can_proceed)}${metric("预览数据行",r.template.summary.row_count)}${metric("门禁问题",g.summary.issue_count,g.summary.issue_count>0)}</div><p class="note">${g.can_proceed?"门禁 PASS，允许导出。演示模板不能直接用于真实系统。":"门禁 FAIL，正式模板导出已阻断。请查看下方问题并修正源资料或配置。"}</p>${filterControl(s,options)}
  <h3>字段映射预览</h3>${table(["源列","原字段 → 标准字段","状态 / 原因","明细"],selected(mappings,s,x=>x.status).map(x=>[esc(x.source_column_letter),`${esc(x.source_field_name)} → ${esc(x.standard_field_name ?? "未映射")}`,`${esc(label(x.status))}<br>${esc(x.reason)}`,detail(x)]))}
  <h3>清洗预览</h3>${table(["来源单元格","标准字段","原值 → 处理值","状态 / 规则","明细"],selected(cleans,s,x=>x.action).map(x=>[esc(`${x.sheet_name}!${x.source_column_letter}${x.row_number}`),esc(x.standard_field_name),`${esc(value(x.original_value))} → ${esc(value(x.cleaned_value))}`,`${esc(label(x.action))}<br>${esc(x.rule_code ?? "无")}`,detail(x)]))}
  <h3>编码映射预览</h3>${table(["来源","类型原值 → 清洗值","系统编码","状态 / 原因","明细"],selected(codes,s,x=>x.status).map(x=>[esc(`${x.sheet_name}!${x.source_column_letter ?? ""}${x.row_number}`),`${esc(value(x.original_value))} → ${esc(value(x.cleaned_value))}`,esc(x.system_code ?? "未生成"),`${esc(label(x.status))}<br>${esc(x.reason)}`,detail(x)]))}
  <h3>质量门禁问题</h3>${table(["严重级别","来源","规则 / 原因","明细"],selected(g.issues,s,x=>x.severity).map(x=>[esc(label(x.severity)),esc(`${x.sheet}!${x.column_letter ?? ""}${x.row_number ?? ""}`),`${esc(x.rule_code)}<br>${esc(x.reason)}`,detail(x)]))}
  <details><summary>模板装配预览（尚未导出）</summary><pre>${pretty(r.template)}</pre></details>`;
}
function verification(s) {
  const r=s.result,m=r.summary;
  const options=r.records.map(x=>x.status);
  const locs = xs => xs.map(x=>`${x.sheet_name}!${x.row_number}`).join(" / ") || "无";
  function recordDetail(x) {
    const find=(rows,locations)=>rows.filter(row=>locations.some(loc=>loc.row_number===row.source_row_number && loc.sheet_name===row.sheet_name));
    return {"核验记录":x,"原始源行（含原值与处理值）":find(r.original_source.rows,x.original_source_rows),"系统导出行（含原值与处理值）":find(r.system_export.rows,x.system_export_rows)};
  }
  return `<div class="metrics">${metric("核验完成状态",m.verification_complete?"已完成":"未完成",!m.verification_complete)}${metric("数据一致性",!m.verification_complete?"无法判定":m.data_consistent?"一致":"不一致",!m.data_consistent)}${metric("匹配对 / 一致行",`${m.matched_row_count} / ${m.unchanged_row_count}`)}${metric("已确认变化字段",m.changed_field_count,m.changed_field_count>0)}${metric("不可判定字段",m.unverifiable_field_count,m.unverifiable_field_count>0)}${metric("异常问题对象",m.issue_count,m.issue_count>0)}</div>
  <p class="note">已确认差异与无法判定问题可同时存在。未完成不等于无差异；问题报告仍可导出。统计单位不可直接相加。</p>
  <details><summary>完整统计（输入行、缺失、新增、冲突组及字段数）</summary><pre>${pretty(m)}</pre></details>${filterControl(s,options)}
  <h3>核验记录与差异</h3>${table(["主键 / 状态","双方来源行","字段原值 → 处理值 / 原因","完整明细"],selected(r.records,s,x=>x.status).map(x=>[
    `${esc(value(x.key_values))}<br><span class="pill ${x.status==="MATCH"?"":"warn"}">${esc(label(x.status))}</span>`,
    `原始源：${esc(locs(x.original_source_rows))}<br>系统导出：${esc(locs(x.system_export_rows))}`,
    x.field_differences.map(d=>`<p><b>${esc(d.standard_field_name)} · ${esc(label(d.status))}</b><br>原始源 ${esc(value(d.original_source.original_value))} → ${esc(value(d.original_source.normalized_value))}<br>系统导出 ${esc(value(d.system_export.original_value))} → ${esc(value(d.system_export.normalized_value))}<br>${esc(d.reason)}</p>`).join("") || "行级结论；展开查看双方原始数据",
    detail(recordDetail(x))]))}
  <h3>异常问题（保留全部问题，不受行状态筛选影响）</h3>${table(["问题编码 / 类别","原因","来源 / 明细"],r.issues.map(x=>[`${esc(x.code)}<br>${esc(x.category)}`,esc(x.message),detail(x)]))}`;
}
function render() {
  const s=states[active], prep=active==="prepare", fileFields=Object.entries(roles[active]).map(([role,name])=>`<div class="field"><label for="file-${role}">${esc(name)}</label><input type="file" id="file-${role}" data-role="${role}" accept="${["source","system"].includes(role)?".xlsx":".json"}" ${s.busy?"disabled":""}><div class="filename">${esc(s.files[role]?.name ?? "尚未选择")}</div>${s.files[role] && !["source","system"].includes(role)?`<details><summary>查看所用 JSON 规则</summary><pre>${pretty(s.files[role].preview)}</pre></details>`:""}</div>`).join("");
  document.querySelectorAll("[data-tab]").forEach(b=>{b.classList.toggle("active",b.dataset.tab===active);b.disabled=s.busy;});
  document.querySelector("#workspace").innerHTML=`<div class="flow-head"><div><h2>${prep?"原始资料整理":"导入后回源核验"}</h2><p>${prep?"探查 → 字段映射 → 清洗与编码 → 质量门禁 → 标准模板":"最初原始源 + 系统实际导出 → 主键匹配 → 差异与异常 → 核验报告"}</p></div><div class="row"><button id="demo" ${s.busy?"disabled":""}>加载合成演示</button>${prep?`<button id="fail-demo" ${s.busy?"disabled":""}>加载门禁失败演示</button>`:""}</div></div>
  <div class="demo">演示配置，尚未适配真实业务模板。${prep?"真实导入前必须确认字段、编码、门禁和模板规则。":"回源基准必须是最初收集的原始源文档，整理后的标准导入模板不能替代。"}</div>
  <div class="columns"><section><div class="card"><h3>01 选择输入与规则</h3>${fileFields}${prep?`<div class="row"><div class="field"><label for="sheet">处理 Sheet</label><select id="sheet" ${s.busy?"disabled":""}><option value="">请选择</option>${(s.files.source?.preview?.sheets || []).map(x=>`<option value="${esc(x.sheet_name)}" ${s.settings.sheet===x.sheet_name?"selected":""}>${esc(x.sheet_name)}</option>`).join("")}</select></div><div class="field"><label for="header">物理表头行</label><input type="number" id="header" min="1" step="1" value="${esc(s.settings.header)}" ${s.busy?"disabled":""}></div></div>`:`<label><input type="checkbox" id="confirm" ${s.confirmed?"checked":""} ${s.busy?"disabled":""}> 我确认所选回源基准是最初收集的原始源文档</label><p class="note">双方 Sheet、物理表头、主键和字段来自接入 / 核验配置；展开 JSON 可查看。</p>`}
  <div class="actions"><button id="run" class="primary" ${s.busy || !canRun(s)?"disabled":""}>${prep?"运行整理与质量门禁":"执行回源核验"}</button><button id="export" ${s.busy || !s.result || !s.exportable?"disabled":""}>${prep?"导出标准模板":"导出核验报告"}</button></div><p class="note">输入完整且范围确认后可运行。生成新文件后由浏览器下载。${prep?"门禁 FAIL 时禁止导出。":"有差异或无法判定问题时仍可导出。"}</p></div>
  <div class="card"><h3>输入结构探查</h3>${structure(s.files.source)}${structure(s.files.system)}${!s.files.source?'<p class="note">选择 Excel 后展示所有 Sheet、候选表头与结构问题。</p>':""}</div></section>
  <section><div id="status" class="status ${s.error?"error":""}" role="status" aria-live="polite">${esc(s.status)}</div><div class="card" id="results"><h3>02 预览、问题与追溯</h3>${configSummary(s)}${s.result?(prep?preparation(s):verification(s)):'<div class="empty"><b>等待运行</b>选定输入并运行后，这里会显示当前输入的结果。</div>'}</div></section></div>`;
  bind();
}
function bind() {
  const s=states[active];
  document.querySelectorAll("input[type=file]").forEach(input=>input.addEventListener("change",()=>{
    const file=input.files[0];if(s.busy || !file)return;invalidate(s);delete s.files[input.dataset.role];if(active==="verify" && input.dataset.role==="source")s.confirmed=false;
    perform("正在读取文件并探查 / 加载规则…",async()=>{
      await s.settingsSync;
      const cleared=await api("invalidate",{role:input.dataset.role});s.revision=cleared.revision;
      if(file.size>32*1024*1024)throw new Error("文件超过 32 MB，请使用单文件小样本");
      const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(",")[1]);reader.onerror=()=>reject(new Error("文件读取失败，请重新选择"));reader.readAsDataURL(file);});
      const response=await api("upload",{role:input.dataset.role,name:file.name,data});s.revision=response.revision;s.files[input.dataset.role]=response;
      if(active==="prepare" && input.dataset.role==="source") {const sheet=response.preview.sheets.find(x=>!x.empty);s.settings={sheet:sheet?.sheet_name ?? "",header:sheet?.header_row ?? 1};const d=await api("settings",{settings:s.settings});s.revision=d.revision;}
      s.status="文件已加载，旧结果已失效。请确认结构与规则后运行。";
    });
  }));
  ["sheet","header"].forEach(id=>document.getElementById(id)?.addEventListener("change",()=>{
    if(s.busy)return;
    const settings={sheet:document.getElementById("sheet").value,header:Number(document.getElementById("header").value)};invalidate(s);s.settings=settings;
    perform("正在更新输入范围…",async()=>{
      if(await s.settingsSync === false){const d=await api("invalidate",{});s.revision=d.revision;}
      const d=await api("settings",{settings});s.revision=d.revision;s.settingsDirty=false;s.settingsSync=null;s.status="Sheet / 表头已更换，旧结果失效，请重新运行";
    });
  }));
  // 首次按键即撤销服务端结果；保留编辑中的输入控件，失焦后同步最终范围。
  document.getElementById("header")?.addEventListener("input",()=>{
    if(s.busy)return;
    if(!s.settingsDirty){
      invalidate(s);s.settingsDirty=true;
      s.settingsSync=api("invalidate",{}).then(d=>{s.revision=d.revision;return true;}).catch(error=>{s.status=error.message;s.error=true;document.getElementById("status").textContent=s.status;return false;});
    }
    document.getElementById("export").disabled=true;document.getElementById("run").disabled=true;
    document.getElementById("status").textContent=s.status;
    document.getElementById("results").innerHTML='<h3>02 预览、问题与追溯</h3><div class="empty"><b>等待运行</b>表头正在修改，旧结果已失效；确认范围后重新运行。</div>';
  });
  document.getElementById("confirm")?.addEventListener("change",e=>{s.confirmed=e.target.checked;render();});
  document.getElementById("filter")?.addEventListener("change",e=>{s.filter=e.target.value;render();});
  document.getElementById("run").onclick=()=>{if(s.busy || !canRun(s))return;invalidate(s);perform("运行中，请等待；结果与导出按钮已锁定…",async()=>{const d=await api("run",{revision:s.revision,original_confirmed:s.confirmed});s.result=d.result;s.revision=d.revision;s.exportable=d.exportable;s.status=active==="prepare"?`处理结束 · 质量门禁 ${s.result.gate.status}`:"核验处理结束 · 完成状态与数据一致性见下方结果";});};
  document.getElementById("export").onclick=()=>perform("正在生成新的 Excel 文件…",async()=>{const blob=await api("export",{revision:s.revision},true);const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=`${active==="prepare"?"标准模板":"回源核验报告"}-${Date.now()}.xlsx`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);s.status="Excel 已生成并交给浏览器下载，请查看下载列表。";});
  const demo=fail=>{if(s.busy)return;invalidate(s);perform("正在生成合成 Excel 并加载演示配置…",async()=>{await s.settingsSync;const d=await api("demo",{fail});s.files=d.files;s.settings=d.settings;s.settingsDirty=false;s.settingsSync=null;s.revision=d.revision;s.confirmed=false;s.status="合成演示已加载。演示配置尚未适配真实业务模板，请运行。";});};
  document.getElementById("demo").onclick=()=>demo(false);
  document.getElementById("fail-demo")?.addEventListener("click",()=>demo(true));
}
document.querySelectorAll("[data-tab]").forEach(button=>button.onclick=()=>{if(states[active].busy)return;active=button.dataset.tab;render();});
render();
