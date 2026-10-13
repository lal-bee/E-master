// P1-12 Windows Edge 实际页面冒烟；标准 Node.js，无第三方依赖。
// 内置 Browser 工具不可用时使用。服务须先启动；不用于启动产品。
import {spawn, execFileSync} from 'node:child_process';
import {mkdtemp, mkdir, readFile, readdir, writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {tmpdir} from 'node:os';
import {resolve, join} from 'node:path';
import assert from 'node:assert/strict';

const base = process.argv[2] || 'http://127.0.0.1:8765';
const root = resolve('.');
const artifacts = await mkdtemp(join(tmpdir(), 'e-master-browser-'));
const screenshots = join(root, 'docs', 'images');
await mkdir(screenshots, {recursive:true});
const demo = JSON.parse(execFileSync(join(root,'.venv','Scripts','python.exe'),
  ['examples/ui_demo.py','--output',join(artifacts,'samples')],{encoding:'utf8'}));
const acceptance = JSON.parse(execFileSync(join(root,'.venv','Scripts','python.exe'),
  ['tests/ui/acceptance_fixture.py','--create',join(artifacts,'acceptance')],{encoding:'utf8'}));
const edge = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',[
  '--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check',
  '--remote-debugging-port=0',`--user-data-dir=${join(artifacts,'edge-profile')}`,base
],{windowsHide:true,stdio:['ignore','ignore','pipe']});
let debug='';
edge.stderr.on('data', chunk=>debug+=chunk);
const sleep = ms=>new Promise(r=>setTimeout(r,ms));
async function until(check, description) {
  const end=Date.now()+20000;
  while(Date.now()<end){if(await check())return;await sleep(100);}
  throw new Error(`超时：${description}`);
}
let socket, secondSocket;
try {
  await until(()=>/DevTools listening on (ws:\/\/[^\s]+)/.test(debug),'浏览器调试连接');
  const endpoint=debug.match(/DevTools listening on (ws:\/\/[^\s]+)/)[1];
  const port=new URL(endpoint).port;
  let page;
  await until(async()=>{page=(await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.type==='page'&&t.url.startsWith(base));return page;},'页面打开');
  socket=new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((res,rej)=>{socket.onopen=res;socket.onerror=rej;});
  let id=0;
  const pending=new Map(),exceptions=[],paused=[],dialogs=[];
  socket.onmessage=e=>{const message=JSON.parse(e.data);if(message.id){const p=pending.get(message.id);if(p){pending.delete(message.id);message.error?p.reject(message.error):p.resolve(message.result);}}else if(message.method==='Runtime.exceptionThrown'){exceptions.push(message.params);}else if(message.method==='Fetch.requestPaused'){paused.push(message.params);}else if(message.method==='Page.javascriptDialogOpening'){dialogs.push(message.params);}};
  const call=(method,params={})=>new Promise((resolve,reject)=>{const next=++id;pending.set(next,{resolve,reject});socket.send(JSON.stringify({id:next,method,params}));});
  const evaluate=async expression=>{const response=await call('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(response.exceptionDetails)throw new Error(JSON.stringify(response.exceptionDetails));return response.result.value;};
  const click=selector=>evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
  const text=()=>evaluate('document.querySelector("#status")?.textContent');
  const ready=()=>evaluate('document.querySelector("#run") && !document.querySelector("#run").disabled');
  const idle=()=>evaluate('document.querySelector("#demo") && !document.querySelector("#demo").disabled');
  const upload=async(role,path)=>{
    const document=await call('DOM.getDocument');
    const node=await call('DOM.querySelector',{nodeId:document.root.nodeId,selector:`#file-${role}`});
    await call('DOM.setFileInputFiles',{nodeId:node.nodeId,files:[path]});
    await until(async()=>await idle() && (await text()).includes('文件已加载'),'上传 '+role);
  };
  const downloaded=async()=> (await readdir(artifacts)).filter(n=>n.endsWith('.xlsx'));
  const downloadReport=async(name,generic=false)=>{
    const state=await evaluate('JSON.stringify(states.verify.result)');
    const snapshot=join(artifacts,name+'.json');
    await writeFile(snapshot,state);
    const before=await downloaded();
    await click('#export');
    await until(async()=> (await downloaded()).length===before.length+1,'报告下载 '+name);
    await until(idle,'报告导出结束');
    const report=(await downloaded()).find(file=>!before.includes(file));
    console.log(execFileSync(join(root,'.venv','Scripts','python.exe'),
      ['tests/ui/acceptance_fixture.py','--report',join(artifacts,report),'--result',snapshot,...(generic?['--generic']:[])],{encoding:'utf8'}).trim());
    return report;
  };
  const shot=async name=>{const data=await call('Page.captureScreenshot',{format:'png'});await writeFile(join(screenshots,name),Buffer.from(data.data,'base64'));};
  await call('Runtime.enable');
  await call('Page.enable');
  await call('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false});
  await call('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:artifacts});
  await until(idle,'初始页面');
  console.log('Initial UI:',await evaluate('document.querySelector("nav").textContent'));
  assert(await evaluate('document.querySelector("#export").disabled'));

  // 通过真正的浏览器 file input 加载每一个文件角色。
  for(const [role,path] of Object.entries(demo.prepare)){
    await upload(role,path);
  }
  assert.equal(await evaluate('document.querySelector("#header").value'),'2');
  // 延迟浏览器请求以观察运行中的锁定状态，不依靠任务恰好足够慢。
  await call('Fetch.enable',{patterns:[{urlPattern:base+'/api/run',requestStage:'Request'}]});
  await click('#run');
  await until(()=>paused.length===1,'拦截运行请求');
  assert(await evaluate('[...document.querySelectorAll("#workspace input, #workspace select, #workspace button, nav button")].every(x=>x.disabled)'));
  await click('#run');await click('#demo');
  assert.equal(paused.length,1);
  await call('Fetch.continueRequest',{requestId:paused.shift().requestId});
  await call('Fetch.disable');
  await until(async()=>await ready() && (await text()).includes('质量门禁 PASS'),'整理 PASS');
  assert.equal(await evaluate('document.querySelector("#export").disabled'),false);
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("D-PUMP")'));
  await shot('p1-12-prepare.png');
  await call('Fetch.enable',{patterns:[{urlPattern:base+'/api/export',requestStage:'Request'}]});
  await click('#export');
  await until(()=>paused.length===1,'拦截下载请求');
  await click('#export');assert.equal(paused.length,1);
  await call('Fetch.continueRequest',{requestId:paused.shift().requestId});
  await call('Fetch.disable');
  await until(async()=> (await readdir(artifacts)).filter(n=>n.endsWith('.xlsx')).length===1,'模板下载');
  await until(ready,'导出结束');
  console.log(execFileSync(join(root,'.venv','Scripts','python.exe'),
    ['tests/ui/acceptance_fixture.py','--template',join(artifacts,(await downloaded())[0])],{encoding:'utf8'}).trim());

  const oldRevision=await evaluate('states.prepare.revision');
  await evaluate('document.querySelector("#header").value="1"; document.querySelector("#header").dispatchEvent(new Event("input",{bubbles:true}));');
  assert(await evaluate('document.querySelector("#export").disabled && document.querySelector("#results").textContent.includes("等待运行")'));
  await until(()=>evaluate(`states.prepare.revision > ${oldRevision}`),'尚未失焦时撤销服务端结果');
  const stale=await evaluate(`fetch('/api/export',{method:'POST',headers:{'Content-Type':'application/json','X-Session-Token':SESSION_TOKEN},body:JSON.stringify({flow:'prepare',revision:${oldRevision}})}).then(r=>r.status)`);
  assert.equal(stale,400);
  await evaluate('document.querySelector("#header").dispatchEvent(new Event("change",{bubbles:true}));');
  await until(async()=>await ready() && (await text()).includes('旧结果失效'),'表头变更');
  assert(await evaluate('document.querySelector("#export").disabled'));
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("等待运行")'));
  await evaluate('document.querySelector("#sheet").value="空白Sheet";document.querySelector("#sheet").dispatchEvent(new Event("change",{bubbles:true}));');
  await until(async()=>await idle() && (await text()).includes('旧结果失效'),'Sheet 变更');
  assert(await evaluate('document.querySelector("#export").disabled'));

  await click('#demo');
  await until(async()=>await ready() && (await text()).includes('合成演示已加载'),'重新加载 PASS');
  await click('#run');
  await until(async()=>await ready() && (await text()).includes('质量门禁 PASS'),'再次 PASS');
  await evaluate('globalThis.NativeFileReader=FileReader;globalThis.FileReader=class {readAsDataURL(){queueMicrotask(()=>this.onerror());}};');
  const doc=await call('DOM.getDocument');
  const failedNode=await call('DOM.querySelector',{nodeId:doc.root.nodeId,selector:'#file-source'});
  await call('DOM.setFileInputFiles',{nodeId:failedNode.nodeId,files:[demo.prepare.source]});
  await until(async()=>await idle() && (await text()).includes('文件读取失败'),'FileReader 失败');
  assert(await evaluate('document.querySelector("#run").disabled && document.querySelector("#export").disabled && states.prepare.result===null'));
  const missing=await evaluate(`fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json','X-Session-Token':SESSION_TOKEN},body:JSON.stringify({flow:'prepare',revision:states.prepare.revision})}).then(r=>r.json())`);
  assert(missing.error.includes('文件未选择完整'));
  await evaluate('globalThis.FileReader=NativeFileReader;');
  await click('#fail-demo');
  await until(async()=>await ready() && (await text()).includes('合成演示已加载'),'FAIL 演示');
  await click('#run');
  await until(async()=>await ready() && (await text()).includes('质量门禁 FAIL'),'整理 FAIL');
  assert(await evaluate('document.querySelector("#export").disabled'));
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("DEMO_DUPLICATE")'));
  await shot('p1-12-gate-fail.png');

  await click('[data-tab="verify"]');
  assert(await evaluate('document.querySelector("#export").disabled'));
  await click('#demo');
  await until(async()=> (await text()).includes('合成演示已加载'),'核验演示');
  await click('#confirm');
  await click('#run');
  await until(async()=>await ready() && (await text()).includes('核验处理结束'),'核验完成');
  assert.equal(await evaluate('document.querySelector("#export").disabled'),false);
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("已确认变化")'));
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("无法判定")'));
  await evaluate('document.querySelector("#filter").value="UNVERIFIABLE";document.querySelector("#filter").dispatchEvent(new Event("change",{bubbles:true}));');
  await evaluate('document.querySelector("tbody details").open=true');
  assert(await evaluate('document.querySelector("tbody").textContent.includes("原始源行（含原值与处理值）")'));
  await shot('p1-12-verification.png');
  await downloadReport('default-ui-result',true);

  for(const [role,path] of Object.entries(acceptance.verify))await upload(role,path);
  await click('#confirm');await click('#run');
  await until(async()=>await ready() && (await text()).includes('核验处理结束'),'特殊文本核验');
  assert.equal(await evaluate('document.querySelectorAll("#workspace img,#workspace svg,#workspace em,[onerror],[onload]").length'),0);
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("<img src=x onerror=alert(7)>")'));
  assert(await evaluate('states.verify.result.original_source.rows[0].fields.find(x=>x.standard_field_name==="long_number").original_value.type === "int"'));
  await shot('p1-12-report-values.png');
  await downloadReport('special-text-ui-result');

  // 两个真实标签页各自获取会话，第二页不得影响第一页输入 / 结果 / 下载。
  const firstSnapshot=await evaluate('JSON.stringify(states.verify.result)');
  const firstToken=await evaluate('SESSION_TOKEN');
  const target=await call('Target.createTarget',{url:base});
  let secondPage;
  await until(async()=>{secondPage=(await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.id===target.targetId);return secondPage;},'第二标签页');
  secondSocket=new WebSocket(secondPage.webSocketDebuggerUrl);
  await new Promise((res,rej)=>{secondSocket.onopen=res;secondSocket.onerror=rej;});
  let secondId=0;const secondPending=new Map();
  secondSocket.onmessage=e=>{const m=JSON.parse(e.data);if(m.id){const p=secondPending.get(m.id);if(p){secondPending.delete(m.id);m.error?p.reject(m.error):p.resolve(m.result);}}else if(m.method==='Runtime.exceptionThrown'){exceptions.push(m.params);}};
  const secondCall=(method,params={})=>new Promise((resolve,reject)=>{const id=++secondId;secondPending.set(id,{resolve,reject});secondSocket.send(JSON.stringify({id,method,params}));});
  const secondEval=async expression=>{const r=await secondCall('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw new Error(JSON.stringify(r.exceptionDetails));return r.result.value;};
  await secondCall('Runtime.enable');
  await until(()=>secondEval('!!document.querySelector("#demo")'),'第二页初始状态');
  assert.notEqual(await secondEval('SESSION_TOKEN'),firstToken);
  await secondEval('document.querySelector("[data-tab=verify]").click();document.querySelector("#demo").click();');
  await until(()=>secondEval('!states.verify.busy && Object.keys(states.verify.files).length===5'),'第二页演示');
  await secondEval('document.querySelector("#confirm").click();document.querySelector("#run").click();');
  await until(()=>secondEval('!states.verify.busy && states.verify.result!==null'),'第二页核验');
  assert.equal(await evaluate('JSON.stringify(states.verify.result)'),firstSnapshot);
  assert.notEqual(await secondEval('states.verify.result.original_source.source_file'),await evaluate('states.verify.result.original_source.source_file'));
  const secondState=join(artifacts,'second-tab-ui-result.json');await writeFile(secondState,await secondEval('JSON.stringify(states.verify.result)'));
  const beforeSecond=await downloaded();
  await secondEval('document.querySelector("#export").click();');
  await until(async()=> (await downloaded()).length===beforeSecond.length+1,'第二页报告下载');
  const secondReport=(await downloaded()).find(n=>!beforeSecond.includes(n));
  console.log(execFileSync(join(root,'.venv','Scripts','python.exe'),['tests/ui/acceptance_fixture.py','--report',join(artifacts,secondReport),'--result',secondState,'--generic'],{encoding:'utf8'}).trim());
  assert.equal(await evaluate('JSON.stringify(states.verify.result)'),firstSnapshot);
  await downloadReport('first-tab-after-second-result');

  await upload('verification',acceptance.invalid_config);
  assert(await evaluate('states.verify.result===null && document.querySelector("#export").disabled'));
  await click('#run');
  await until(async()=>await idle() && await evaluate('states.verify.error'),'运行失败提示');
  assert(await evaluate('states.verify.result===null && document.querySelector("#export").disabled && document.querySelector("#results").textContent.includes("等待运行")'));
  assert.equal(await evaluate('document.querySelectorAll("#workspace img,#workspace svg,#workspace em,[onerror],[onload]").length'),0);
  await shot('p1-12-run-failed.png');
  await click('#demo');
  await until(async()=> (await text()).includes('合成演示已加载'),'换输入');
  assert(await evaluate('document.querySelector("#export").disabled'));
  assert(await evaluate('document.querySelector("#workspace").textContent.includes("等待运行")'));
  await shot('p1-12-input-changed.png');
  assert.deepEqual(exceptions,[]);
  assert.deepEqual(dialogs,[]);
  for(const [path,hash] of Object.entries(acceptance.hashes))assert.equal(createHash('sha256').update(await readFile(path)).digest('hex'),hash);
  const outputs=(await readdir(artifacts)).filter(n=>n.endsWith('.xlsx'));
  const evidence={result:'PASS',artifacts,outputs,checks:['file uploads','run and download input locking','duplicate clicks','PASS template download','header input invalidation before blur','Sheet invalidation','FileReader failure clears server role','FAIL export disabled','verification differences and issues','filter and details','four-Sheet report readback against UI result','leading zeros long codes zero False and formula strings','HTML text escaping','two browser tabs with isolated inputs results and downloads','configuration change and run failure','input invalidation','input hashes unchanged','no page exceptions or dialogs']};
  await writeFile(join(artifacts,'acceptance-evidence.json'),JSON.stringify(evidence,null,2));
  console.log(JSON.stringify(evidence,null,2));
  await call('Browser.close');
} finally {
  socket?.close();
  secondSocket?.close();
  edge.kill();
  // 保留验证样本、下载及截图，不执行文件清理。
}
