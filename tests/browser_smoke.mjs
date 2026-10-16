// Windows Edge 浏览器验收，使用 Node 内置 CDP 客户端，无第三方依赖。
import assert from 'node:assert/strict';
import {spawn, execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {mkdir, mkdtemp, readFile, readdir, writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join, resolve} from 'node:path';

const root = resolve('.');
const python = join(root, '.venv', 'Scripts', 'python.exe');
const artifacts = await mkdtemp(join(tmpdir(), 'e-master-converter-acceptance-'));
const images = join(root, 'docs', 'images');
await mkdir(images, {recursive: true});
const fixtures = JSON.parse(execFileSync(python,
  ['-m', 'tests.browser_fixture', '--create', join(artifacts, 'samples')], {encoding: 'utf8'}));
const preserved = [...Object.values(fixtures), join(root, 'docs', '设备小类编码字段映射配置文件.xlsx'),
  join(root, 'docs', '安全阀台账模版-2026.10.09.xlsx')];
const hash = async path => createHash('sha256').update(await readFile(path)).digest('hex');
const hashes = await Promise.all(preserved.map(hash));
const server = spawn(python, ['-m', 'tests.browser_fixture', '--serve'],
  {cwd: root, windowsHide: true, stdio: ['pipe', 'pipe', 'pipe']});
let serverText = '', serverErrors = '';
server.stdout.on('data', chunk => serverText += chunk);
server.stderr.on('data', chunk => serverErrors += chunk);
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
async function until(check, label) {
  const end = Date.now() + 20000;
  while (Date.now() < end) { if (await check()) return; await sleep(100); }
  throw new Error(`超时：${label}\n${serverErrors}`);
}

let edge, socket;
const exceptions = [], dialogs = [], paused = [], downloads = [];
try {
  await until(() => /http:\/\/127\.0\.0\.1:\d+/.test(serverText), '本地服务启动');
  const base = serverText.match(/http:\/\/127\.0\.0\.1:\d+/)[0];
  edge = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', [
    '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
    '--remote-debugging-port=0', `--user-data-dir=${join(artifacts, 'edge-profile')}`, base
  ], {windowsHide: true, stdio: ['ignore', 'ignore', 'pipe']});
  let debug = '';
  edge.stderr.on('data', chunk => debug += chunk);
  await until(() => /DevTools listening on (ws:\/\/[^\s]+)/.test(debug), '浏览器启动');
  const port = new URL(debug.match(/DevTools listening on (ws:\/\/[^\s]+)/)[1]).port;
  let page;
  await until(async () => {
    page = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json())
      .find(item => item.type === 'page' && item.url.startsWith(base));
    return page;
  }, '浏览器页面');
  socket = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
  let id = 0;
  const pending = new Map();
  socket.onmessage = event => {
    const message = JSON.parse(event.data);
    if (message.id) {
      const item = pending.get(message.id);
      if (item) {
        pending.delete(message.id);
        clearTimeout(item.timeout);
        message.error ? item.reject(new Error(JSON.stringify(message.error))) : item.resolve(message.result);
      }
    } else if (message.method === 'Runtime.exceptionThrown') exceptions.push(message.params);
    else if (message.method === 'Page.javascriptDialogOpening') dialogs.push(message.params);
    else if (message.method === 'Fetch.requestPaused') paused.push(message.params);
    else if (message.method === 'Browser.downloadProgress') downloads.push(message.params);
  };
  const call = (method, params = {}) => new Promise((resolve, reject) => {
    const next = ++id;
    const timeout = setTimeout(() => { pending.delete(next); reject(new Error(`CDP 超时：${method}`)); }, 20000);
    pending.set(next, {resolve, reject, timeout});
    socket.send(JSON.stringify({id: next, method, params}));
  });
  const evaluate = async expression => {
    const response = await call('Runtime.evaluate', {expression, returnByValue: true, awaitPromise: true});
    if (response.exceptionDetails) throw new Error(JSON.stringify(response.exceptionDetails));
    return response.result.value;
  };
  const click = selector => evaluate(`document.querySelector(${JSON.stringify(selector)}).click()`);
  const upload = async path => {
    const doc = await call('DOM.getDocument');
    const node = await call('DOM.querySelector', {nodeId: doc.root.nodeId, selector: '#file-input'});
    await call('DOM.setFileInputFiles', {nodeId: node.nodeId, files: [path]});
    await until(() => evaluate('!document.querySelector("#start-button").disabled'), '文件选择');
  };
  const complete = () => until(() => evaluate('!document.querySelector("#start-button").disabled && !document.querySelector("#precheck").hidden'), '处理结束');
  const shot = async name => {
    const metrics = await call('Page.getLayoutMetrics');
    const data = await call('Page.captureScreenshot', {format: 'png', captureBeyondViewport: true,
      clip: {x: 0, y: 0, width: metrics.cssContentSize.width, height: metrics.cssContentSize.height, scale: 1}});
    await writeFile(join(images, name), Buffer.from(data.data, 'base64'));
  };

  await call('Runtime.enable');
  await call('Page.enable');
  await call('Browser.setDownloadBehavior', {behavior: 'allow', downloadPath: artifacts, eventsEnabled: true});
  await call('Emulation.setDeviceMetricsOverride', {width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false});
  await until(() => evaluate('!!document.querySelector("#start-button")'), '页面初始化');
  assert(await evaluate('document.querySelector("#start-button").disabled && document.querySelector("#results").hidden'));
  await shot('converter-ready.png');

  await upload(fixtures.good);
  await call('Fetch.enable', {patterns: [{urlPattern: base + '/api/convert*', requestStage: 'Request'}]});
  await click('#start-button');
  await until(() => paused.length === 1, '暂停处理请求');
  assert(await evaluate('document.querySelector("#start-button").disabled && document.querySelector("#file-input").disabled'));
  await click('#start-button');
  assert.equal(paused.length, 1);
  await call('Fetch.continueRequest', {requestId: paused.shift().requestId});
  await call('Fetch.disable');
  await complete();
  assert(await evaluate('!document.querySelector("#results").hidden && document.querySelector("#check-badge").textContent === "检查通过"'));
  assert.deepEqual(await evaluate('[...document.querySelectorAll("#stats strong")].map(x => Number(x.textContent))'), [1, 3, 2, 5, 1, 1]);
  assert.equal(await evaluate('document.querySelector("#output-name").textContent'), '合成台账_处理结果.xlsx');
  await shot('converter-success.png');
  await click('#download-button');
  await until(() => downloads.some(item => item.state === 'completed'), '浏览器下载完成');
  const filename = (await readdir(artifacts)).find(name => name.endsWith('.xlsx'));
  assert(filename);
  const readback = JSON.parse(execFileSync(python, ['-m', 'tests.browser_fixture', '--verify', join(artifacts, filename)], {encoding: 'utf8'}));
  assert.equal(readback.download_readback, 'PASS');

  await upload(fixtures.bad);
  assert(await evaluate('document.querySelector("#results").hidden && document.querySelector("#precheck").hidden && !document.querySelector("#download-button").hasAttribute("href")'));
  await click('#start-button');
  await complete();
  assert.equal(await evaluate('document.querySelectorAll("#issue-rows tr").length'), 7);
  assert(await evaluate('[...document.querySelectorAll("#issue-rows tr")].every(row => /^[A-Z]+[0-9]+$/.test(row.cells[2].textContent))'));
  assert(await evaluate('document.querySelector("#results").hidden && !document.querySelector("#download-button").hasAttribute("href")'));
  await shot('converter-blocked.png');

  await upload(fixtures.hostile);
  await click('#start-button');
  await complete();
  assert(await evaluate('document.querySelector("#issue-rows").textContent.includes("<img src=x onerror=alert(7)>") && !document.querySelector("#issue-rows img")'));

  await upload(fixtures.invalid);
  await click('#start-button');
  await until(() => evaluate('!document.querySelector("#start-button").disabled && !document.querySelector("#status").hidden'), '损坏文件提示');
  assert(await evaluate('document.querySelector("#status").classList.contains("error") && document.querySelector("#results").hidden && !document.querySelector("#download-button").hasAttribute("href")'));

  const real = preserved.at(-1);
  await upload(real);
  await click('#start-button');
  await complete();
  assert.equal(await evaluate('document.querySelectorAll("#issue-rows tr").length'), 39);
  assert(await evaluate('[...document.querySelectorAll("#issue-rows tr")].every(row => row.cells[4].textContent === "安全阀")'));
  assert(await evaluate('document.querySelector("#results").hidden'));

  // 拖拽上传成功，并在窄屏检查溢出。
  await evaluate(`(async () => {
    const data = await fetch(${JSON.stringify(base + '/')}).then(r => r.arrayBuffer());
    const transfer = new DataTransfer();
    transfer.items.add(new File([data], 'invalid.txt'));
    document.querySelector('#dropzone').dispatchEvent(new DragEvent('drop', {dataTransfer: transfer, bubbles: true}));
  })()`);
  assert(await evaluate('document.querySelector("#start-button").disabled && document.querySelector("#results").hidden'));
  await upload(fixtures.good);
  await evaluate(`(() => {
    const transfer = new DataTransfer();
    transfer.items.add(document.querySelector('#file-input').files[0]);
    document.querySelector('#dropzone').dispatchEvent(new DragEvent('drop', {dataTransfer: transfer, bubbles: true}));
  })()`);
  await click('#start-button');
  await complete();
  await call('Emulation.setDeviceMetricsOverride', {width: 390, height: 844, deviceScaleFactor: 1, mobile: true});
  assert(await evaluate('document.documentElement.scrollWidth <= window.innerWidth'));
  await shot('converter-mobile.png');
  assert.deepEqual(exceptions, []);
  assert.deepEqual(dialogs, []);
  assert.deepEqual(await Promise.all(preserved.map(hash)), hashes);
  const evidence = {result: 'PASS', artifacts, readback, checks: [
    '真实 file input 上传', '处理中锁定及重复点击', '统计数值', '浏览器下载及 Excel 回读',
    '换文件立即清除旧结果', '跨表七项阻断及单元格定位', '特殊文本安全显示',
    '损坏输入清除下载', '真实安全阀台账39行未匹配阻断', '拖拽上传', '手机宽度无页面溢出',
    '原件与映射哈希不变', '无页面异常或弹窗'
  ]};
  await writeFile(join(artifacts, 'acceptance-evidence.json'), JSON.stringify(evidence, null, 2));
  console.log(JSON.stringify(evidence, null, 2));
  await call('Browser.close');
} finally {
  socket?.close();
  edge?.kill();
  server.stdin.end('\n');
  await until(() => server.exitCode !== null, '验收服务停止并清理临时输出');
}
