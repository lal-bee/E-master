'use strict';

const $ = selector => document.querySelector(selector);
const input = $('#file-input');
const dropzone = $('#dropzone');
let selectedFile = null;
let busy = false;

function status(message, error = false) {
  $('#status').textContent = message;
  $('#status').classList.toggle('error', error);
  $('#status').hidden = !message;
}

function clearResult() {
  $('#precheck').hidden = true;
  $('#results').hidden = true;
  $('#download-button').removeAttribute('href');
  $('#issue-rows').replaceChildren();
  $('#stats').replaceChildren();
}

function selectFile(file) {
  if (busy) return;
  selectedFile = null;
  clearResult();
  status('');
  dropzone.classList.remove('selected');
  $('#file-label').textContent = '将 Excel 文件拖拽至此，或点击选择';
  $('#file-detail').textContent = '支持 .xlsx 格式，最大 32 MB';
  $('#start-button').disabled = true;
  $('#action-hint').textContent = '选择文件后即可开始';
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) {
    status('仅支持 .xlsx 文件，请重新选择。', true);
    input.value = '';
    return;
  }
  if (!file.size || file.size > 32 * 1024 * 1024) {
    status('请选择非空且在 32 MB 以内的 Excel 文件。', true);
    input.value = '';
    return;
  }
  selectedFile = file;
  dropzone.classList.add('selected');
  $('#file-label').textContent = file.name;
  $('#file-detail').textContent = `${(file.size / 1024).toFixed(1)} KB · 点击可重新选择`;
  $('#start-button').disabled = false;
  $('#action-hint').textContent = '准备就绪，将先检查所有业务工作表';
}

input.addEventListener('change', () => selectFile(input.files[0]));
dropzone.addEventListener('keydown', event => {
  if (!busy && (event.key === 'Enter' || event.key === ' ')) {
    event.preventDefault();
    input.click();
  }
});
for (const eventName of ['dragenter', 'dragover']) dropzone.addEventListener(eventName, event => {
  event.preventDefault();
  if (!busy) dropzone.classList.add('dragging');
});
for (const eventName of ['dragleave', 'drop']) dropzone.addEventListener(eventName, event => {
  event.preventDefault();
  dropzone.classList.remove('dragging');
});
dropzone.addEventListener('drop', event => {
  if (busy) return;
  if (event.dataTransfer.files.length !== 1) {
    selectFile(null);
    status('每次请选择一个 .xlsx 文件。', true);
    return;
  }
  input.value = '';
  selectFile(event.dataTransfer.files[0]);
});
window.addEventListener('dragover', event => event.preventDefault());
window.addEventListener('drop', event => event.preventDefault());

function showResult(result) {
  $('#precheck').hidden = false;
  $('#precheck').classList.toggle('blocked', !result.success);
  $('#check-mark').textContent = result.success ? '✓' : '!';
  $('#check-badge').textContent = result.success ? '检查通过' : '检查未通过';
  $('#check-description').textContent = result.success
    ? `已检查 ${result.sheets.length} 个工作表、${result.rows} 行数据，没有阻断问题。`
    : `处理已终止，请修正后重新执行。完整检查发现 ${result.issues.length} 个阻断问题。`;
  $('#issue-container').hidden = result.success;
  if (!result.success) {
    const fragment = document.createDocumentFragment();
    for (const issue of result.issues) {
      const row = document.createElement('tr');
      for (const value of [issue.sheet, issue.row, issue.cell, issue.field, issue.value || '（空）', issue.message]) {
        const cell = document.createElement('td');
        cell.textContent = String(value);
        row.append(cell);
      }
      fragment.append(row);
    }
    $('#issue-rows').append(fragment);
    $('#action-hint').textContent = '修正原文件或补充映射表后，重新选择文件并处理';
    return;
  }
  $('#results').hidden = false;
  $('#processed-sheets').textContent = result.sheets.join('、');
  const stats = [
    ['处理工作表', result.sheets.length], ['数据行数', result.rows],
    ['小类编码填充数', result.subcategory_filled], ['日期转换数', result.dates_converted],
    ['日期置空数', result.dates_cleared], ['设备代码格式化数', result.codes_formatted]
  ];
  for (const [label, value] of stats) {
    const item = document.createElement('div');
    item.className = 'stat';
    const number = document.createElement('strong');
    number.textContent = value;
    const caption = document.createElement('span');
    caption.textContent = label;
    item.append(number, caption);
    $('#stats').append(item);
  }
  $('#output-name').textContent = result.filename;
  $('#download-button').href = result.download_url;
  $('#download-button').download = result.filename;
  $('#action-hint').textContent = '处理完成，可以下载新的 Excel 文件';
}

$('#start-button').addEventListener('click', async () => {
  if (!selectedFile || busy) return;
  busy = true;
  input.disabled = true;
  $('#start-button').disabled = true;
  $('#start-label').textContent = '正在处理…';
  dropzone.setAttribute('aria-disabled', 'true');
  clearResult();
  status('正在完整预检查。检查通过后自动转换，请稍候…');
  try {
    const response = await fetch(`/api/convert?filename=${encodeURIComponent(selectedFile.name)}`, {
      method: 'POST', headers: {'Content-Type': 'application/octet-stream'}, body: selectedFile
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || '处理失败，请重新执行。');
    status('');
    showResult(result);
  } catch (error) {
    status(error.message || '连接失败，请确认本地工具仍在运行。', true);
    $('#action-hint').textContent = '处理已终止，请修正后重新执行';
  } finally {
    busy = false;
    input.disabled = false;
    $('#start-button').disabled = !selectedFile;
    $('#start-label').textContent = '开始处理';
    dropzone.removeAttribute('aria-disabled');
  }
});
