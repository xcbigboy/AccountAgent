'use strict';
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="app-token"]').content;
let templates = [], selected = 'example', excel = null, busy = false, aiReady = false, pendingJob = null;
function clearResult() { $('statusPanel').hidden = true; $('reviewPanel').hidden = true; pendingJob = null; $('acceptMapping').checked = false; $('step3').classList.remove('current'); $('failureDownloads').replaceChildren(); }
function setBusy(value, message = '正在检查数据并生成报告…') {
  busy = value;
  document.querySelectorAll('button,input').forEach(el => el.disabled = value);
  $('useAI').disabled = value || !aiReady;
  $('generate').disabled = value || !excel || !selected;
  $('confirmMapping').disabled = value || !$('acceptMapping').checked;
  if (value) { $('statusPanel').hidden = false; $('busyPanel').hidden = false; $('successPanel').hidden = true; $('errorPanel').hidden = true; $('warningsPanel').hidden = true; $('busyText').textContent = message; }
  else $('busyPanel').hidden = true;
}
function error(messages) {
  $('statusPanel').hidden = false; $('errorPanel').hidden = false; $('successPanel').hidden = true;
  $('failureDownloads').replaceChildren();
  $('errors').replaceChildren(...messages.map(message => { const li = document.createElement('li'); li.textContent = message; return li; }));
}
async function post(url, data) {
  const response = await fetch(url, {method:'POST', headers:{'Content-Type':'application/json','X-App-Token':token}, body:JSON.stringify(data)});
  const body = await response.json();
  if (!response.ok) throw new Error(body.error || '处理失败，请重试。');
  return body;
}
function fileData(file) {
  if (!file || file.size === 0 || file.size > 10 * 1024 * 1024) return Promise.reject(new Error('请选择非空文件，单个文件最多 10 MB。'));
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve({name:file.name, data:reader.result.split(',')[1]}); reader.onerror = () => reject(new Error('文件读取失败，请重新选择。')); reader.readAsDataURL(file); });
}
function chooseExcel(file) {
  if (busy || !file) return;
  clearResult();
  if (!/\.(xlsx|xlsm)$/i.test(file.name) || file.size > 10 * 1024 * 1024 || file.size === 0) { excel = null; $('generate').disabled = true; $('fileTitle').textContent = '把数据表拖到这里'; $('fileInfo').textContent = '支持 .xlsx / .xlsm，单个文件不超过 10 MB'; error(['请选择 .xlsx / .xlsm 文件，单个文件最多 10 MB。']); return; }
  excel = file;
  $('fileTitle').textContent = file.name;
  $('fileInfo').textContent = `${(file.size / 1024).toFixed(1)} KB · 已准备好`;
  $('chooseFile').textContent = '更换文件';
  $('readyText').textContent = '数据表已就绪，选择模板后点击生成。';
  $('step2').classList.add('current');
  $('generate').disabled = !selected;
}
function renderTemplates() {
  $('templateList').replaceChildren(...templates.map(item => {
    const label = document.createElement('label'); label.className = 'template-option';
    const input = document.createElement('input'); input.type = 'radio'; input.name = 'template'; input.value = item.id; input.checked = item.id === selected; input.disabled = busy;
    input.addEventListener('change', () => { selected = item.id; clearResult(); });
    const text = document.createElement('span'), title = document.createElement('strong'), description = document.createElement('small');
    title.textContent = item.name; description.textContent = item.description; text.append(title, description); label.append(input, text); return label;
  }));
}
async function loadTemplates() {
  const response = await fetch('/api/templates'); if (!response.ok) throw new Error('无法加载模板，请重新启动界面。');
  const data = await response.json(); templates = data.templates; aiReady = data.ai_ready;
  renderTemplates(); $('useAI').disabled = busy || !aiReady;
  $('aiDescription').textContent = aiReady ? '本机已配置 DeepSeek，使用 deepseek-v4-pro。' : '本机尚未设置 DEEPSEEK_API_KEY。设置后重启界面即可启用；普通填报无需密钥。';
}
function result(data) {
  $('statusPanel').hidden = false;
  if (data.status === 'needs_review') {
    pendingJob = data.id; $('statusPanel').hidden = true; $('reviewPanel').hidden = false;
    $('reviewRows').replaceChildren(...data.matches.map(item => { const tr = document.createElement('tr'); [item.label + ' · ' + ({end:'期末',begin:'期初'}[item.period] || item.period), item.label_cell, `${data.sheet}!${item.value_cell} · 表头 ${item.header_cell}`, item.reason].forEach(value => { const td = document.createElement('td'); td.textContent = value; tr.append(td); }); return tr; }));
    $('reviewPanel').scrollIntoView({behavior:'smooth', block:'start'}); return;
  }
  $('reviewPanel').hidden = true; pendingJob = null;
  if (data.status === 'failed') {
    error(data.errors.length ? data.errors : ['数据校验未通过。']);
    $('failureDownloads').replaceChildren(...(data.files || []).map(name => { const link = document.createElement('a'); link.href = `/download/${data.id}/${name}?token=${encodeURIComponent(token)}`; link.download = name; link.textContent = name === 'audit.json' ? '下载核对日志 JSON' : '下载核对日志 CSV'; link.className = 'download-link'; return link; }));
  }
  else {
    $('successPanel').hidden = false; $('errorPanel').hidden = true;
    $('successSummary').textContent = `${data.records} 项金额已填报，${data.checks} 项核对全部通过。请打开 Word 确认最终版面。`;
    const names = {'report.docx':'下载 Word 报告','audit.json':'核对日志 JSON','audit.csv':'核对日志 CSV','report-and-audit.zip':'打包下载'};
    $('downloads').replaceChildren(...data.files.map(name => { const link = document.createElement('a'); link.href = `/download/${data.id}/${name}?token=${encodeURIComponent(token)}`; link.download = name; link.textContent = names[name]; link.className = name === 'report.docx' ? 'primary' : 'download-link'; return link; }));
    $('step3').classList.add('current');
  }
  $('warningsPanel').hidden = !data.warnings?.length;
  $('warnings').replaceChildren(...(data.warnings || []).map(text => { const li = document.createElement('li'); li.textContent = text; return li; }));
  $('statusPanel').scrollIntoView({behavior:'smooth',block:'nearest'});
}
async function generate(demo = false) {
  if (busy || (!demo && !excel)) return;
  clearResult(); setBusy(true, $('useAI').checked && !demo ? 'DeepSeek 正在识别表格位置，完成后请确认映射…' : '正在核对金额并填写报告…');
  try { const data = await post('/api/run', {template:demo ? 'example' : selected, demo, use_ai:!demo && $('useAI').checked, excel:demo ? null : await fileData(excel)}); setBusy(false); result(data); }
  catch (e) { setBusy(false); error([e.message]); }
}
$('chooseFile').addEventListener('click', () => $('excelFile').click());
$('excelFile').addEventListener('change', e => chooseExcel(e.target.files[0]));
['dragenter','dragover'].forEach(type => $('dropzone').addEventListener(type, e => {e.preventDefault(); if (!busy) $('dropzone').classList.add('dragging');}));
['dragleave','drop'].forEach(type => $('dropzone').addEventListener(type, e => {e.preventDefault(); $('dropzone').classList.remove('dragging');}));
$('dropzone').addEventListener('drop', e => chooseExcel(e.dataTransfer.files[0]));
$('generate').addEventListener('click', () => generate());
$('useAI').addEventListener('change', clearResult);
$('demoButton').addEventListener('click', () => generate(true));
$('addTemplate').addEventListener('click', () => {$('templateError').textContent=''; $('templateDialog').showModal();});
$('closeDialog').addEventListener('click', () => $('templateDialog').close());
$('templateForm').addEventListener('submit', async e => {
  e.preventDefault(); if (busy) return; setBusy(true, '正在保存模板…'); $('templateError').textContent = '';
  try { const data = await post('/api/template', {name:$('templateName').value.trim(), word:await fileData($('wordFile').files[0]), config:await fileData($('configFile').files[0])}); selected = data.id; await loadTemplates(); $('templateDialog').close(); $('templateForm').reset(); clearResult(); setBusy(false); }
  catch (e) { setBusy(false); $('statusPanel').hidden = true; $('templateError').textContent = e.message; }
});
$('acceptMapping').addEventListener('change', () => $('confirmMapping').disabled = busy || !$('acceptMapping').checked);
$('cancelReview').addEventListener('click', clearResult);
$('confirmMapping').addEventListener('click', async () => {
  if (busy || !pendingJob || !$('acceptMapping').checked) return;
  setBusy(true, '正在验证映射并生成报告…');
  try { const data = await post('/api/approve', {id:pendingJob, accept:true}); setBusy(false); result(data); }
  catch (e) { setBusy(false); error([e.message]); }
});
loadTemplates().catch(e => error([e.message]));
