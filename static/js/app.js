/* static/js/app.js — 捷捷寶寶粥訂單轉換系統前端邏輯 */

// ── 工具函式 ─────────────────────────────────────────────────────────────

function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const colors = {
    success: '#22c55e', danger: '#ef4444', warning: '#f97316', info: '#3b82f6',
  };
  const id = 'toast-' + Date.now();
  const html = `
    <div id="${id}" class="toast align-items-center border-0 show"
         style="background:rgba(255,255,255,0.88);backdrop-filter:blur(12px);
                border-left:4px solid ${colors[type] || colors.info} !important;
                box-shadow:0 8px 24px rgba(0,0,0,.12);min-width:260px">
      <div class="d-flex align-items-center px-3 py-2 gap-2">
        <span style="color:${colors[type] || colors.info};font-size:1.1rem">
          ${type === 'success' ? '✓' : type === 'danger' ? '✕' : 'ℹ'}
        </span>
        <div class="me-auto small fw-semibold">${message}</div>
        <button type="button" class="btn-close btn-close-sm" onclick="this.closest('.toast').remove()"></button>
      </div>
    </div>`;
  container.insertAdjacentHTML('beforeend', html);
  setTimeout(() => document.getElementById(id)?.remove(), 3500);
}

function animateCount(elId, target) {
  const el = document.getElementById(elId);
  if (!el) return;
  const start    = parseInt(el.textContent) || 0;
  const duration = 600;
  const startTs  = performance.now();
  const reduced  = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (reduced) { el.textContent = target; return; }
  function step(ts) {
    const progress = Math.min((ts - startTs) / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3); // easeOutCubic
    el.textContent = Math.round(start + (target - start) * ease);
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

// ── 資料夾瀏覽器 Modal ────────────────────────────────────────────────────

let _fbTargetInputId = null;
let _fbParentPath    = null;

async function browseFolderPicker(targetId) {
  _fbTargetInputId = targetId;
  await _fbNavigate('');
  new bootstrap.Modal(document.getElementById('folderBrowserModal')).show();
}

async function _fbNavigate(path) {
  const url = path ? `/browse-path?path=${encodeURIComponent(path)}` : '/browse-path';
  try {
    const res  = await fetch(url);
    const data = await res.json();
    if (!res.ok) { showToast(data.error || '無法讀取路徑', 'danger'); return; }

    _fbParentPath = data.parent;
    document.getElementById('fbCurrentPath').textContent = data.current;
    document.getElementById('fbCurrentPath').title       = data.current;
    document.getElementById('fbBtnUp').disabled = !data.parent;

    const list = document.getElementById('fbDirList');
    if (data.dirs.length === 0) {
      list.innerHTML = '<div class="text-muted small py-2 px-1">（沒有子資料夾）</div>';
    } else {
      list.innerHTML = data.dirs.map(d => `
        <div class="file-item visible d-flex align-items-center gap-2 py-2 px-2 rounded fb-dir-item"
             style="cursor:pointer" data-path="${d.path.replace(/"/g, '&quot;')}">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="#f97316" style="flex-shrink:0">
            <path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V7z"/>
          </svg>
          <span class="small text-truncate">${d.name}</span>
        </div>`).join('');
      list.querySelectorAll('.fb-dir-item').forEach(el => {
        el.addEventListener('click', () => _fbNavigate(el.dataset.path));
      });
    }
  } catch (e) {
    showToast('讀取失敗', 'danger');
  }
}

function _fbSelect() {
  const pathEl = document.getElementById('fbCurrentPath');
  const input  = document.getElementById(_fbTargetInputId);
  if (input && pathEl.textContent) input.value = pathEl.textContent;
  bootstrap.Modal.getInstance(document.getElementById('folderBrowserModal'))?.hide();
}

// ── 掃描邏輯 ─────────────────────────────────────────────────────────────

let _scanResult = null;

async function doScan() {
  const folder = document.getElementById('folderInput')?.value?.trim();
  if (!folder) { showToast('請先輸入資料夾路徑', 'warning'); return; }

  // UI 狀態
  const btnScan = document.getElementById('btnScan');
  const scanIcon    = document.getElementById('scanIcon');
  const scanSpinner = document.getElementById('scanSpinner');
  btnScan.disabled  = true;
  scanIcon?.classList.add('d-none');
  scanSpinner?.classList.remove('d-none');

  document.getElementById('emptyState')?.classList.add('d-none');
  document.getElementById('fileListArea')?.classList.add('d-none');
  document.getElementById('resultArea')?.classList.add('d-none');
  document.getElementById('btnConvert') && (document.getElementById('btnConvert').disabled = true);

  try {
    const res  = await fetch('/scan', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({folder}),
    });
    const data = await res.json();
    if (!res.ok) { showToast(data.error || '掃描失敗', 'danger'); return; }

    // 若後端自動 fallback 到上層目錄，更新輸入框並提示使用者
    if (data.fallback_folder) {
      const input = document.getElementById('folderInput');
      if (input) input.value = data.fallback_folder;
      showToast(`原路徑不存在，已自動切換至：${data.fallback_folder}`, 'warning');
      try { localStorage.setItem('default_folder', data.fallback_folder); } catch(e) {}
    } else {
      try { localStorage.setItem('default_folder', folder); } catch(e) {}
    }

    _scanResult = data;
    renderFileList(data.files);
  } catch (e) {
    showToast('掃描發生錯誤：' + e.message, 'danger');
  } finally {
    btnScan.disabled = false;
    scanIcon?.classList.remove('d-none');
    scanSpinner?.classList.add('d-none');
  }
}

function renderFileList(files) {
  const listEl  = document.getElementById('fileList');
  const countBadge = document.getElementById('fileCountBadge');
  const listArea   = document.getElementById('fileListArea');
  const unknownPanel = document.getElementById('unknownPanel');
  const btnConvert   = document.getElementById('btnConvert');

  listArea.classList.remove('d-none');
  countBadge.textContent = `${files.length} 個檔案`;

  const badgeMap = {
    shopee:     ['badge-shopee',     '蝦皮'],
    a1baby:     ['badge-a1baby',     'A1婦幼展'],
    leage:      ['badge-leage',      '樂齡網'],
    a1leage:    ['badge-a1leage',    'A1樂齡官網'],
    jjofficial: ['badge-jjofficial', '捷捷官網'],
    yodee:      ['badge-yodee',      '優迪通路'],
    kadomo:     ['badge-kadomo',     '卡多摩'],
    licai:      ['badge-licai',      '麗兒采家'],
    xuantu:     ['badge-xuantu',     '炫兔團購'],
    tuanma:     ['badge-tuanma',     '其他團媽'],
    chocho:     ['badge-chocho',     'CHOCHO通路'],
  };

  let pendingCount = 0;

  if (files.length === 0) {
    listEl.innerHTML = `<div class="text-muted small py-3 px-1 text-center">
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" stroke-width="1.5" class="mb-2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/></svg>
      <div>資料夾中沒有可處理的 .xlsx / .pdf / .csv 檔案</div>
      <div class="text-secondary" style="font-size:.75rem;margin-top:.25rem">請確認檔案格式是否為 .xlsx 或 .csv（舊版 .xls 不支援）</div>
    </div>`;
    unknownPanel.classList.add('d-none');
    btnConvert && (btnConvert.disabled = true);
    return;
  }

  listEl.innerHTML = files.map((f, i) => {
    const isArchived = f.status === 'archived';
    const src        = f.source_type;
    if (!isArchived) pendingCount++;

    const [badgeClass, badgeLabel] = badgeMap[src] || ['badge-unknown', '未識別'];
    const statusHtml = isArchived
      ? '<span class="badge bg-secondary bg-opacity-20 text-secondary-emphasis rounded-pill" style="font-size:.72rem">已處理</span>'
      : `<span class="badge-source ${badgeClass}" style="font-size:.72rem">${f.source_label || badgeLabel}</span>`;

    const checkbox = isArchived
      ? `<input type="checkbox" class="file-checkbox form-check-input" disabled
              style="flex-shrink:0;width:16px;height:16px;cursor:not-allowed;opacity:.4">`
      : `<input type="checkbox" class="file-checkbox form-check-input" checked
              data-path="${f.path}" style="flex-shrink:0;width:16px;height:16px;cursor:pointer"
              onclick="event.stopPropagation()" onchange="onFileCheckChange()">`;

    return `
      <div class="file-item d-flex align-items-center gap-2 py-2 px-3 rounded${isArchived ? ' archived' : ''}"
           style="animation-delay:${i * 40}ms"
           ${isArchived ? '' : 'onclick="this.querySelector(\'.file-checkbox\').click()"'}>
        ${checkbox}
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
             stroke="${isArchived ? '#94a3b8' : '#f97316'}" stroke-width="2" style="flex-shrink:0">
          <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/>
        </svg>
        <span class="small text-truncate flex-grow-1" style="max-width:280px" title="${f.name}">${f.name}</span>
        ${statusHtml}
      </div>`;
  }).join('');

  // 觸發 slide-in animation（stagger）
  requestAnimationFrame(() => {
    listEl.querySelectorAll('.file-item').forEach((el, i) => {
      setTimeout(() => el.classList.add('visible'), i * 40);
    });
  });

  unknownPanel.classList.toggle('d-none', pendingCount === 0);
  // 重設為自動偵測
  const forcedSelect = document.getElementById('forcedSource');
  if (forcedSelect) forcedSelect.value = '';
  _updateConvertButton();

  if (pendingCount === 0 && files.length > 0) {
    showToast('所有檔案都已處理過，無需重複轉換', 'info');
  }
}

function _updateConvertButton() {
  const checked = document.querySelectorAll('.file-checkbox:not(:disabled):checked').length;
  const btn = document.getElementById('btnConvert');
  if (btn) btn.disabled = checked === 0;
}

function toggleAllFiles(check) {
  document.querySelectorAll('.file-checkbox:not(:disabled)').forEach(cb => cb.checked = check);
  _updateConvertButton();
}

function onFileCheckChange() {
  _updateConvertButton();
}

// ── 轉換邏輯 ─────────────────────────────────────────────────────────────

let _eventSource = null;

async function doConvert() {
  const folder   = document.getElementById('folderInput')?.value?.trim();
  const operator = document.getElementById('operatorInput')?.value?.trim();
  const forced   = document.getElementById('forcedSource')?.value || null;

  // 收集勾選的檔案路徑
  const selectedPaths = Array.from(
    document.querySelectorAll('.file-checkbox:not(:disabled):checked')
  ).map(cb => cb.dataset.path);

  if (!folder) { showToast('請先輸入資料夾路徑', 'warning'); return; }

  const btnConvert   = document.getElementById('btnConvert');
  const progressArea = document.getElementById('progressArea');
  const progressBar  = document.getElementById('progressBar');
  const progressLabel = document.getElementById('progressLabel');
  const progressPct  = document.getElementById('progressPct');
  const progressLog  = document.getElementById('progressLog');

  btnConvert.disabled = true;
  progressArea.classList.remove('d-none');
  progressBar.style.width = '5%';
  progressLabel.textContent = '準備中...';
  progressPct.textContent   = '0%';
  progressLog.innerHTML = '';
  document.getElementById('resultArea')?.classList.add('d-none');

  // 啟動 SSE 監聽
  if (_eventSource) { _eventSource.close(); _eventSource = null; }
  _eventSource = new EventSource('/convert/status');

  _eventSource.onmessage = (e) => {
    const msg = JSON.parse(e.data);
    handleProgress(msg, progressBar, progressLabel, progressPct, progressLog);
  };
  _eventSource.onerror = () => {
    _eventSource.close(); _eventSource = null;
    btnConvert.disabled = false;
    showToast('SSE 連線中斷', 'danger');
  };

  // 啟動轉換
  try {
    const res = await fetch('/convert', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({folder, operator, source_type: forced,
                            selected_files: selectedPaths}),
    });
    const data = await res.json();
    if (!res.ok || !data.ok) {
      showToast(data.error || '轉換請求失敗', 'danger');
      btnConvert.disabled = false;
      _eventSource?.close();
    }
  } catch (e) {
    showToast('發生錯誤：' + e.message, 'danger');
    btnConvert.disabled = false;
    _eventSource?.close();
  }
}

function handleProgress(msg, bar, label, pct, log) {
  if (msg.type === 'heartbeat') return;

  if (msg.type === 'progress') {
    const p = Math.min((parseInt(bar.style.width) || 5) + 20, 85);
    bar.style.width = p + '%';
    pct.textContent = p + '%';
    label.textContent = msg.message || '處理中...';
    const line = document.createElement('div');
    line.textContent = '› ' + msg.message;
    log.appendChild(line);
    log.scrollTop = log.scrollHeight;
    return;
  }

  if (msg.type === 'done') {
    bar.style.width  = '100%';
    pct.textContent  = '100%';
    label.textContent = msg.fail > 0 ? '轉換完成（有失敗項目）' : '轉換完成！';
    _eventSource?.close(); _eventSource = null;
    document.getElementById('btnConvert').disabled = false;

    if (msg.fail > 0) {
      _fdShowDecision(msg);  // 有失敗 → 先問使用者
    } else {
      showResultCard(msg);
      showToast('轉換成功', 'success');
    }
    return;
  }

  if (msg.type === 'error') {
    bar.style.width   = '100%';
    bar.style.background = '#ef4444';
    label.textContent = '發生錯誤';
    _eventSource?.close(); _eventSource = null;
    document.getElementById('btnConvert').disabled = false;
    showToast(msg.message || '轉換失敗', 'danger');
  }
}

function showResultCard(msg) {
  const resultArea = document.getElementById('resultArea');
  resultArea.classList.remove('d-none');

  animateCount('resSuccess', msg.success || 0);
  animateCount('resFail',    msg.fail    || 0);
  animateCount('resManual',  msg.manual  || 0);

  // 觸發 spring 動效
  resultArea.querySelectorAll('.result-card').forEach((el, i) => {
    el.style.animationDelay = (i * 80) + 'ms';
    el.classList.remove('animate');
    void el.offsetWidth; // reflow
    el.classList.add('animate');
  });

  // 操作按鈕
  const actionsEl = document.getElementById('resultActions');
  if (actionsEl && msg.log_id) {
    actionsEl.innerHTML = `
      ${msg.fail > 0 ? `<a href="/errors/${msg.log_id}" class="btn btn-sm btn-glass-secondary">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" class="me-1">
          <path d="M12 2a10 10 0 100 20A10 10 0 0012 2zm1 14h-2v-2h2v2zm0-4h-2V7h2v5z"/>
        </svg>查看錯誤</a>` : ''}
      <a href="/download/${msg.log_id}" class="btn btn-sm btn-glass-primary">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" class="me-1">
          <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
        </svg>下載結果</a>`;
  }
}

// ── 轉換失敗決策 Modal ────────────────────────────────────────────────────

let _fdMsg = null;

async function _fdShowDecision(msg) {
  _fdMsg = msg;

  // 載入錯誤列表
  let errors = [];
  try {
    const res = await fetch(`/errors/${msg.log_id}/data`);
    errors = await res.json();
  } catch(e) {}

  // 摘要
  document.getElementById('fdSummary').innerHTML =
    `成功 <strong>${msg.success}</strong> 筆，失敗 <strong>${msg.fail}</strong> 筆。` +
    `<br>失敗的品項不會出現在輸出檔案中。`;

  // 錯誤列表
  const listEl = document.getElementById('fdErrorList');
  if (errors.length) {
    listEl.innerHTML = errors.map(e => {
      const fileName = e.source_file ? e.source_file.replace(/^.*[\\/]/, '') : '';
      const orig = e.original_value ? `（${e.original_value}）` : '';
      return `
      <div class="py-2 border-bottom" style="border-color:rgba(0,0,0,.06)!important">
        <div class="d-flex gap-2 align-items-start">
          <span class="badge bg-light text-dark border" style="flex-shrink:0;font-size:.75rem">${fileName || '未知檔案'}</span>
          <span class="text-muted" style="flex-shrink:0">第${e.row_number}列</span>
        </div>
        <div class="mt-1" style="color:#c2410c;font-size:.9rem" title="${e.reason}">
          <span class="fw-semibold">${orig}</span>
          <span>${e.reason}</span>
        </div>
      </div>`;
    }).join('');
  } else {
    listEl.innerHTML = '';
  }

  document.getElementById('fdAliasArea').classList.add('d-none');
  new bootstrap.Modal(document.getElementById('failDecisionModal')).show();
}

function _fdAccept() {
  bootstrap.Modal.getInstance(document.getElementById('failDecisionModal'))?.hide();
  if (_fdMsg) showResultCard(_fdMsg);
  showToast('已接受部分結果', 'warning');
}

async function _fdRollback() {
  if (!_fdMsg?.log_id) return;
  try {
    await fetch(`/convert/rollback/${_fdMsg.log_id}`, {method: 'POST'});
    bootstrap.Modal.getInstance(document.getElementById('failDecisionModal'))?.hide();
    // 重置 UI
    document.getElementById('progressArea')?.classList.add('d-none');
    document.getElementById('resultArea')?.classList.add('d-none');
    showToast('已放棄此次轉換，檔案已還原', 'info');
  } catch(e) {
    showToast('放棄失敗：' + e.message, 'danger');
  }
}

function _fdShowAlias() {
  const aliasArea = document.getElementById('fdAliasArea');
  const aliasList = document.getElementById('fdAliasList');
  aliasArea.classList.remove('d-none');

  // 從錯誤列表取出所有缺少的品號（去重）
  const items = document.querySelectorAll('#fdErrorList .d-flex');
  const seen  = new Set();
  const entries = [];
  items.forEach(el => {
    const reason = el.querySelector('[title]')?.getAttribute('title') || '';
    const m     = reason.match(/品號 '([^']+)'/);
    const nameM = reason.match(/（(.+)）在品號/);
    if (m && !seen.has(m[1])) {
      seen.add(m[1]);
      entries.push({ sku: m[1], name: nameM ? nameM[1] : '' });
    }
  });

  if (!entries.length) {
    aliasList.innerHTML = '<div class="text-muted small">無法解析錯誤品號</div>';
    return;
  }

  aliasList.innerHTML = entries.map((e, i) => `
    <div class="mb-3 fd-alias-entry" data-sku="${e.sku}" data-target-no="" data-idx="${i}">
      <div class="small fw-semibold mb-1" style="color:#c2410c">${e.sku}</div>
      ${e.name ? `<div class="small text-muted mb-1" style="font-size:.78rem">${e.name}</div>` : ''}
      <div class="position-relative">
        <input class="glass-input form-control form-control-sm fd-alias-search"
               data-idx="${i}" placeholder="輸入關鍵字搜尋對應品項…" autocomplete="off">
        <div class="fd-dropdown position-absolute w-100 d-none"
             style="top:100%;left:0;z-index:9999;background:rgba(255,255,255,0.97);
                    border:1px solid rgba(0,0,0,.1);border-radius:6px;
                    box-shadow:0 4px 16px rgba(0,0,0,.12);max-height:180px;overflow-y:auto"></div>
      </div>
      <div class="fd-selected small mt-1 d-none" style="color:#15803d"></div>
    </div>`).join('');

  // 綁定每個搜尋框的 input 事件
  aliasList.querySelectorAll('.fd-alias-search').forEach(input => {
    let timer;
    input.addEventListener('input', () => {
      clearTimeout(timer);
      timer = setTimeout(() => _fdSearchProduct(input), 300);
    });
  });
}

async function _fdSearchProduct(input) {
  const idx      = input.dataset.idx;
  const entry    = document.querySelector(`.fd-alias-entry[data-idx="${idx}"]`);
  const dropdown = entry.querySelector('.fd-dropdown');
  const q        = input.value.trim();
  if (q.length < 1) { dropdown.classList.add('d-none'); return; }

  try {
    const source = _fdMsg?.source_type || '';
    const res  = await fetch(`/alias/search?q=${encodeURIComponent(q)}&source_type=${source}`);
    const data = await res.json();
    if (!data.length) {
      dropdown.innerHTML = '<div class="px-3 py-2 small text-muted">找不到符合的品項</div>';
      dropdown.classList.remove('d-none');
      return;
    }
    dropdown.innerHTML = data.map(p => `
      <div class="fd-dropdown-item px-3 py-2 small"
           style="cursor:pointer;border-bottom:1px solid rgba(0,0,0,.05)"
           data-no="${p['品號']}" data-name="${p['品名']}">
        <span style="color:#c2410c;font-weight:600">${p['品號']}</span>
        <span class="ms-2 text-muted">${p['品名']}</span>
        <span class="ms-1" style="color:#6b7280;font-size:.78rem">$${p['商品結帳價']}</span>
      </div>`).join('');
    dropdown.classList.remove('d-none');

    dropdown.querySelectorAll('.fd-dropdown-item').forEach(item => {
      item.addEventListener('mouseenter', () => item.style.background = 'rgba(249,115,22,.08)');
      item.addEventListener('mouseleave', () => item.style.background = '');
      item.addEventListener('click', () => {
        entry.dataset.targetNo = item.dataset.no;
        input.value = `${item.dataset.no} — ${item.dataset.name}`;
        const sel = entry.querySelector('.fd-selected');
        sel.textContent = `✓ 已選擇：${item.dataset.no} ${item.dataset.name}`;
        sel.classList.remove('d-none');
        dropdown.classList.add('d-none');
      });
    });
  } catch(e) {
    dropdown.classList.add('d-none');
  }
}

async function _fdSubmitAliases() {
  const entries = document.querySelectorAll('.fd-alias-entry');
  let addedCount = 0;
  const source = _fdMsg?.source_type || '';

  for (const entry of entries) {
    const sku    = entry.dataset.sku;
    const target = entry.dataset.targetNo;
    if (!target) continue;
    try {
      const res = await fetch('/alias', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({alias_name: sku, target_no: target, source_type: source}),
      });
      const data = await res.json();
      if (data.ok) addedCount++;
    } catch(e) {}
  }
  await _fdRollback();
  if (addedCount > 0) {
    showToast(`已新增 ${addedCount} 筆品號別名，請重新選取檔案轉換`, 'success');
  }
}

// ── 報表頁 count-up（頁面載入後觸發）────────────────────────────────────

function initCountUpObserver() {
  const els = document.querySelectorAll('.count-up[data-target]');
  if (!els.length) return;
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const el = entry.target;
        animateCount(el.id, parseInt(el.dataset.target) || 0);
        observer.unobserve(el);
      }
    });
  }, {threshold: 0.3});
  els.forEach(el => observer.observe(el));
}

// ── 初始化 ───────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  initCountUpObserver();

  // 從 localStorage 還原上次的資料夾路徑與執行者
  try {
    const savedFolder = localStorage.getItem('default_folder');
    if (savedFolder) {
      const folderInput = document.getElementById('folderInput');
      if (folderInput) folderInput.value = savedFolder;
    }
  } catch(e) {}

  // Enter 鍵觸發掃描
  document.getElementById('folderInput')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') doScan();
  });
});
