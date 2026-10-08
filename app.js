// ============================================================
// SMART FBT — app.js
// Automated Benefit Tax Reconciliation & Upload Tool
// Phase 1: SAP vs PeopleSoft Reconciliation Engine
// ============================================================

// ---- STATE ----
let currentDataType = 'MBA';
let rawPsRows  = [];   // Raw parsed rows from PeopleSoft file
let rawSapRows = [];   // Raw parsed rows from SAP file
let psFileState  = null; // { file, cleanRows, profile, fmt }
let sapFileState = null; // { file, cleanRows, profile, fmt }
let reconData  = [];   // Final reconciliation result array
let filteredData = []; // Data after filter/search
let currentFilter = 'all';
let searchQuery   = '';
let currentPage   = 1;
let pageSize      = 10;
let sortCol = 'none';
let sortAsc = true;

// ---- MAPS ----
let psMap  = new Map(); // key: HMS_INVOICE_NBR -> { emplid, totalReimb, items[] }
let sapMap = new Map(); // key: REFERENCE      -> { persNo, totalGL, lines[] }

// ============================================================
// TAB / PANEL NAVIGATION
// ============================================================
function switchTab(tab) {
  if (tab !== 'upload' && reconData.length === 0) {
    showToast('⚠️ Silakan upload file dan jalankan rekonsiliasi terlebih dahulu.');
    return;
  }

  document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
  document.querySelectorAll('.tab-nav').forEach(b => {
    b.classList.remove('border-blue-400','text-blue-300');
    b.classList.add('border-transparent','text-navy-400');
  });

  document.getElementById('panel-' + tab).classList.remove('hidden');
  const btn = document.getElementById('tab-' + tab);
  if (btn) {
    btn.classList.remove('border-transparent','text-navy-400');
    btn.classList.add('border-blue-400','text-blue-300');
  }
  updateStepper(tab);
}

function handleStepperClick(tab, stepNum) {
  if (stepNum > 1 && reconData.length === 0) {
    showToast('⚠️ Silakan upload file dan jalankan rekonsiliasi terlebih dahulu.');
    return;
  }
  switchTab(tab);
}

// ============================================================
// PHASE STEPPER
// ============================================================
function updateStepper(tab) {
  // Map tab → which step number is currently "active"
  const stepMap = { 'upload': 1, 'compare': reconData.length > 0 ? 3 : 2, 'summary': 4 };
  const currentStep = stepMap[tab] || 1;

  const checkSVG = `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M5 13l4 4L19 7"/>
  </svg>`;

  for (let i = 1; i <= 4; i++) {
    const item   = document.getElementById('stepper-' + i);
    const circle = document.getElementById('step-circle-' + i);
    if (!item || !circle) continue;

    if (reconData.length === 0) {
      if (i === 1) {
        item.className = 'stepper-item active';
        circle.innerHTML = '<span>1</span>';
      } else {
        item.className = 'stepper-item disabled';
        circle.innerHTML = '<span>' + i + '</span>';
      }
    } else {
      // Data is present
      if (i < currentStep) {
        // COMPLETED — green circle + checkmark
        item.className   = 'stepper-item completed';
        circle.innerHTML = checkSVG;
      } else if (i === currentStep) {
        // ACTIVE — navy blue circle with number
        item.className   = 'stepper-item active';
        circle.innerHTML = '<span>' + i + '</span>';
      } else {
        // UNLOCKED FUTURE STEP
        item.className   = 'stepper-item';
        circle.innerHTML = '<span>' + i + '</span>';
      }
    }
  }

  // Update subtitles dynamically
  const sub1 = document.getElementById('step-sub-1');
  if (sub1) {
    if (reconData.length > 0) sub1.textContent = 'Uploaded & Ready';
    else sub1.textContent = 'SAP + PeopleSoft';
  }
  if (tab === 'compare' && reconData.length > 0) {
    const sub3 = document.getElementById('step-sub-3');
    if (sub3) sub3.textContent = reconData.length + ' records';
  }
  if (tab === 'summary') {
    const sub4 = document.getElementById('step-sub-4');
    if (sub4) sub4.textContent = '2 reports ready';
  }
}

// ============================================================
// COLLAPSIBLE HOW-TO PANEL
// ============================================================
function toggleHowTo() {
  const panel = document.getElementById('howto-panel');
  const btn   = document.getElementById('howto-btn');
  const wrapper = document.getElementById('howto-wrapper');
  const isOpen = panel.classList.contains('open');
  if (isOpen) {
    panel.classList.remove('open');
    btn.classList.remove('howto-open');
    wrapper.classList.remove('border-blue-400');
  } else {
    panel.classList.add('open');
    btn.classList.add('howto-open');
    wrapper.classList.add('border-blue-400');
  }
}

// ============================================================
// DOWNLOAD SAMPLE TEMPLATE
// ============================================================
function downloadSampleTemplate(type) {
  if (typeof XLSX === 'undefined') {
    alert('SheetJS library belum dimuat. Pastikan koneksi internet aktif.');
    return;
  }

  const wb = XLSX.utils.book_new();
  let headers, sheetName, fileName;

  if (type === 'sap') {
    headers = ['HMS_FBT_SEQNBR','HMS_COMPANY_CD','HMS_DOC_NBR','CURRENCY_CD','DOC_TYPE',
               'CREATED_BY_USER','HMS_HEADER_TEXT','REFERENCE','DOCUMENT_DT','POSTING_DT',
               'HMS_SUPPL_ACT','HMS_SUP_POST_KEY','HMS_SUPL_AMOUNT','HMS_ASSIGNMENT',
               'HMS_PERS_NO','HMS_PAYMENT_BLOCK','HMS_CLEAR_DT','HMS_CLEAR_DOC',
               'HMSA_ACCT_NO','HMS_GL_POST_KEY','HMS_COST_CENTER','HMS_GL_AMOUNT','TAX_CODE','TEXT254',
               'HMS_TAXABLE','HMS_FBT_TYPE','PRCSINSTANCE','CREATEDBY','CREATEDDTTM','LASTUPDBY','LASTUPDDTTM'];
    sheetName = 'SAP_Template';
    fileName  = 'HDC_FBT_SAP_TEMPLATE.xlsx';
  } else if (type === 'ps_mba') {
    headers = ['EMPLID','RECEIPT_DT','SEQNO','HMS_MEDICAL_CD','ORG_RECEIPT_DT','HMS_MED_ENTLT_PRD',
               'HMS_MBA_CLAIM_CAT','DEPENDENT_BENEF','ACCTG_ENTRY_FLG','HMS_MED_GRANT_LTR',
               'HMS_Memo_LTR_NO','HMS_RECEIPT_AMT','HMS_Reimburse_AMT','HMS_FORWARD_STATUS',
               'HMS_INVOICE_NBR','HMS_ACC_DATE'];
    sheetName = 'PeopleSoft_MBA';
    fileName  = 'HDC_FBT_PS_MBA_TEMPLATE.xlsx';
  } else if (type === 'ps_fsa') {
    headers = ['HMS_FSA_TRANS_ID','EMPLID','HMS_FLX_YEAR','HMS_FSA_STAT','DESCRLONG',
               'SEQNBR','RECEIPT_DT','HMS_FSA_ITEM','HMS_FLX_UNIT','LASTUPDDTTM'];
    sheetName = 'PeopleSoft_FSA';
    fileName  = 'HDC_FBT_PS_FSA_TEMPLATE.xlsx';
  } else if (type === 'ps_cb') {
    headers = ['ECS_RMB_TRANSID','EMPLID','ECS_GROUP_ID','EXPENSE_TYPE','PURPOSE','ECS_RMB_STAT',
               'ECS_RMB_SAP_REF','ECS_RMB_P1_DTTM','SEQNBR','RECEIPT_DT','ECS_RMB_EXP_ID',
               'EXPENSE_ITEM','ECS_AMT_RMB','COMMENTS'];
    sheetName = 'PeopleSoft_CB';
    fileName  = 'HDC_FBT_PS_CB_TEMPLATE.xlsx';
  } else {
    return; // Fallback
  }

  const data = [headers]; // Only headers, no data rows
  const ws = XLSX.utils.aoa_to_sheet(data);

  // Style header row (bold + navy background)
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (!ws[addr]) continue;
    ws[addr].s = {
      fill: { patternType: 'solid', fgColor: { rgb: type === 'sap' ? '0F2247' : '1e3f7a' } },
      font: { bold: true, color: { rgb: 'FFFFFF' } }
    };
  }

  // Column widths
  ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 4, 14) }));

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
  showToast('✓ Template downloaded: ' + fileName);
}

// ============================================================
// SMART MULTI-VARIANT FILE DETECTION & NORMALIZATION
// ============================================================
function validateFileName(fileName, type) {
  // Expected: HDC_FBT_SAP_[TYPE]_[PERIOD].ext  or  HDC_FBT_PS_[TYPE]_[PERIOD].ext
  const prefix = type === 'sap' ? /^HDC_FBT_SAP_/i : /^HDC_FBT_PS_/i;
  const ext = fileName.match(/\.([a-z]+)$/i);
  const validExt = ext && ['xlsx','xls','csv'].includes(ext[1].toLowerCase());
  const validName = prefix.test(fileName);
  return { validName, validExt, ok: validName && validExt };
}

function normalizeHeaderName(raw) {
  if (!raw) return '';
  let str = String(raw).trim().toUpperCase();
  // Strip parenthetical notes, e.g. "HMS_MEDICAL_CD (Cek TER...)" -> "HMS_MEDICAL_CD"
  const parenIdx = str.indexOf('(');
  if (parenIdx > -1) {
    const mainPart = str.substring(0, parenIdx).trim();
    if (mainPart) str = mainPart;
  }
  return str.replace(/\s+/g, '_');
}

function findHeaderRowIndex(rows) {
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const row = rows[i];
    if (!Array.isArray(row)) continue;
    const normalized = row.map(normalizeHeaderName);
    const isPS = normalized.some(h => 
      h === 'EMPLID' || h === 'EMPL_ID' || h === 'HMS_INVOICE_NBR' || h === 'ECS_RMB_SAP_REF' || h === 'ECS_RMB_TRANSID' || h === 'HMS_FSA_TRANS_ID'
    );
    const isSAP = normalized.some(h => 
      h === 'HMS_FBT_SEQNBR' || h === 'REFERENCE' || h === 'REFERENCE_NUMBER' || h === 'HMS_PERS_NO' || h === 'HMS_GL_AMOUNT' || h === 'HMS_DOC_NBR'
    );
    if (isPS || isSAP) return i;
  }
  return 0;
}

function detectFileProfile(rows) {
  if (!rows || rows.length < 1) {
    return { valid: false, system: 'unknown', reason: 'File kosong atau tidak memiliki data.' };
  }

  const hdrIdx = findHeaderRowIndex(rows);
  const headerRow = rows[hdrIdx] || [];
  const normalized = headerRow.map(normalizeHeaderName);

  // Check PeopleSoft signatures
  const hasEmplid   = normalized.some(h => h === 'EMPLID' || h === 'EMPL_ID');
  const hasPsMba    = normalized.some(h => h === 'HMS_INVOICE_NBR');
  const hasPsCb     = normalized.some(h => h === 'ECS_RMB_SAP_REF' || h === 'ECS_RMB_TRANSID');
  const hasPsFsa    = normalized.some(h => h === 'HMS_FSA_TRANS_ID');
  const hasPsAnyMed = normalized.some(h => h === 'HMS_MEDICAL_CD' || h === 'HMS_MBA_CLAIM_CAT' || h === 'HMS_REIMBURSE_AMT');

  // Check SAP signatures
  const hasSapRef    = normalized.some(h => h === 'REFERENCE' || h === 'REFERENCE_NUMBER' || h === 'REFERENCE_CODE');
  const hasSapAssign = normalized.some(h => h === 'HMS_ASSIGNMENT');
  const hasSapPers   = normalized.some(h => h === 'HMS_PERS_NO');
  const hasSapGl     = normalized.some(h => h === 'HMS_GL_AMOUNT');
  const hasSapSeq    = normalized.some(h => h === 'HMS_FBT_SEQNBR' || h === 'HMS_DOC_NBR');

  let psScore = 0;
  let psModule = 'MBA';
  if (hasEmplid) psScore += 2;
  if (hasPsMba) { psScore += 4; psModule = 'MBA'; }
  if (hasPsCb)  { psScore += 4; psModule = 'CB'; }
  if (hasPsFsa) { psScore += 4; psModule = 'FSA'; }
  if (hasPsAnyMed) psScore += 2;

  let sapScore = 0;
  if (hasSapRef) sapScore += 3;
  if (hasSapAssign) sapScore += 2;
  if (hasSapPers) sapScore += 2;
  if (hasSapGl) sapScore += 3;
  if (hasSapSeq) sapScore += 2;

  if (psScore >= 3 && psScore > sapScore) {
    const keyCol = hasPsMba ? 'HMS_INVOICE_NBR' : (hasPsCb ? 'ECS_RMB_SAP_REF' : (hasPsFsa ? 'HMS_FSA_TRANS_ID' : 'EMPLID'));
    return {
      valid: true,
      system: 'ps',
      module: psModule,
      headerIndex: hdrIdx,
      label: `PeopleSoft ${psModule}`,
      keyCol: keyCol,
      summary: `PeopleSoft (${psModule}) · Kunci: ${keyCol}`
    };
  }

  if (sapScore >= 3 && sapScore > psScore) {
    const keyCol = hasSapRef ? 'REFERENCE' : (hasSapAssign ? 'HMS_ASSIGNMENT' : 'REFERENCE');
    const sapMod = hasSapAssign && !hasSapRef ? 'FSA' : psModule;
    return {
      valid: true,
      system: 'sap',
      module: sapMod,
      headerIndex: hdrIdx,
      label: 'SAP Finance',
      keyCol: keyCol,
      summary: `SAP Finance · Kunci: ${keyCol}`
    };
  }

  return {
    valid: false,
    system: 'unknown',
    headerIndex: hdrIdx,
    reason: 'Kolom tidak dikenali sebagai format FBT PeopleSoft maupun SAP Finance.'
  };
}

function updateDropzoneUI(type) {
  const st = type === 'ps' ? psFileState : sapFileState;
  const dz = document.getElementById('dz-' + type);
  const badge = document.getElementById(type + '-badge');
  const msg = document.getElementById(type + '-validation-msg');
  const fname = document.getElementById(type + '-fname');
  const fdesc = document.getElementById(type + '-fdesc');

  if (!st || !st.profile) {
    resetDropzoneUI(type);
    return;
  }

  const validRows = Math.max(0, st.cleanRows.length - 1 - st.profile.headerIndex);
  const sizeFmt = (st.file.size / 1024 / 1024).toFixed(2) + ' MB';

  if (fname) fname.textContent = st.file.name;
  if (fdesc) fdesc.textContent = `${st.fmt} · ${sizeFmt} · ${validRows.toLocaleString('id-ID')} rows`;

  dz.classList.remove('dz-valid','dz-invalid','dz-warn','border-blue-300','bg-blue-50/40');
  msg.classList.remove('hidden');

  if (!st.profile.valid) {
    // Completely invalid file
    dz.classList.add('dz-invalid');
    badge.innerHTML = `<span class="text-rose-700 font-bold flex items-center gap-1.5"><svg class="w-4 h-4 text-rose-600 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg> ❌ Format Kolom Tidak Dikenali</span>`;
    msg.className = 'mt-2 text-[11px] text-rose-800 bg-rose-50 border border-rose-200 rounded-xl p-2.5 text-left';
    msg.innerHTML = `⚠️ <strong>Bukan file FBT valid</strong>: Kolom transaksi FBT tidak ditemukan. Harap pastikan file memuat kolom ${type === 'ps' ? 'EMPLID dan HMS_INVOICE_NBR / ECS_RMB_SAP_REF' : 'REFERENCE / HMS_ASSIGNMENT dan HMS_PERS_NO'}.`;
  } else if (st.profile.system !== type) {
    // Swapped file!
    dz.classList.add('dz-warn');
    badge.innerHTML = `<span class="text-amber-800 font-bold flex items-center gap-1.5"><svg class="w-4 h-4 text-amber-600 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg> ⚠️ Terdeteksi: ${st.profile.label}</span>`;
    msg.className = 'mt-2 text-[11px] text-amber-900 bg-amber-50 border border-amber-300 rounded-xl p-2.5 flex flex-col sm:flex-row items-center justify-between gap-2 shadow-sm text-left';
    msg.innerHTML = `
      <div class="flex items-center gap-1.5">
        <span class="text-base shrink-0">⚠️</span>
        <span>File ini berisi data <strong>${st.profile.label}</strong>, bukan ${type === 'ps' ? 'PeopleSoft' : 'SAP Finance'}.</span>
      </div>
      <button type="button" onclick="event.stopPropagation(); swapUploadedFiles();" 
        class="px-2.5 py-1 bg-amber-700 hover:bg-amber-800 text-white font-bold text-[10px] rounded-lg shadow transition whitespace-nowrap shrink-0">
        🔀 Tukar Posisi File
      </button>
    `;
  } else {
    // Valid and matches!
    dz.classList.add('dz-valid');
    badge.innerHTML = `<span class="text-emerald-700 font-semibold flex items-center gap-1.5"><svg class="w-4 h-4 text-emerald-600 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg> ✓ ${st.profile.summary} (${validRows.toLocaleString('id-ID')} baris)</span>`;
    msg.className = 'mt-2 text-[11px] text-emerald-700 font-semibold flex items-center gap-1';
    msg.innerHTML = `✅ Terverifikasi: Data ${st.profile.label} cocok untuk kotak ini.`;
  }
}

function resetDropzoneUI(type) {
  const dz = document.getElementById('dz-' + type);
  const msg = document.getElementById(type + '-validation-msg');
  const fname = document.getElementById(type + '-fname');
  const fdesc = document.getElementById(type + '-fdesc');
  const badge = document.getElementById(type + '-badge');

  if (dz) {
    dz.classList.remove('dz-valid','dz-invalid','dz-warn','border-blue-300','bg-blue-50/40');
    dz.classList.add('border-blue-300','bg-blue-50/40');
  }
  if (fname) fname.textContent = type === 'ps' ? 'Click or Drag PeopleSoft File' : 'Click or Drag SAP File';
  if (fdesc) fdesc.innerHTML = `Format: <span class="font-bold text-blue-700">HDC_FBT_${type.toUpperCase()}_[TYPE]_[PERIOD].xlsx</span>`;
  if (badge) badge.innerHTML = `<span>Waiting for ${type === 'ps' ? 'PeopleSoft' : 'SAP'} file...</span>`;
  if (msg) { msg.textContent = ''; msg.classList.add('hidden'); }
}

function swapUploadedFiles() {
  const temp = psFileState;
  psFileState = sapFileState;
  sapFileState = temp;

  rawPsRows = psFileState ? psFileState.cleanRows : [];
  rawSapRows = sapFileState ? sapFileState.cleanRows : [];

  updateDropzoneUI('ps');
  updateDropzoneUI('sap');
  checkReady();
  showToast('🔀 Posisi kedua file berhasil ditukar!');
}

// ============================================================
// CSV PARSER (RFC-4180 compliant)
// ============================================================
function parseCSV(text) {
  // Auto-detect delimiter
  const firstLineIdx = text.indexOf('\n');
  const firstLine = firstLineIdx > -1 ? text.substring(0, firstLineIdx) : text;
  const delimiter = firstLine.includes(';') ? ';' : ',';

  const rows = [];
  let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') {
        if (text[i+1] === '"') { field += '"'; i++; }
        else inQ = false;
      } else { field += c; }
    } else {
      if (c === '"') { inQ = true; }
      else if (c === delimiter) { row.push(field.trim()); field = ''; }
      else if (c === '\r') { if (text[i+1] === '\n') i++; row.push(field.trim()); rows.push(row); row = []; field = ''; }
      else if (c === '\n') { row.push(field.trim()); rows.push(row); row = []; field = ''; }
      else { field += c; }
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field.trim()); rows.push(row); }
  return rows;
}

// ============================================================
// FILE UPLOAD HANDLERS
// ============================================================
function onDragOver(e) { e.preventDefault(); e.currentTarget.classList.add('dropzone-active'); }
function onDragLeave(e) { e.preventDefault(); e.currentTarget.classList.remove('dropzone-active'); }
function onDrop(e, type) {
  e.preventDefault();
  e.currentTarget.classList.remove('dropzone-active');
  if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0], type);
}
function onFileSelect(e, type) {
  if (e.target.files.length) processFile(e.target.files[0], type);
}

function processFile(file, type) {
  const name = file.name.toLowerCase();
  const isExcel = name.endsWith('.xlsx') || name.endsWith('.xls');

  showToast('Reading file: ' + file.name + '...');

  if (isExcel) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        handleParsedRows(rows, file, type, 'Excel');
      } catch(err) {
        showToast('❌ Failed to read Excel: ' + err.message); console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rows = parseCSV(e.target.result);
      handleParsedRows(rows, file, type, 'CSV');
    };
    reader.readAsText(file);
  }
}

function handleParsedRows(rows, file, type, fmt) {
  // Filter out blank/ghost rows (rows where all cells are empty or whitespace)
  const cleanRows = rows.filter(r => 
    Array.isArray(r) && r.some(c => c !== null && c !== undefined && String(c).trim() !== '')
  );

  const profile = detectFileProfile(cleanRows);
  const stateObj = { file, cleanRows, fmt, profile };

  if (type === 'ps') {
    psFileState = stateObj;
    rawPsRows = cleanRows;
  } else {
    sapFileState = stateObj;
    rawSapRows = cleanRows;
  }

  updateDropzoneUI(type);

  // If the other dropzone already has a file, also refresh its UI
  const otherType = type === 'ps' ? 'sap' : 'ps';
  const otherState = otherType === 'ps' ? psFileState : sapFileState;
  if (otherState) {
    updateDropzoneUI(otherType);
  }

  checkReady();
  showToast('✓ File ' + file.name + ' berhasil dimuat.');
}

function checkReady() {
  const btn = document.getElementById('btn-process');
  const dot = document.getElementById('status-dot');
  const txt = document.getElementById('status-text');

  const psReady = psFileState && psFileState.profile && psFileState.profile.valid && psFileState.profile.system === 'ps';
  const sapReady = sapFileState && sapFileState.profile && sapFileState.profile.valid && sapFileState.profile.system === 'sap';

  const ready = psReady && sapReady;
  btn.disabled = !ready;

  if (ready) {
    dot.className = 'w-2 h-2 rounded-full bg-emerald-500 pulse-dot';
    txt.textContent = `Kedua file valid (${psFileState.profile.label} & ${sapFileState.profile.label})! Klik tombol "Compare & Reconcile" untuk memulai.`;
    txt.className = 'text-emerald-700 font-semibold text-xs';
  } else {
    const psHasFile = !!psFileState;
    const sapHasFile = !!sapFileState;

    if (psHasFile && sapHasFile) {
      if (psFileState.profile?.system === 'sap' && sapFileState.profile?.system === 'ps') {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500 animate-pulse';
        txt.innerHTML = '⚠️ <strong>File tertukar!</strong> PeopleSoft dan SAP terbalik posisinya. Klik tombol "Tukar Posisi File".';
        txt.className = 'text-amber-800 text-xs';
      } else if (!psFileState.profile?.valid || !sapFileState.profile?.valid) {
        dot.className = 'w-2 h-2 rounded-full bg-rose-500';
        txt.textContent = 'Salah satu atau kedua file tidak memiliki struktur kolom FBT yang valid.';
        txt.className = 'text-rose-700 text-xs';
      } else {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500';
        txt.textContent = 'Harap periksa posisi file: Tempatkan PeopleSoft di kotak PeopleSoft dan SAP di kotak SAP.';
        txt.className = 'text-amber-700 text-xs';
      }
    } else if (psHasFile) {
      if (psFileState.profile?.system !== 'ps') {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500';
        txt.textContent = 'File di kotak PeopleSoft terdeteksi sebagai SAP Finance. Klik tukar atau unggah file PeopleSoft.';
        txt.className = 'text-amber-700 text-xs';
      } else {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500';
        txt.textContent = `File PeopleSoft (${psFileState.profile.module}) siap. Menunggu file SAP Finance...`;
        txt.className = 'text-amber-700 text-xs';
      }
    } else if (sapHasFile) {
      if (sapFileState.profile?.system !== 'sap') {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500';
        txt.textContent = 'File di kotak SAP terdeteksi sebagai PeopleSoft. Klik tukar atau unggah file SAP.';
        txt.className = 'text-amber-700 text-xs';
      } else {
        dot.className = 'w-2 h-2 rounded-full bg-amber-500';
        txt.textContent = 'File SAP Finance siap. Menunggu file PeopleSoft...';
        txt.className = 'text-amber-700 text-xs';
      }
    } else {
      dot.className = 'w-2 h-2 rounded-full bg-slate-300';
      txt.textContent = 'Silakan upload kedua file untuk memulai rekonsiliasi.';
      txt.className = 'text-slate-500 text-xs';
    }
  }
}

// ============================================================
// CORE RECONCILIATION ENGINE
// ============================================================
function runComparison() {
  const btn = document.getElementById('btn-process');
  const spinner = document.getElementById('status-spinner');
  const statusText = document.getElementById('status-text');
  const statusDot = document.getElementById('status-dot');

  // UI: Processing state
  btn.disabled = true;
  btn.innerHTML = 'Processing...';
  spinner.classList.remove('hidden');
  statusDot.className = 'w-2 h-2 rounded-full bg-amber-400 animate-pulse';
  statusText.textContent = 'Processing reconciliation, matching records...';
  statusText.className = 'text-amber-700 font-bold';
  
  // Yield thread to browser to render the spinner
  setTimeout(() => {

  // --- Process PeopleSoft ---
  psMap.clear();
  const psHdrIdx = (psFileState && psFileState.profile && psFileState.profile.headerIndex !== undefined)
    ? psFileState.profile.headerIndex
    : findHeaderRowIndex(rawPsRows);
  const psHeader = rawPsRows[psHdrIdx] || [];

  if (psFileState && psFileState.profile && psFileState.profile.module) {
    currentDataType = psFileState.profile.module;
  } else {
    const psHeaderStr = psHeader.map(normalizeHeaderName).join(' ');
    if (psHeaderStr.includes('HMS_FSA_TRANS_ID')) {
      currentDataType = 'FSA';
    } else if (psHeaderStr.includes('ECS_RMB_SAP_REF') || psHeaderStr.includes('ECS_RMB_TRANSID')) {
      currentDataType = 'CB';
    } else {
      currentDataType = 'MBA';
    }
  }

  let ps_inv, ps_emp, ps_reimb, ps_med, ps_Memo, ps_cat, ps_date, ps_entitle;
  if (currentDataType === 'FSA') {
    ps_inv   = findColMulti(psHeader, ['HMS_FSA_TRANS_ID', 'FSA_TRANS_ID', 'TRANS_ID']);
    ps_emp   = findColMulti(psHeader, ['EMPLID', 'EMPL_ID', 'EMPLOYEE_ID', 'PERS_NO']);
    ps_reimb = findColMulti(psHeader, ['HMS_FLX_UNIT', 'FLX_UNIT', 'AMOUNT']);
    ps_med   = findColMulti(psHeader, ['HMS_FSA_ITEM', 'FSA_ITEM']);
    ps_Memo  = findColMulti(psHeader, ['DESCRLONG', 'DESCRIPTION', 'COMMENTS']);
    ps_cat   = findColMulti(psHeader, ['HMS_FLX_YEAR', 'FLX_YEAR']);
    ps_date  = findColMulti(psHeader, ['RECEIPT_DT', 'RECEIPT_DATE', 'DATE']);
    ps_entitle = ps_cat;
  } else if (currentDataType === 'CB') {
    ps_inv   = findColMulti(psHeader, ['ECS_RMB_SAP_REF', 'ECS_RMB_TRANSID', 'SAP_REF', 'TRANSID']);
    ps_emp   = findColMulti(psHeader, ['EMPLID', 'EMPL_ID', 'EMPLOYEE_ID', 'PERS_NO']);
    ps_reimb = findColMulti(psHeader, ['ECS_AMT_RMB', 'AMT_RMB', 'AMOUNT']);
    ps_med   = findColMulti(psHeader, ['EXPENSE_TYPE', 'EXP_TYPE']);
    ps_Memo  = findColMulti(psHeader, ['COMMENTS', 'COMMENT', 'DESCRLONG']);
    ps_cat   = findColMulti(psHeader, ['EXPENSE_ITEM', 'EXP_ITEM']);
    ps_date  = findColMulti(psHeader, ['RECEIPT_DT', 'RECEIPT_DATE', 'DATE']);
    ps_entitle = ps_date;
  } else {
    ps_inv   = findColMulti(psHeader, ['HMS_INVOICE_NBR', 'INVOICE_NBR', 'INVOICE_NO', 'INVOICE']);
    ps_emp   = findColMulti(psHeader, ['EMPLID', 'EMPL_ID', 'EMPLOYEE_ID', 'PERS_NO']);
    ps_reimb = findColMulti(psHeader, ['HMS_REIMBURSE_AMT', 'HMS_RECEIPT_AMT', 'REIMBURSE_AMT', 'AMOUNT']);
    ps_med   = findColMulti(psHeader, ['HMS_MEDICAL_CD', 'MEDICAL_CD']);
    ps_Memo  = findColMulti(psHeader, ['HMS_MEMO_LTR_NO', 'HMS_MED_GRANT_LTR', 'MEMO_LTR_NO']);
    ps_cat   = findColMulti(psHeader, ['HMS_MBA_CLAIM_CAT', 'CLAIM_CAT']);
    ps_date  = findColMulti(psHeader, ['RECEIPT_DT', 'ORG_RECEIPT_DT', 'DATE']);
    ps_entitle = findColMulti(psHeader, ['HMS_MED_ENTLT_PRD', 'ENTLT_PRD']);
  }

  // Safety fallbacks if any critical column was not matched
  if (ps_inv < 0) ps_inv = findColMulti(psHeader, ['HMS_INVOICE_NBR', 'ECS_RMB_SAP_REF', 'ECS_RMB_TRANSID', 'HMS_FSA_TRANS_ID']);
  if (ps_emp < 0) ps_emp = findColMulti(psHeader, ['EMPLID', 'EMPL_ID', 'EMPLOYEE_ID', 'PERS_NO']);
  if (ps_reimb < 0) ps_reimb = findColMulti(psHeader, ['HMS_REIMBURSE_AMT', 'ECS_AMT_RMB', 'HMS_FLX_UNIT', 'HMS_RECEIPT_AMT', 'AMOUNT']);

  let minTime = Infinity, maxTime = -Infinity;
  for (let i = psHdrIdx + 1; i < rawPsRows.length; i++) {
    const row = rawPsRows[i];
    if (!row || ps_inv < 0 || row.length <= ps_inv) continue;
    const inv = String(row[ps_inv] || '').trim();
    if (!inv) continue;

    const reimb = parseNum(ps_reimb >= 0 ? row[ps_reimb] : 0);
    const emplid = ps_emp >= 0 ? String(row[ps_emp] || '').trim() : '';
    const medCd = ps_med >= 0 ? String(row[ps_med] || '') : '';
    const claimCat = ps_cat >= 0 ? String(row[ps_cat] || '') : '';
    const memo = ps_Memo >= 0 ? String(row[ps_Memo] || '') : '';
    const date = ps_date >= 0 ? String(row[ps_date] || '') : '';
    const entitle = ps_entitle >= 0 ? String(row[ps_entitle] || '') : '';

    if (date) {
      const t = new Date(date).getTime();
      if (!isNaN(t) && t > 0) {
        if (t < minTime) minTime = t;
        if (t > maxTime) maxTime = t;
      }
    }

    if (!psMap.has(inv)) {
      psMap.set(inv, {
        inv, emplid,
        totalReimb: 0, items: [],
        medCd, claimCat
      });
    }
    const obj = psMap.get(inv);
    obj.totalReimb += reimb;
    obj.items.push({
      row, reimb,
      Memo   : memo,
      medCd  : medCd,
      cat    : claimCat,
      date   : date,
      entitle: entitle
    });
  }

  // --- Process SAP ---
  sapMap.clear();
  const sapHdrIdx = (sapFileState && sapFileState.profile && sapFileState.profile.headerIndex !== undefined)
    ? sapFileState.profile.headerIndex
    : findHeaderRowIndex(rawSapRows);
  const sapHeader = rawSapRows[sapHdrIdx] || [];

  const sap_ref  = currentDataType === 'FSA'
    ? findColMulti(sapHeader, ['HMS_ASSIGNMENT', 'REFERENCE', 'REFERENCE_NUMBER'])
    : findColMulti(sapHeader, ['REFERENCE', 'REFERENCE_NUMBER', 'HMS_ASSIGNMENT']);
  const sap_gl   = findColMulti(sapHeader, ['HMS_GL_AMOUNT', 'GL_AMOUNT', 'HMS_SUPL_AMOUNT', 'AMOUNT']);
  const sap_pers = findColMulti(sapHeader, ['HMS_PERS_NO', 'PERS_NO', 'PERNR', 'EMPLID']);
  const sap_doc  = findColMulti(sapHeader, ['HMS_DOC_NBR', 'DOC_NBR', 'DOCUMENT_NUMBER', 'BELNR']);
  const sap_txt  = findColMulti(sapHeader, ['TEXT254', 'HMS_HEADER_TEXT', 'SGTXT', 'TEXT']);
  const sap_tax  = findColMulti(sapHeader, ['TAX_CODE', 'MWSKZ']);

  for (let i = sapHdrIdx + 1; i < rawSapRows.length; i++) {
    const row = rawSapRows[i];
    if (!row || sap_ref < 0 || row.length <= sap_ref) continue;
    const ref = String(row[sap_ref] || '').trim();
    if (!ref) continue;

    const glAmt = parseNum(sap_gl >= 0 ? row[sap_gl] : 0);
    const persNo = sap_pers >= 0 ? String(row[sap_pers] || '').trim() : '';
    const docNbr = sap_doc >= 0 ? String(row[sap_doc] || '') : '';
    const text = sap_txt >= 0 ? String(row[sap_txt] || '') : '';
    const taxCode = sap_tax >= 0 ? String(row[sap_tax] || '') : '';

    if (!sapMap.has(ref)) {
      sapMap.set(ref, {
        ref, persNo,
        docNbr,
        totalGL: 0, lines: []
      });
    }
    const obj = sapMap.get(ref);
    obj.totalGL += glAmt;
    obj.lines.push({ 
      row, 
      glAmt, 
      text, 
      docNbr,
      taxCode
    });
  }

  // --- Reconcile ---
  reconData = [];
  const allKeys = new Set([...psMap.keys(), ...sapMap.keys()]);

  allKeys.forEach(key => {
    const p = psMap.get(key);
    const s = sapMap.get(key);

    const psTotal  = p ? p.totalReimb : 0;
    const sapTotal = s ? s.totalGL    : 0;
    const diff     = sapTotal - psTotal;  // positive = SAP > PS (Excess), negative = PS > SAP

    let status, statusLabel, statusClass;
    let isNikMismatch = false;

    if (p && s) {
      if (p.emplid !== s.persNo) {
        isNikMismatch = true;
        status = 'mismatch'; statusLabel = 'NIK Mismatch'; statusClass = 'badge-excess-ps';
      } else if (Math.abs(diff) < 1) {
        status = 'match'; statusLabel = 'MATCH'; statusClass = 'badge-matched';
      } else if (diff > 0) {
        status = 'mismatch'; statusLabel = 'SAP > PS'; statusClass = 'badge-excess-sap';
      } else {
        status = 'mismatch'; statusLabel = 'PS > SAP'; statusClass = 'badge-excess-ps';
      }
    } else if (p && !s) {
      status = 'notfound'; statusLabel = 'PSFT Only'; statusClass = 'badge-only-ps';
    } else {
      status = 'notfound'; statusLabel = 'SAP Only'; statusClass = 'badge-only-sap';
    }

    reconData.push({
      invoice: key,
      emplid : p ? p.emplid : (s ? s.persNo : '—'),
      psTotal, sapTotal, diff,
      status, statusLabel, statusClass,
      isNikMismatch,
      isDiff : status !== 'match',
      pData  : p || null,
      sData  : s || null,
      dataType: currentDataType
    });
  });

  // Set Date Range & Type labels
  let periodStr = "Periode";
  if (minTime !== Infinity && maxTime !== -Infinity) {
    const minD = new Date(minTime), maxD = new Date(maxTime);
    const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const p1 = `${minD.getDate()}${m[minD.getMonth()]}`;
    const p2 = `${maxD.getDate()}${m[maxD.getMonth()]}${String(maxD.getFullYear()).slice(-2)}`;
    periodStr = (p1 === p2.slice(0, p1.length)) ? p2 : `${p1}_${p2}`;
  }
  const elPeriode = document.getElementById('inp-periode');
  if (elPeriode) {
    elPeriode.value = periodStr;
    const ev = new Event('input');
    elPeriode.dispatchEvent(ev);
  }
  const sumType = document.getElementById('sum-Type');
  if (sumType) sumType.textContent = currentDataType;
  const dlTypePs = document.getElementById('dl-Type-ps');
  if (dlTypePs) dlTypePs.textContent = currentDataType;
  const dlTypeSap = document.getElementById('dl-Type-sap');
  if (dlTypeSap) dlTypeSap.textContent = currentDataType;

  // Update Table Headers dynamically based on currentDataType
  const hdrClaimCat = document.querySelector('th.col-ps-claimcat');
  const hdrMedType  = document.querySelector('th.col-ps-medtype');
  const hdrMemo     = document.querySelector('th.col-ps-memo');
  const hdrEntitle  = document.querySelector('th.col-ps-entitle');

  const lblClaimCat = document.querySelector('#chk-ps-claimcat + span');
  const lblMedType  = document.querySelector('#chk-ps-medtype + span');
  const lblMemo     = document.querySelector('#chk-ps-memo + span');
  const lblEntitle  = document.querySelector('#chk-ps-entitle + span');

  if (currentDataType === 'MBA') {
    if (hdrClaimCat) hdrClaimCat.innerHTML = 'Claim Cat <span class="text-slate-300">⇅</span>';
    if (hdrMedType)  hdrMedType.innerHTML  = 'Medical Type';
    if (hdrMemo)     hdrMemo.innerHTML     = 'Memo / RS';
    if (hdrEntitle)  hdrEntitle.innerHTML  = 'Entitlement Prd';
    if (lblClaimCat) lblClaimCat.textContent = 'HMS_MBA_CLAIM_CAT';
    if (lblMedType)  lblMedType.textContent  = 'HMS_MEDICAL_CD';
    if (lblMemo)     lblMemo.textContent     = 'HMS_MEMO_LTR_NO';
    if (lblEntitle)  lblEntitle.textContent  = 'HMS_MED_ENTLT_PRD';
  } else if (currentDataType === 'CB') {
    if (hdrClaimCat) hdrClaimCat.innerHTML = 'Expense Item <span class="text-slate-300">⇅</span>';
    if (hdrMedType)  hdrMedType.innerHTML  = 'Expense Type';
    if (hdrMemo)     hdrMemo.innerHTML     = 'Comments';
    if (hdrEntitle)  hdrEntitle.innerHTML  = 'Receipt Date';
    if (lblClaimCat) lblClaimCat.textContent = 'EXPENSE_ITEM';
    if (lblMedType)  lblMedType.textContent  = 'EXPENSE_TYPE';
    if (lblMemo)     lblMemo.textContent     = 'COMMENTS';
    if (lblEntitle)  lblEntitle.textContent  = 'RECEIPT_DT';
  } else if (currentDataType === 'FSA') {
    if (hdrClaimCat) hdrClaimCat.innerHTML = 'Flex Year <span class="text-slate-300">⇅</span>';
    if (hdrMedType)  hdrMedType.innerHTML  = 'FSA Item';
    if (hdrMemo)     hdrMemo.innerHTML     = 'Description';
    if (hdrEntitle)  hdrEntitle.innerHTML  = 'Receipt Date';
    if (lblClaimCat) lblClaimCat.textContent = 'HMS_FLX_YEAR';
    if (lblMedType)  lblMedType.textContent  = 'HMS_FSA_ITEM';
    if (lblMemo)     lblMemo.textContent     = 'DESCRLONG';
    if (lblEntitle)  lblEntitle.textContent  = 'RECEIPT_DT';
  }

  // Sort: not-found first, then mismatch, then match; by abs(diff) desc within groups
  reconData.sort((a, b) => {
    const order = { notfound: 0, mismatch: 1, match: 2 };
    if (order[a.status] !== order[b.status]) return order[a.status] - order[b.status];
    return Math.abs(b.diff) - Math.abs(a.diff);
  });

  updateKPIs();
  updateSummaryPage();
  currentFilter = 'all';
  searchQuery   = '';
  currentPage   = 1;
  const searchInput = document.getElementById('search-input');
  if (searchInput) searchInput.value = '';
  applyFilter();
  // UI: Success state
  spinner.classList.add('hidden');
  statusDot.className = 'w-2 h-2 rounded-full bg-emerald-500';
  statusText.textContent = 'Success! ' + reconData.length + ' invoices reconciled.';
  statusText.className = 'text-emerald-700 font-bold';
  showToast('✓ Reconciliation complete! ' + reconData.length + ' invoices processed.');
  
  // Switch tab after showing success for a brief moment
  setTimeout(() => {
    const exportBtn = document.getElementById('btn-export-all');
    if (exportBtn) exportBtn.disabled = false;

    switchTab('compare');
    // Reset button and status for next time
    btn.innerHTML = `
      <svg id="btn-icon-compare" class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
      </svg>
      Compare & Reconcile
    `;
    btn.disabled = false;
    statusText.className = '';
    statusDot.className = 'w-2 h-2 rounded-full bg-slate-300';
    statusText.textContent = 'Please upload both files to start reconciliation.';
  }, 1000);

  }, 100); // Small delay to let the browser paint the loading state
}

// ---- FORMATTER ----
function formatShortIDR(val) {
  const absVal = Math.abs(val);
  if (absVal >= 1e9) {
    return 'Rp ' + (absVal / 1e9).toFixed(2).replace('.', ',') + ' M';
  } else if (absVal >= 1e6) {
    return 'Rp ' + (absVal / 1e6).toFixed(2).replace('.', ',') + ' Jt';
  } else {
    return 'Rp ' + absVal.toLocaleString('id-ID');
  }
}

// ---- KPI UPDATER ----
function updateKPIs() {
  const total     = reconData.length;
  const matched   = reconData.filter(d => d.status === 'match').length;
  const mismatch  = reconData.filter(d => d.status === 'mismatch').length;
  const notfound  = reconData.filter(d => d.status === 'notfound').length;
  const grandPS   = reconData.reduce((s, d) => s + d.psTotal,  0);
  const grandSAP  = reconData.reduce((s, d) => s + d.sapTotal, 0);
  const grandDiff = grandSAP - grandPS;

  const matchPct = total > 0 ? ((matched / total) * 100).toFixed(1) : 0;

  setText('kpi-match', matched.toLocaleString('id-ID').replace(/,/g, '.'));
  setText('kpi-match-sub', `✓ ${matchPct}% (${total.toLocaleString('id-ID').replace(/,/g, '.')} Total Unik)`);
  
  setText('kpi-ps-short', formatShortIDR(grandPS));
  setText('kpi-ps-total', 'Rp ' + Math.round(grandPS).toLocaleString('id-ID').replace(/,/g, '.') + ',00');
  
  setText('kpi-sap-short', formatShortIDR(grandSAP));
  setText('kpi-sap-total', 'Rp ' + Math.round(grandSAP).toLocaleString('id-ID').replace(/,/g, '.') + ',00');

  setText('kpi-mismatch-count', `${mismatch} Invoice`);
  setText('kpi-diff-short', formatShortIDR(grandDiff));
  
  const diffPrefix = grandDiff >= 0 ? '+' : '-';
  setText('kpi-diff', `${diffPrefix}Rp ${Math.round(Math.abs(grandDiff)).toLocaleString('id-ID').replace(/,/g, '.')},00`);

  setText('cnt-all',      total);
  setText('cnt-match',    matched);
  setText('cnt-mismatch', mismatch);
  setText('cnt-notfound', notfound);
}

// ---- SUMMARY PAGE ----
function updateSummaryPage() {
  const elJenis = document.getElementById('inp-jenis');
  const elPeriode = document.getElementById('inp-periode');
  const jenis   = (elJenis ? elJenis.value : null) || 'MBA';
  const periode = (elPeriode ? elPeriode.value : null) || '10Agt_9Sept26';

  const total    = reconData.length;
  const matched  = reconData.filter(d => d.status === 'match').length;
  const mismatch = reconData.filter(d => d.status === 'mismatch').length;
  const notfound = reconData.filter(d => d.status === 'notfound').length;
  const matchPct    = total > 0 ? ((matched  / total) * 100).toFixed(1) : 0;
  const mismatchPct = total > 0 ? ((mismatch / total) * 100).toFixed(1) : 0;

  const matchedRows = reconData.filter(d => d.status === 'match');
  const mmRows      = reconData.filter(d => d.status === 'mismatch');
  const grandPS     = reconData.reduce((s, d) => s + d.psTotal,  0);
  const grandSAP    = reconData.reduce((s, d) => s + d.sapTotal, 0);
  const mmPS        = mmRows.reduce((s, d) => s + d.psTotal,  0);
  const mmSAP       = mmRows.reduce((s, d) => s + d.sapTotal, 0);
  const matchPS     = matchedRows.reduce((s, d) => s + d.psTotal,  0);
  const matchSAP    = matchedRows.reduce((s, d) => s + d.sapTotal, 0);

  // Hero section
  setText('sum-Period', periode.replace(/_/g, ' '));
  setText('sum-Type',   jenis);
  setText('sum-ts',     new Date().toLocaleString('id-ID'));

  // KPI Cards
  setText('s-total',      total);
  setText('s-match',      matched);
  setText('s-match-pct',  matchPct + '%');
  setText('s-mismatch',   mismatch);
  setText('s-mismatch-pct', mismatchPct + '%');
  setText('s-notfound',   notfound);

  // Progress bars
  const matchBar    = document.getElementById('s-match-bar');
  const mismatchBar = document.getElementById('s-mismatch-bar');
  const notfoundBar = document.getElementById('s-notfound-bar');
  if (matchBar)    matchBar.style.width    = matchPct + '%';
  if (mismatchBar) mismatchBar.style.width = mismatchPct + '%';
  if (notfoundBar) notfoundBar.style.width = (total > 0 ? ((notfound / total) * 100).toFixed(1) : 0) + '%';

  // Financial Summary
  setText('fin-match-ps',   'IDR ' + Math.round(matchPS).toLocaleString('id-ID'));
  setText('fin-match-sap',  'IDR ' + Math.round(matchSAP).toLocaleString('id-ID'));
  setText('fin-mm-ps',      'IDR ' + Math.round(mmPS).toLocaleString('id-ID'));
  setText('fin-mm-sap',     'IDR ' + Math.round(mmSAP).toLocaleString('id-ID'));
  setText('fin-mm-diff',    (mmSAP - mmPS >= 0 ? '+' : '') + 'IDR ' + Math.round(mmSAP - mmPS).toLocaleString('id-ID'));
  setText('fin-total-ps',   'IDR ' + Math.round(grandPS).toLocaleString('id-ID'));
  setText('fin-total-sap',  'IDR ' + Math.round(grandSAP).toLocaleString('id-ID'));
  const grandDiff = grandSAP - grandPS;
  const diffStr = (grandDiff >= 0 ? '+' : '') + 'IDR ' + Math.round(grandDiff).toLocaleString('id-ID');
  setText('fin-total-diff', diffStr);
  
  // Set the Net Difference KPI Card in Download Tab
  const diffPrefix = grandDiff >= 0 ? '+' : '-';
  setText('sum-netdiff', `${diffPrefix}Rp ${Math.round(Math.abs(grandDiff)).toLocaleString('id-ID').replace(/,/g, '.')},00`);

  // File card labels
  ['ps','sap'].forEach(t => {
    const el_t = document.getElementById('dl-Type-' + t);
    const el_p = document.getElementById('dl-Period-' + t);
    if (el_t) el_t.textContent = jenis;
    if (el_p) el_p.textContent = periode;
  });

  setText('dl-ps-rows',  rawPsRows.length - 1 + ' rows · ' + psMap.size + ' unique invoices');
  setText('dl-sap-rows', rawSapRows.length - 1 + ' rows · ' + sapMap.size + ' unique invoices');
}

// ============================================================
// FILTER & SEARCH
// ============================================================
function setFilter(f) {
  currentFilter = f;
  currentPage   = 1;
  document.querySelectorAll('.flt-btn').forEach(b => {
    b.className = 'flt-btn px-3 py-1 text-xs font-medium rounded-lg text-slate-600 hover:bg-white/70';
  });
  const activeBtn = document.getElementById('flt-' + f);
  if (activeBtn) activeBtn.className = 'flt-btn px-3 py-1 text-xs font-semibold rounded-lg bg-white text-navy-700 shadow-sm';
  applyFilter();
}

function doSearch(e) {
  searchQuery = e.target.value.toLowerCase().trim();
  currentPage = 1;
  applyFilter();
}

function applyFilter() {
  filteredData = reconData.filter(d => {
    if (currentFilter === 'match'    && d.status !== 'match')    return false;
    if (currentFilter === 'mismatch' && d.status !== 'mismatch') return false;
    if (currentFilter === 'notfound' && d.status !== 'notfound') return false;
    if (searchQuery) {
      const q = searchQuery;
      const inInv  = d.invoice.toLowerCase().includes(q);
      const inEmp  = d.emplid.toLowerCase().includes(q);
      const inStat = d.statusLabel.toLowerCase().includes(q);
      if (!inInv && !inEmp && !inStat) return false;
    }
    return true;
  });

  if (sortCol !== 'none') {
    filteredData.sort((a, b) => {
      let valA = a[sortCol];
      let valB = b[sortCol];
      if (typeof valA === 'string') valA = valA.toLowerCase();
      if (typeof valB === 'string') valB = valB.toLowerCase();
      
      if (valA < valB) return sortAsc ? -1 : 1;
      if (valA > valB) return sortAsc ? 1 : -1;
      return 0;
    });
  }

  renderTable();
}

function setSort(col) {
  if (sortCol === col) {
    sortAsc = !sortAsc;
  } else {
    sortCol = col;
    sortAsc = true;
  }
  applyFilter();
}

// ============================================================
// COLUMN VISIBILITY
// ============================================================
// Map: colId -> which CSS class to toggle on all matching th/td
const COL_CLASSES = {
  'ps-receipt' : 'col-ps-receipt',
  'ps-claimcat': 'col-ps-claimcat',
  'ps-medtype' : 'col-ps-medtype',
  'ps-memo'    : 'col-ps-memo',
  'ps-entitle' : 'col-ps-entitle',
  'ps-lines'   : 'col-ps-lines',
  'sap-docnbr' : 'col-sap-docnbr',
  'sap-postdt' : 'col-sap-postdt',
  'sap-company': 'col-sap-company',
  'sap-gltext' : 'col-sap-gltext',
  'sap-costctr': 'col-sap-costctr',
  'sap-gllines': 'col-sap-gllines',
};

function toggleColumnDropdown(e) {
  e.stopPropagation();
  const menu = document.getElementById('col-dropdown');
  menu.classList.toggle('dropdown-hidden');
}

// Close dropdown if clicked outside
document.addEventListener('click', (e) => {
  const menu = document.getElementById('col-dropdown');
  if (menu && !menu.classList.contains('dropdown-hidden') && !e.target.closest('.relative')) {
    menu.classList.add('dropdown-hidden');
  }
});

function toggleColumn(colId) {
  const cls   = COL_CLASSES[colId];
  const chk   = document.getElementById('chk-' + colId);
  if (!cls || !chk) return;

  const isChecked = chk.checked;
  document.querySelectorAll('.' + cls).forEach(el => {
    if (isChecked) el.classList.remove('col-hidden');
    else           el.classList.add('col-hidden');
  });

  updateGroupColspan();
}

function updateGroupColspan() {
  // PS group
  const psVisible  = ['col-ps-receipt','col-ps-claimcat','col-ps-medtype','col-ps-memo','col-ps-entitle','col-ps-lines','col-ps-total']
    .filter(c => document.querySelector('.' + c + ':not(.col-hidden)') !== null).length;
  // SAP group
  const sapVisible = ['col-sap-docnbr','col-sap-postdt','col-sap-company','col-sap-gltext','col-sap-costctr','col-sap-gllines','col-sap-total']
    .filter(c => document.querySelector('.' + c + ':not(.col-hidden)') !== null).length;

  const grpPS  = document.getElementById('grp-ps');
  const grpSAP = document.getElementById('grp-sap');
  if (grpPS)  grpPS.colSpan  = Math.max(1, psVisible);
  if (grpSAP) grpSAP.colSpan = Math.max(1, sapVisible);
}

// ============================================================
// TABLE RENDER — col-* classes on every td for toggle support
// ============================================================
function renderTable() {
  const tbody = document.getElementById('compare-tbody');

  // Determine current visibility state from checkboxes
  const vis = {};
  for (const [colId, cls] of Object.entries(COL_CLASSES)) {
    const chk = document.getElementById('chk-' + colId);
    // If checkbox is checked, it's visible (no col-hidden). If unchecked, it's hidden.
    vis[cls] = chk ? chk.checked : false;
  }

  if (!filteredData.length) {
    tbody.innerHTML = `<tr><td colspan="18" class="text-center py-16 text-slate-400">No data matches the filter/search.</td></tr>`;
    updatePagination(0, 1, 0, 0);
    return;
  }

  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  currentPage = Math.max(1, Math.min(currentPage, totalPages));
  const start = (currentPage - 1) * pageSize;
  const end   = Math.min(start + pageSize, filteredData.length);
  const slice = filteredData.slice(start, end);

  tbody.innerHTML = slice.map((d, i) => {
    // ---- Row accent ----
    const rowAccent = d.status === 'match'             ? ''             :
                      d.statusLabel === 'SAP > PS'     ? 'row-excess-sap' :
                      d.statusLabel === 'PS > SAP'     ? 'row-excess-ps'  :
                      d.statusLabel === 'PSFT Only'    ? 'row-only-ps'    : 'row-only-sap';

    // ---- PS fields ----
    const psReceipt  = d.pData && d.pData.items[0] ? d.pData.items[0].date : '—';
    const psClaimCat = d.pData ? (d.pData.items.map(x => x.cat).filter((v,i,a) => a.indexOf(v)===i).join(', ') || '—') : '—';
    const psMedTypes = d.pData ? [...new Set(d.pData.items.map(x => currentDataType === 'MBA' ? medLabel(x.medCd) : x.medCd))].join(', ') || '—' : '—';
    const psMemo     = d.pData ? (d.pData.items.map(x => x.Memo).filter(Boolean).join(' / ') || '—') : '—';
    const psEntitle  = d.pData && d.pData.items[0] ? d.pData.items[0].entitle : '—';
    const psLines    = d.pData ? d.pData.items.length : 0;
    const psTotalFmt = d.psTotal > 0 ? Math.round(d.psTotal).toLocaleString('id-ID') : '—';

    // ---- SAP fields ----
    const sapDoc     = d.sData ? [...new Set(d.sData.lines.map(x => x.docNbr))].join(', ') || '—' : '—';
    const sapPostDt  = d.sData && d.sData.lines[0] ? getColVal(rawSapRows, d.sData.lines[0].row, 'POSTING_DT', 9) : '—';
    const sapCompany = d.sData && d.sData.lines[0] ? getColVal(rawSapRows, d.sData.lines[0].row, 'HMS_COMPANY_CD', 1) : '—';
    const sapGLText  = d.sData ? [...new Set(d.sData.lines.map(x => x.text).filter(Boolean))].slice(0,2).join(' | ') || '—' : '—';
    const sapCostCtr = d.sData && d.sData.lines[0] ? getColVal(rawSapRows, d.sData.lines[0].row, 'HMS_COST_CENTER', 20) : '—';
    const sapLines   = d.sData ? d.sData.lines.length : 0;
    const sapTotalFmt= d.sapTotal > 0 ? Math.round(d.sapTotal).toLocaleString('id-ID') : '—';

    // ---- Diff ----
    const hasDiff       = Math.abs(d.diff) >= 1;
    const diffAbs       = Math.round(Math.abs(d.diff)).toLocaleString('id-ID');
    const diffSign      = d.diff > 0 ? '+' : (d.diff < 0 ? '−' : '');
    const diffStr       = hasDiff ? `${diffSign} ${diffAbs}` : '—';
    const diffCellClass = d.diff > 0 ? 'text-amber-700 font-bold' :
                          d.diff < 0 ? 'text-rose-700 font-bold'  : 'text-slate-400';
    const diffArrow     = d.diff > 0 ? '<span class="text-amber-500 ml-1">↑</span>' :
                          d.diff < 0 ? '<span class="text-rose-500 ml-1">↓</span>' : '';

    // ---- Keterangan ----
    const ket = buildKetSAP(d);

    // ---- Chips on total cells ----
    const psTotalCell  = d.pData
      ? `<span class="font-semibold text-emerald-800">${psTotalFmt}</span>${hasDiff ? `<span class="diff-chip ml-1">${d.psTotal < d.sapTotal ? '↓ kecil' : '↑ besar'}</span>` : '<span class="diff-chip chip-ok ml-1">✓</span>'}`
      : `<span class="text-slate-300 italic">—</span>`;
    const sapTotalCell = d.sData
      ? `<span class="font-semibold text-blue-800">${sapTotalFmt}</span>${hasDiff ? `<span class="diff-chip ml-1">${d.sapTotal > d.psTotal ? '↑ besar' : '↓ kecil'}</span>` : '<span class="diff-chip chip-ok ml-1">✓</span>'}`
      : `<span class="text-slate-300 italic">—</span>`;

    // ---- Freeze bg ----
    const fbg = d.status === 'match'             ? '#fff'    :
                d.statusLabel === 'SAP > PS'     ? '#fffbeb' :
                d.statusLabel === 'PS > SAP'     ? '#fff1f2' :
                d.statusLabel === 'PSFT Only'    ? '#faf5ff' : '#eff6ff';

    return `
      <tr class="data-row ${rowAccent} transition hover:brightness-95">
        <td class="col-no  tbl-freeze-no  text-center text-slate-400 font-sans"         style="background:${fbg}">${start+i+1}</td>
        <td class="col-inv tbl-freeze-inv font-bold text-navy-700 font-mono"            style="background:${fbg}; max-width:150px;">
          <div class="truncate" title="${d.invoice}">${d.invoice}</div>
          ${!d.pData ? '<div class="text-[9px] text-blue-500 font-semibold">SAP only</div>' : !d.sData ? '<div class="text-[9px] text-purple-500 font-semibold">PS only</div>' : ''}
        </td>
        <td class="col-emp tbl-freeze-emp font-semibold text-slate-700"                 style="background:${fbg}">${d.emplid}</td>

        <!-- PS columns -->
        <td class="col-ps-receipt ${vis['col-ps-receipt'] ? '' : 'col-hidden'} border-l border-green-100 text-slate-600 font-mono text-[10px]">
          ${psReceipt !== '—' ? psReceipt : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-ps-claimcat ${vis['col-ps-claimcat'] ? '' : 'col-hidden'} border-l border-green-100 text-emerald-700 font-medium">
          ${psClaimCat !== '—' ? `<span class="bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded text-[10px]">${psClaimCat}</span>` : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-ps-medtype ${vis['col-ps-medtype'] ? '' : 'col-hidden'} border-l border-green-100 text-slate-600">
          <div class="truncate max-w-[90px]" title="${psMedTypes}">${psMedTypes !== '—' ? psMedTypes : '<span class="text-slate-300">—</span>'}</div>
        </td>
        <td class="col-ps-memo ${vis['col-ps-memo'] ? '' : 'col-hidden'} border-l border-green-100 text-slate-600">
          <div class="truncate max-w-[120px]" title="${psMemo}">${psMemo !== '—' ? psMemo : '<span class="text-slate-300">—</span>'}</div>
        </td>
        <td class="col-ps-entitle ${vis['col-ps-entitle'] ? '' : 'col-hidden'} border-l border-green-100 text-slate-600 text-center font-mono">
          ${psEntitle !== '—' ? psEntitle : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-ps-lines ${vis['col-ps-lines'] ? '' : 'col-hidden'} border-l border-green-100 text-center">
          ${psLines > 0 ? `<span class="bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded-full text-[10px]">${psLines}</span>` : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-ps-total border-l border-green-100 text-right">${psTotalCell}</td>

        <!-- SAP columns -->
        <td class="col-sap-docnbr ${vis['col-sap-docnbr'] ? '' : 'col-hidden'} border-l border-blue-100 font-mono text-blue-700">
          <div class="truncate max-w-[90px]" title="${sapDoc}">${sapDoc !== '—' ? sapDoc : '<span class="text-slate-300">—</span>'}</div>
        </td>
        <td class="col-sap-postdt ${vis['col-sap-postdt'] ? '' : 'col-hidden'} border-l border-blue-100 text-slate-600 font-mono text-[10px]">
          ${sapPostDt !== '—' ? sapPostDt : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-sap-company ${vis['col-sap-company'] ? '' : 'col-hidden'} border-l border-blue-100 text-center">
          ${sapCompany !== '—' ? `<span class="bg-blue-50 border border-blue-200 px-1.5 py-0.5 rounded text-[10px] font-mono">${sapCompany}</span>` : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-sap-gltext ${vis['col-sap-gltext'] ? '' : 'col-hidden'} border-l border-blue-100 text-slate-600">
          <div class="truncate max-w-[120px]" title="${sapGLText}">${sapGLText !== '—' ? sapGLText : '<span class="text-slate-300">—</span>'}</div>
        </td>
        <td class="col-sap-costctr ${vis['col-sap-costctr'] ? '' : 'col-hidden'} border-l border-blue-100 text-slate-600 text-center font-mono text-[10px]">
          ${sapCostCtr !== '—' ? sapCostCtr : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-sap-gllines ${vis['col-sap-gllines'] ? '' : 'col-hidden'} border-l border-blue-100 text-center">
          ${sapLines > 0 ? `<span class="bg-blue-100 text-blue-800 font-bold px-2 py-0.5 rounded-full text-[10px]">${sapLines}</span>` : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-sap-total border-l border-blue-100 text-right">${sapTotalCell}</td>

        <!-- Recon columns -->
        <td class="col-rec-diff border-l border-amber-100 text-right font-mono ${diffCellClass}${hasDiff ? ' bg-amber-50/50' : ''}">
          ${hasDiff ? `${diffStr}${diffArrow}` : '<span class="text-slate-300">—</span>'}
        </td>
        <td class="col-rec-ket border-l border-amber-100 text-slate-600 align-top">
          <div class="whitespace-normal break-words w-48 text-[10px] leading-snug">${ket}</div>
        </td>
        <td class="col-rec-status border-l border-amber-100 text-center">
          <span class="px-2 py-0.5 font-bold text-[10px] ${d.statusClass} whitespace-nowrap" style="border-radius:20px;">${d.statusLabel}</span>
        </td>
        <td class="col-rec-detail border-l border-amber-100 text-center">
          <button onclick="openDetailModal('${d.invoice}')" title="Lihat baris detail"
            class="inline-flex items-center gap-1 px-2 py-1 bg-navy-50 hover:bg-navy-100 text-navy-600 border border-navy-200 rounded-lg text-[10px] font-semibold transition">
            <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
            Rows
          </button>
        </td>
      </tr>`;
  }).join('');

  updatePagination(filteredData.length, totalPages, start, end);
}

function getColVal(rawRows, row, colName, fallbackIdx) {
  const sapHdrIdx = (sapFileState?.profile?.headerIndex !== undefined) ? sapFileState.profile.headerIndex : 0;
  const hdr = rawRows[sapHdrIdx] || rawRows[0] || [];
  const idx = findCol(hdr, colName, fallbackIdx);
  return idx >= 0 ? (String(row[idx] || '').trim() || '—') : '—';
}

function changePageSize(val) {
  pageSize = parseInt(val, 10) || 10;
  currentPage = 1;
  renderTable();
}

function goToPage(page) {
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  currentPage = Math.max(1, Math.min(page, totalPages));
  renderTable();
}

function prevPage() {
  if (currentPage > 1) {
    currentPage--;
    renderTable();
  }
}

function nextPage() {
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  if (currentPage < totalPages) {
    currentPage++;
    renderTable();
  }
}

function updatePagination(total, totalPages, start, end) {
  const infoEl = document.getElementById('pg-entries-info');
  if (infoEl) {
    if (total === 0) {
      infoEl.textContent = '0 - 0 of 0 entries';
    } else {
      infoEl.textContent = `${start + 1} - ${end} of ${total.toLocaleString('id-ID')} entries`;
    }
  }

  const ctrlEl = document.getElementById('pagination-controls');
  if (!ctrlEl) return;

  if (total === 0) {
    ctrlEl.innerHTML = `
      <button disabled class="w-8 h-8 flex items-center justify-center rounded-xl text-slate-300 cursor-not-allowed">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"/></svg>
      </button>
      <button class="w-8 h-8 flex items-center justify-center rounded-xl bg-blue-600 text-white font-bold text-xs shadow-sm">1</button>
      <button disabled class="w-8 h-8 flex items-center justify-center rounded-xl text-slate-300 cursor-not-allowed">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/></svg>
      </button>
    `;
    return;
  }

  let html = '';

  // Prev Button
  const prevDisabled = currentPage <= 1;
  html += `
    <button onclick="prevPage()" ${prevDisabled ? 'disabled' : ''} title="Previous Page"
      class="w-8 h-8 flex items-center justify-center rounded-xl ${prevDisabled ? 'text-slate-300 cursor-not-allowed' : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 transition cursor-pointer'}">
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7" />
      </svg>
    </button>
  `;

  // Page Numbers
  const pages = getPageNumbers(currentPage, totalPages);
  pages.forEach(p => {
    if (p === '...') {
      html += `<span class="w-8 h-8 flex items-center justify-center text-slate-400 text-xs select-none">...</span>`;
    } else {
      const isActive = p === currentPage;
      if (isActive) {
        html += `
          <button class="w-8 h-8 flex items-center justify-center rounded-xl bg-blue-600 text-white font-bold text-xs shadow-sm" aria-current="page">
            ${p}
          </button>
        `;
      } else {
        html += `
          <button onclick="goToPage(${p})"
            class="w-8 h-8 flex items-center justify-center rounded-xl text-slate-600 hover:text-slate-900 hover:bg-slate-200/70 text-xs font-medium transition cursor-pointer">
            ${p}
          </button>
        `;
      }
    }
  });

  // Next Button
  const nextDisabled = currentPage >= totalPages;
  html += `
    <button onclick="nextPage()" ${nextDisabled ? 'disabled' : ''} title="Next Page"
      class="w-8 h-8 flex items-center justify-center rounded-xl ${nextDisabled ? 'text-slate-300 cursor-not-allowed' : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 transition cursor-pointer'}">
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7" />
      </svg>
    </button>
  `;

  ctrlEl.innerHTML = html;
}

function getPageNumbers(current, total) {
  if (total <= 10) {
    return Array.from({ length: total }, (_, i) => i + 1);
  }
  if (current <= 4) {
    return [1, 2, 3, 4, 5, '...', total];
  }
  if (current >= total - 3) {
    return [1, '...', total - 4, total - 3, total - 2, total - 1, total];
  }
  return [1, '...', current - 1, current, current + 1, '...', total];
}

// ============================================================
// MODAL DETAIL VIEW
// ============================================================
function openDetailModal(inv) {
  const d = reconData.find(r => r.invoice === inv);
  if (!d) return;

  const diffStr = d.diff === 0 ? '—' : (d.diff > 0 ? '+' : '') + 'IDR ' + Math.round(d.diff).toLocaleString('id-ID');
  const diffColor = d.diff > 0 ? 'text-amber-600' : d.diff < 0 ? 'text-pink-600' : 'text-emerald-600';

  document.getElementById('modal-title').textContent = 'Invoice Detail: ' + inv;
  document.getElementById('modal-sub').innerHTML =
    `Employee ID: <strong class="text-white">${d.emplid}</strong>
    &nbsp;·&nbsp; Status: <strong class="text-white">${d.statusLabel}</strong>
    &nbsp;·&nbsp; Difference: <strong class="${diffColor}">${diffStr}</strong>`;

  document.getElementById('modal-body').innerHTML = generateDetailHTML(d);
  document.getElementById('modal').classList.remove('hidden');
}

function closeModal() {
  document.getElementById('modal').classList.add('hidden');
}

function generateDetailHTML(d) {
  const diffColor = d.diff > 0 ? 'text-amber-600' : d.diff < 0 ? 'text-pink-600' : 'text-emerald-600';

  // PS Rows
  let psHtml = '';
  if (d.pData && d.pData.items.length) {
    psHtml = d.pData.items.map((it, i) => `
      <tr class="hover:bg-slate-50">
        <td class="p-2 border border-slate-200">${i+1}</td>
        <td class="p-2 border border-slate-200">${it.medCd} (${medLabel(it.medCd)})</td>
        <td class="p-2 border border-slate-200 text-center">${it.cat}</td>
        <td class="p-2 border border-slate-200 text-right font-bold text-emerald-700">IDR ${Math.round(it.reimb).toLocaleString('id-ID')}</td>
        <td class="p-2 border border-slate-200 font-sans text-slate-600 text-[10px] whitespace-normal break-words">${it.Memo || '—'}</td>
      </tr>`).join('');
  } else {
    psHtml = `<tr><td colspan="5" class="p-3 text-center text-slate-400">No PeopleSoft data for this invoice.</td></tr>`;
  }

  // SAP Rows
  let sapHtml = '';
  if (d.sData && d.sData.lines.length) {
    sapHtml = d.sData.lines.map((ln, i) => {
      const isExcess = ln.text.toLowerCase().includes('excess');
      return `
      <tr class="hover:bg-slate-50 ${isExcess ? 'bg-amber-50/50' : ''}">
        <td class="p-2 border border-slate-200">${i+1}</td>
        <td class="p-2 border border-slate-200 font-bold">${ln.docNbr}</td>
        <td class="p-2 border border-slate-200 text-right font-bold ${isExcess ? 'text-amber-600' : 'text-slate-800'}">
          IDR ${Math.round(ln.glAmt).toLocaleString('id-ID')}
        </td>
        <td class="p-2 border border-slate-200 font-sans text-slate-600 text-[10px] whitespace-normal break-words">${ln.text || '—'}</td>
        <td class="p-2 border border-slate-200 text-center">
          ${isExcess
            ? '<span class="bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded text-[9px] font-bold">Excess MBA</span>'
            : '<span class="text-slate-400 text-[9px]">GL Expense</span>'}
        </td>
      </tr>`;
    }).join('');
  } else {
    sapHtml = `<tr><td colspan="5" class="p-3 text-center text-slate-400">No SAP data for this invoice.</td></tr>`;
  }

  return `
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
      <!-- PS Table -->
      <div>
        <h4 class="font-bold text-[10px] uppercase tracking-wider text-slate-700 mb-2 font-sans flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
          PeopleSoft HR Data (${d.pData ? d.pData.items.length : 0} claim lines)
        </h4>
        <div class="overflow-x-auto border border-slate-200 rounded-lg bg-white shadow-sm">
          <table class="w-full text-left text-[11px]">
            <thead class="bg-slate-100 text-slate-600">
              <tr>
                <th class="p-2 border border-slate-200 w-8">#</th>
                <th class="p-2 border border-slate-200">${currentDataType === 'MBA' ? 'Medical Type' : (currentDataType === 'CB' ? 'Expense Type' : 'FSA Item')}</th>
                <th class="p-2 border border-slate-200 text-center">${currentDataType === 'MBA' ? 'Category' : (currentDataType === 'CB' ? 'Expense Item' : 'Flex Year')}</th>
                <th class="p-2 border border-slate-200 text-right">Reimburse</th>
                <th class="p-2 border border-slate-200">${currentDataType === 'MBA' ? 'Memo' : (currentDataType === 'CB' ? 'Comments' : 'Description')}</th>
              </tr>
            </thead>
            <tbody>${psHtml}</tbody>
          </table>
        </div>
      </div>

      <!-- SAP Table -->
      <div>
        <h4 class="font-bold text-[10px] uppercase tracking-wider text-slate-700 mb-2 font-sans flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full bg-blue-500"></span>
          SAP Finance Data (${d.sData ? d.sData.lines.length : 0} G/L Lines)
        </h4>
        <div class="overflow-x-auto border border-slate-200 rounded-lg bg-white shadow-sm">
          <table class="w-full text-left text-[11px]">
            <thead class="bg-slate-100 text-slate-600">
              <tr>
                <th class="p-2 border border-slate-200 w-8">#</th>
                <th class="p-2 border border-slate-200">SAP Doc</th>
                <th class="p-2 border border-slate-200 text-right">GL Amount</th>
                <th class="p-2 border border-slate-200">Text / TEXT254</th>
                <th class="p-2 border border-slate-200 text-center">Type</th>
              </tr>
            </thead>
            <tbody>${sapHtml}</tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function closeModal() { document.getElementById('modal').classList.add('hidden'); }

// ============================================================
// EXCEL DOWNLOAD (REAL .xlsx)
// ============================================================
function downloadReport(source) {
  if (!reconData.length) { showToast('⚠️ No data to download.'); return; }

  const elPeriode = document.getElementById('inp-periode');
  const jenis   = currentDataType || 'MBA';
  const periode = ((elPeriode ? elPeriode.value : null) || 'Periode').replace(/\s/g, '_');
  const srcLabel = source === 'ps' ? 'PeopleSoft' : 'SAP';
  const filename = `Rekon_Fringe_Benefit_Tax_-${jenis}-_${srcLabel}_${periode}.xlsx`;

  showToast('Creating file ' + filename + '...');

  const wb = XLSX.utils.book_new();
  let sheetData = [];

  if (source === 'ps') {
    // Base PS: one row per original PS row, with recon columns appended
    const psHdrIdx = (psFileState?.profile?.headerIndex !== undefined) ? psFileState.profile.headerIndex : findHeaderRowIndex(rawPsRows);
    const psHeader = rawPsRows[psHdrIdx] || [];
    let ps_inv = currentDataType === 'FSA'
      ? findColMulti(psHeader, ['HMS_FSA_TRANS_ID', 'FSA_TRANS_ID'])
      : (currentDataType === 'CB'
          ? findColMulti(psHeader, ['ECS_RMB_SAP_REF', 'ECS_RMB_TRANSID'])
          : findColMulti(psHeader, ['HMS_INVOICE_NBR', 'INVOICE_NBR']));
    if (ps_inv < 0) ps_inv = findColMulti(psHeader, ['HMS_INVOICE_NBR', 'ECS_RMB_SAP_REF', 'ECS_RMB_TRANSID', 'HMS_FSA_TRANS_ID']);

    sheetData.push([
      ...psHeader,
      'REKON_STATUS', 'TOTAL_SAP_GL', 'DIFFERENCE_SAP_vs_PS', 'REMARKS'
    ]);
    for (let i = psHdrIdx + 1; i < rawPsRows.length; i++) {
      const row = rawPsRows[i];
      const inv = (ps_inv >= 0 && row) ? String(row[ps_inv] || '').trim() : '';
      const d   = reconData.find(r => r.invoice === inv);

      let status = '—', sapTotal = '—', selisih = '—', ket = '—';
      if (d) {
        status   = d.statusLabel;
        sapTotal = d.sapTotal;
        selisih  = d.diff;
        ket      = buildKetPS(d);
      }
      sheetData.push([...row, status, sapTotal, selisih, ket]);
    }
  } else {
    // Base SAP: one row per original SAP line (G/L row), with recon columns appended
    const sapHdrIdx = (sapFileState?.profile?.headerIndex !== undefined) ? sapFileState.profile.headerIndex : findHeaderRowIndex(rawSapRows);
    const sapHeader = rawSapRows[sapHdrIdx] || [];
    let sap_ref = currentDataType === 'FSA'
      ? findColMulti(sapHeader, ['HMS_ASSIGNMENT', 'REFERENCE'])
      : findColMulti(sapHeader, ['REFERENCE', 'HMS_ASSIGNMENT']);
    if (sap_ref < 0) sap_ref = findColMulti(sapHeader, ['REFERENCE', 'HMS_ASSIGNMENT', 'REFERENCE_NUMBER']);

    sheetData.push([
      ...sapHeader,
      'REKON_STATUS', 'TOTAL_PS_REIMB', 'DIFFERENCE_SAP_vs_PS', 'REMARKS'
    ]);
    for (let i = sapHdrIdx + 1; i < rawSapRows.length; i++) {
      const row = rawSapRows[i];
      const ref = (sap_ref >= 0 && row) ? String(row[sap_ref] || '').trim() : '';
      const d   = reconData.find(r => r.invoice === ref);

      let status = '—', psTotal = '—', selisih = '—', ket = '—';
      if (d) {
        status  = d.statusLabel;
        psTotal = d.psTotal;
        selisih = d.diff;
        ket     = buildKetSAP(d);
      }
      sheetData.push([...row, status, psTotal, selisih, ket]);
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  // Style header row
  const headerRange = XLSX.utils.decode_range(ws['!ref']);
  for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: col });
    if (!ws[addr]) continue;
    ws[addr].s = {
      fill: { patternType: 'solid', fgColor: { rgb: '0F2247' } },
      font: { bold: true, color: { rgb: 'FFFFFF' } }
    };
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Rekonsiliasi_' + srcLabel);
  XLSX.writeFile(wb, filename);
  showToast('✓ File downloaded successfully: ' + filename);
}

function downloadBothReports() {
  downloadReport('ps');
  setTimeout(() => downloadReport('sap'), 800);
}

function buildKetPS(d) {
  return buildKetSAP(d);
}

function buildKetSAP(d) {
  if (d.status === 'match') return 'Matched - Nominal dan NIK sesuai';
  if (d.statusLabel === 'SAP Only') return 'Data tidak ditemukan di PeopleSoft.';
  if (d.statusLabel === 'PSFT Only') return 'Data tidak ditemukan di SAP.';
  
  if (d.isNikMismatch) {
    return `Selisih NIK: PS (${d.pData.emplid}) vs SAP (${d.sData.persNo}). Pembayaran berpotensi salah sasaran!`;
  }

  // DETECTIVE LOGIC
  const diffAbs = Math.round(Math.abs(d.diff));
  
  // 1. Cek Text SAP
  if (d.sData && d.sData.lines) {
    for (const ln of d.sData.lines) {
      const txt = (ln.text || '').toLowerCase();
      const taxCode = (ln.taxCode || 'V0').toUpperCase();
      
      if (txt.includes('excess') || txt.includes('kelebihan') || txt.includes('limit') || txt.includes('plafon')) {
        if (Math.round(ln.glAmt) === diffAbs || d.dataType === 'MBA') {
          return `Terdapat potongan Limit/Excess sebesar Rp ${diffAbs.toLocaleString('id-ID')}`;
        }
      }
      if (txt.includes('tax') || txt.includes('pajak') || txt.includes('pph') || (taxCode !== 'V0' && taxCode !== '')) {
        if (Math.round(ln.glAmt) === diffAbs) {
          return `Terdapat potongan Pajak (Withholding Tax) sebesar Rp ${diffAbs.toLocaleString('id-ID')}`;
        }
      }
      if (txt.includes('kurs') || txt.includes('exc rate')) {
        return `Terdapat penyesuaian selisih kurs sebesar Rp ${diffAbs.toLocaleString('id-ID')}`;
      }
    }
  }

  // 2. Cek PS Partial Reject (khusus C&B/FSA)
  if ((d.dataType === 'CB' || d.dataType === 'FSA') && d.pData && d.pData.items) {
    for (const it of d.pData.items) {
      if (Math.round(it.reimb) === diffAbs) {
        const desc = d.dataType === 'CB' ? it.cat : it.medCd;
        return `Kuitansi untuk [${desc}] senilai Rp ${diffAbs.toLocaleString('id-ID')} kemungkinan ditolak (Rejected) atau tidak dibayarkan.`;
      }
    }
  }

  // 3. Fallback: Tampilkan text SAP
  let fallbackText = 'Selisih tidak dapat diidentifikasi secara otomatis.';
  if (d.sData && d.sData.lines.length > 0) {
    const uniqueTexts = [...new Set(d.sData.lines.map(x => x.text).filter(Boolean))];
    if (uniqueTexts.length > 0) {
      fallbackText = `Selisih Rp ${diffAbs.toLocaleString('id-ID')}. Indikasi SAP: ` + uniqueTexts.join(' | ');
    }
  }

  return fallbackText;
}

// ============================================================
// DEMO DATA LOADER
// ============================================================
function loadDemoData() {
  // PS Data: HMS_INVOICE_NBR, EMPLID, HMS_MEDICAL_CD, HMS_MBA_CLAIM_CAT, HMS_Reimburse_AMT, HMS_Memo_LTR_NO
  const psHeader = ['EMPLID','RECEIPT_DT','SEQNO','HMS_MEDICAL_CD','ORG_RECEIPT_DT','HMS_MED_ENTLT_PRD','HMS_MBA_CLAIM_CAT','DEPENDENT_BENEF','ACCTG_ENTRY_FLG','HMS_MED_GRANT_LTR','HMS_Memo_LTR_NO','HMS_RECEIPT_AMT','HMS_Reimburse_AMT','HMS_FORWARD_STATUS','HMS_INVOICE_NBR','HMS_ACC_DATE'];
  rawPsRows = [psHeader,
    ['80434680','2023-08-31','0','0002','2023-08-15','2023','D','04','Y','','RANAP MITRA KELUARGA','7735273.00','7735273.00','ACTG','230901MB671076','2023-09-01'],
    ['80415263','2023-08-31','0','0001','2023-08-21','2023','D','04','Y','','','798600.00','798600.00','ACTG','230901MB671079','2023-09-01'],
    ['80431379','2023-08-31','0','0001','2023-08-29','2023','D','02','Y','','','440117.00','440117.00','ACTG','230901MB671085','2023-09-01'],
    ['80008599','2023-08-31','1','0002','2023-08-23','2023','E','','Y','','RANAP RS WAVA HUSADA','44325813.00','44325813.00','ACTG','230901MB671099','2023-09-01'],
    ['80012422','2023-08-31','0','0004','2023-08-30','2023','D','02','Y','','','200000.00','200000.00','ACTG','230901MB671136','2023-09-01'],
    ['80012439','2023-08-31','0','0001','2023-08-25','2023','D','01','Y','','','290000.00','290000.00','ACTG','230901MB671137','2023-09-01'],
    ['80001234','2023-08-31','0','0001','2023-08-20','2023','E','','Y','','','1200000.00','1200000.00','ACTG','230901MB671200','2023-09-01'],  // Only PS
  ];

  // SAP Data: REFERENCE, HMS_PERS_NO, HMS_GL_AMOUNT, TEXT254, HMS_DOC_NBR
  const sapHeader = ['HMS_FBT_SEQNBR','HMS_COMPANY_CD','HMS_DOC_NBR','CURRENCY_CD','DOC_TYPE','CREATED_BY_USER','HMS_HEADER_TEXT','REFERENCE','DOCUMENT_DT','POSTING_DT','HMS_SUPPL_ACT','HMS_SUP_POST_KEY','HMS_SUPL_AMOUNT','HMS_ASSIGNMENT','HMS_PERS_NO','HMS_PAYMENT_BLOCK','HMS_CLEAR_DT','HMS_CLEAR_DOC','HMSA_ACCT_NO','HMS_GL_POST_KEY','HMS_COST_CENTER','HMS_GL_AMOUNT','TAX_CODE','TEXT254'];
  rawSapRows = [sapHeader,
    ['1001','1616','2500001','IDR','XU','ID-BATCH','PMB-434680','230901MB671076','01-SEP-2023','02-SEP-2023','ID00434680','31','7735273','00000000','80434680','','','','80320100','40','1616315000','7735273.00','V0','Ranap Mitra Keluarga'],
    ['1002','1616','2500002','IDR','XU','ID-BATCH','PMB-415263','230901MB671079','01-SEP-2023','02-SEP-2023','ID00415263','31','798600','00000000','80415263','','','','80320100','40','1616315000','798600.00','V0','Outpatient'],
    ['1003','1616','2500003','IDR','XU','ID-BATCH','PMB-431379','230901MB671085','01-SEP-2023','02-SEP-2023','ID00431379','31','440117','00000000','80431379','','','','80320100','40','1616315000','440117.00','V0','Outpatient'],
    ['1004','1616','2500004','IDR','XU','ID-BATCH','PMB-008599-RANAP','230901MB671099','01-SEP-2023','02-SEP-2023','ID00008599','31','44325813','00000000','80008599','','','','80320100','40','1616315000','44325813.00','V0','Ranap RS Wava Husada'],
    ['1005','1616','2500004B','IDR','XU','ID-BATCH','PMB-008599-EXCESS','230901MB671099','01-SEP-2023','02-SEP-2023','ID00008599','31','44325813','00000000','80008599','','','','80320100','40','1616315000','3200000.00','V0','Excess MBA-80008599'], // Excess line
    ['1006','1616','2500005','IDR','XU','ID-BATCH','PMB-012422','230901MB671136','01-SEP-2023','02-SEP-2023','ID00012422','31','200000','00000000','80012422','','','','80320100','40','1616315000','200000.00','V0','Glasses'],
    ['1007','1616','2500006','IDR','XU','ID-BATCH','PMB-012439','230901MB671137','01-SEP-2023','02-SEP-2023','ID00012439','31','450000','00000000','80012439','','','','80320100','40','1616315000','450000.00','V0','Outpatient'], // SAP > PS (mismatch)
    ['1008','1616','2500007','IDR','XU','ID-BATCH','PMB-SAP-ONLY','230901MB999999','01-SEP-2023','02-SEP-2023','ID00099999','31','500000','00000000','80099999','','','','80320100','40','1616315000','500000.00','V0','Jurnal Manual AP'],  // Only SAP
  ];

  handleParsedRows(rawPsRows, { name: 'Demo_PeopleSoft.csv', size: 1024 }, 'ps', 'Demo');
  handleParsedRows(rawSapRows, { name: 'Demo_SAP.csv', size: 1024 }, 'sap', 'Demo');
  showToast('✓ Sample data loaded successfully. Click "Compare & Reconcile".');
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function findCol(headers, name, fallback = -1) {
  if (!Array.isArray(headers)) return fallback;
  const target = normalizeHeaderName(name);
  const idx = headers.findIndex(h => normalizeHeaderName(h) === target);
  return idx >= 0 ? idx : fallback;
}
function findColMulti(headers, names, fallback = -1) {
  if (!Array.isArray(headers)) return fallback;
  const targets = names.map(normalizeHeaderName);
  for (const target of targets) {
    const idx = headers.findIndex(h => normalizeHeaderName(h) === target);
    if (idx >= 0) return idx;
  }
  return fallback;
}
function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  return parseFloat(String(val).replace(/,/g, '')) || 0;
}
function formatRp(val) {
  if (val >= 1e9) return (val / 1e9).toFixed(2) + ' M';
  if (val >= 1e6) return (val / 1e6).toFixed(2) + ' Jt';
  return Math.round(val).toLocaleString('id-ID');
}
function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}
function medLabel(code) {
  const map = { '0001': 'Outpatient', '0002': 'Inpatient', '0003': 'Dental', '0004': 'Glasses', '9001': 'Vitamin' };
  return map[code] || code;
}
function resetAll() {
  rawPsRows = []; rawSapRows = []; reconData = []; filteredData = [];
  psFileState = null; sapFileState = null;
  psMap.clear(); sapMap.clear();
  currentPage = 1;
  pageSize = 10;
  const selPageSize = document.getElementById('select-page-size');
  if (selPageSize) selPageSize.value = '10';

  const exportBtn = document.getElementById('btn-export-all');
  if (exportBtn) exportBtn.disabled = true;

  document.getElementById('fi-ps').value = '';
  document.getElementById('fi-sap').value = '';

  resetDropzoneUI('ps');
  resetDropzoneUI('sap');

  checkReady();
  switchTab('upload');
  showToast('🔄 Reset. Ready to process new files.');
}

// ---- TOAST ----
let toastTimer;
function showToast(msg) {
  const t = document.getElementById('toast');
  document.getElementById('toast-text').textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}

// ---- INIT ----
document.addEventListener('DOMContentLoaded', () => {
  switchTab('upload');
  // Sync period labels when inputs change
  const elP = document.getElementById('inp-periode');
  if (elP) {
    elP.addEventListener('input', () => {
      const periode = elP.value || 'Periode';
      const sumPeriod = document.getElementById('sum-Period');
      if (sumPeriod) sumPeriod.textContent = periode;
      ['ps','sap'].forEach(t => {
        const el_p = document.getElementById('dl-Period-' + t);
        if (el_p) el_p.textContent = periode;
      });
    });
  }

  // Close modal on ESC key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
  });
});

