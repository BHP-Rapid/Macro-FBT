// ============================================================
// SMART FBT — app.js
// Automated Benefit Tax Reconciliation & Upload Tool
// Phase 1: SAP vs PeopleSoft Reconciliation Engine
// ============================================================

// ---- STATE ----
let currentDataType = 'MBA';
let rawPsRows  = [];   // Raw parsed rows from PeopleSoft file
let rawSapRows = [];   // Raw parsed rows from SAP file
let reconData  = [];   // Final reconciliation result array
let filteredData = []; // Data after filter/search
let currentFilter = 'all';
let searchQuery   = '';
let currentPage   = 1;
const PAGE_SIZE   = 15;
let sortCol = 'none';
let sortAsc = true;

// ---- MAPS ----
let psMap  = new Map(); // key: HMS_INVOICE_NBR -> { emplid, totalReimb, items[] }
let sapMap = new Map(); // key: REFERENCE      -> { persNo, totalGL, lines[] }

// ============================================================
// TAB / PANEL NAVIGATION
// ============================================================
function switchTab(tab) {
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

    if (i < currentStep) {
      // COMPLETED — green circle + checkmark
      item.className   = 'stepper-item completed';
      circle.innerHTML = checkSVG;
    } else if (i === currentStep) {
      // ACTIVE — navy blue circle with number
      item.className   = 'stepper-item active';
      circle.innerHTML = '<span>' + i + '</span>';
    } else {
      // PENDING — gray circle with number
      item.className   = 'stepper-item';
      circle.innerHTML = '<span>' + i + '</span>';
    }
  }

  // Update subtitles dynamically
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
  let headers, rows, sheetName, fileName;

  if (type === 'sap') {
    headers = ['HMS_FBT_SEQNBR','HMS_COMPANY_CD','HMS_DOC_NBR','CURRENCY_CD','DOC_TYPE',
               'CREATED_BY_USER','HMS_HEADER_TEXT','REFERENCE','DOCUMENT_DT','POSTING_DT',
               'HMS_SUPPL_ACT','HMS_SUP_POST_KEY','HMS_SUPL_AMOUNT','HMS_ASSIGNMENT',
               'HMS_PERS_NO','HMS_PAYMENT_BLOCK','HMS_CLEAR_DT','HMS_CLEAR_DOC',
               'HMSA_ACCT_NO','HMS_GL_POST_KEY','HMS_COST_CENTER','HMS_GL_AMOUNT','TAX_CODE','TEXT254'];
    rows = [
      ['1001','1616','2500001','IDR','XU','ID-BATCH','PMB-434680','230901MB671076','01-SEP-2026','02-SEP-2026','ID00434680','31','7735273','00000000','80434680','','','','80320100','40','1616315000','7735273','V0','Ranap Mitra Keluarga'],
      ['1002','1616','2500002','IDR','XU','ID-BATCH','PMB-415263','230901MB671079','01-SEP-2026','02-SEP-2026','ID00415263','31','798600','00000000','80415263','','','','80320100','40','1616315000','798600','V0','Outpatient'],
      ['1003','1616','2500003','IDR','XU','ID-BATCH','PMB-431379','230901MB671085','01-SEP-2026','02-SEP-2026','ID00431379','31','440117','00000000','80431379','','','','80320100','40','1616315000','440117','V0','Outpatient'],
      ['1004','1616','2500004','IDR','XU','ID-BATCH','PMB-008599','230901MB671099','01-SEP-2026','02-SEP-2026','ID00008599','31','44325813','00000000','80008599','','','','80320100','40','1616315000','44325813','V0','Ranap RS Wava Husada'],
      ['1005','1616','2500005','IDR','XU','ID-BATCH','PMB-012422','230901MB671136','01-SEP-2026','02-SEP-2026','ID00012422','31','200000','00000000','80012422','','','','80320100','40','1616315000','200000','V0','Glasses'],
    ];
    sheetName = 'SAP_Template';
    fileName  = 'HDC_FBT_SAP_MBA_10Agt_9Sept26_TEMPLATE.xlsx';
  } else {
    headers = ['EMPLID','RECEIPT_DT','SEQNO','HMS_MEDICAL_CD','ORG_RECEIPT_DT','HMS_MED_ENTLT_PRD',
               'HMS_MBA_CLAIM_CAT','DEPENDENT_BENEF','ACCTG_ENTRY_FLG','HMS_MED_GRANT_LTR',
               'HMS_Memo_LTR_NO','HMS_RECEIPT_AMT','HMS_Reimburse_AMT','HMS_FORWARD_STATUS',
               'HMS_INVOICE_NBR','HMS_ACC_DATE'];
    rows = [
      ['80434680','2026-08-31','0','0002','2026-08-15','2026','D','04','Y','','RANAP MITRA KELUARGA','7735273','7735273','ACTG','230901MB671076','2026-09-01'],
      ['80415263','2026-08-31','0','0001','2026-08-21','2026','D','04','Y','','','798600','798600','ACTG','230901MB671079','2026-09-01'],
      ['80431379','2026-08-31','0','0001','2026-08-29','2026','D','02','Y','','','440117','440117','ACTG','230901MB671085','2026-09-01'],
      ['80008599','2026-08-31','1','0002','2026-08-23','2026','E','','Y','','RANAP RS WAVA HUSADA','44325813','44325813','ACTG','230901MB671099','2026-09-01'],
      ['80012422','2026-08-31','0','0004','2026-08-30','2026','D','02','Y','','','200000','200000','ACTG','230901MB671136','2026-09-01'],
    ];
    sheetName = 'PeopleSoft_Template';
    fileName  = 'HDC_FBT_PS_MBA_10Agt_9Sept26_TEMPLATE.xlsx';
  }

  const data = [headers, ...rows];
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
// FILE NAME VALIDATION
// ============================================================
function validateFileName(fileName, type) {
  // Expected: HDC_FBT_SAP_[TYPE]_[PERIOD].ext  or  HDC_FBT_PS_[TYPE]_[PERIOD].ext
  const prefix = type === 'sap' ? /^HDC_FBT_SAP_/i : /^HDC_FBT_PS_/i;
  const ext = fileName.match(/\.([a-z]+)$/i);
  const validExt = ext && ['xlsx','xls','csv'].includes(ext[1].toLowerCase());
  const validName = prefix.test(fileName);
  return { validName, validExt, ok: validName && validExt };
}

function showFileValidation(type, fileName) {
  const dz  = document.getElementById('dz-' + type);
  const msg = document.getElementById(type + '-validation-msg');
  if (!dz || !msg) return;

  const result = validateFileName(fileName, type);
  dz.classList.remove('dz-valid','dz-invalid','border-blue-300','bg-blue-50/40');
  msg.classList.remove('hidden','text-emerald-600','text-rose-600');

  if (result.ok) {
    dz.classList.add('dz-valid');
    msg.className  = 'mt-2 text-[11px] text-emerald-700 font-semibold flex items-center gap-1';
    msg.innerHTML  = '✅ Nama file sesuai format.';
  } else {
    dz.classList.add('dz-invalid');
    msg.className  = 'mt-2 text-[11px] text-rose-700 font-semibold flex items-center gap-1 flex-wrap';
    const expectedPrefix = type === 'sap' ? 'HDC_FBT_SAP_' : 'HDC_FBT_PS_';
    msg.innerHTML  = `⚠️ Nama file tidak sesuai format. Format yang diharapkan: <span class="font-mono bg-rose-100 px-1 rounded">${expectedPrefix}[TYPE]_[PERIOD].xlsx</span>`;
  }
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
  const sizeFmt = (file.size / 1024 / 1024).toFixed(2) + ' MB';
  const validRows = Math.max(0, rows.length - 1);

  // Show file name validation
  showFileValidation(type, file.name);

  if (type === 'ps') {
    rawPsRows = rows;
    document.getElementById('ps-fname').textContent = file.name;
    document.getElementById('ps-fdesc').textContent = `${fmt} · ${sizeFmt} · ${validRows.toLocaleString('id-ID')} rows`;
    document.getElementById('ps-badge').innerHTML =
      `<span class="text-emerald-700 font-semibold">✓ ${validRows.toLocaleString('id-ID')} PeopleSoft Rows Ready</span>`;
  } else {
    rawSapRows = rows;
    document.getElementById('sap-fname').textContent = file.name;
    document.getElementById('sap-fdesc').textContent = `${fmt} · ${sizeFmt} · ${validRows.toLocaleString('id-ID')} rows`;
    document.getElementById('sap-badge').innerHTML =
      `<span class="text-blue-700 font-semibold">✓ ${validRows.toLocaleString('id-ID')} SAP Rows Ready</span>`;
  }
  checkReady();
  showToast('✓ File loaded successfully: ' + file.name);
}

function checkReady() {
  const ready = rawPsRows.length > 0 && rawSapRows.length > 0;
  const btn = document.getElementById('btn-process');
  const dot = document.getElementById('status-dot');
  const txt = document.getElementById('status-text');

  btn.disabled = !ready;
  if (ready) {
    dot.className = 'w-2 h-2 rounded-full bg-emerald-500 pulse-dot';
    txt.textContent = 'Kedua file siap! Klik tombol "Bandingkan & Rekonsiliasi" untuk memulai.';
    txt.className = 'text-emerald-700 font-semibold text-xs';
  } else if (rawPsRows.length > 0) {
    dot.className = 'w-2 h-2 rounded-full bg-amber-500';
    txt.textContent = 'PeopleSoft file ready. Waiting for SAP file...';
    txt.className = 'text-amber-700 text-xs';
  } else if (rawSapRows.length > 0) {
    dot.className = 'w-2 h-2 rounded-full bg-amber-500';
    txt.textContent = 'SAP file ready. Waiting for PeopleSoft file...';
    txt.className = 'text-amber-700 text-xs';
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
  const psHeader = rawPsRows[0] || [];

  const psHeaderStr = psHeader.join(' ');
  if (psHeaderStr.includes('HMS_FSA_TRANS_ID')) {
    currentDataType = 'FSA';
  } else if (psHeaderStr.includes('ECS_RMB_SAP_REF')) {
    currentDataType = 'CB';
  } else {
    currentDataType = 'MBA';
  }

  let ps_inv, ps_emp, ps_reimb, ps_med, ps_Memo, ps_cat, ps_date, ps_entitle;
  if (currentDataType === 'FSA') {
    ps_inv   = findCol(psHeader, 'HMS_FSA_TRANS_ID',  0);
    ps_emp   = findCol(psHeader, 'EMPLID',            1);
    ps_reimb = findCol(psHeader, 'HMS_FLX_UNIT',      8);
    ps_med   = findCol(psHeader, 'HMS_FSA_ITEM',      7);
    ps_Memo  = findCol(psHeader, 'DESCRLONG',         4);
    ps_cat   = findCol(psHeader, 'HMS_FLX_YEAR',      2);
    ps_date  = findCol(psHeader, 'RECEIPT_DT',        6);
    ps_entitle = findCol(psHeader, 'HMS_FLX_YEAR',    2); // Dummy for FSA
  } else if (currentDataType === 'CB') {
    ps_inv   = findCol(psHeader, 'ECS_RMB_SAP_REF',   6);
    ps_emp   = findCol(psHeader, 'EMPLID',            1);
    ps_reimb = findCol(psHeader, 'ECS_AMT_RMB',       12);
    ps_med   = findCol(psHeader, 'EXPENSE_TYPE',      3);
    ps_Memo  = findCol(psHeader, 'COMMENTS',          13);
    ps_cat   = findCol(psHeader, 'EXPENSE_ITEM',      11);
    ps_date  = findCol(psHeader, 'RECEIPT_DT',        9);
    ps_entitle = findCol(psHeader, 'RECEIPT_DT',      9); // Dummy for CB
  } else {
    ps_inv   = findCol(psHeader, 'HMS_INVOICE_NBR',   14);
    ps_emp   = findCol(psHeader, 'EMPLID',            0);
    ps_reimb = findCol(psHeader, 'HMS_Reimburse_AMT', 12);
    ps_med   = findCol(psHeader, 'HMS_MEDICAL_CD',    3);
    ps_Memo  = findCol(psHeader, 'HMS_Memo_LTR_NO',   10);
    ps_cat   = findCol(psHeader, 'HMS_MBA_CLAIM_CAT', 6);
    ps_date  = findCol(psHeader, 'RECEIPT_DT',        1);
    ps_entitle = findCol(psHeader, 'HMS_MED_ENTLT_PRD', 5);
  }

  for (let i = 1; i < rawPsRows.length; i++) {
    const row = rawPsRows[i];
    if (!row || row.length <= ps_inv) continue;
    const inv = String(row[ps_inv] || '').trim();
    if (!inv) continue;

    const reimb = parseNum(row[ps_reimb]);

    if (!psMap.has(inv)) {
      psMap.set(inv, {
        inv, emplid: String(row[ps_emp] || '').trim(),
        totalReimb: 0, items: [],
        medCd: String(row[ps_med] || ''), claimCat: String(row[ps_cat] || '')
      });
    }
    const obj = psMap.get(inv);
    obj.totalReimb += reimb;
    obj.items.push({
      row, reimb,
      Memo   : String(row[ps_Memo] || ''),
      medCd  : String(row[ps_med]  || ''),
      cat    : String(row[ps_cat]  || ''),
      date   : String(row[ps_date] || ''),
      entitle: String(row[ps_entitle] || '')
    });
  }

  // --- Process SAP ---
  sapMap.clear();
  let sapHdrIdx = 0;
  if (rawSapRows[1] && String(rawSapRows[1]).includes('HMS_FBT_SEQNBR')) sapHdrIdx = 1;
  const sapHeader = rawSapRows[sapHdrIdx] || [];

  const sap_ref  = currentDataType === 'FSA' ? findCol(sapHeader, 'HMS_ASSIGNMENT', 13) : findCol(sapHeader, 'REFERENCE', 7);
  const sap_gl   = findCol(sapHeader, 'HMS_GL_AMOUNT', 21);
  const sap_pers = findCol(sapHeader, 'HMS_PERS_NO',   14);
  const sap_doc  = findCol(sapHeader, 'HMS_DOC_NBR',   2);
  const sap_txt  = findColMulti(sapHeader, ['TEXT254','HMS_HEADER_TEXT'], 23);
  const sap_tax  = findCol(sapHeader, 'TAX_CODE', 22);

  for (let i = sapHdrIdx + 1; i < rawSapRows.length; i++) {
    const row = rawSapRows[i];
    if (!row || row.length <= sap_ref) continue;
    const ref = String(row[sap_ref] || '').trim();
    if (!ref) continue;

    const glAmt = parseNum(row[sap_gl]);

    if (!sapMap.has(ref)) {
      sapMap.set(ref, {
        ref, persNo: String(row[sap_pers] || '').trim(),
        docNbr: String(row[sap_doc] || ''),
        totalGL: 0, lines: []
      });
    }
    const obj = sapMap.get(ref);
    obj.totalGL += glAmt;
    obj.lines.push({ 
      row, 
      glAmt, 
      text: String(row[sap_txt] || ''), 
      docNbr: String(row[sap_doc] || ''),
      taxCode: String(row[sap_tax] || '')
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
  setText('fin-total-diff', (grandSAP - grandPS >= 0 ? '+' : '') + 'IDR ' + Math.round(grandSAP - grandPS).toLocaleString('id-ID'));

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
    updatePagination(0);
    return;
  }

  const totalPages = Math.ceil(filteredData.length / PAGE_SIZE);
  currentPage = Math.max(1, Math.min(currentPage, totalPages));
  const start = (currentPage - 1) * PAGE_SIZE;
  const end   = Math.min(start + PAGE_SIZE, filteredData.length);
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

// Helper: get column value from a raw row by column name
function getColVal(rawRows, row, colName, fallbackIdx) {
  const hdr = rawRows[0] || [];
  const idx = findCol(hdr, colName, fallbackIdx);
  return String(row[idx] || '').trim() || '—';
}

function updatePagination(total, totalPages, start, end) {
  setText('pg-start', total === 0 ? '0' : start + 1);
  setText('pg-end',   end || 0);
  setText('pg-total', total);
  setText('pg-info',  `Page ${currentPage || 1}/${totalPages || 1}`);
  const prevBtn = document.getElementById('btn-prev');
  const nextBtn = document.getElementById('btn-next');
  if (prevBtn) prevBtn.disabled = currentPage <= 1;
  if (nextBtn) nextBtn.disabled = currentPage >= (totalPages || 1);
}

function prevPage() { if (currentPage > 1) { currentPage--; renderTable(); } }
function nextPage() {
  const totalPages = Math.ceil(filteredData.length / PAGE_SIZE);
  if (currentPage < totalPages) { currentPage++; renderTable(); }
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

  const elJenis = document.getElementById('inp-jenis');
  const elPeriode = document.getElementById('inp-periode');
  const jenis   = ((elJenis ? elJenis.value : null) || 'MBA').replace(/\s/g, '_');
  const periode = ((elPeriode ? elPeriode.value : null) || '10Agt_9Sept26').replace(/\s/g, '_');
  const srcLabel = source === 'ps' ? 'PeopleSoft' : 'SAP';
  const filename = `Rekon_Fringe_Benefit_Tax_-${jenis}-_${srcLabel}_${periode}.xlsx`;

  showToast('Creating file ' + filename + '...');

  const wb = XLSX.utils.book_new();
  let sheetData = [];

  if (source === 'ps') {
    // Base PS: one row per original PS row, with recon columns appended
    sheetData.push([
      ...rawPsRows[0],
      'REKON_STATUS', 'TOTAL_SAP_GL', 'DIFFERENCE_SAP_vs_PS', 'REMARKS'
    ]);
    for (let i = 1; i < rawPsRows.length; i++) {
      const row = rawPsRows[i];
      const psHeader = rawPsRows[0];
      const ps_inv = findCol(psHeader, 'HMS_INVOICE_NBR', 14);
      const inv = String(row[ps_inv] || '').trim();
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
    const sapHdrIdx = rawSapRows[1] && String(rawSapRows[1]).includes('HMS_FBT_SEQNBR') ? 1 : 0;
    sheetData.push([
      ...rawSapRows[sapHdrIdx],
      'REKON_STATUS', 'TOTAL_PS_REIMB', 'DIFFERENCE_SAP_vs_PS', 'REMARKS'
    ]);
    for (let i = sapHdrIdx + 1; i < rawSapRows.length; i++) {
      const row = rawSapRows[i];
      const sapHeader = rawSapRows[sapHdrIdx];
      const sap_ref = findCol(sapHeader, 'REFERENCE', 7);
      const ref = String(row[sap_ref] || '').trim();
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
function findCol(headers, name, fallback) {
  const nameUpper = name.toUpperCase();
  const idx = headers.findIndex(h => typeof h === 'string' && h.trim().toUpperCase() === nameUpper);
  return idx >= 0 ? idx : fallback;
}
function findColMulti(headers, names, fallback) {
  const namesUpper = names.map(n => n.toUpperCase());
  for (const nameUpper of namesUpper) {
    const idx = headers.findIndex(h => typeof h === 'string' && h.trim().toUpperCase() === nameUpper);
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
  psMap.clear(); sapMap.clear();
  document.getElementById('ps-fname').textContent = 'Click or Drag PeopleSoft File';
  document.getElementById('ps-fdesc').innerHTML = 'Format: <span class="font-bold text-blue-700">HDC_FBT_PS_[TYPE]_[PERIOD].xlsx</span>';
  document.getElementById('sap-fname').textContent = 'Click or Drag SAP File';
  document.getElementById('sap-fdesc').innerHTML = 'Format: <span class="font-bold text-blue-700">HDC_FBT_SAP_[TYPE]_[PERIOD].xlsx</span>';
  document.getElementById('ps-badge').innerHTML = '<span>Waiting for PeopleSoft file...</span>';
  document.getElementById('sap-badge').innerHTML = '<span>Waiting for SAP file...</span>';
  document.getElementById('fi-ps').value = '';
  document.getElementById('fi-sap').value = '';
  // Reset validation
  ['ps','sap'].forEach(t => {
    const dz = document.getElementById('dz-' + t);
    const msg = document.getElementById(t + '-validation-msg');
    if (dz) dz.classList.remove('dz-valid','dz-invalid');
    if (msg) { msg.textContent = ''; msg.classList.add('hidden'); }
  });
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
  // Sync period/jenis labels when inputs change
  ['inp-jenis','inp-periode'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => {
        const elJ = document.getElementById('inp-jenis');
        const elP = document.getElementById('inp-periode');
        const jenis   = (elJ ? elJ.value : null) || 'MBA';
        const periode = (elP ? elP.value : null) || '10Agt_9Sept26';
        ['ps','sap'].forEach(t => {
          const el_j = document.getElementById('dl-jenis-' + t);
          const el_p = document.getElementById('dl-periode-' + t);
          if (el_j) el_j.textContent = jenis;
          if (el_p) el_p.textContent = periode;
        });
      });
    }
  });

  // Close modal on ESC key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
  });
});

