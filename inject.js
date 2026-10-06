const fs = require('fs');

let appJS = fs.readFileSync('app.js', 'utf8');

// Replace dummy data
const dummyDataStart = appJS.indexOf('const sapData = [');
const dummyDataEnd = appJS.indexOf('/* ===== POPULATE TABLES ===== */');
appJS = appJS.substring(0, dummyDataStart) + 
`let sapData = [];
let psData = [];

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(file);
  });
}

const fmtIDR = (val) => {
    if (!val || isNaN(val)) return '0';
    return Number(val).toLocaleString('en-US');
}

async function doRealReconciliation() {
  try {
    const sapBuf = await readFileAsArrayBuffer(selectedFile);
    const sapWb = XLSX.read(sapBuf, {type: 'array'});
    const sapSheet = sapWb.Sheets[sapWb.SheetNames[0]];
    const sapJsonRaw = XLSX.utils.sheet_to_json(sapSheet, {header: 1, defval: ''});
    
    let sapHeaders = sapJsonRaw[0];
    let sapRows = sapJsonRaw.slice(1);
    if (!sapHeaders.includes('HMS_PERS_NO') && sapJsonRaw.length > 1) {
        sapHeaders = sapJsonRaw[1];
        sapRows = sapJsonRaw.slice(2);
    }
    
    const psBuf = await readFileAsArrayBuffer(psFile);
    const psWb = XLSX.read(psBuf, {type: 'array'});
    const psSheet = psWb.Sheets[psWb.SheetNames[0]];
    const psJsonRaw = XLSX.utils.sheet_to_json(psSheet, {header: 1, defval: ''});
    const psHeaders = psJsonRaw[0];
    const psRows = psJsonRaw.slice(1);

    const sapEmpIdx = sapHeaders.indexOf('HMS_PERS_NO');
    const sapAmtIdx = sapHeaders.indexOf('HMS_SUPL_AMOUNT (Pake ini)');
    const sapTypeIdx = sapHeaders.indexOf('HMS_FBT_TYPE');
    
    const psEmpIdx = psHeaders.indexOf('EMPLID');
    const psAmtIdx = psHeaders.indexOf('Total Payment PSFT');
    const psTypeIdx = psHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_MEDICAL_CD'));
    
    let parsedSAP = sapRows.filter(r => r[sapEmpIdx]).map(row => ({
      id: String(row[sapEmpIdx] || '').trim(),
      name: 'Pegawai ' + String(row[sapEmpIdx] || '').trim(),
      period: 'Aug-26',
      type: row[sapTypeIdx] || 'Medical',
      amountNum: parseFloat(row[sapAmtIdx]) || 0,
      amount: fmtIDR(row[sapAmtIdx]),
      raw: row
    }));

    let parsedPS = psRows.filter(r => r[psEmpIdx]).map(row => ({
      id: String(row[psEmpIdx] || '').trim(),
      name: 'Pegawai ' + String(row[psEmpIdx] || '').trim(),
      period: 'Aug-26',
      type: row[psTypeIdx] || 'Medical',
      amountNum: parseFloat(row[psAmtIdx]) || 0,
      amount: fmtIDR(row[psAmtIdx]),
      raw: row
    }));

    sapData = parsedSAP.map(s => {
      let match = parsedPS.find(p => p.id === s.id);
      let diff = 0, status = 'unmatched', ket = 'Not Found in PS', diffStr = '—';
      if (match) {
          diff = s.amountNum - match.amountNum;
          if (diff === 0) { status = 'matched'; ket = 'Matched dengan PS'; diffStr = '0'; }
          else { status = 'unmatched'; ket = 'Amount Mismatch: SAP beda ' + fmtIDR(Math.abs(diff)); diffStr = diff > 0 ? '+' + fmtIDR(diff) : fmtIDR(diff); }
      }
      return { ...s, diff: diffStr, status, ket };
    });

    psData = parsedPS.map(p => {
      let match = parsedSAP.find(s => s.id === p.id);
      let diff = 0, status = 'unmatched', ket = 'Not Found in SAP', diffStr = '—';
      if (match) {
          diff = p.amountNum - match.amountNum;
          if (diff === 0) { status = 'matched'; ket = 'Matched dengan SAP'; diffStr = '0'; }
          else { status = 'unmatched'; ket = 'Amount Mismatch: PS beda ' + fmtIDR(Math.abs(diff)); diffStr = diff > 0 ? '+' + fmtIDR(diff) : fmtIDR(diff); }
      }
      return { ...p, diff: diffStr, status, ket };
    });
    
    return true;
  } catch (e) {
    console.error(e);
    return false;
  }
}

` + appJS.substring(dummyDataEnd);

// Replace runProcessing to await doRealReconciliation
appJS = appJS.replace('function runProcessing() {', 'async function runProcessing() {');
appJS = appJS.replace('tick();', 'await doRealReconciliation(); tick();');
appJS = appJS.replace('setTimeout(tick, 950);', 'setTimeout(tick, 200);'); // speed up animation

// Replace simulateDownload to use SheetJS real export
const dlFn = `function simulateDownload(filename) {
      showToast('Downloading: ' + filename);
      const isSAP = filename.includes('SAP');
      const data = isSAP ? sapData : psData;
      
      const ws = XLSX.utils.json_to_sheet(data.map(d => ({
        'Emp ID': d.id,
        'Name': d.name,
        'Period': d.period,
        'Claim Type': d.type,
        'Amount': d.amountNum,
        'Selisih': d.diff,
        'Status / Keterangan': d.ket
      })));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Reconciliation");
      XLSX.writeFile(wb, filename);
    }`;

const oldDlStart = appJS.indexOf('function simulateDownload');
const oldDlEnd = appJS.indexOf('}', oldDlStart) + 1;
appJS = appJS.substring(0, oldDlStart) + dlFn + appJS.substring(oldDlEnd);

fs.writeFileSync('app.js', appJS);
