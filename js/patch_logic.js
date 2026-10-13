const fs = require('fs');

let appJS = fs.readFileSync('app.js', 'utf8');

const regex = /async function doRealReconciliation\(\) \{[\s\S]*?return true;\s*\} catch \(e\) \{[\s\S]*?return false;\s*\}/;

const newLogic = `async function doRealReconciliation() {
  try {
    const sapBuf = await readFileAsArrayBuffer(selectedFile);
    const sapWb = XLSX.read(sapBuf, {type: 'array'});
    const sapSheet = sapWb.Sheets[sapWb.SheetNames[0]];
    const sapJsonRaw = XLSX.utils.sheet_to_json(sapSheet, {header: 1, defval: ''});
    
    // SAP Data
    let sapHeaders = sapJsonRaw[0] || [];
    let sapRows = sapJsonRaw.slice(1);
    if (!sapHeaders.includes('REFERENCE') && sapJsonRaw.length > 1) {
        sapHeaders = sapJsonRaw[1];
        sapRows = sapJsonRaw.slice(2);
    }
    
    const psBuf = await readFileAsArrayBuffer(psFile);
    const psWb = XLSX.read(psBuf, {type: 'array'});
    const psSheet = psWb.Sheets[psWb.SheetNames[0]];
    const psJsonRaw = XLSX.utils.sheet_to_json(psSheet, {header: 1, defval: ''});
    
    let psHeaders = psJsonRaw[0] || [];
    let psRows = psJsonRaw.slice(1);

    // Indexes PS
    const psInvIdx = psHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_INVOICE_NBR'));
    const psEmpIdx = psHeaders.findIndex(h => typeof h === 'string' && h.includes('EMPLID'));
    const psAmtIdx = psHeaders.findIndex(h => typeof h === 'string' && (h.includes('HMS_REIMBURSE_AMT') || h.includes('Total Payment')));
    const psMedIdx = psHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_MEDICAL_CD'));

    // Indexes SAP
    const sapRefIdx = sapHeaders.findIndex(h => typeof h === 'string' && h.includes('REFERENCE'));
    const sapEmpIdx = sapHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_PERS_NO'));
    const sapGlAmtIdx = sapHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_GL_AMOUNT'));
    const sapTypeIdx = sapHeaders.findIndex(h => typeof h === 'string' && h.includes('HMS_FBT_TYPE'));

    if (psInvIdx === -1 || sapRefIdx === -1) {
        throw new Error("Kolom HMS_INVOICE_NBR atau REFERENCE tidak ditemukan!");
    }

    // 1. Group PS by Invoice
    let psMap = new Map();
    for(let row of psRows) {
        const inv = String(row[psInvIdx] || '').trim();
        if(!inv) continue;
        const amt = parseFloat(row[psAmtIdx]) || 0;
        
        if(!psMap.has(inv)) {
            psMap.set(inv, {
                inv: inv,
                empId: row[psEmpIdx] || '',
                type: row[psMedIdx] || '',
                totalAmt: 0
            });
        }
        psMap.get(inv).totalAmt += amt;
    }

    // 2. Group SAP by Reference
    let sapMap = new Map();
    for(let row of sapRows) {
        const ref = String(row[sapRefIdx] || '').trim();
        if(!ref) continue;
        const amt = parseFloat(row[sapGlAmtIdx]) || 0;
        
        if(!sapMap.has(ref)) {
            sapMap.set(ref, {
                ref: ref,
                empId: row[sapEmpIdx] || '',
                type: row[sapTypeIdx] || '',
                totalAmt: 0
            });
        }
        sapMap.get(ref).totalAmt += amt;
    }

    // 3. Recon SAP Base
    sapData = Array.from(sapMap.values()).map(s => {
        let match = psMap.get(s.ref);
        let diff = 0, status = 'unmatched', ket = 'Not Found in PS', diffStr = '—';
        if (match) {
            diff = s.totalAmt - match.totalAmt;
            if (Math.abs(diff) < 0.01) { 
                status = 'matched'; ket = 'Matched dengan PS'; diffStr = '0'; 
            } else { 
                status = 'unmatched'; 
                ket = 'Excess/Mismatch: SAP beda ' + fmtIDR(Math.abs(diff)); 
                diffStr = diff > 0 ? '+' + fmtIDR(diff) : fmtIDR(diff); 
            }
        }
        return {
            id: s.empId,
            name: s.ref, // Tampilkan Reference / Invoice di kolom name
            period: 'Aug-26',
            type: s.type,
            amountNum: s.totalAmt,
            amount: fmtIDR(s.totalAmt),
            diff: diffStr,
            status: status,
            ket: ket
        };
    });

    // 4. Recon PS Base
    psData = Array.from(psMap.values()).map(p => {
        let match = sapMap.get(p.inv);
        let diff = 0, status = 'unmatched', ket = 'Not Found in SAP', diffStr = '—';
        if (match) {
            diff = p.totalAmt - match.totalAmt;
            if (Math.abs(diff) < 0.01) { 
                status = 'matched'; ket = 'Matched dengan SAP'; diffStr = '0'; 
            } else { 
                status = 'unmatched'; 
                ket = 'Excess/Mismatch: PS beda ' + fmtIDR(Math.abs(diff)); 
                diffStr = diff > 0 ? '+' + fmtIDR(diff) : fmtIDR(diff); 
            }
        }
        return {
            id: p.empId,
            name: p.inv, // Tampilkan Reference / Invoice di kolom name
            period: 'Aug-26',
            type: p.type,
            amountNum: p.totalAmt,
            amount: fmtIDR(p.totalAmt),
            diff: diffStr,
            status: status,
            ket: ket
        };
    });
    
    return true;
  } catch (e) {
    console.error(e);
    return false;
  }`;

if(regex.test(appJS)) {
    appJS = appJS.replace(regex, newLogic);
    fs.writeFileSync('app.js', appJS);
    console.log("Updated app.js");
} else {
    console.log("Could not find regex match!");
}
