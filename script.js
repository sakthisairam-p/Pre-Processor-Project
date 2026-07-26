let rawRows = [], headers = [], colTypes = {}, processedRows = null, processedHeaders = null;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const app = document.getElementById('app');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e => {
  e.preventDefault(); dropzone.classList.remove('drag');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

document.getElementById('resetBtn').addEventListener('click', () => {
  app.style.display = 'none';
  dropzone.style.display = 'block';
  fileInput.value = '';
  setStep(1);
});

function setStep(n){
  for(let i=1;i<=4;i++) document.getElementById('step'+i).classList.toggle('active', i===n);
}

function handleFile(file){
  setStep(2);
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx' || ext === 'xls'){
    const reader = new FileReader();
    reader.onload = function(e){
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' , raw: false});
      headers = json.length ? Object.keys(json[0]) : [];
      rawRows = json;
      finishLoad(file.name);
    };
    reader.readAsArrayBuffer(file);
  } else {
    Papa.parse(file, {
      header: true,
      dynamicTyping: false,
      skipEmptyLines: true,
      complete: function(results){
        rawRows = results.data;
        headers = results.meta.fields || [];
        finishLoad(file.name);
      }
    });
  }
}

function finishLoad(fileName){
  inferTypes();
  renderSummary(fileName);
  populateEncodeColumns();
  dropzone.style.display = 'none';
  app.style.display = 'block';
  renderPreviewTable(rawRows, headers, 'raw');
  setStep(3);
}

function populateEncodeColumns(){
  const box = document.getElementById('encodeColSelect');
  box.innerHTML = '';
  const textCols = headers.filter(h => colTypes[h] === 'text');
  if (!textCols.length){
    box.innerHTML = '<div class="hint" style="color:var(--muted);">No text columns detected in this file.</div>';
    return;
  }
  textCols.forEach(h => {
    const uniq = uniqueCount(h);
    const div = document.createElement('div');
    div.className = 'checkline';
    div.innerHTML = `<input type="checkbox" class="encodeColCheck" value="${h}" id="enc_${h}"><label for="enc_${h}">${h} <span style="color:var(--muted);">(${uniq} unique)</span></label>`;
    box.appendChild(div);
  });
}

document.getElementById('optEncode').addEventListener('change', updateEncodeVisibility);
document.getElementById('optEncodeScope').addEventListener('change', updateEncodeVisibility);
function updateEncodeVisibility(){
  const method = document.getElementById('optEncode').value;
  const scope = document.getElementById('optEncodeScope').value;
  document.getElementById('optEncodeScope').style.display = method === 'none' ? 'none' : 'block';
  document.getElementById('encodeColSelect').style.display = (method !== 'none' && scope === 'specific') ? 'block' : 'none';
}

function inferTypes(){
  colTypes = {};
  headers.forEach(h => {
    let numCount = 0, total = 0;
    rawRows.forEach(r => {
      const v = r[h];
      if (v === undefined || v === null || v === '') return;
      total++;
      if (!isNaN(parseFloat(v)) && isFinite(v)) numCount++;
    });
    // majority-based: a column counts as numeric if most non-null values are numeric.
    // the rest get caught by mismatch detection instead of flipping the whole column to "text".
    colTypes[h] = (total > 0 && (numCount / total) >= 0.8) ? 'number' : 'text';
  });
}

function countNulls(h){
  return rawRows.filter(r => r[h] === undefined || r[h] === null || r[h] === '').length;
}
function uniqueCount(h){
  return new Set(rawRows.map(r => r[h])).size;
}
function countDupeRows(rows){
  const seen = new Set(); let dupes = 0;
  rows.forEach(r => { const key = JSON.stringify(r); if (seen.has(key)) dupes++; else seen.add(key); });
  return dupes;
}
function isMismatch(h, v){
  if (colTypes[h] !== 'number') return false;
  if (v === undefined || v === null || v === '') return false;
  return isNaN(parseFloat(v)) || !isFinite(v);
}
function countMismatches(h){
  return rawRows.filter(r => isMismatch(h, r[h])).length;
}
function totalMismatches(rows){
  let n = 0;
  headers.forEach(h => { if (colTypes[h] === 'number') n += rows.filter(r => isMismatch(h, r[h])).length; });
  return n;
}

function renderSummary(fileName){
  document.getElementById('statFile').textContent = fileName;
  document.getElementById('statRows').textContent = rawRows.length;
  document.getElementById('statCols').textContent = headers.length;
  document.getElementById('statDupes').textContent = countDupeRows(rawRows);
  let totalNulls = 0; headers.forEach(h => totalNulls += countNulls(h));
  document.getElementById('statNulls').textContent = totalNulls;

  const mismatches = totalMismatches(rawRows);
  document.getElementById('statMismatch').textContent = mismatches;

  const dupeCount = countDupeRows(rawRows);
  const statusEl = document.getElementById('statStatus');
  if (totalNulls > 0 || dupeCount > 0 || mismatches > 0){
    statusEl.textContent = 'Needs cleaning';
    statusEl.style.color = 'var(--warn)';
  } else {
    statusEl.textContent = 'Clean';
    statusEl.style.color = 'var(--accent)';
  }

  const colList = document.getElementById('colList');
  colList.innerHTML = '';
  headers.forEach(h => {
    const nulls = countNulls(h);
    const uniq = uniqueCount(h);
    const mismatch = countMismatches(h);
    const div = document.createElement('div');
    div.className = 'col-item';
    div.innerHTML = `
      <div class="name">${h}</div>
      <div class="meta">
        <span class="badge ${colTypes[h]==='number'?'num':'text'}">${colTypes[h]}</span>
        <span class="badge ${nulls>0?'nulls':''}">${nulls} null${nulls===1?'':'s'}</span>
        <span class="badge">${uniq} unique</span>
        ${mismatch>0 ? `<span class="badge mismatch">${mismatch} mismatch${mismatch===1?'':'es'}</span>` : ''}
      </div>`;
    colList.appendChild(div);
  });
}

function renderPreviewTable(rows, cols, mode){
  const table = document.getElementById('previewTable');
  document.getElementById('previewLabel').textContent = `(first 10 rows${mode==='processed' ? ' — processed' : ''})`;
  if (!rows.length){ table.innerHTML = '<tr><td class="empty">No rows remaining after preprocessing.</td></tr>'; return; }
  let html = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>';
  rows.slice(0,10).forEach(r => {
    html += '<tr>' + cols.map(c => `<td>${r[c] !== undefined && r[c] !== null ? r[c] : ''}</td>`).join('') + '</tr>';
  });
  html += '</tbody>';
  table.innerHTML = html;
}

document.getElementById('applyBtn').addEventListener('click', () => {
  let rows = rawRows.map(r => ({...r}));
  const missingMode = document.getElementById('optMissing').value;
  const dedupe = document.getElementById('optDedupe').checked;
  const scaleMode = document.getElementById('optScale').value;
  const encodeMode = document.getElementById('optEncode').value;
  const mismatchMode = document.getElementById('optMismatch').value;
  let cols = [...headers];
  const logLines = [];
  const startRows = rows.length;

  // type mismatches (e.g. "N/A" sitting in a numeric column)
  if (mismatchMode === 'drop'){
    const before = rows.length;
    rows = rows.filter(r => !cols.some(h => isMismatch(h, r[h])));
    logLines.push(`Dropped rows with type mismatches: ${before - rows.length} removed`);
  } else if (mismatchMode === 'convert'){
    let n = 0;
    rows.forEach(r => { cols.forEach(h => { if (isMismatch(h, r[h])){ r[h] = ''; n++; } }); });
    if (n > 0) logLines.push(`Converted ${n} mismatched value(s) to null (handled by missing-value rule)`);
  }

  // cast numerics
  cols.forEach(h => {
    if (colTypes[h] === 'number'){
      rows.forEach(r => { if (r[h] !== '' && r[h] !== undefined && r[h] !== null && !isNaN(parseFloat(r[h])) && isFinite(r[h])) r[h] = parseFloat(r[h]); });
    }
  });

  // missing values
  if (missingMode === 'drop'){
    const beforeDrop = rows.length;
    rows = rows.filter(r => cols.every(h => r[h] !== '' && r[h] !== undefined && r[h] !== null && !(colTypes[h]==='number' && isNaN(r[h]))));
    logLines.push(`Dropped rows with missing values: ${beforeDrop - rows.length} removed`);
  } else if (missingMode === 'fillMeanMode' || missingMode === 'fillZero'){
    cols.forEach(h => {
      if (colTypes[h] === 'number'){
        const nums = rows.map(r => r[h]).filter(v => v !== '' && v !== undefined && v !== null && !isNaN(v));
        const mean = nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0;
        const fillVal = missingMode === 'fillZero' ? 0 : Math.round(mean*1000)/1000;
        rows.forEach(r => { if (r[h] === '' || r[h] === undefined || r[h] === null || isNaN(r[h])) r[h] = fillVal; });
      } else {
        const counts = {};
        rows.forEach(r => { const v = r[h]; if (v !== '' && v !== undefined && v !== null){ counts[v]=(counts[v]||0)+1; } });
        let mode = '';
        let max = 0;
        Object.keys(counts).forEach(k => { if (counts[k] > max){ max = counts[k]; mode = k; } });
        const fillVal = missingMode === 'fillZero' ? '' : mode;
        rows.forEach(r => { if (r[h] === '' || r[h] === undefined || r[h] === null) r[h] = fillVal; });
      }
    });
    logLines.push(`Filled missing values (${missingMode === 'fillZero' ? '0 / empty string' : 'mean / mode'})`);
  }

  // dedupe
  if (dedupe){
    const seen = new Set(); const deduped = [];
    rows.forEach(r => { const key = JSON.stringify(r); if (!seen.has(key)){ seen.add(key); deduped.push(r); } });
    logLines.push(`Removed duplicate rows: ${rows.length - deduped.length} removed`);
    rows = deduped;
  }

  // scaling
  if (scaleMode !== 'none'){
    cols.forEach(h => {
      if (colTypes[h] !== 'number') return;
      const nums = rows.map(r => r[h]).filter(v => typeof v === 'number' && !isNaN(v));
      if (!nums.length) return;
      if (scaleMode === 'minmax'){
        const min = Math.min(...nums), max = Math.max(...nums);
        const range = max - min || 1;
        rows.forEach(r => { if (typeof r[h] === 'number') r[h] = Math.round(((r[h]-min)/range)*10000)/10000; });
      } else if (scaleMode === 'zscore'){
        const mean = nums.reduce((a,b)=>a+b,0)/nums.length;
        const variance = nums.reduce((a,b)=>a+(b-mean)**2,0)/nums.length;
        const std = Math.sqrt(variance) || 1;
        rows.forEach(r => { if (typeof r[h] === 'number') r[h] = Math.round(((r[h]-mean)/std)*10000)/10000; });
      }
    });
    logLines.push(`Applied ${scaleMode === 'minmax' ? 'min-max normalization' : 'z-score standardization'} to numeric columns`);
  }

  // encoding
  const encodeScope = document.getElementById('optEncodeScope').value;
  let encodeTargets = cols.filter(h => colTypes[h] === 'text');
  if (encodeMode !== 'none' && encodeScope === 'specific'){
    encodeTargets = Array.from(document.querySelectorAll('.encodeColCheck:checked')).map(el => el.value);
  }

  if (encodeMode === 'onehot'){
    let newCols = [...cols];
    encodeTargets.forEach(h => {
      const cats = [...new Set(rows.map(r => r[h]))].filter(v => v !== '' && v !== undefined && v !== null);
      if (cats.length > 15) return;
      newCols = newCols.filter(c => c !== h);
      cats.forEach(cat => {
        const colName = `${h}_${cat}`;
        newCols.push(colName);
        rows.forEach(r => { r[colName] = r[h] === cat ? 1 : 0; });
      });
    });
    cols = newCols;
    logLines.push(`One-hot encoded: ${encodeTargets.join(', ') || '(no columns selected)'}`);
  } else if (encodeMode === 'label'){
    encodeTargets.forEach(h => {
      const cats = [...new Set(rows.map(r => r[h]))];
      const map = {}; cats.forEach((c,i) => map[c] = i);
      rows.forEach(r => { r[h] = map[r[h]]; });
    });
    logLines.push(`Label-encoded: ${encodeTargets.join(', ') || '(no columns selected)'}`);
  }

  processedRows = rows;
  processedHeaders = cols;

  // final validation pass — confirm the output is actually clean
  let remainingNulls = 0, remainingMismatches = 0;
  cols.forEach(h => {
    rows.forEach(r => {
      const v = r[h];
      if (v === '' || v === undefined || v === null) remainingNulls++;
      else if (typeof v === 'number' && isNaN(v)) remainingMismatches++;
    });
  });
  const remainingDupes = countDupeRows(rows);
  const isClean = remainingNulls === 0 && remainingMismatches === 0 && remainingDupes === 0;

  const logBox = document.getElementById('logBox');
  logBox.style.display = 'block';
  let html = `✓ Preprocessing applied — ${startRows} → ${rows.length} rows, ${cols.length} columns<br>`;
  html += logLines.map(l => '• ' + l).join('<br>');
  html += `<br><br>${isClean
    ? '✓ Validation passed — no nulls, no duplicates, no type mismatches. Dataset is ML-ready.'
    : `⚠ Validation found remaining issues — nulls: ${remainingNulls}, duplicates: ${remainingDupes}, mismatches: ${remainingMismatches}. Adjust the options above (e.g. switch "Leave as-is" to a fill/drop rule) and re-apply before using this for model training.`}`;
  logBox.innerHTML = html;
  logBox.style.color = isClean ? 'var(--accent)' : 'var(--warn)';
  logBox.style.borderColor = isClean ? 'var(--accent-dim)' : 'var(--warn-border)';
  logBox.style.background = isClean ? 'var(--accent-tint)' : 'var(--warn-tint)';

  renderPreviewTable(rows, cols, 'processed');
  document.getElementById('downloadBtn').style.display = 'inline-block';
  setStep(4);
});

document.getElementById('downloadBtn').addEventListener('click', () => {
  if (!processedRows) return;
  const csv = Papa.unparse({ fields: processedHeaders, data: processedRows.map(r => processedHeaders.map(h => r[h])) });
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'processed_dataset.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

setStep(1);