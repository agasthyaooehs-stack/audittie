// ==========================================
// 100% LOCAL MATRIX ENGINE + FAST EMBEDDINGS
// ==========================================
import { pipeline, cos_sim } from 'https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.1';

let masterWorkbook = null;
let currentResults = [];
let aiExtractor = null; // Holds the local AI model

// DOM Elements
const btnRun = document.getElementById('btn-run');
const btnExport = document.getElementById('btn-export');
const selectSummary = document.getElementById('sheet-summary');
const selectTBs = document.getElementById('sheet-tbs');
const tableResultsBody = document.querySelector('#table-results tbody');
const tableSupportBody = document.querySelector('#table-support-breakdown tbody');
const uploadHint = document.getElementById('upload-hint');
const llmToggle = document.getElementById('llm-toggle');

// 1. File Upload & Auto-Detection
document.getElementById('file-wb').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  uploadHint.textContent = '⏳ Parsing workbook... Please wait.';
  uploadHint.style.color = 'var(--warning)';

  const reader = new FileReader();
  reader.onload = (evt) => {
    setTimeout(() => {
      try {
        const data = new Uint8Array(evt.target.result);
        masterWorkbook = XLSX.read(data, { type: 'array', cellStyles: false, cellFormula: false, cellHTML: false });
        populateDropdowns(masterWorkbook.SheetNames);
        document.getElementById('mapping-container').style.display = 'grid';
        uploadHint.textContent = 'Workbook loaded successfully.';
        uploadHint.style.color = 'var(--success)';
        btnRun.disabled = false;
      } catch (err) {
        console.error(err);
        uploadHint.textContent = 'Error: File too large or corrupted.';
        uploadHint.style.color = 'var(--danger)';
      }
    }, 50);
  };
  reader.readAsArrayBuffer(file);
});

function populateDropdowns(sheetNames) {
  selectSummary.innerHTML = ''; selectTBs.innerHTML = '';
  sheetNames.forEach(name => {
    const optSum = document.createElement('option'); optSum.value = name; optSum.textContent = name;
    if (name.toUpperCase().includes('GRP')) optSum.selected = true;
    selectSummary.appendChild(optSum);

    const optTB = document.createElement('option'); optTB.value = name; optTB.textContent = name;
    if (name.toUpperCase().startsWith('TB-')) optTB.selected = true;
    selectTBs.appendChild(optTB);
  });
}

function parseAccountingNumber(val) {
  if (typeof val === 'number') return val;
  const str = String(val || '0').trim();
  const isNegative = (str.startsWith('(') && str.endsWith(')')) || str.startsWith('-');
  const cleaned = str.replace(/[^\d.-]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : (isNegative ? -Math.abs(num) : Math.abs(num));
}

// ==========================================
// LOCAL AI INITIALIZATION & MATCHING
// ==========================================
async function loadLocalAI() {
  if (!aiExtractor) {
    uploadHint.textContent = '📥 Downloading local AI model (~22MB, first time only)...';
    uploadHint.style.color = 'var(--accent)';
    // Loads a tiny, ultra-fast embedding model into browser cache
    aiExtractor = await pipeline('feature-extraction', 'Xenova/all-MiniLM-L6-v2'); 
    uploadHint.textContent = '🧠 AI Model cached successfully.';
  }
}

async function getVector(text) {
  const output = await aiExtractor([text], { pooling: 'mean', normalize: true });
  return output.tolist()[0];
}

// 2. The Core Execution Engine
btnRun.addEventListener('click', async () => {
  const tolerance = parseFloat(document.getElementById('tolerance-input').value) || 1.00;
  const useLLM = llmToggle.checked;

  const summarySheetName = selectSummary.value;
  const selectedTBNames = Array.from(selectTBs.selectedOptions).map(opt => opt.value);

  if (!summarySheetName || selectedTBNames.length === 0) {
    alert("Please select a summary sheet and at least one TB sheet.");
    return;
  }

  btnRun.disabled = true;
  btnRun.textContent = "Running Match Engine...";

  try {
    if (useLLM) await loadLocalAI(); // Boot up local AI if toggled on

    const grpGrid = XLSX.utils.sheet_to_json(masterWorkbook.Sheets[summarySheetName], { header: 1, blankrows: false });
    
    let headerRowIdx = -1;
    let columnMapping = {}; 
    
    for (let r = 0; r < Math.min(30, grpGrid.length); r++) {
      const row = grpGrid[r];
      if (!row) continue;
      let foundMappings = 0;
      let tempMapping = {};

      for (let c = 0; c < row.length; c++) {
        const match = String(row[c] || '').match(/\b(\d{4})\b/);
        if (match) {
          const matchingTB = selectedTBNames.find(tb => tb.includes(match[1]));
          if (matchingTB) { tempMapping[c] = matchingTB; foundMappings++; }
        }
      }
      if (foundMappings >= 1) { headerRowIdx = r; columnMapping = tempMapping; break; }
    }

    if (headerRowIdx === -1) {
      alert("Could not map GRP columns to TB sheets. Ensure headers have matching 4-digit codes.");
      btnRun.disabled = false; btnRun.textContent = "Run Auto Matrix Tie-Out"; return;
    }

    const firstMappedCol = Math.min(...Object.keys(columnMapping).map(Number));
    const descColIdx = firstMappedCol > 0 ? firstMappedCol - 1 : 0;

    // Pre-parse TBs and calculate their AI vectors upfront to save time
    const preParsedTBs = {};
    for (const tbName of selectedTBNames) {
      const rawData = XLSX.utils.sheet_to_json(masterWorkbook.Sheets[tbName], { defval: 0, blankrows: false });
      if (rawData.length === 0) continue;

      const keys = Object.keys(rawData[0]);
      const descKey = keys.find(k => /acct|item|long|desc|particular/i.test(k)) || keys[0];
      let amtKey = keys.find(k => /accumulated|net|closing/i.test(k)) || [...keys].reverse().find(k => /balance|amount|total/i.test(k)) || keys[keys.length - 1];

      let parsedRows = rawData.map(r => ({
        rawDesc: String(r[descKey] || ''),
        cleanDesc: String(r[descKey] || '').toLowerCase().replace(/[^a-z0-9]/g, ''),
        amount: parseAccountingNumber(r[amtKey]),
        used: false,
        vector: null 
      })).filter(r => r.cleanDesc !== '' && r.amount !== 0); 

      // If AI is on, generate vectors for all valid TB rows instantly
      if (useLLM) {
        uploadHint.textContent = `🧠 Generating semantic vectors for ${tbName}...`;
        for (let row of parsedRows) {
          row.vector = await getVector(row.rawDesc);
        }
      }
      preParsedTBs[tbName] = parsedRows;
    }

    currentResults = [];

    // Loop over GRP Rows
    for (let r = headerRowIdx + 1; r < grpGrid.length; r++) {
      const row = grpGrid[r];
      if (!row) continue;
      
      const grpDescRaw = String(row[descColIdx] || '').trim();
      const grpDescClean = grpDescRaw.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (!grpDescClean) continue;

      for (const cIdx of Object.keys(columnMapping)) {
        const targetTBName = columnMapping[cIdx];
        const grpAmount = parseAccountingNumber(row[cIdx]);
        const tbRows = preParsedTBs[targetTBName] || [];
        
        let matchedTbRows = [];
        let aiReasoning = '';
        let isAiMatch = false;

        // 1. HARD DETERMINISTIC CHECKS FIRST
        if (Math.abs(grpAmount) < 0.01) {
          // If FS row is $0, it is a presentation header. Ignore automatically.
          aiReasoning = 'Deterministic: Ignored $0 Presentation Header';
        } else {
          // Look for exact string match first
          matchedTbRows = tbRows.filter(t => !t.used && t.cleanDesc === grpDescClean);
          
          if (matchedTbRows.length === 0) {
            // Look for substring match, but ONLY accept it if the math ties exactly to prevent false positives
            const substringMatches = tbRows.filter(t => !t.used && grpDescClean.length >= 4 && (t.cleanDesc.includes(grpDescClean) || grpDescClean.includes(t.cleanDesc)));
            const subSum = substringMatches.reduce((s, t) => s + t.amount, 0);
            if (substringMatches.length > 0 && Math.abs(Math.abs(grpAmount) - Math.abs(subSum)) <= tolerance) {
              matchedTbRows = substringMatches;
              aiReasoning = 'Deterministic: Substring sum matched exactly';
            }
          } else {
            aiReasoning = 'Deterministic: Exact string match';
          }
        }

        // 2. LOCAL AI SEMANTIC FALLBACK
        // If deterministic failed, and the line has a dollar value, use vector math
        if (matchedTbRows.length === 0 && Math.abs(grpAmount) >= 0.01 && useLLM) {
          const grpVector = await getVector(grpDescRaw);
          
          let bestMatch = null;
          let highestScore = 0;

          for (const tb of tbRows) {
            if (tb.used || !tb.vector) continue;
            const score = cos_sim(grpVector, tb.vector);
            if (score > highestScore) {
              highestScore = score;
              bestMatch = tb;
            }
          }

          // If similarity is greater than 70%, we consider it a confident semantic match
          if (highestScore >= 0.70) {
            matchedTbRows = [bestMatch];
            isAiMatch = true;
            aiReasoning = `Semantic AI Match (${(highestScore * 100).toFixed(1)}% confidence)`;
          } else {
            aiReasoning = 'No valid matches found';
          }
        }

        const tbSum = matchedTbRows.reduce((sum, t) => sum + t.amount, 0);
        const variance = Math.abs(Math.abs(grpAmount) - Math.abs(tbSum));

        // Skip rendering completely blank rows to keep the UI clean
        if (Math.abs(grpAmount) < 0.01 && matchedTbRows.length === 0) continue;

        // Lock the rows so they aren't reused
        if (variance <= tolerance) matchedTbRows.forEach(t => t.used = true);

        let status = 'TIED';
        let tickMark = isAiMatch ? '🤖 AI Matched' : '✓ Agreed';

        if (Math.abs(grpAmount) < 0.01 && matchedTbRows.length === 0) {
          status = 'TIED'; tickMark = 'Header'; 
        } else if (variance > tolerance && variance <= 1.0) { 
          status = 'ROUNDING'; tickMark += ' (Rounding)'; 
        } else if (variance > 1.0) { 
          status = 'VARIANCE'; tickMark = '✕ Variance'; 
        }

        currentResults.push({
          accountDesc: grpDescRaw,
          targetTB: targetTBName,
          grpAmount: grpAmount,
          tbSum: tbSum,
          variance: Math.abs(grpAmount) < 0.01 ? 0 : variance, 
          status: status,
          tickMark: tickMark,
          aiReasoning: aiReasoning,
          matchedRows: matchedTbRows
        });
      }
    }

    renderResults(currentResults);
    btnExport.disabled = false;
    uploadHint.textContent = 'Matrix Tie-Out Complete.';
    uploadHint.style.color = 'var(--success)';

  } catch (err) {
    console.error(err);
    uploadHint.textContent = 'Error processing files.';
    uploadHint.style.color = 'var(--danger)';
  } finally {
    btnRun.disabled = false;
    btnRun.textContent = "Run Auto Matrix Tie-Out";
  }
});

// 3. Render Outputs
function renderResults(results) {
  tableResultsBody.innerHTML = '';
  if (results.length === 0) {
    tableResultsBody.innerHTML = `<tr><td colspan="8" class="empty-state">No relevant accounting matches found.</td></tr>`;
    return;
  }

  results.forEach((res, index) => {
    const tr = document.createElement('tr');
    let badgeClass = res.status === 'ROUNDING' ? 'badge-warn' : (res.status === 'VARIANCE' ? 'badge-fail' : 'badge-pass');
    
    tr.innerHTML = `
      <td><span class="badge ${badgeClass}">${res.status}</span></td>
      <td><strong>${res.accountDesc}</strong></td>
      <td><span style="color:var(--text-muted);">${res.targetTB}</span></td>
      <td>$${res.grpAmount.toLocaleString(undefined, {minimumFractionDigits: 2})}</td>
      <td>$${res.tbSum.toLocaleString(undefined, {minimumFractionDigits: 2})}</td>
      <td style="color: ${res.variance > 0 ? 'var(--danger)' : 'inherit'}">$${res.variance.toFixed(2)}</td>
      <td><code>${res.tickMark}</code></td>
      <td class="ai-reasoning-cell">${res.aiReasoning}</td>
    `;
    
    tr.addEventListener('click', () => {
      document.querySelectorAll('#table-results tr').forEach(r => r.classList.remove('selected'));
      tr.classList.add('selected');
      inspectDetail(res);
    });
    tableResultsBody.appendChild(tr);
  });
}

function inspectDetail(result) {
  document.getElementById('detail-subtitle').innerText = `${result.accountDesc} -> ${result.targetTB}`;
  tableSupportBody.innerHTML = '';

  if (!result.matchedRows.length) {
    tableSupportBody.innerHTML = `<tr><td colspan="2" class="empty-state">No supporting rows attached.</td></tr>`;
    return;
  }

  result.matchedRows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.rawDesc}</td>
      <td>$${row.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}</td>
    `;
    tableSupportBody.appendChild(tr);
  });
}

// 4. Export
btnExport.addEventListener('click', () => {
  if (!currentResults.length) return;
  const exportData = currentResults.map(r => ({
    'Status': r.status,
    'GRP Account': r.accountDesc,
    'Target TB': r.targetTB,
    'GRP Value': r.grpAmount,
    'TB Sum': r.tbSum,
    'Absolute Variance': r.variance,
    'Tick Mark': r.tickMark,
    'Matching Logic': r.aiReasoning
  }));
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Tie-Out');
  XLSX.writeFile(workbook, `Matrix_TieOut_${new Date().toISOString().split('T')[0]}.xlsx`);
});