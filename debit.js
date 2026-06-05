// script.js — robust header detection + debit note extraction
(() => {
  const fileInput = document.getElementById('fileInput');
  const statusEl = document.getElementById('status');
  const supplierListEl = document.getElementById('supplierList');
  const generateSelectedBtn = document.getElementById('generateSelectedBtn');
  const downloadSelectedBtn = document.getElementById('downloadSelectedPDFBtn');
  const downloadAllBtn = document.getElementById('downloadAllPDFBtn');
  const notesContainer = document.getElementById('notesContainer');

  let suppliersMap = {};
  let supplierOrder = [];

  fileInput.addEventListener('change', handleFile, false);
  generateSelectedBtn.addEventListener('click', generateSelected);
  downloadSelectedBtn.addEventListener('click', downloadSelectedAsPDF);
  downloadAllBtn.addEventListener('click', downloadAllAsPDF);


  function setStatus(msg, isError = false) {
    statusEl.textContent = msg;
    statusEl.style.color = isError ? '#c62828' : '';
  }

  function handleFile(e) {
  const f = e.target.files && e.target.files[0];
  if (!f) {
    setStatus('No file selected.');
    return;
  }

  setStatus(`Reading file: ${f.name}...`);
  suppliersMap = {};
  supplierOrder = [];
  notesContainer.innerHTML = '';
  supplierListEl.innerHTML = 'Reading file...';

  const reader = new FileReader();

  reader.onload = (ev) => {
    try {
      let rowsRaw = [];
      const ext = f.name.split('.').pop().toLowerCase();

      if (ext === 'csv') {
        // parse CSV
        const text = ev.target.result;
        const workbook = XLSX.read(text, { type: 'string' });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        rowsRaw = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
      } else {
        // default Excel parsing
        const data = new Uint8Array(ev.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        rowsRaw = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
      }

      if (!rowsRaw || rowsRaw.length === 0) {
        setStatus('File is empty or unrecognized.', true);
        supplierListEl.innerHTML = 'No valid rows found.';
        return;
      }

      const rawHeaders = Object.keys(rowsRaw[0] || {});
      const headerMap = buildHeaderMap(rawHeaders);

      if (!headerMap.debitNote) {
        const fallback = rawHeaders.find(h => normalizeKey(h).includes('debit'));
        if (fallback) headerMap.debitNote = fallback;
      }

      console.log('Detected headerMap:', headerMap);
      setStatus(`Headers mapped (debit): ${headerMap.debitNote || 'NOT FOUND'}. Parsing rows...`);

      const rows = normalizeRows(rowsRaw, headerMap);

      let explicitCount = 0;
      rows.forEach(r => {
        const debit = computeDebitForRow(r);
        r.debitAmount = debit;
        if (parseNumber(r.rawDebitValue) > 0.01) explicitCount++;
        if (debit > 0.01) {
          const sup = r.supplier || 'Unknown Supplier';
          if (!suppliersMap[sup]) {
            suppliersMap[sup] = { name: sup, state: r.state || '', items: [], totalDebit: 0 };
            supplierOrder.push(sup);
          }
          suppliersMap[sup].items.push(r);
          suppliersMap[sup].totalDebit += debit;
        }
      });

      renderSupplierList();
      setStatus(`Loaded ${rows.length} rows. Header (Debit): ${headerMap.debitNote || 'NOT FOUND'}. Rows with debit: ${explicitCount}. Found ${Object.keys(suppliersMap).length} supplier(s) requiring debit.`);
      console.log('Suppliers map:', suppliersMap);

    } catch (err) {
      console.error(err);
      setStatus('Failed to read or parse the file. Open console for details.', true);
    }
  };

  if (f.name.toLowerCase().endsWith('.csv')) {
    reader.readAsText(f); // CSV as text
  } else {
    reader.readAsArrayBuffer(f); // Excel
  }
}


  // normalizeKey: strip diacritics, lower, remove non-alphanumerics
  function normalizeKey(s) {
    if (s === null || s === undefined) return '';
    try {
      return String(s)
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '') // remove diacritics
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
    } catch (e) {
      return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
    }
  }

  // function buildHeaderMap(rawHeaders) {
  //   const map = {};
  //   const tryFind = (candidates) => {
  //     for (const h of rawHeaders) {
  //       const n = normalizeKey(h);
  //       for (const cand of candidates) {
  //         if (n.includes(cand)) return h;
  //       }
  //     }
  //     return null;
  //   };

  //   map.month = tryFind(['month']);
  //   map.state = tryFind(['state']);
  //   map.supplier = tryFind(['supplier','suppliername']);
  //   map.item = tryFind(['item','description','product']);
  //   map.unitRate = tryFind(['unitrate','finalplannedrate','unit']);
  //   map.plannedUnits = tryFind(['plannedunits','plannedunit']);
  //   map.plannedAmount = tryFind(['plannedamount','plannedamountvated','planned']);
  //   map.invoicedUnits = tryFind(['invoicedunits','invoicedunit']);
  //   map.invoicedUnitRate = tryFind(['invoicedunitrate','invoicedrate']);
  //   map.invAmount = tryFind(['invamount','invamountvated','invoicedamount','invoiceamount']);
  //   map.invoiceNo = tryFind(['invoiceno','invno','invnumber']);
  //   map.receivedUnits = tryFind(['receivedunits','receivedunit']);
  //   map.receivedAmount = tryFind(['receivedunitsamount','receivedamount']);
  //   map.unitsPayable = tryFind(['unitspayable','payableunits']);
  //   map.actualPayable = tryFind(['actualpayable','actualpayableamount']);
  //   map.debitNote = tryFind(['debitnote','debit']);
  //   map.remarks = tryFind(['remarks','remark']);

  //   return map;
  // }


  function buildHeaderMap(rawHeaders) {
  // Pre-normalize headers once
  const norm = rawHeaders.map(h => ({ raw: h, n: normalizeKey(h) }));

  // Prefer exact -> startsWith -> includes, and prioritize candidate order
  const pick = (cands) => {
    const cs = cands.map(c => normalizeKey(c));

    // exact
    for (const c of cs) {
      const hit = norm.find(h => h.n === c);
      if (hit) return hit.raw;
    }
    // startsWith
    for (const c of cs) {
      const hit = norm.find(h => h.n.startsWith(c));
      if (hit) return hit.raw;
    }
    // includes
    for (const c of cs) {
      const hit = norm.find(h => h.n.includes(c));
      if (hit) return hit.raw;
    }
    return null;
  };

  const map = {};
  map.month            = pick(['month']);
  map.state            = pick(['state']);
  map.supplier         = pick(['supplier','suppliername']);
  map.item             = pick(['item','description','product']);

  // IMPORTANT: remove generic 'unit' so it doesn't match Planned Units, etc.
  map.unitRate         = pick(['final planned rate (vated)','final planned rate','unit rate','rate']);

  map.plannedUnits     = pick(['planned units','plannedunit','planned qty','plannedquantity']);

  // IMPORTANT: remove generic 'planned' so it can't hit "Final Planned rate"
  map.plannedAmount    = pick([
    'planned amount (vated)',
    'planned amount',
    'plannedamountvated',
    'plannedamount'
  ]);

  map.invoicedUnits    = pick(['invoiced units','invoicedunit']);
  map.invoicedUnitRate = pick(['invoiced unit rate (vated)','invoiced unit rate','invoicedunitrate','invoicedrate']);

  map.invAmount        = pick([
    'inv amount (vated)',
    'inv amount',
    'invoiced amount (vated)',
    'invoiced amount',
    'invoice amount',
    'invamountvated',
    'invamount'
  ]);

  map.invoiceNo        = pick(['invoice no','invoiceno','invno','inv number','invnumber']);
  map.receivedUnits    = pick(['received units','receivedunit']);
  map.receivedAmount   = pick(['received units amount (vated)','received units amount','receivedamount']);
  map.unitsPayable     = pick(['units payable','unitspayable','payable units','payableunits']);
  map.actualPayable    = pick(['actual payable amount','actual payable amount (₦)','actualpayableamount','actualpayable']);
  map.debitNote        = pick(['debit note','debitnote','debit']);
  map.remarks          = pick(['remarks','remark']);

  return map;
}



  function normalizeRows(rowsRaw, map) {
    return rowsRaw.map(r => {
      const get = key => {
        const h = map[key];
        return h ? (r[h] === undefined ? '' : r[h]) : '';
      };
      // store raw debit cell value as well for debug
      const rawDebit = get('debitNote');
      return {
        month: String(get('month') || '').trim(),
        state: String(get('state') || '').trim(),
        supplier: String(get('supplier') || '').trim(),
        item: String(get('item') || '').trim(),
        unitRate: String(get('unitRate') || '').trim(),
        plannedUnits: get('plannedUnits') || 0,
        plannedAmount: get('plannedAmount') || 0,
        invoicedUnits: get('invoicedUnits') || 0,
        invoicedUnitRate: get('invoicedUnitRate') || 0,
        invAmount: get('invAmount') || 0,
        invoiceNo: String(get('invoiceNo') || '').trim(),
        receivedUnits: get('receivedUnits') || 0,
        receivedAmount: get('receivedAmount') || 0,
        unitsPayable: get('unitsPayable') || 0,
        actualPayable: get('actualPayable') || 0,
        rawDebitValue: rawDebit,
        debitNote: rawDebit || '', // keep raw string too
        remarks: String(get('remarks') || '').trim()
      };
    });
  }

  // parse messy number strings (commas, currency, spaces)
  function parseNumber(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return val;
    const s = String(val).replace(/[^0-9.\-]/g, '');
    if (s === '' || s === '.' || s === '-.' ) return 0;
    const n = Number(s);
    return isNaN(n) ? 0 : n;
  }

  // compute debit: explicit raw debit cell > 0.01 wins; else planned-actual diff if positive
  function computeDebitForRow(row) {
    const explicit = parseNumber(row.rawDebitValue);
    const planned = parseNumber(row.plannedAmount);
    const actual = parseNumber(row.actualPayable);

    if (explicit > 0.01) return round2(explicit);

    // fallback: planned - actual if meaningful
    if (planned > 0 && actual >= 0) {
      const diff = planned - actual;
      if (diff > 0.01) return round2(diff);
    }
    return 0;
  }

  function round2(n){ return Math.round((n + Number.EPSILON)*100)/100; }

  function formatCurrency(n) {
    const num = Number(n) || 0;
    try {
      return new Intl.NumberFormat('en-NG', { style:'currency', currency:'NGN', maximumFractionDigits:2 }).format(num);
    } catch (e) {
      return '₦' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    }
  }

  function renderSupplierList() {
    supplierListEl.innerHTML = '';
    const keys = Object.keys(suppliersMap);
    if (keys.length === 0) {
      supplierListEl.innerHTML = '<div class="muted">No suppliers require debit notes.</div>';
      generateSelectedBtn.disabled = true;
      downloadSelectedBtn.disabled = true;
      downloadAllBtn.disabled = true;
      return;
    }

    keys.forEach((supName) => {
      const sup = suppliersMap[supName];
      const el = document.createElement('div');
      el.className = 'supplier-item';
      el.innerHTML = `
        <div class="supplier-left">
          <label style="display:flex; gap:8px; align-items:center;">
            <input type="checkbox" class="supplier-checkbox" data-supplier="${escapeHtml(supName)}" checked />
            <div style="min-width:0">
              <div class="supplier-name">${escapeHtml(supName)}</div>
              <div class="supplier-meta">${sup.items.length} item(s) — Total debit ${formatCurrency(sup.totalDebit)}</div>
            </div>
          </label>
        </div>
        <div class="supplier-actions">
          <button class="btn small-btn generate-btn" data-supplier="${escapeHtml(supName)}">Generate</button>
        </div>
      `;
      supplierListEl.appendChild(el);
    });

    Array.from(supplierListEl.querySelectorAll('.generate-btn')).forEach(btn => {
      btn.addEventListener('click', (ev) => {
        const sup = ev.currentTarget.dataset.supplier;
        generateForSupplier(sup);
      });
    });

    const checkboxes = Array.from(supplierListEl.querySelectorAll('.supplier-checkbox'));
    checkboxes.forEach(cb => cb.addEventListener('change', updateActionButtons));
    updateActionButtons();
  }

  function updateActionButtons() {
  const checks = supplierListEl.querySelectorAll('.supplier-checkbox:checked');

  // Control state of buttons
  generateSelectedBtn.disabled = checks.length === 0;
  downloadSelectedPDFBtn.disabled = checks.length === 0;
  downloadAllPDFBtn.disabled = Object.keys(suppliersMap).length === 0;
}

function generateSelected() {
  const checked = Array.from(
    supplierListEl.querySelectorAll('.supplier-checkbox:checked')
  ).map(c => c.dataset.supplier);

  if (checked.length === 0) return;

  notesContainer.innerHTML = '';
  checked.forEach(sup => generateForSupplier(sup));
}

async function downloadSelectedAsPDF() {
  const checked = Array.from(
    supplierListEl.querySelectorAll('.supplier-checkbox:checked')
  ).map(c => c.dataset.supplier);

  if (checked.length === 0) return;

  for (const supName of checked) {
    const noteEl = document.querySelector(
      `.note-card[data-supplier="${supName}"] .note-body`
    );
    if (noteEl) {
      await exportNoteToPDF(noteEl, supName);
    }
  }
}

async function downloadAllAsPDF() {
  const keys = Object.keys(suppliersMap);
  if (keys.length === 0) return;

  for (const supName of keys) {
    const noteEl = document.querySelector(
      `.note-card[data-supplier="${supName}"] .note-body`
    );
    if (noteEl) {
      await exportNoteToPDF(noteEl, supName);
    }
  }
}

function generateForSupplier(supplierName) {
  const sup = suppliersMap[supplierName];
  if (!sup) return;

  const noteHTML = createDebitNoteHTML(sup);

  const wrapper = document.createElement('div');
  wrapper.className = 'note-card';
  wrapper.dataset.supplier = supplierName; // <-- IMPORTANT: added this

  wrapper.innerHTML = `
    <div class="note-head">
      <h3>Recova Debit Note — ${escapeHtml(supplierName)}</h3>
      <div class="note-meta">${formatCurrency(sup.totalDebit)} — ${sup.items.length} item(s)</div>
    </div>
    <div class="note-body">${noteHTML}</div>
    <div class="note-footer">
      <div class="note-actions">
        <button class="btn download-note">Download PDF</button>
        <button class="btn print-note">Print</button>
      </div>
      <div style="color:var(--muted); font-size:12px;">Generated: ${new Date().toLocaleString()}</div>
    </div>
  `;
  notesContainer.appendChild(wrapper);

  // Per-supplier PDF download
  wrapper.querySelector('.download-note').addEventListener('click', async () => {
    const noteEl = wrapper.querySelector('.note-body');
    await exportNoteToPDF(noteEl, supplierName);
  });

  // Per-supplier print
  wrapper.querySelector('.print-note').addEventListener('click', () => {
    const w = window.open('', '_blank');
    const full = wrapStandaloneHtml(noteHTML);
    w.document.open();
    w.document.write(full);
    w.document.close();
    w.onload = () => {
      w.print();
    };
  });



  }

//   function createDebitNoteHTML(supObj) {
//   const items = supObj.items;
//   const totalPlanned = items.reduce((s,i) => s + parseNumber(i.plannedAmount), 0);
//   const totalActual = items.reduce((s,i) => s + parseNumber(i.actualPayable), 0);
//   const totalDebit = items.reduce((s,i) => s + parseNumber(i.debitAmount), 0);

//   const reasons = Array.from(new Set(items.map(it => it.remarks || '').filter(Boolean)));
//   const reasonsText = reasons.length ? reasons.join('; ') : 'Not specified';

//   const rowsHtml = items.map(it => `
//     <tr>
//       <td>${escapeHtml(it.item || '')}</td>
//       <td>${escapeHtml(it.invoiceNo || '')}</td>
//       <td>${escapeHtml(it.plannedUnits || '')}</td>
//       <td>${formatCurrency(parseNumber(it.plannedAmount))}</td>
//       <td>${escapeHtml(it.invoicedUnits || '')}</td>
//       <td>${formatCurrency(parseNumber(it.invAmount))}</td>
//       <td>${escapeHtml(it.receivedUnits || '')}</td>
//       <td>${formatCurrency(parseNumber(it.actualPayable))}</td>
//       <td>${formatCurrency(parseNumber(it.debitAmount))}</td>
//     </tr>
//   `).join('\n');

//   const month = items.map(i => i.month).find(Boolean) || '';
//   const addressLine = `Dear ${escapeHtml(supObj.name)},`;

//   const html = `
//     <div style="font-size:14px;color:#222;">
//       <div style="margin-bottom:12px;">
//         <div style="font-weight:700;"></div>
//         <div style="color:#666;font-size:13px;">Invoice Month: ${escapeHtml(month)} &nbsp; | &nbsp; Supplier State: ${escapeHtml(supObj.state)}</div>
//       </div>

//       <p style="margin:8px 0 6px 0">${addressLine}</p>

//       <p style="margin:6px 0;">
//         This is to notify you that a <strong>Debit Note of ${formatCurrency(totalDebit)}</strong> is being issued to your account in relation to supplies listed below.
//         The total planned amount across these deliveries was <strong>${formatCurrency(totalPlanned)}</strong>, while the actual payable amount recorded is <strong>${formatCurrency(totalActual)}</strong>.
//         The resulting shortfall (debit) is <strong>${formatCurrency(totalDebit)}</strong>.
//       </p>

//       <p style="margin:6px 0;"><strong>Reason(s):</strong> ${escapeHtml(reasonsText)}</p>

//       <!-- table wrapper ensures mobile horizontal scroll instead of overflow -->
//       <div class="table-wrapper" style="width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;margin-top:12px;padding-bottom:6px;">
//         <table class="note-table" style="width:100%;border-collapse:collapse;font-size:13px;min-width:0;">
//           <thead>
//             <tr>
//               <th style="white-space:nowrap">Item</th>
//               <th style="white-space:nowrap">Invoice No</th>
//               <th style="white-space:nowrap">Planned Units</th>
//               <th style="white-space:nowrap">Planned Amount</th>
//               <th style="white-space:nowrap">Invoiced Units</th>
//               <th style="white-space:nowrap">Invoiced Amount</th>
//               <th style="white-space:nowrap">Received Units</th>
//               <th style="white-space:nowrap">Actual Payable</th>
//               <th style="white-space:nowrap">Debit Amount</th>
//             </tr>
//           </thead>
//           <tbody>
//             ${rowsHtml}
//             <tr style="font-weight:700;">
//               <td colspan="3" style="text-align:left;">Totals</td>
//               <td>${formatCurrency(totalPlanned)}</td>
//               <td></td>
//               <td></td>
//               <td></td>
//               <td>${formatCurrency(totalActual)}</td>
//               <td>${formatCurrency(totalDebit)}</td>
//             </tr>
//           </tbody>
//         </table>
//       </div>

//       <p style="margin:12px 0 6px 0;">
//         Please adjust your account and provide confirmation. If you disagree with the values shown here, contact accounts immediately with supporting documentation.
//       </p>

//       <p style="margin:6px 0;">Regards,<br><strong>Procurement / Accounts</strong></p>
//     </div>
//   `;
//   return html;
// }



function createDebitNoteHTML(supObj) {
  const items = supObj.items;

  const totalPlanned = items.reduce((s,i) => s + parseNumber(i.plannedAmount), 0);
  const totalActual = items.reduce((s,i) => s + parseNumber(i.actualPayable), 0);
  const totalDebit  = items.reduce((s,i) => s + parseNumber(i.debitAmount), 0);

  // NEW: totals for invoiced units and amount
  const totalInvoicedUnits  = items.reduce((s,i) => s + parseNumber(i.invoicedUnits), 0);
  const totalInvoicedAmount = items.reduce((s,i) => s + parseNumber(i.invAmount), 0);

  const reasons = Array.from(new Set(items.map(it => it.remarks || '').filter(Boolean)));
  const reasonsText = reasons.length ? reasons.join('; ') : 'Not specified';

  // helper just for displaying units nicely
  const fmtUnits = (n) => {
    const v = Number(n) || 0;
    try { return v.toLocaleString(); } catch { return String(v); }
  };

  const rowsHtml = items.map(it => `
    <tr>
      <td>${escapeHtml(it.item || '')}</td>
      <td>${escapeHtml(it.invoiceNo || '')}</td>
      <td>${escapeHtml(it.plannedUnits || '')}</td>
      <td>${formatCurrency(parseNumber(it.plannedAmount))}</td>
      <td>${escapeHtml(it.invoicedUnits || '')}</td>
      <td>${formatCurrency(parseNumber(it.invAmount))}</td>
      <td>${escapeHtml(it.receivedUnits || '')}</td>
      <td>${formatCurrency(parseNumber(it.actualPayable))}</td>
      <td>${formatCurrency(parseNumber(it.debitAmount))}</td>
    </tr>
  `).join('\n');

  const month = items.map(i => i.month).find(Boolean) || '';
  const addressLine = `Dear ${escapeHtml(supObj.name)},`;

  const html = `
    <div style="font-size:14px;color:#222;">
      <div style="margin-bottom:12px;">
        <div style="font-weight:700;"></div>
        <div style="color:#666;font-size:13px;">Invoice Month: ${escapeHtml(month)} &nbsp; | &nbsp; Supplier State: ${escapeHtml(supObj.state)}</div>
      </div>

      <p style="margin:8px 0 6px 0">${addressLine}</p>

      <p style="margin:6px 0;">
        This is to notify you that a <strong>Debit Note of ${formatCurrency(totalDebit)}</strong> is being issued to your account in relation to supplies listed below.
        The total planned amount across these deliveries was <strong>${formatCurrency(totalPlanned)}</strong>, while the actual payable amount recorded is <strong>${formatCurrency(totalActual)}</strong>.
      </p>

      <!-- NEW: invoiced explanation -->
      <p style="margin:6px 0;">
        You invoiced <strong>${fmtUnits(totalInvoicedUnits)} unit(s)</strong> for a total of
        <strong>${formatCurrency(totalInvoicedAmount)}</strong>. After reconciliation, the approved payable is
        <strong>${formatCurrency(totalActual)}</strong>. The difference results in a debit of
        <strong>${formatCurrency(totalDebit)}</strong>.
      </p>

      <p style="margin:6px 0;"><strong>Reason(s):</strong> ${escapeHtml(reasonsText)}</p>

      <!-- table wrapper ensures mobile horizontal scroll instead of overflow -->
      <div class="table-wrapper" style="width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;margin-top:12px;padding-bottom:6px;">
        <table class="note-table" style="width:100%;border-collapse:collapse;font-size:13px;min-width:0;">
          <thead>
            <tr>
              <th style="white-space:nowrap">Item</th>
              <th style="white-space:nowrap">Invoice No</th>
              <th style="white-space:nowrap">Planned Units</th>
              <th style="white-space:nowrap">Planned Amount</th>
              <th style="white-space:nowrap">Invoiced Units</th>
              <th style="white-space:nowrap">Invoiced Amount</th>
              <th style="white-space:nowrap">Received Units</th>
              <th style="white-space:nowrap">Actual Payable</th>
              <th style="white-space:nowrap">Debit Amount</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
            <tr style="font-weight:700;">
              <td colspan="3" style="text-align:left;">Totals</td>
              <td>${formatCurrency(totalPlanned)}</td>
              <td></td>
              <td>${formatCurrency(totalInvoicedAmount)}</td>
              <td></td>
              <td>${formatCurrency(totalActual)}</td>
              <td>${formatCurrency(totalDebit)}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <p style="margin:12px 0 6px 0;">
        Please adjust your account and provide confirmation. If you disagree with the values shown here, contact accounts immediately with supporting documentation.
      </p>

      <p style="margin:6px 0;">Regards,<br><strong>Procurement / Accounts</strong></p>
    </div>
  `;
  return html;
}


function wrapStandaloneHtml(innerHtml) {
  const inlineCSS = `
    /* reset & base */
    * { box-sizing: border-box; }
    html,body { height: 100%; margin:0; padding:0; }
    body {
      font-family: Arial, Helvetica, sans-serif;
      background: #fff;
      color: #111;
      padding: 20px;
      line-height: 1.45;
      -webkit-font-smoothing:antialiased;
    }
    p { margin: 10px 0; }

    /* table base */
    .table-wrapper {
      width: 100%;
      overflow-x: auto;
      -webkit-overflow-scrolling: touch;
    }
    .note-table {
      border-collapse: collapse;
      width: 100%;
      margin-top: 0;
      font-size: 13px;
      min-width: 0; /* default: allow shrink on large screens */
    }
    .note-table th, .note-table td {
      border: 1px solid #ddd;
      padding: 8px;
      text-align: center;
      vertical-align: middle;
      word-break: break-word;
      overflow-wrap: anywhere;
    }
    .note-table th {
      background: #f7f7f7;
      font-weight: 600;
      white-space: nowrap;
    }

    /* Make the table readable on small screens:
       - keep the table inside a horizontal scroll container instead of overflowing the screen
       - on mobile, set a reasonable min-width so columns don't squish into unreadable lines
       - reduce font-size slightly on small devices
    */
    @media (max-width: 768px) {
      body { font-size: 13px; padding: 12px; }
      .note-table { font-size: 12px; min-width: 680px; } /* triggers horizontal scroll */
    }

    @media (max-width: 480px) {
      body { font-size: 12px; padding: 8px; }
      .note-table { font-size: 11px; min-width: 540px; } /* smaller phones: still scroll but smaller min width */
      .note-table th, .note-table td { padding: 6px; }
    }

    /* small visual niceties for printing/PDF */
    @media print {
      body { padding: 8mm; }
      .table-wrapper { overflow: visible !important; }
      .note-table { min-width: 0 !important; font-size: 11px; }
    }
  `;
  return `<!doctype html>
  <html>
    <head>
      <meta charset="utf-8">
      <title>Debit Notes</title>
      <meta name="viewport" content="width=device-width,initial-scale=1" />
      <style>${inlineCSS}</style>
    </head>
    <body>${innerHtml}</body>
  </html>`;
}


  function downloadBlob(content, filename) {
    const blob = new Blob([content], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, function (c) {
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }

  function slugify(s) {
    return String(s || '').toLowerCase().replace(/\s+/g,'-').replace(/[^a-z0-9\-]/g,'').substr(0,60);
  }

  // helpers exposed to console (for debugging if needed)
  window.__dn_debug = {
    parseNumber,
    normalizeKey,
    buildHeaderMap
  };


  // ====== PDF EXPORT ======
async function downloadSelectedAsPDF() {
  const checkboxes = document.querySelectorAll('#supplierList input[type=checkbox]:checked');
  if (checkboxes.length === 0) {
    alert('Please select at least one supplier.');
    return;
  }
  for (const cb of checkboxes) {
    const supplierName = cb.value;
    const noteEl = document.querySelector(`.debit-note[data-supplier="${supplierName}"]`);
    if (noteEl) {
      await exportNoteToPDF(noteEl, supplierName);
    }
  }
}

async function downloadAllAsPDF() {
  for (const supplierName of supplierOrder) {
    const noteEl = document.querySelector(`.debit-note[data-supplier="${supplierName}"]`);
    if (noteEl) {
      await exportNoteToPDF(noteEl, supplierName);
    }
  }
}



async function exportNoteToPDF(noteEl, supplierName) {
  const { jsPDF } = window.jspdf;

  // Clone into fixed-width wrapper
  const wrapper = document.createElement('div');
  wrapper.style.position = 'fixed';
  wrapper.style.left = '-9999px';
  wrapper.style.top = '0';
  wrapper.style.width = '794px'; // ~A4 width at 96dpi
  wrapper.style.background = '#fff';

  // Clone the note
  const clone = noteEl.cloneNode(true);

  // --- SHRINK TABLE TO FIT ---
  const table = clone.querySelector('.note-table');
  if (table) {
    table.style.transformOrigin = 'top left';
    table.style.transform = 'scale(0.97)'; // scale down table slightly
    table.style.display = 'inline-block';
  }

  wrapper.appendChild(clone);
  document.body.appendChild(wrapper);

  const canvas = await html2canvas(wrapper, {
    scale: 3,
    useCORS: true,
  });
  document.body.removeChild(wrapper);

  const imgData = canvas.toDataURL('image/png');
  const pdf = new jsPDF('p', 'mm', 'a4');

  // --- ADD PDF HEADING ---
  pdf.setFontSize(16);
  pdf.setFont('helvetica', 'bold');
  pdf.text('RECOVA - DEBIT NOTE', pdf.internal.pageSize.getWidth() / 2, 15, { align: 'center' });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imgProps = pdf.getImageProperties(imgData);

  const margin = 10;
  const usableWidth = pageWidth - margin * 2;

  const pdfWidth = usableWidth;
  const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

  let heightLeft = pdfHeight;
  let position = 25; // start below heading

  pdf.addImage(imgData, 'PNG', margin, position, pdfWidth, pdfHeight);
  heightLeft -= pageHeight - position;

  while (heightLeft > 0) {
    position = heightLeft - pdfHeight;
    pdf.addPage();
    pdf.addImage(imgData, 'PNG', margin, position, pdfWidth, pdfHeight);
    heightLeft -= pageHeight;
  }

  pdf.save(`DebitNote_${supplierName}.pdf`);
}

})();

