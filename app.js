/* =========================================================
   Maestro Inventarios | app.js
   FIX robusto + Diagnóstico:
   - Detecta sucursales desde headers "Inv-..."
   - Encuentra columnas aunque cambien espacios / guiones / acentos
   - Corrige H ILUSTRES y SAN AGUST (san agust / san agustin / etc.)
   - Diagnóstico: imprime qué columnas detectó por sucursal
   ========================================================= */

const RIESGO_DIAS = 15;
const SOBREINV_DIAS = 60;

let chartHist = null;

const state = {
  fileName: null,
  outputRows: [],
  warehouses: [],           // [{key, label, suffix}]
  selectedWarehouseKey: null,
  sort: { col: null, dir: "desc" },
  filters: {
    clasif: "ALL",
    q: "",
    onlyInv: false,
    covMin: null,
    covMax: null
  },
  diag: {
    headers: [],
    perWarehouseCols: {} // key -> {invCol, clsCol, promCol, covMesCol, covDiaCol}
  }
};

// ----------------------- Helpers -----------------------

function stripAccents(s) {
  return String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function norm(s) {
  return stripAccents(String(s ?? ""))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normKeyCompact(s) {
  return stripAccents(String(s ?? ""))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "")
    .trim();
}

function canonicalWarehouseKeyFromSuffix(suffixRaw) {
  const k = normKeyCompact(suffixRaw);

  if (k.includes("ilustres")) return "ilustres";
  if (k.startsWith("san") && k.includes("agust")) return "san_agust";

  if (k.includes("adelitas")) return "adelitas";
  if (k.includes("express")) return "express";
  if (k.includes("general")) return "general";

  return k || "unknown";
}

function warehouseLabelFromKey(key) {
  switch (key) {
    case "adelitas": return "adelitas";
    case "express": return "express";
    case "general": return "general";
    case "ilustres": return "h ilustres";
    case "san_agust": return "san agust";
    default: return key;
  }
}

function toNum(v) {
  const n = Number(String(v ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function toInt(v) {
  return Math.trunc(toNum(v));
}

function median(values) {
  const a = values.filter(x => Number.isFinite(x)).slice().sort((x,y)=>x-y);
  if (!a.length) return 0;
  const mid = Math.floor(a.length/2);
  return a.length % 2 ? a[mid] : (a[mid-1] + a[mid]) / 2;
}

function fmtInt(n) {
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits: 0 }).format(n ?? 0);
}

function fmtPct(frac) {
  const n = (Number.isFinite(frac) ? frac : 0) * 100;
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits: 1 }).format(n) + "%";
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

// ----------------------- DOM -----------------------

const el = (id) => document.getElementById(id);

const fileInput = el("fileInput");
const btnClear = el("btnClear");
const tabSummary = el("tabSummary");
const tabDetail = el("tabDetail");
const summaryView = el("summaryView");
const detailView = el("detailView");
const statusMsg = el("statusMsg");
const statusMeta = el("statusMeta");

const warehouseSelect = el("warehouseSelect");
const classSelect = el("classSelect");
const searchInput = el("searchInput");
const onlyInvToggle = el("onlyInvToggle");
const covMin = el("covMin");
const covMax = el("covMax");
const btnExport = el("btnExport");

const gSkus = el("gSkus");
const gInv = el("gInv");
const gRisk = el("gRisk");
const gOver = el("gOver");

const insights = el("insights");

const dataTable = el("dataTable");
const thead = dataTable.querySelector("thead");
const tbody = dataTable.querySelector("tbody");
const tableHint = el("tableHint");

// Diagnóstico (si existe el div en index)
const diagBox = el("diagBox");

// ----------------------- Status / UI -----------------------

function setStatus(okText, metaText = "") {
  statusMsg.textContent = okText;
  statusMeta.textContent = metaText;
}

function setEnabled(enabled) {
  btnClear.disabled = !enabled;
  tabSummary.disabled = !enabled;
  tabDetail.disabled = !enabled;

  warehouseSelect.disabled = !enabled;
  classSelect.disabled = !enabled;
  searchInput.disabled = !enabled;
  onlyInvToggle.disabled = !enabled;
  covMin.disabled = !enabled;
  covMax.disabled = !enabled;
  btnExport.disabled = !enabled;
}

function showView(which) {
  if (which === "summary") {
    tabSummary.classList.add("active");
    tabDetail.classList.remove("active");
    summaryView.classList.remove("hidden");
    detailView.classList.add("hidden");
  } else {
    tabDetail.classList.add("active");
    tabSummary.classList.remove("active");
    detailView.classList.remove("hidden");
    summaryView.classList.add("hidden");
  }
}

// ----------------------- Sheet parsing -----------------------

function pickOutputSheet(workbook) {
  const names = workbook.SheetNames || [];
  const preferred = ["OUTPUT", "Output", "output"];

  for (const p of preferred) {
    const hit = names.find(n => norm(n) === norm(p));
    if (hit) return hit;
  }
  return names[0];
}

function readWorkbookXlsx(arrayBuffer) {
  return XLSX.read(arrayBuffer, { type: "array" });
}

function sheetToRows(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function detectWarehousesFromHeaders(headers) {
  const invHeaders = headers
    .map(h => String(h))
    .filter(h => norm(h).startsWith("inv ")); // "Inv-xxx" => "inv xxx"

  const suffixes = invHeaders.map(h => {
    const raw = String(h);
    const idx = raw.indexOf("-");
    if (idx >= 0) return raw.slice(idx + 1).trim();
    return raw.replace(/^inv[\s-]+/i, "").trim();
  });

  const byKey = new Map();
  for (const suffix of suffixes) {
    const key = canonicalWarehouseKeyFromSuffix(suffix);
    if (!byKey.has(key)) {
      byKey.set(key, { key, label: warehouseLabelFromKey(key), suffix });
    }
  }

  const order = ["adelitas","express","general","ilustres","san_agust"];
  return Array.from(byKey.values()).sort((a,b) => {
    const ia = order.indexOf(a.key);
    const ib = order.indexOf(b.key);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib) || a.label.localeCompare(b.label);
  });
}

function normalizeClasif(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
  if (s.toUpperCase() === "A") return "A";
  if (s.toUpperCase() === "B") return "B";
  if (s.toUpperCase() === "C") return "C";
  if (norm(s).includes("sin")) return "Sin Mov";
  return s;
}

function buildHeaderIndex(headers) {
  const idx = {};
  for (const h of headers) {
    const k = normKeyCompact(h);
    if (!idx[k]) idx[k] = h;
  }
  return idx;
}

function findCol(headersIndex, headersList, prefixVariants, whSuffix) {
  const suffixN = normKeyCompact(whSuffix);
  const prefNs = prefixVariants.map(v => normKeyCompact(v));

  // 1) Intentos exactos
  for (const p of prefixVariants) {
    const tries = [
      `${p}-${whSuffix}`,
      `${p} -${whSuffix}`,
      `${p}- ${whSuffix}`,
      `${p} - ${whSuffix}`,
      `${p} ${whSuffix}`,
    ];
    for (const t of tries) {
      const k = normKeyCompact(t);
      if (headersIndex[k]) return headersIndex[k];
    }
  }

  // 2) Fallback por contiene (prefijo + sucursal)
  for (const h of headersList) {
    const hk = normKeyCompact(h);
    const hasPrefix = prefNs.some(pn => hk.includes(pn));
    const hasSuffix = hk.includes(suffixN);
    if (hasPrefix && hasSuffix) return h;
  }

  return null;
}

function buildOutputModel(rows) {
  if (!rows.length) return { warehouses: [], outputRows: [] };

  const headers = Object.keys(rows[0] || {});
  const headersIndex = buildHeaderIndex(headers);
  const warehouses = detectWarehousesFromHeaders(headers);

  // Guardamos para diagnóstico
  state.diag.headers = headers.slice();

  const baseCodigo = headers.find(h => norm(h) === "codigo") || "Codigo";
  const baseDesc = headers.find(h => norm(h) === "desc prod" || norm(h) === "desc_prod") || "desc_prod";
  const baseMeses = headers.find(h => norm(h) === "mesesusados") || "MesesUsados";

  // Pre-detect columnas por sucursal (una sola vez)
  state.diag.perWarehouseCols = {};
  for (const wh of warehouses) {
    const invCol = findCol(headersIndex, headers, ["Inv"], wh.suffix);
    const clsCol = findCol(headersIndex, headers, ["Clasificacion","Clasificación"], wh.suffix);
    const promCol = findCol(headersIndex, headers, ["Promedio Vta Mes","Promedio Vta. Mes","Prom Vta Mes"], wh.suffix);
    const covMesCol = findCol(headersIndex, headers, ["Cobertura (Mes)","Cobertura Mes","Cobertura(Mes)"], wh.suffix);
    const covDiaCol = findCol(headersIndex, headers, ["Cobertura Dias (30)","Cobertura Días (30)","Cobertura Dias(30)","Cobertura Días(30)"], wh.suffix);

    state.diag.perWarehouseCols[wh.key] = { invCol, clsCol, promCol, covMesCol, covDiaCol };
  }

  const out = rows.map(r => {
    const item = {
      meses: toInt(r[baseMeses]),
      codigo: String(r[baseCodigo] ?? "").trim(),
      desc: String(r[baseDesc] ?? "").trim(),
      byWh: {}
    };

    for (const wh of warehouses) {
      const cols = state.diag.perWarehouseCols[wh.key] || {};
      const inv = toInt(cols.invCol ? r[cols.invCol] : 0);
      const clasif = normalizeClasif(cols.clsCol ? r[cols.clsCol] : "");
      const prom = toNum(cols.promCol ? r[cols.promCol] : 0);
      const covMes = toNum(cols.covMesCol ? r[cols.covMesCol] : 0);
      const covDias = toNum(cols.covDiaCol ? r[cols.covDiaCol] : 0);

      item.byWh[wh.key] = { inv, clasif, prom, covMes, covDias };
    }

    return item;
  });

  return { warehouses, outputRows: out };
}

// ----------------------- Diagnóstico -----------------------

function renderDiag() {
  if (!diagBox) return; // si no existe en el html, no pasa nada

  if (!state.warehouses.length || !Object.keys(state.diag.perWarehouseCols || {}).length) {
    diagBox.textContent = "—";
    return;
  }

  const lines = [];
  lines.push("DIAGNÓSTICO DE COLUMNAS DETECTADAS");
  lines.push("=================================");
  lines.push(`Headers totales: ${state.diag.headers.length}`);
  lines.push("");

  for (const wh of state.warehouses) {
    const cols = state.diag.perWarehouseCols[wh.key] || {};
    lines.push(`Sucursal: ${wh.label}  (suffix detectado: "${wh.suffix}")`);
    lines.push(`  Inv:              ${cols.invCol || "❌ NO ENCONTRADO"}`);
    lines.push(`  Clasificacion:     ${cols.clsCol || "❌ NO ENCONTRADO"}`);
    lines.push(`  Promedio Vta Mes:  ${cols.promCol || "❌ NO ENCONTRADO"}`);
    lines.push(`  Cobertura (Mes):   ${cols.covMesCol || "❌ NO ENCONTRADO"}`);
    lines.push(`  Cobertura Dias:    ${cols.covDiaCol || "❌ NO ENCONTRADO"}`);
    lines.push("");
  }

  diagBox.textContent = lines.join("\n");
}

// ----------------------- Rendering -----------------------

function renderWarehouseSelect() {
  warehouseSelect.innerHTML = "";
  for (const w of state.warehouses) {
    const opt = document.createElement("option");
    opt.value = w.key;
    opt.textContent = w.label;
    warehouseSelect.appendChild(opt);
  }
  state.selectedWarehouseKey = state.selectedWarehouseKey || (state.warehouses[0]?.key ?? null);
  warehouseSelect.value = state.selectedWarehouseKey ?? "";
}

function applyFilters() {
  const whKey = state.selectedWarehouseKey;
  if (!whKey) return [];

  const q = norm(state.filters.q);
  const clasif = state.filters.clasif;

  const min = (state.filters.covMin == null || state.filters.covMin === "")
    ? null : Number(state.filters.covMin);
  const max = (state.filters.covMax == null || state.filters.covMax === "")
    ? null : Number(state.filters.covMax);

  return state.outputRows.filter(row => {
    const cell = row.byWh[whKey];
    if (!cell) return false;

    if (clasif !== "ALL") {
      if (clasif === "Sin Mov") {
        if (cell.clasif !== "Sin Mov") return false;
      } else {
        if (cell.clasif !== clasif) return false;
      }
    }

    if (state.filters.onlyInv && !(cell.inv > 0)) return false;

    if (q) {
      const hay = norm(row.codigo).includes(q) || norm(row.desc).includes(q);
      if (!hay) return false;
    }

    if (min != null && Number.isFinite(min)) {
      if (!(cell.covDias >= min)) return false;
    }
    if (max != null && Number.isFinite(max)) {
      if (!(cell.covDias <= max)) return false;
    }

    return true;
  });
}

function renderKPIs(filtered) {
  const whKey = state.selectedWarehouseKey;
  const invs = filtered.map(r => r.byWh[whKey]?.inv ?? 0);
  const covs = filtered.map(r => r.byWh[whKey]?.covDias ?? 0).filter(x => x > 0);

  const invTotal = invs.reduce((a,b)=>a+b,0);
  const covMed = median(covs);

  el("kpiMeses").textContent = fmtInt(Math.max(...filtered.map(r => r.meses || 0), 0));
  el("kpiSkus").textContent = fmtInt(filtered.length);
  el("kpiInv").textContent = fmtInt(invTotal);
  el("kpiCobMed").textContent = new Intl.NumberFormat("es-MX", { maximumFractionDigits: 1 }).format(covMed || 0);
}

function renderTable(filtered) {
  const whKey = state.selectedWarehouseKey;
  const { col, dir } = state.sort;
  const sorted = filtered.slice();

  const getSortVal = (r) => {
    const c = r.byWh[whKey];
    switch (col) {
      case "codigo": return r.codigo;
      case "desc": return r.desc;
      case "clasif": return c?.clasif ?? "";
      case "inv": return c?.inv ?? 0;
      case "prom": return c?.prom ?? 0;
      case "covDias": return c?.covDias ?? 0;
      default: return c?.inv ?? 0;
    }
  };

  if (col) {
    sorted.sort((a,b) => {
      const va = getSortVal(a);
      const vb = getSortVal(b);
      const mult = dir === "asc" ? 1 : -1;
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * mult;
      return String(va).localeCompare(String(vb), "es", { sensitivity:"base" }) * mult;
    });
  }

  const cols = [
    { key:"codigo", label:"Código" },
    { key:"desc", label:"Descripción" },
    { key:"clasif", label:"Clasificación" },
    { key:"inv", label:"Existencia" },
    { key:"prom", label:"Prom. venta mes" },
    { key:"covDias", label:"Cobertura (días)" },
  ];

  thead.innerHTML = "";
  const trh = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    th.textContent = c.label;
    th.onclick = () => {
      if (state.sort.col === c.key) {
        state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
      } else {
        state.sort.col = c.key;
        state.sort.dir = "desc";
      }
      render();
    };
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  tbody.innerHTML = "";
  const view = sorted.slice(0, 5000);
  for (const r of view) {
    const c = r.byWh[whKey];
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.codigo)}</td>
      <td>${escapeHtml(r.desc)}</td>
      <td>${escapeHtml(c?.clasif ?? "")}</td>
      <td>${fmtInt(c?.inv ?? 0)}</td>
      <td>${new Intl.NumberFormat("es-MX",{maximumFractionDigits:1}).format(c?.prom ?? 0)}</td>
      <td>${new Intl.NumberFormat("es-MX",{maximumFractionDigits:0}).format(c?.covDias ?? 0)}</td>
    `;
    tbody.appendChild(tr);
  }

  tableHint.textContent = `Mostrando ${fmtInt(view.length)} de ${fmtInt(sorted.length)} filas`;
}

function renderGeneral() {
  const whKeys = state.warehouses.map(w => w.key);
  const skus = state.outputRows.length;

  let invTotal = 0;
  let riskCount = 0;
  let overCount = 0;
  let covCount = 0;

  for (const r of state.outputRows) {
    for (const k of whKeys) {
      const c = r.byWh[k];
      if (!c) continue;
      invTotal += c.inv || 0;

      const d = c.covDias || 0;
      if (d > 0) {
        covCount++;
        if (d < RIESGO_DIAS) riskCount++;
        if (d > SOBREINV_DIAS) overCount++;
      }
    }
  }

  gSkus.textContent = fmtInt(skus);
  gInv.textContent = fmtInt(invTotal);
  gRisk.textContent = fmtPct(covCount ? (riskCount / covCount) : 0);
  gOver.textContent = fmtPct(covCount ? (overCount / covCount) : 0);

  const bins = [
    { label:"0–15", min:0, max:15 },
    { label:"16–30", min:16, max:30 },
    { label:"31–60", min:31, max:60 },
    { label:"61–120", min:61, max:120 },
    { label:">120", min:121, max:1e9 },
  ];
  const counts = bins.map(()=>0);

  for (const r of state.outputRows) {
    for (const k of whKeys) {
      const d = r.byWh[k]?.covDias ?? 0;
      if (!(d > 0)) continue;
      const idx = bins.findIndex(b => d >= b.min && d <= b.max);
      if (idx >= 0) counts[idx]++;
    }
  }

  const ctx = document.getElementById("chartHist").getContext("2d");
  if (chartHist) chartHist.destroy();
  chartHist = new Chart(ctx, {
    type: "bar",
    data: {
      labels: bins.map(b=>b.label),
      datasets: [{ label:"Conteo", data: counts }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display:false } },
      scales: { y: { ticks: { precision:0 } } }
    }
  });

  insights.innerHTML = "";
  const li1 = document.createElement("li");
  li1.textContent = `Inventario total (todas las sucursales): ${fmtInt(invTotal)} piezas.`;
  const li2 = document.createElement("li");
  li2.textContent = `Riesgo (<${RIESGO_DIAS} días): ${fmtPct(covCount ? riskCount/covCount : 0)} • Sobreinv (>${SOBREINV_DIAS} días): ${fmtPct(covCount ? overCount/covCount : 0)}.`;
  insights.appendChild(li1);
  insights.appendChild(li2);
}

function render() {
  if (!state.outputRows.length) return;
  renderGeneral();
  renderDiag(); // ✅ aquí se actualiza el panel

  const filtered = applyFilters();
  renderKPIs(filtered);
  renderTable(filtered);
}

// ----------------------- Export -----------------------

function exportCsv(rows) {
  const whKey = state.selectedWarehouseKey;
  const header = ["Codigo","Descripcion","Clasificacion","Existencia","Prom_venta_mes","Cobertura_dias"];
  const lines = [header.join(",")];

  for (const r of rows) {
    const c = r.byWh[whKey];
    lines.push([
      `"${String(r.codigo).replaceAll('"','""')}"`,
      `"${String(r.desc).replaceAll('"','""')}"`,
      `"${String(c?.clasif ?? "").replaceAll('"','""')}"`,
      String(c?.inv ?? 0),
      String(c?.prom ?? 0),
      String(c?.covDias ?? 0),
    ].join(","));
  }

  const blob = new Blob([lines.join("\n")], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `vista_${whKey}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ----------------------- CSV parser -----------------------

function csvToObjects(csvText) {
  const lines = csvText.split(/\r?\n/).filter(l => l.trim().length);
  if (!lines.length) return [];

  const headers = splitCsvLine(lines[0]).map(h => h.trim());
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const parts = splitCsvLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = parts[idx] ?? "";
    });
    rows.push(obj);
  }
  return rows;
}

function splitCsvLine(line) {
  const out = [];
  let cur = "";
  let inQ = false;

  for (let i=0;i<line.length;i++) {
    const ch = line[i];
    if (ch === '"' ) {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === "," && !inQ) {
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

// ----------------------- Events -----------------------

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  state.fileName = file.name;
  setStatus("⏳ Leyendo archivo...", "");

  try {
    if (file.name.toLowerCase().endsWith(".csv")) {
      const text = await file.text();
      const rows = csvToObjects(text);
      const model = buildOutputModel(rows);

      state.outputRows = model.outputRows;
      state.warehouses = model.warehouses;

    } else {
      const buf = await file.arrayBuffer();
      const wb = readWorkbookXlsx(buf);
      const outName = pickOutputSheet(wb);
      const rows = sheetToRows(wb, outName);

      const model = buildOutputModel(rows);
      state.outputRows = model.outputRows;
      state.warehouses = model.warehouses;
    }

    if (!state.warehouses.length) {
      setStatus("❌ No detecté sucursales. Revisa que existan columnas tipo 'Inv-...'", "");
      setEnabled(false);
      return;
    }

    renderWarehouseSelect();
    setEnabled(true);

    setStatus(
      "✅ Archivo cargado correctamente",
      `Sucursales detectadas: ${state.warehouses.map(w=>w.label).join(", ")} • Filas (SKUs): ${fmtInt(state.outputRows.length)}`
    );

    state.selectedWarehouseKey = state.warehouses[0].key;
    warehouseSelect.value = state.selectedWarehouseKey;

    showView("summary");
    render();

  } catch (err) {
    console.error(err);
    setStatus("❌ Error al leer el archivo", String(err?.message ?? err));
    setEnabled(false);
  }
});

btnClear.addEventListener("click", () => {
  fileInput.value = "";
  setStatus("Carga un archivo para comenzar.", "");
  setEnabled(false);

  state.fileName = null;
  state.outputRows = [];
  state.warehouses = [];
  state.selectedWarehouseKey = null;

  state.diag.headers = [];
  state.diag.perWarehouseCols = {};

  thead.innerHTML = "";
  tbody.innerHTML = "";
  if (diagBox) diagBox.textContent = "—";
  if (chartHist) { chartHist.destroy(); chartHist = null; }
});

tabSummary.addEventListener("click", () => showView("summary"));
tabDetail.addEventListener("click", () => showView("detail"));

warehouseSelect.addEventListener("change", () => {
  state.selectedWarehouseKey = warehouseSelect.value;
  render();
});

classSelect.addEventListener("change", () => {
  state.filters.clasif = classSelect.value;
  render();
});

searchInput.addEventListener("input", () => {
  state.filters.q = searchInput.value || "";
  render();
});

onlyInvToggle.addEventListener("change", () => {
  state.filters.onlyInv = !!onlyInvToggle.checked;
  render();
});

covMin.addEventListener("input", () => {
  state.filters.covMin = covMin.value;
  render();
});
covMax.addEventListener("input", () => {
  state.filters.covMax = covMax.value;
  render();
});

btnExport.addEventListener("click", () => {
  const filtered = applyFilters();
  exportCsv(filtered);
});
