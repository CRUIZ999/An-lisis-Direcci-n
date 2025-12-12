// ==================== CONFIG ====================

// m² por almacén
const M2_POR_ALMACEN = {
  "GENERAL": 1538,
  "EXPRESS": 369,
  "SAN AGUST": 870,
  "ADELITAS": 348,
  "H ILUSTRES": 100,
  "Todas": 3225
};

// Mapeos de columnas (tolerante por normalizeKey)
const RENAME_FACTURAS = {
  "No_fac": "Factura",
  "Falta_fac": "Fecha",
  "Hora_fac": "Hora",
  "Lugar": "Almacen",
  "Nom_cte": "Cliente",
  "Nom_age": "Vendedor",
  "Desc_prod": "Producto",
  "Cant_surt": "Cantidad",
  "Subt_fac": "Subtotal",
  "Cto_ent": "Costo",
  "Des_tial": "Marca",
  "Cse_prod": "ID Categoria",
  "Categoria": "CategoriaNombre",
  "Tipo de Factura": "TipoFacturaTexto"
};

const RENAME_NOTAS = {
  "No_fac": "Nota",
  "Falta_fac": "Fecha",
  "Hravta": "Hora",
  "Lugar": "Almacen",
  "Nom_fac": "Cliente",
  "Nom_age": "Vendedor",
  "Desc_prod": "Producto",
  "Cant_surt": "Cantidad",
  "Subt_prod": "Subtotal",
  "Cto_ent": "Costo",
  "Des_tial": "Marca",
  "Cse_prod": "ID Categoria",
  "Categoria": "CategoriaNombre",
  "Cve_suc": "Albaranes"
};

// Detalle: columnas disponibles (paleta)
const ALL_COLS = [
  "Año","Fecha","Hora","Almacén",
  "Factura/Nota","Origen",
  "Cliente","Categoría","Marca","Vendedor",
  "Producto","Cantidad",
  "Subtotal","Costo","Utilidad","Margen %","Precio sin IVA"
];

// Detalle: 7 columnas por default
const DEFAULT_SELECTED_COLS = ["Año","Almacén","Categoría","Costo","Subtotal","Margen %","Utilidad"];

// localStorage keys
const LS_KEY = "cedro_dashboard_view_v1";

// ==================== UTILS ====================

function normalizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(/_/g, "");
}

function buildNormalizedMap(mapping) {
  const out = {};
  for (const k in mapping) out[normalizeKey(k)] = mapping[k];
  return out;
}

const NORM_FACT = buildNormalizedMap(RENAME_FACTURAS);
const NORM_NOTAS = buildNormalizedMap(RENAME_NOTAS);

function toNumber(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return v;
  const s = v.toString().replace(/\s/g, "").replace(/,/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function formatCurrency(v) {
  return Number(v || 0).toLocaleString("es-MX", {
    style: "currency",
    currency: "MXN",
    maximumFractionDigits: 0
  });
}

function formatPercent(d) {
  if (!isFinite(d)) return "0.0%";
  return (d * 100).toFixed(1) + "%";
}

function parseFecha(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    const base = new Date(Date.UTC(1899, 11, 30));
    return new Date(base.getTime() + v * 86400000);
  }
  const s = v.toString().trim();
  const d1 = new Date(s);
  if (!isNaN(d1)) return d1;
  const parts = s.split(/[\/\-]/);
  if (parts.length === 3) {
    const [a, b, c] = parts.map(x => parseInt(x, 10));
    if (a > 1900) return new Date(a, b - 1, c);
    if (c > 1900) return new Date(c, b - 1, a);
  }
  return null;
}

function sumField(arr, field) {
  return arr.reduce((acc, r) => acc + (r[field] || 0), 0);
}

function unique(arr) {
  return Array.from(new Set(arr));
}

function debounce(fn, delay) {
  let t;
  return function (...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(null, args), delay);
  };
}

function renameRow(row, normMap) {
  const out = {};
  for (const key in row) {
    const dest = normMap[normalizeKey(key)];
    if (dest) out[dest] = row[key];
  }
  return out;
}

function safeUpper(x){ return (x || "").toString().trim().toUpperCase(); }
function safeStr(x){ return (x || "").toString().trim(); }

function yyyymmdd(d){
  if (!(d instanceof Date)) return "";
  return d.toISOString().slice(0,10);
}

// ==================== STATE ====================

let records = [];
let yearsDisponibles = [];
let almacenesDisponibles = [];
let categoriasDisponibles = []; // nombres
let categoriaNameById = {}; // opcional

let charts = {
  mensual: null,
  almacen: null,
  credit: null,
  cats: null,
};

let highlightNeg = true;

// detalle builder
let selectedCols = DEFAULT_SELECTED_COLS.slice(); // 7 columnas visibles
let detalleSort = { col: null, asc: true };
let detalleFiltros = {};
let detalleBusqueda = "";

// modal payloads
let lastMetricModal = null;   // {metricId,yPrev,yCur,dimLabel,rows}
let lastFacturaModal = null;  // {key,rows}

// explorer
let explorerChart = null;
let expLastTable = [];
let expLastDim = null;

// ==================== DOM (set in init) ====================

let fileInput, fileNameSpan, errorDiv;
let filterYear, filterStore, filterType, filterCategory;
let toggleNeg, btnResetView;

let kpiVentas, kpiVentasSub, kpiUtilidad, kpiUtilidadSub, kpiMargen, kpiMargenSub, kpiM2, kpiM2Sub, kpiTrans, kpiClientes;

let tablaTopClientes, tablaTopVendedores;
let thYearPrev, thYearCurrent, tablaYoYBody;

let tablaCredit, chartCreditCanvas;

let chartCatsCanvas, tablaCats;

let chipsPool, detalleHeaderRow, detalleFilterRow, detalleTableBody, searchGlobalInput, btnExportDetalle;

let modalBackdrop, modalClose, modalTitle, modalSub, modalYearPrev, modalYearCurrent, modalDimCol, modalTableBody;
let btnModalToDetalle, btnModalExportXlsx, btnModalExportPdf;

let modal2Backdrop, modal2Close, modal2Title, modal2Sub, modal2Tbody;
let btnModal2ToDetalle, btnModal2ExportXlsx, btnModal2ExportPdf;

let expDim, expMetric, expType, expGenerate, expExportXlsx, expToDetalle, expTableBody, expColDim, expCanvas;

let btnPrint, btnClearStorage;

// ==================== VIEW STORAGE ====================

function loadView() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return;
    const v = JSON.parse(raw);

    if (Array.isArray(v.selectedCols) && v.selectedCols.length === 7) {
      selectedCols = v.selectedCols.slice();
    }
    if (v.highlightNeg !== undefined) highlightNeg = !!v.highlightNeg;

    if (filterYear && v.filtros?.year !== undefined) filterYear.value = v.filtros.year;
    if (filterStore && v.filtros?.store !== undefined) filterStore.value = v.filtros.store;
    if (filterType && v.filtros?.tipo !== undefined) filterType.value = v.filtros.tipo;
    if (filterCategory && v.filtros?.categoria !== undefined) filterCategory.value = v.filtros.categoria;

    if (toggleNeg) toggleNeg.checked = highlightNeg;
  } catch (e) {}
}

function saveView() {
  const f = getFiltrosRaw();
  const payload = { selectedCols, highlightNeg, filtros: f };
  localStorage.setItem(LS_KEY, JSON.stringify(payload));
}

function resetView() {
  selectedCols = DEFAULT_SELECTED_COLS.slice();
  highlightNeg = true;
  if (toggleNeg) toggleNeg.checked = true;

  if (filterYear) filterYear.value = "all";
  if (filterStore) filterStore.value = "all";
  if (filterType) filterType.value = "both";
  if (filterCategory) filterCategory.value = "all";

  detalleSort = { col: null, asc: true };
  detalleFiltros = {};
  detalleBusqueda = "";
  if (searchGlobalInput) searchGlobalInput.value = "";

  localStorage.removeItem(LS_KEY);

  initDetalleHeaders();
  actualizarTodo();
}

// ==================== FILE LOAD ====================

function handleFile(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  if (fileNameSpan) fileNameSpan.textContent = file.name;
  if (errorDiv) errorDiv.textContent = "";
  records = [];
  categoriaNameById = {};

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });

      const sheetFactName =
        wb.SheetNames.find(n => n.toLowerCase().includes("factura")) ||
        wb.SheetNames[0];
      const sheetNotasName =
        wb.SheetNames.find(n => n.toLowerCase().includes("nota")) ||
        wb.SheetNames[1];

      const sheetFact = wb.Sheets[sheetFactName];
      const sheetNotas = wb.Sheets[sheetNotasName];

      if (!sheetFact || !sheetNotas) {
        throw new Error("No se encontraron hojas 'facturas' y 'notas'.");
      }

      // header para localizar columna BB (excluir “factura del dia”)
      const headerRows = XLSX.utils.sheet_to_json(sheetFact, { header: 1, range: 0, raw: true });
      const header = headerRows[0] || [];
      const idxBB = XLSX.utils.decode_col("BB");
      const colBB = header[idxBB];

      const rawFact = XLSX.utils.sheet_to_json(sheetFact, { defval: null });
      const rawNotas = XLSX.utils.sheet_to_json(sheetNotas, { defval: null });

      const factRen = rawFact.map(r => renameRow(r, NORM_FACT));
      const notasRen = rawNotas.map(r => renameRow(r, NORM_NOTAS));

      // FACTURAS
      for (let i = 0; i < rawFact.length; i++) {
        const rowOrig = rawFact[i];
        const row = factRen[i];

        const flagTexto = colBB
          ? (rowOrig[colBB] ?? "").toString().trim().toLowerCase()
          : "";

        if (flagTexto === "factura del dia") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;

        const anio = fecha.getFullYear();
        const mes = fecha.getMonth() + 1;

        const subtotal = toNumber(row["Subtotal"]);
        const costo = toNumber(row["Costo"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad = !(flagTexto === "rango" || flagTexto === "base");

        // crédito vs contado (tu lógica original)
        let esCredito = false;
        if (toNumber(row["Descuento"]) === 0 && flagTexto !== "base") esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = safeUpper(row["Almacen"]);
        const folio = safeStr(row["Factura"]);
        const hora = safeStr(row["Hora"]);
        const cliente = safeStr(row["Cliente"]);
        const vendedor = safeStr(row["Vendedor"]);
        const marca = safeStr(row["Marca"]);
        const producto = safeStr(row["Producto"]);
        const cantidad = toNumber(row["Cantidad"]);

        const catId = safeStr(row["ID Categoria"]);
        const catName = safeStr(row["CategoriaNombre"]) || catId || "(Sin categoría)";
        if (catId && catName) categoriaNameById[catId] = catName;

        records.push({
          origen: "factura",
          anio, mes,
          fecha, hora,
          almacen,
          folio,
          cliente,
          vendedor,
          marca,
          producto,
          cantidad,
          categoriaId: catId,
          categoriaNombre: catName,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          esCredito,
          tipoFactura
        });
      }

      // NOTAS (solo REM si existe ese criterio)
      for (let i = 0; i < rawNotas.length; i++) {
        const row = notasRen[i];

        const albaran = safeStr(row["Albaranes"]).toLowerCase();
        if (albaran && albaran !== "rem") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;

        const anio = fecha.getFullYear();
        const mes = fecha.getMonth() + 1;

        const subtotal = toNumber(row["Subtotal"]);
        const costo = toNumber(row["Costo"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad = true;

        let esCredito = false;
        if (toNumber(row["Descuento"]) === 0) esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = safeUpper(row["Almacen"]);
        const folio = safeStr(row["Nota"]);
        const hora = safeStr(row["Hora"]);
        const cliente = safeStr(row["Cliente"]);
        const vendedor = safeStr(row["Vendedor"]);
        const marca = safeStr(row["Marca"]);
        const producto = safeStr(row["Producto"]);
        const cantidad = toNumber(row["Cantidad"]);

        const catId = safeStr(row["ID Categoria"]);
        const catName = safeStr(row["CategoriaNombre"]) || catId || "(Sin categoría)";
        if (catId && catName) categoriaNameById[catId] = catName;

        records.push({
          origen: "nota",
          anio, mes,
          fecha, hora,
          almacen,
          folio,
          cliente,
          vendedor,
          marca,
          producto,
          cantidad,
          categoriaId: catId,
          categoriaNombre: catName,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          esCredito,
          tipoFactura
        });
      }

      if (!records.length) throw new Error("No se generaron registros válidos.");

      yearsDisponibles = unique(records.map(r => r.anio)).sort((a,b)=>a-b);
      almacenesDisponibles = unique(records.map(r => r.almacen || "")).filter(Boolean).sort();
      categoriasDisponibles = unique(records.map(r => r.categoriaNombre || "(Sin categoría)")).sort();

      poblarFiltros();

      // carga vista guardada ya con selects habilitados
      loadView();

      initDetalleChips();
      initDetalleHeaders();
      initExplorerUI();

      actualizarTodo();
      renderExplorer();
    } catch (err) {
      console.error(err);
      if (errorDiv) errorDiv.textContent = "Error al procesar el archivo: " + err.message;
    }
  };

  reader.onerror = function () {
    if (errorDiv) errorDiv.textContent = "No se pudo leer el archivo.";
  };

  reader.readAsArrayBuffer(file);
}

// ==================== FILTERS ====================

function poblarFiltros() {
  if (!filterYear || !filterStore || !filterType || !filterCategory) return;

  filterYear.innerHTML = "<option value='all'>Todos los años</option>";
  yearsDisponibles.forEach(y => {
    const opt = document.createElement("option");
    opt.value = String(y);
    opt.textContent = String(y);
    filterYear.appendChild(opt);
  });

  filterStore.innerHTML = "<option value='all'>Todos los almacenes</option>";
  almacenesDisponibles.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    filterStore.appendChild(opt);
  });

  filterCategory.innerHTML = "<option value='all'>Todas las categorías</option>";
  categoriasDisponibles.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    filterCategory.appendChild(opt);
  });

  filterYear.disabled = false;
  filterStore.disabled = false;
  filterType.disabled = false;
  filterCategory.disabled = false;
}

function getFiltrosRaw() {
  return {
    year: filterYear ? filterYear.value : "all",
    store: filterStore ? filterStore.value : "all",
    tipo: filterType ? filterType.value : "both",
    categoria: filterCategory ? filterCategory.value : "all"
  };
}

function getFiltros() {
  const raw = getFiltrosRaw();
  return {
    year: raw.year === "all" ? null : parseInt(raw.year, 10),
    store: raw.store === "all" ? null : raw.store,
    tipo: raw.tipo || "both",
    categoria: raw.categoria === "all" ? null : raw.categoria
  };
}

function filtrarRecords(ignorarYear) {
  const f = getFiltros();
  return records.filter(r => {
    if (!ignorarYear && f.year && r.anio !== f.year) return false;
    if (f.store && r.almacen !== f.store) return false;
    if (f.categoria && r.categoriaNombre !== f.categoria) return false;
    if (f.tipo === "contado" && r.tipoFactura !== "contado") return false;
    if (f.tipo === "credito" && r.tipoFactura !== "credito") return false;
    return true;
  });
}

function onFilterChange() {
  saveView();
  actualizarTodo();
  renderExplorer();
}

// ==================== TABS ====================

function bindTabs() {
  Array.from(document.querySelectorAll(".tab-btn")).forEach(btn => {
    btn.addEventListener("click", () => {
      const tabId = btn.dataset.tab;
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      const panel = document.getElementById(tabId);
      if (panel) panel.classList.add("active");
    });
  });
}

// ==================== KPIs + RESUMEN CHARTS ====================

function actualizarTodo() {
  if (!records.length) return;
  const datos = filtrarRecords(false);

  actualizarKpis(datos);
  actualizarChartsResumen(datos);
  actualizarTop(datos);
  actualizarYoY();
  actualizarCreditoVsContado(datos);

  // ✅ CATEGORÍAS: SOLO FACTURAS SIN FILTROS (IGNORA "datos")
  actualizarCategoriasSoloFacturasSinFiltros();

  renderDetalle();
}

function actualizarKpis(data) {
  const ventas = sumField(data, "subtotal");
  const utilData = data.filter(r => r.incluirUtilidad);
  const utilidad = sumField(utilData, "utilidad");
  const margen = ventas > 0 ? utilidad / ventas : 0;

  const filtros = getFiltros();
  const yearTxt = filtros.year || "todos los años";
  const almTxt = filtros.store || "todas las sucursales";

  if (kpiVentas) kpiVentas.textContent = formatCurrency(ventas);
  if (kpiVentasSub) kpiVentasSub.textContent = `Periodo: ${yearTxt}, almacén: ${almTxt}`;
  if (kpiUtilidad) kpiUtilidad.textContent = formatCurrency(utilidad);
  if (kpiUtilidadSub) kpiUtilidadSub.textContent = `Basado en ${utilData.length.toLocaleString("es-MX")} filas válidas`;
  if (kpiMargen) kpiMargen.textContent = formatPercent(margen);
  if (kpiMargenSub) kpiMargenSub.textContent = `Margen promedio del periodo filtrado`;

  let m2 = 0;
  if (filtros.store && M2_POR_ALMACEN[filtros.store]) m2 = M2_POR_ALMACEN[filtros.store];
  else m2 = M2_POR_ALMACEN["Todas"];
  const ventasM2 = m2 > 0 ? ventas / m2 : 0;

  if (kpiM2) kpiM2.textContent = `${formatCurrency(ventasM2)}/m²`;
  if (kpiM2Sub) kpiM2Sub.textContent = `${filtros.store || "Todas"} – ${m2.toLocaleString("es-MX")} m²`;

  const opsSet = new Set(data.map(r => r.origen + "|" + r.folio));
  if (kpiTrans) kpiTrans.textContent = opsSet.size.toLocaleString("es-MX");

  const clientesSet = new Set(data.map(r => r.cliente || "").filter(Boolean));
  if (kpiClientes) kpiClientes.textContent = clientesSet.size.toLocaleString("es-MX");
}

function actualizarChartsResumen(data) {
  const canvasMensual = document.getElementById("chart-mensual");
  const canvasAlm = document.getElementById("chart-almacen");
  if (!canvasMensual || !canvasAlm) return;

  const filtros = getFiltros();
  let base = data;

  if (!filtros.year) {
    const maxYear = Math.max(...yearsDisponibles);
    base = data.filter(r => r.anio === maxYear);
  }

  // mensual
  const ventasMes = new Array(12).fill(0);
  base.forEach(r => { const idx = r.mes - 1; if (idx>=0 && idx<12) ventasMes[idx]+=r.subtotal; });
  const labelsMes = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

  if (charts.mensual) charts.mensual.destroy();
  charts.mensual = new Chart(canvasMensual.getContext("2d"), {
    type:"bar",
    data:{ labels:labelsMes, datasets:[{ label:"Ventas", data:ventasMes }]},
    options:{
      responsive:true,
      plugins:{ tooltip:{ callbacks:{ label:(ctx)=> formatCurrency(ctx.raw) } } },
      scales:{ y:{ ticks:{ callback:v=> v.toLocaleString("es-MX") } } }
    }
  });

  // por almacén
  const porAlm = {};
  data.forEach(r => {
    const a = r.almacen || "(sin)";
    if (!porAlm[a]) porAlm[a] = { ventas:0 };
    porAlm[a].ventas += r.subtotal;
  });
  const labs = Object.keys(porAlm);
  const ventas = labs.map(a => porAlm[a].ventas);
  const ventasM2 = labs.map(a => {
    const m2 = M2_POR_ALMACEN[a] || 0;
    return m2>0 ? porAlm[a].ventas / m2 : 0;
  });

  if (charts.almacen) charts.almacen.destroy();
  charts.almacen = new Chart(canvasAlm.getContext("2d"), {
    type:"bar",
    data:{ labels:labs, datasets:[
      { label:"Ventas", data:ventas },
      { label:"Ventas por m²", data:ventasM2 }
    ]},
    options:{
      responsive:true,
      plugins:{ tooltip:{ callbacks:{ label:(ctx)=>{
        const lab = ctx.dataset.label || "";
        if (lab.toLowerCase().includes("m²")) return `${lab}: ${formatCurrency(ctx.raw)}/m²`;
        return `${lab}: ${formatCurrency(ctx.raw)}`;
      }}}},
      scales:{ y:{ ticks:{ callback:v=> v.toLocaleString("es-MX") } } }
    }
  });
}

function actualizarTop(data) {
  if (!tablaTopClientes || !tablaTopVendedores) return;

  const clientes = {};
  data.forEach(r => {
    const c = r.cliente || "(sin cliente)";
    if (!clientes[c]) clientes[c] = { ventas:0, util:0, mAcum:0, n:0 };
    clientes[c].ventas += r.subtotal;
    if (r.incluirUtilidad) {
      clientes[c].util += r.utilidad;
      clientes[c].mAcum += r.subtotal>0 ? r.utilidad/r.subtotal : 0;
      clientes[c].n++;
    }
  });

  const listaC = Object.entries(clientes).map(([nombre,d])=>({
    nombre, ventas:d.ventas, utilidad:d.util, margen:d.n? d.mAcum/d.n : 0
  })).sort((a,b)=>b.ventas-a.ventas).slice(0,5);

  tablaTopClientes.innerHTML = "";
  listaC.forEach((r,idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${r.nombre}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right">${formatCurrency(r.utilidad)}</td>
      <td class="text-right">${formatPercent(r.margen)}</td>
    `;
    tablaTopClientes.appendChild(tr);
  });

  const vend = {};
  const opsPorVend = {};
  data.forEach(r => {
    const v = r.vendedor || "(sin vendedor)";
    if (!vend[v]) vend[v] = { ventas:0, util:0 };
    vend[v].ventas += r.subtotal;
    if (r.incluirUtilidad) vend[v].util += r.utilidad;

    const key = r.origen + "|" + r.folio;
    if (!opsPorVend[v]) opsPorVend[v] = new Set();
    opsPorVend[v].add(key);
  });

  const listaV = Object.entries(vend).map(([nombre,d])=>({
    nombre, ventas:d.ventas, utilidad:d.util, ops:(opsPorVend[nombre]||new Set()).size
  })).sort((a,b)=>b.ventas-a.ventas).slice(0,5);

  tablaTopVendedores.innerHTML = "";
  listaV.forEach((r,idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${r.nombre}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right">${formatCurrency(r.utilidad)}</td>
      <td class="text-right">${r.ops.toLocaleString("es-MX")}</td>
    `;
    tablaTopVendedores.appendChild(tr);
  });
}

// ==================== YOY (MODAL) ====================
// (tu código de YoY + abrirDetalleMetrica queda igual)
// --- Pega aquí tu bloque existente de actualizarYoY() y abrirDetalleMetrica() tal como lo tienes ---
// Nota: No lo cambié porque está bien.

// ==================== CRÉDITO VS CONTADO ====================
// (tu código existente de actualizarCreditoVsContado queda igual)

// ==================== CATEGORÍAS (SOLO FACTURAS, SIN FILTROS) ====================

function actualizarCategoriasSoloFacturasSinFiltros() {
  if (!chartCatsCanvas || !tablaCats) return;
  if (!records.length) return;

  const data = records.filter(r => r.origen === "factura"); // ✅ SOLO FACTURAS
  const map = {};
  data.forEach(r=>{
    const c = r.categoriaNombre || "(Sin categoría)";
    if (!map[c]) map[c] = { ventas:0, util:0 };
    map[c].ventas += r.subtotal;
    if (r.incluirUtilidad) map[c].util += r.utilidad;
  });

  const rows = Object.entries(map).map(([cat,d])=>{
    const m = d.ventas>0 ? d.util/d.ventas : 0;
    return { cat, ventas:d.ventas, util:d.util, margen:m };
  }).sort((a,b)=>b.ventas-a.ventas);

  // table
  tablaCats.innerHTML = "";
  rows.slice(0,50).forEach(r=>{
    const tr = document.createElement("tr");
    tr.style.cursor = "pointer";
    tr.innerHTML = `
      <td>${r.cat}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right ${highlightNeg && r.util<0 ? "neg" : ""}">${formatCurrency(r.util)}</td>
      <td class="text-right ${highlightNeg && r.margen<0 ? "neg" : ""}">${formatPercent(r.margen)}</td>
    `;
    tr.addEventListener("click", ()=>{
      // manda al detalle filtrado por categoría (ahora sí respetando tus filtros globales si quieres)
      if (filterCategory) filterCategory.value = r.cat;
      onFilterChange();
      document.querySelector(`.tab-btn[data-tab="tab-detalle"]`)?.click();
    });
    tablaCats.appendChild(tr);
  });

  const top = rows.slice(0,12);
  if (charts.cats) charts.cats.destroy();
  charts.cats = new Chart(chartCatsCanvas.getContext("2d"), {
    type:"bar",
    data:{ labels: top.map(x=>x.cat), datasets:[{ label:"Ventas", data: top.map(x=>x.ventas) }]},
    options:{
      responsive:true,
      onClick:(evt,elements)=>{
        if (!elements?.length) return;
        const idx = elements[0].index;
        const cat = charts.cats.data.labels[idx];
        if (filterCategory) filterCategory.value = cat;
        onFilterChange();
        document.querySelector(`.tab-btn[data-tab="tab-detalle"]`)?.click();
      },
      plugins:{ tooltip:{ callbacks:{ label:(ctx)=> formatCurrency(ctx.raw) } } },
      scales:{ y:{ ticks:{ callback:v=> v.toLocaleString("es-MX") } } }
    }
  });
}

// ==================== DETALLE (BUILDER 7 COLS) ====================

function initDetalleChips() {
  if (!chipsPool) return;
  chipsPool.innerHTML = "";

  ALL_COLS.forEach(col=>{
    const chip = document.createElement("div");
    chip.className = "chip";
    chip.textContent = col;
    chip.draggable = true;
    chip.dataset.col = col;

    chip.addEventListener("dragstart", (e)=>{
      chip.classList.add("dragging");
      e.dataTransfer.setData("text/plain", col);
    });
    chip.addEventListener("dragend", ()=>{
      chip.classList.remove("dragging");
    });

    chipsPool.appendChild(chip);
  });
}

function initDetalleHeaders() {
  if (!detalleHeaderRow || !detalleFilterRow) return;

  // ✅ NO borres filtros cada vez
  const prevFiltros = { ...detalleFiltros };

  detalleHeaderRow.innerHTML = "";
  detalleFilterRow.innerHTML = "";

  selectedCols.forEach((col, idx)=>{
    // header
    const th = document.createElement("th");
    th.classList.add("sortable","th-drop-target");
    th.dataset.slot = String(idx);

    th.innerHTML = `
      <div class="th-label">
        <span>${col}</span>
        <span class="hint">⇅</span>
      </div>
    `;

    th.addEventListener("click", ()=> sortDetalle(col));

    th.addEventListener("dragover", (e)=>{
      e.preventDefault();
      th.classList.add("drag-over");
    });
    th.addEventListener("dragleave", ()=> th.classList.remove("drag-over"));
    th.addEventListener("drop", (e)=>{
      e.preventDefault();
      th.classList.remove("drag-over");
      const incoming = e.dataTransfer.getData("text/plain");
      if (!incoming) return;
      selectedCols[idx] = incoming;
      saveView();
      initDetalleHeaders();
      renderDetalle();
    });

    detalleHeaderRow.appendChild(th);

    // filter row
    const thF = document.createElement("th");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "filter-input";
    inp.placeholder = "Filtrar...";

    // ✅ recupera filtro anterior para esa col
    inp.value = prevFiltros[col] || "";

    inp.addEventListener("input", debounce(()=>{
      detalleFiltros[col] = inp.value.toLowerCase();
      renderDetalle();
    }, 150));

    thF.appendChild(inp);
    detalleFilterRow.appendChild(thF);
  });

  // ✅ reconstruye el objeto de filtros solo con las columnas actuales (sin perder valores)
  detalleFiltros = {};
  selectedCols.forEach(col => { detalleFiltros[col] = (prevFiltros[col] || ""); });
}

function getValorDetalle(r, col) {
  switch (col) {
    case "Año": return r.anio;
    case "Fecha": return yyyymmdd(r.fecha);
    case "Hora": return r.hora || "";
    case "Almacén": return r.almacen || "";
    case "Factura/Nota": return r.folio || "";
    case "Origen": return r.origen;
    case "Cliente": return r.cliente || "";
    case "Categoría": return r.categoriaNombre || "";
    case "Marca": return r.marca || "";
    case "Vendedor": return r.vendedor || "";
    case "Producto": return r.producto || "";
    case "Cantidad": return r.cantidad || 0;
    case "Subtotal": return r.subtotal || 0;
    case "Costo": return r.costo || 0;
    case "Utilidad": return r.utilidad || 0;
    case "Margen %": return r.subtotal > 0 ? (r.utilidad / r.subtotal) : 0;
    case "Precio sin IVA":
      return (r.cantidad > 0) ? (r.subtotal / r.cantidad) : 0;
    default: return "";
  }
}

function isNumericCol(col){
  return ["Cantidad","Subtotal","Costo","Utilidad","Margen %","Precio sin IVA","Año"].includes(col);
}

function sortDetalle(col) {
  if (detalleSort.col === col) detalleSort.asc = !detalleSort.asc;
  else { detalleSort.col = col; detalleSort.asc = true; }
  renderDetalle();
}

function renderDetalle() {
  if (!detalleTableBody) return;
  if (!records.length) { detalleTableBody.innerHTML = ""; return; }

  let arr = filtrarRecords(false).slice();

  // filtros por columna
  arr = arr.filter(r=>{
    for (const col of selectedCols) {
      const txt = (detalleFiltros[col] || "").trim();
      if (!txt) continue;
      const v = getValorDetalle(r, col);
      const s = (typeof v === "number") ? String(v) : String(v || "").toLowerCase();
      if (!s.includes(txt)) return false;
    }
    return true;
  });

  // global
  if (detalleBusqueda) {
    arr = arr.filter(r=>{
      const campos = [
        r.cliente, r.vendedor, r.folio, r.almacen,
        r.categoriaNombre, r.producto, r.marca
      ];
      return campos.some(c => c && c.toString().toLowerCase().includes(detalleBusqueda));
    });
  }

  // sort
  if (detalleSort.col) {
    const col = detalleSort.col;
    const asc = detalleSort.asc;
    arr.sort((a,b)=>{
      const va = getValorDetalle(a,col);
      const vb = getValorDetalle(b,col);

      if (isNumericCol(col)) return asc ? (toNumber(va)-toNumber(vb)) : (toNumber(vb)-toNumber(va));
      const sa = String(va || "");
      const sb = String(vb || "");
      return asc ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }

  detalleTableBody.innerHTML = "";
  arr.forEach(r=>{
    const tr = document.createElement("tr");

    selectedCols.forEach(col=>{
      const td = document.createElement("td");
      let v = getValorDetalle(r, col);

      if (col === "Factura/Nota") {
        td.style.cursor = "pointer";
        td.style.textDecoration = "underline";
        td.addEventListener("click", ()=> abrirDetalleFactura(r.origen, r.folio, r.cliente));
      }

      if (col === "Subtotal" || col === "Costo" || col === "Utilidad" || col === "Precio sin IVA") {
        td.classList.add("text-right");
        const num = toNumber(v);
        td.textContent = formatCurrency(num);
        if (highlightNeg && num < 0) td.classList.add("neg");
      } else if (col === "Margen %") {
        td.classList.add("text-right");
        td.textContent = formatPercent(v);
        if (highlightNeg && toNumber(v) < 0) td.classList.add("neg");
      } else if (col === "Cantidad" || col === "Año") {
        td.classList.add("text-right");
        td.textContent = Number(v || 0).toLocaleString("es-MX");
      } else {
        td.textContent = v || "";
      }

      tr.appendChild(td);
    });

    detalleTableBody.appendChild(tr);
  });
}

// ==================== MODAL FACTURA (LÍNEAS) ====================
// (tu código existente queda igual)

// ==================== EXPLORER ====================
// (tu código existente queda igual)

// ==================== DOWNLOAD CHART PNG ====================

function downloadCanvasPNG(canvasId) {
  const c = document.getElementById(canvasId);
  if (!c) return;
  const a = document.createElement("a");
  a.href = c.toDataURL("image/png");
  a.download = `${canvasId}.png`;
  a.click();
}

// ==================== INIT (DOM READY) ====================

window.addEventListener("DOMContentLoaded", () => {
  // DOM refs
  fileInput = document.getElementById("file-input");
  fileNameSpan = document.getElementById("file-name");
  errorDiv = document.getElementById("error");

  filterYear = document.getElementById("filter-year");
  filterStore = document.getElementById("filter-store");
  filterType = document.getElementById("filter-type");
  filterCategory = document.getElementById("filter-category");

  toggleNeg = document.getElementById("toggle-neg");
  btnResetView = document.getElementById("btn-reset-view");

  kpiVentas = document.getElementById("kpi-ventas");
  kpiVentasSub = document.getElementById("kpi-ventas-sub");
  kpiUtilidad = document.getElementById("kpi-utilidad");
  kpiUtilidadSub = document.getElementById("kpi-utilidad-sub");
  kpiMargen = document.getElementById("kpi-margen");
  kpiMargenSub = document.getElementById("kpi-margen-sub");
  kpiM2 = document.getElementById("kpi-m2");
  kpiM2Sub = document.getElementById("kpi-m2-sub");
  kpiTrans = document.getElementById("kpi-trans");
  kpiClientes = document.getElementById("kpi-clientes");

  tablaTopClientes = document.getElementById("tabla-top-clientes");
  tablaTopVendedores = document.getElementById("tabla-top-vendedores");

  thYearPrev = document.getElementById("th-year-prev");
  thYearCurrent = document.getElementById("th-year-current");
  tablaYoYBody = document.getElementById("tabla-yoy");

  tablaCredit = document.getElementById("tabla-credit");
  chartCreditCanvas = document.getElementById("chart-credit");

  chartCatsCanvas = document.getElementById("chart-cats");
  tablaCats = document.getElementById("tabla-cats");

  chipsPool = document.getElementById("chips-pool");
  detalleHeaderRow = document.getElementById("detalle-header-row");
  detalleFilterRow = document.getElementById("detalle-filter-row");
  detalleTableBody = document.getElementById("tabla-detalle");
  searchGlobalInput = document.getElementById("search-global");
  btnExportDetalle = document.getElementById("btn-export-detalle");

  modalBackdrop = document.getElementById("modal-backdrop");
  modalClose = document.getElementById("modal-close");
  modalTitle = document.getElementById("modal-title");
  modalSub = document.getElementById("modal-sub");
  modalYearPrev = document.getElementById("modal-year-prev");
  modalYearCurrent = document.getElementById("modal-year-current");
  modalDimCol = document.getElementById("modal-dim-col");
  modalTableBody = document.querySelector("#modal-table tbody");

  btnModalToDetalle = document.getElementById("btn-modal-to-detalle");
  btnModalExportXlsx = document.getElementById("btn-modal-export-xlsx");
  btnModalExportPdf = document.getElementById("btn-modal-export-pdf");

  modal2Backdrop = document.getElementById("modal2-backdrop");
  modal2Close = document.getElementById("modal2-close");
  modal2Title = document.getElementById("modal2-title");
  modal2Sub = document.getElementById("modal2-sub");
  modal2Tbody = document.querySelector("#modal2-table tbody");

  btnModal2ToDetalle = document.getElementById("btn-modal2-to-detalle");
  btnModal2ExportXlsx = document.getElementById("btn-modal2-export-xlsx");
  btnModal2ExportPdf = document.getElementById("btn-modal2-export-pdf");

  expDim = document.getElementById("exp-dim");
  expMetric = document.getElementById("exp-metric");
  expType = document.getElementById("exp-type");
  expGenerate = document.getElementById("exp-generate");
  expExportXlsx = document.getElementById("exp-export-xlsx");
  expToDetalle = document.getElementById("exp-to-detalle");
  expTableBody = document.getElementById("exp-table");
  expColDim = document.getElementById("exp-col-dim");
  expCanvas = document.getElementById("chart-explorer");

  btnPrint = document.getElementById("btn-print");
  btnClearStorage = document.getElementById("btn-clear-storage");

  // listeners base
  bindTabs();

  if (fileInput) fileInput.addEventListener("change", handleFile);

  if (filterYear) filterYear.addEventListener("change", onFilterChange);
  if (filterStore) filterStore.addEventListener("change", onFilterChange);
  if (filterType) filterType.addEventListener("change", onFilterChange);
  if (filterCategory) filterCategory.addEventListener("change", onFilterChange);

  if (toggleNeg) {
    toggleNeg.addEventListener("change", () => {
      highlightNeg = !!toggleNeg.checked;
      saveView();
      renderDetalle();
      actualizarCategoriasSoloFacturasSinFiltros(); // ✅ refresca tab categorías también
    });
  }

  if (btnResetView) btnResetView.addEventListener("click", resetView);

  // ✅ listener global search SOLO UNA VEZ (no en initDetalleHeaders)
  if (searchGlobalInput) {
    searchGlobalInput.addEventListener("input", debounce(()=>{
      detalleBusqueda = searchGlobalInput.value.toLowerCase();
      renderDetalle();
    }, 200));
  }

  // download buttons
  document.querySelectorAll("[data-download-chart]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const id = btn.getAttribute("data-download-chart");
      downloadCanvasPNG(id);
    });
  });

  if (btnPrint) btnPrint.addEventListener("click", () => window.print());

  if (btnClearStorage) {
    btnClearStorage.addEventListener("click", () => {
      localStorage.removeItem(LS_KEY);
      alert("Configuración borrada. Recarga la página para reiniciar.");
    });
  }

  // init UI (sin archivo)
  initDetalleChips();
  initDetalleHeaders();
  initExplorerUI();
  renderDetalle();
});
