// ==================== CONFIGURACIÓN ====================

// m² por almacén
const M2_POR_ALMACEN = {
  "GENERAL": 1538,
  "EXPRESS": 369,
  "SAN AGUST": 870,
  "ADELITAS": 348,
  "H ILUSTRES": 100,
  "Todas": 3225
};

// Mapeos de columnas (acepta variaciones por normalización)
const RENAME_FACTURAS = {
  "No_fac": "Factura",
  "Falta_fac": "Fecha",
  "Hora_fac": "Hora",
  "Descuento": "Descuento ($)",
  "Subt_fac": "Subtotal",
  "Total_fac": "Total Factura",
  "Iva": "IVA",
  "Cve_cte": "ID Cliente",
  "Nom_cte": "Cliente",
  "Cse_prod": "ID Categoria",
  "Categoria": "Categoria Nombre",
  "Nom_cat": "Categoria Nombre",
  "Cve_prod": "Clave",
  "Desc_prod": "Producto",
  "Cant_surt": "Cantidad",
  "Pz": "Cantidad",
  "Dcto1": "Descuento (%)",
  "Cve_age": "ID Vendedor",
  "Nom_age": "Vendedor",
  "Lugar": "Almacen",
  "Des_tial": "Marca",
  "Cto_ent": "Costo"
};

const RENAME_NOTAS = {
  "No_fac": "Nota",
  "Cve_suc": "Albaranes",
  "Falta_fac": "Fecha",
  "Hravta": "Hora",
  "Lugar": "Almacen",
  "Descuento": "Descuento ($)",
  "Cve_prod": "Clave",
  "Desc_prod": "Producto",
  "Cant_surt": "Cantidad",
  "Subt_prod": "Subtotal",
  "Iva_prod": "IVA",
  "Total_fac": "Total Nota",
  "Dcto1": "Descuento (%)",
  "Cse_prod": "ID Categoria",
  "Categoria": "Categoria Nombre",
  "Nom_cat": "Categoria Nombre",
  "Nom_fac": "Cliente",
  "Nom_age": "Vendedor",
  "Des_tial": "Marca",
  "Cto_ent": "Costo Entrada",
  "Cto_entr": "Costo Entrada",
  "Cte_fac": "ID Cliente"
};

// DETALLE: paleta completa
const DETALLE_ALL_COLS = [
  "Año",
  "Fecha",
  "Hora",
  "Almacén",
  "Factura/Nota",
  "Cliente",
  "Categoría",
  "Producto",
  "Tipo factura",
  "Subtotal",
  "Costo",
  "Descuento",
  "Utilidad",
  "Margen %",
  "Marca",
  "Vendedor"
];

const DETALLE_DEFAULT_COLS = [
  "Año",
  "Cliente",
  "Producto",
  "Costo",
  "Subtotal",
  "Margen %",
  "Utilidad"
];

// ==================== UTILIDADES ====================

function $(id) { return document.getElementById(id); }

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
  if (v === null || v === undefined || v === "") return 0;
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
  const x = Number(d || 0);
  if (!isFinite(x)) return "0.0%";
  return (x * 100).toFixed(1) + "%";
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

function monthLabels() {
  return ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
}

function downloadDataUrl(dataUrl, filename) {
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function chartToPng(chart, filename) {
  if (!chart) return;
  const dataUrl = chart.toBase64Image();
  downloadDataUrl(dataUrl, filename);
}

function exportJsonToXlsx(rows, filename, sheetName = "Datos") {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function exportTableToPdf(title, headers, rows, filename) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  doc.setFontSize(12);
  doc.text(title, 40, 30);

  doc.autoTable({
    startY: 45,
    head: [headers],
    body: rows,
    styles: { fontSize: 8 },
    headStyles: { fillColor: [10, 20, 40] },
    theme: "grid"
  });

  doc.save(filename);
}

// ==================== ESTADO ====================

let records = [];
let yearsDisponibles = [];
let almacenesDisponibles = [];
let categoriasDisponibles = [];

let charts = {
  mensual: null,
  almacen: null,
  credito: null,
  categorias: null,
  explorer: null
};

// Vista guardada (config rápida + performance)
const LS_KEYS = {
  detailLayout: "cedro_detalle_layout_v1",
  quickConfig: "cedro_quick_config_v1",
  perfConfig: "cedro_perf_config_v1"
};

let cfg = {
  maxDetalleRows: 5000,
  saveView: true
};

let quick = {
  metric: "ventas",
  chart: "bar",
  mode: "normal",
  negOnly: false
};

// DETALLE configurable
let detalleColsLayout = [];
let detalleSort = { col: null, asc: true };
let detalleFiltros = {};
let detalleBusqueda = "";
let detalleDragCol = null;

// Modal (YoY) contexto
let modalCtx = {
  metricId: null,
  yPrev: null,
  yCur: null,
  rows: [],
  dimTitle: "Categoría"
};

// Modal (Factura/Nota) contexto
let docCtx = {
  origen: null,
  folio: null,
  rows: [],
  meta: {}
};

// ==================== DOM ====================

const fileInput = $("file-input");
const fileNameSpan = $("file-name");
const errorDiv = $("error");

const filterYear = $("filter-year");
const filterStore = $("filter-store");
const filterType = $("filter-type");
const filterCategory = $("filter-category");

const qcMetric = $("qc-metric");
const qcChart = $("qc-chart");
const qcMode = $("qc-mode");
const qcNegOnly = $("qc-neg-only");

const kpiVentas = $("kpi-ventas");
const kpiVentasSub = $("kpi-ventas-sub");
const kpiUtilidad = $("kpi-utilidad");
const kpiUtilidadSub = $("kpi-utilidad-sub");
const kpiMargen = $("kpi-margen");
const kpiM2 = $("kpi-m2");
const kpiM2Sub = $("kpi-m2-sub");
const kpiTrans = $("kpi-trans");
const kpiClientes = $("kpi-clientes");

const tablaTopClientes = $("tabla-top-clientes");
const tablaTopVendedores = $("tabla-top-vendedores");

const thYearPrev = $("th-year-prev");
const thYearCurrent = $("th-year-current");
const tablaYoYBody = $("tabla-yoy");

const tablaCreditoBody = $("tabla-credito");
const tablaCategoriasBody = $("tabla-categorias");

const detalleChipsContainer = $("detalle-chips");
const detalleHeaderRow = $("detalle-header-row");
const detalleFilterRow = $("detalle-filter-row");
const detalleTableBody = $("tabla-detalle");
const searchGlobalInput = $("search-global");

const modalBackdrop = $("modal-backdrop");
const modalClose = $("modal-close");
const modalTitle = $("modal-title");
const modalSub = $("modal-sub");
const modalYearPrev = $("modal-year-prev");
const modalYearCurrent = $("modal-year-current");
const modalDimTitle = $("modal-dim-title");
const modalTableBody = document.querySelector("#modal-table tbody");

const modalExportXlsx = $("modal-export-xlsx");
const modalExportPdf = $("modal-export-pdf");
const modalGoDetalle = $("modal-go-detalle");

const docBackdrop = $("doc-backdrop");
const docClose = $("doc-close");
const docTitle = $("doc-title");
const docSub = $("doc-sub");
const docBody = $("doc-body");
const docExportXlsx = $("doc-export-xlsx");
const docExportPdf = $("doc-export-pdf");
const docGoDetalle = $("doc-go-detalle");

const btnExportChartMensual = $("btn-export-chart-mensual");
const btnExportChartAlmacen = $("btn-export-chart-almacen");
const btnExportChartCredito = $("btn-export-chart-credito");
const btnExportChartCategorias = $("btn-export-chart-categorias");

const chartMainTitle = $("chart-main-title");

const expDim = $("exp-dim");
const expMet = $("exp-met");
const expType = $("exp-type");
const expTop = $("exp-top");
const expRefresh = $("exp-refresh");
const expExportPng = $("exp-export-png");
const expTitle = $("exp-title");
const expThDim = $("exp-th-dim");
const expThMet = $("exp-th-met");
const tablaExplorer = $("tabla-explorer");

const cfgMaxRows = $("cfg-max-rows");
const cfgSaveView = $("cfg-save-view");
const cfgResetView = $("cfg-reset-view");
const cfgClearStorage = $("cfg-clear-storage");
const cfgExportDetalleXlsx = $("cfg-export-detalle-xlsx");
const cfgExportResumenPdf = $("cfg-export-resumen-pdf");

// ==================== LOCALSTORAGE ====================

function loadConfigs() {
  // performance
  try {
    const raw = localStorage.getItem(LS_KEYS.perfConfig);
    if (raw) {
      const v = JSON.parse(raw);
      if (typeof v.maxDetalleRows === "number") cfg.maxDetalleRows = v.maxDetalleRows;
      if (typeof v.saveView === "boolean") cfg.saveView = v.saveView;
    }
  } catch {}

  // quick config
  try {
    const raw = localStorage.getItem(LS_KEYS.quickConfig);
    if (raw) {
      const v = JSON.parse(raw);
      if (v.metric) quick.metric = v.metric;
      if (v.chart) quick.chart = v.chart;
      if (v.mode) quick.mode = v.mode;
      if (typeof v.negOnly === "boolean") quick.negOnly = v.negOnly;
    }
  } catch {}
}

function saveConfigs() {
  if (!cfg.saveView) return;
  try {
    localStorage.setItem(LS_KEYS.perfConfig, JSON.stringify(cfg));
    localStorage.setItem(LS_KEYS.quickConfig, JSON.stringify(quick));
    localStorage.setItem(LS_KEYS.detailLayout, JSON.stringify(detalleColsLayout));
  } catch {}
}

function clearSaved() {
  localStorage.removeItem(LS_KEYS.perfConfig);
  localStorage.removeItem(LS_KEYS.quickConfig);
  localStorage.removeItem(LS_KEYS.detailLayout);
}

// ==================== ARCHIVO ====================

if (fileInput) fileInput.addEventListener("change", handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  fileNameSpan.textContent = file.name;
  errorDiv.textContent = "";
  records = [];

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });

      const sheetFactName =
        wb.SheetNames.find(n => n.toLowerCase().includes("factura")) || wb.SheetNames[0];
      const sheetNotasName =
        wb.SheetNames.find(n => n.toLowerCase().includes("nota")) || wb.SheetNames[1];

      const sheetFact = wb.Sheets[sheetFactName];
      const sheetNotas = wb.Sheets[sheetNotasName];

      if (!sheetFact || !sheetNotas) throw new Error("No se encontraron hojas 'facturas' y 'notas'.");

      // Encabezados para localizar columna BB (Factura del día / Rango / Base)
      const headerRows = XLSX.utils.sheet_to_json(sheetFact, { header: 1, range: 0, raw: true });
      const header = headerRows[0] || [];
      const idxBB = XLSX.utils.decode_col("BB");
      const colBB = header[idxBB];

      const rawFact = XLSX.utils.sheet_to_json(sheetFact, { defval: null, raw: true });
      const rawNotas = XLSX.utils.sheet_to_json(sheetNotas, { defval: null, raw: true });

      const factRen = rawFact.map(r => renameRow(r, NORM_FACT));
      const notasRen = rawNotas.map(r => renameRow(r, NORM_NOTAS));

      // FACTURAS (línea a línea)
      for (let i = 0; i < rawFact.length; i++) {
        const rowOrig = rawFact[i];
        const row = factRen[i];

        const flagTexto = colBB ? (rowOrig[colBB] ?? "").toString().trim().toLowerCase() : "";
        if (flagTexto === "factura del dia") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;

        const anio = fecha.getFullYear();
        const mes = fecha.getMonth() + 1;

        const subtotal = toNumber(row["Subtotal"]);
        const costo = toNumber(row["Costo"]);
        const descuentoMonto = toNumber(row["Descuento ($)"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad = !(flagTexto === "rango" || flagTexto === "base");

        let esCredito = false;
        if (descuentoMonto === 0 && flagTexto !== "base") esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (row["Almacen"] || "").toString().trim().toUpperCase();

        const categoriaId = (row["ID Categoria"] || "").toString().trim();
        const categoriaNombre = (row["Categoria Nombre"] || "").toString().trim();
        const categoriaLabel = categoriaNombre || categoriaId || "(Sin categoría)";

        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Factura"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();
        const hora = (row["Hora"] || "").toString().trim();
        const producto = (row["Producto"] || "").toString().trim();
        const cantidad = toNumber(row["Cantidad"]);

        records.push({
          origen: "factura",
          anio, mes, fecha,
          hora,
          almacen,
          categoriaId,
          categoriaNombre,
          categoria: categoriaLabel,
          cliente,
          vendedor,
          marca,
          folio,
          producto,
          cantidad,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura
        });
      }

      // NOTAS (solo REM)
      for (let i = 0; i < rawNotas.length; i++) {
        const row = notasRen[i];

        const albaran = (row["Albaranes"] || "").toString().trim().toLowerCase();
        if (albaran && albaran !== "rem") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;

        const anio = fecha.getFullYear();
        const mes = fecha.getMonth() + 1;

        const subtotal = toNumber(row["Subtotal"]);
        const costo = toNumber(row["Costo Entrada"]);
        const descuentoMonto = toNumber(row["Descuento ($)"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad = true;

        let esCredito = false;
        if (descuentoMonto === 0) esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (row["Almacen"] || "").toString().trim().toUpperCase();

        const categoriaId = (row["ID Categoria"] || "").toString().trim();
        const categoriaNombre = (row["Categoria Nombre"] || "").toString().trim();
        const categoriaLabel = categoriaNombre || categoriaId || "(Sin categoría)";

        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Nota"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();
        const hora = (row["Hora"] || "").toString().trim();
        const producto = (row["Producto"] || "").toString().trim();
        const cantidad = toNumber(row["Cantidad"]);

        records.push({
          origen: "nota",
          anio, mes, fecha,
          hora,
          almacen,
          categoriaId,
          categoriaNombre,
          categoria: categoriaLabel,
          cliente,
          vendedor,
          marca,
          folio,
          producto,
          cantidad,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura
        });
      }

      if (!records.length) throw new Error("No se generaron registros válidos.");

      yearsDisponibles = unique(records.map(r => r.anio)).sort((a, b) => a - b);
      almacenesDisponibles = unique(records.map(r => r.almacen || "").filter(Boolean));
      categoriasDisponibles = unique(records.map(r => r.categoria || "").filter(Boolean)).sort((a,b)=>a.localeCompare(b));

      habilitarUI();
      poblarFiltros();
      initDetalleUI();
      initExplorerUI();
      actualizarTodo();

    } catch (err) {
      console.error(err);
      errorDiv.textContent = "Error al procesar el archivo: " + err.message;
    }
  };

  reader.onerror = function () {
    errorDiv.textContent = "No se pudo leer el archivo.";
  };

  reader.readAsArrayBuffer(file);
}

// ==================== FILTROS ====================

function habilitarUI() {
  if (filterYear) filterYear.disabled = false;
  if (filterStore) filterStore.disabled = false;
  if (filterType) filterType.disabled = false;
  if (filterCategory) filterCategory.disabled = false;

  if (qcMetric) qcMetric.disabled = false;
  if (qcChart) qcChart.disabled = false;
  if (qcMode) qcMode.disabled = false;

  if (expDim) expDim.disabled = false;
  if (expMet) expMet.disabled = false;
  if (expType) expType.disabled = false;
  if (expTop) expTop.disabled = false;
  if (expRefresh) expRefresh.disabled = false;
  if (expExportPng) expExportPng.disabled = false;
}

function poblarFiltros() {
  filterYear.innerHTML = "<option value='all'>Todos los años</option>";
  yearsDisponibles.forEach(y => {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
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
}

function getFiltros() {
  return {
    year: filterYear && filterYear.value !== "all" ? parseInt(filterYear.value, 10) : null,
    store: filterStore && filterStore.value !== "all" ? filterStore.value : null,
    tipo: filterType ? filterType.value : "both",
    categoria: filterCategory && filterCategory.value !== "all" ? filterCategory.value : null
  };
}

function filtrarRecords(ignorarYear) {
  const f = getFiltros();
  return records.filter(r => {
    if (!ignorarYear && f.year && r.anio !== f.year) return false;
    if (f.store && r.almacen !== f.store) return false;
    if (f.categoria && r.categoria !== f.categoria) return false;
    if (f.tipo === "contado" && r.tipoFactura !== "contado") return false;
    if (f.tipo === "credito" && r.tipoFactura !== "credito") return false;
    return true;
  });
}

if (filterYear) filterYear.addEventListener("change", actualizarTodo);
if (filterStore) filterStore.addEventListener("change", actualizarTodo);
if (filterType) filterType.addEventListener("change", actualizarTodo);
if (filterCategory) filterCategory.addEventListener("change", actualizarTodo);

// ==================== CONFIG RÁPIDA ====================

function applyQuickToUI() {
  if (qcMetric) qcMetric.value = quick.metric;
  if (qcChart) qcChart.value = quick.chart;
  if (qcMode) qcMode.value = quick.mode;
  if (qcNegOnly) qcNegOnly.checked = quick.negOnly;

  document.body.classList.toggle("compact-mode", quick.mode === "compact");
}

function initQuickConfig() {
  if (qcMetric) qcMetric.addEventListener("change", () => {
    quick.metric = qcMetric.value;
    saveConfigs();
    actualizarTodo();
  });

  if (qcChart) qcChart.addEventListener("change", () => {
    quick.chart = qcChart.value;
    saveConfigs();
    actualizarTodo();
  });

  if (qcMode) qcMode.addEventListener("change", () => {
    quick.mode = qcMode.value;
    applyQuickToUI();
    saveConfigs();
    renderDetalle();
  });

  if (qcNegOnly) qcNegOnly.addEventListener("change", () => {
    quick.negOnly = !!qcNegOnly.checked;
    saveConfigs();
    renderDetalle();
  });
}

// ==================== KPIs / GRÁFICAS / TOP ====================

function actualizarTodo() {
  if (!records.length) return;

  const datos = filtrarRecords(false);
  actualizarKpis(datos);
  actualizarGraficasResumen(datos);
  actualizarTop(datos);
  actualizarYoY();
  actualizarCredito(datos);
  actualizarCategorias(datos);
  renderDetalle();
  actualizarExplorador();
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
  if (kpiUtilidadSub) kpiUtilidadSub.textContent = `Basado en ${utilData.length} filas válidas`;
  if (kpiMargen) kpiMargen.textContent = formatPercent(margen);

  let m2 = 0;
  if (filtros.store && M2_POR_ALMACEN[filtros.store]) m2 = M2_POR_ALMACEN[filtros.store];
  else m2 = M2_POR_ALMACEN["Todas"];

  const ventasM2 = m2 > 0 ? ventas / m2 : 0;
  if (kpiM2) kpiM2.textContent = formatCurrency(ventasM2) + "/m²";
  if (kpiM2Sub) kpiM2Sub.textContent = `${filtros.store || "Todas"} – ${m2.toLocaleString("es-MX")} m²`;

  const opsSet = new Set(data.map(r => r.origen + "|" + r.folio));
  if (kpiTrans) kpiTrans.textContent = opsSet.size.toLocaleString("es-MX");

  const clientesSet = new Set(data.map(r => r.cliente || "").filter(x => x));
  if (kpiClientes) kpiClientes.textContent = clientesSet.size.toLocaleString("es-MX");
}

function buildMainMetric(data) {
  // para gráfica mensual principal
  if (quick.metric === "ventas") return { title: "Ventas por mes", get: r => r.subtotal, fmt: formatCurrency };
  if (quick.metric === "utilidad") return { title: "Utilidad por mes", get: r => (r.incluirUtilidad ? r.utilidad : 0), fmt: formatCurrency };
  if (quick.metric === "margen") return {
    title: "Margen % por mes",
    get: null,
    fmt: formatPercent
  };
  if (quick.metric === "credito") return {
    title: "Ventas a crédito por mes",
    get: r => (r.esCredito ? r.subtotal : 0),
    fmt: formatCurrency
  };
  return { title: "Ventas por mes", get: r => r.subtotal, fmt: formatCurrency };
}

function actualizarGraficasResumen(data) {
  const canvasMensual = $("chart-mensual");
  const canvasAlm = $("chart-almacen");
  if (!canvasMensual || !canvasAlm) return;

  // base por año: si no elige año, usa el máximo
  const filtros = getFiltros();
  let base = data;
  let y = filtros.year;
  if (!y) y = Math.max(...yearsDisponibles);
  base = data.filter(r => r.anio === y);

  const metric = buildMainMetric(base);
  if (chartMainTitle) chartMainTitle.textContent = metric.title;

  const labelsMes = monthLabels();

  // Mensual
  let values = new Array(12).fill(0);
  if (quick.metric === "margen") {
    // margen mes = utilidad/ventas
    const ventasMes = new Array(12).fill(0);
    const utilMes = new Array(12).fill(0);
    base.forEach(r => {
      const idx = r.mes - 1;
      if (idx < 0 || idx > 11) return;
      ventasMes[idx] += r.subtotal;
      if (r.incluirUtilidad) utilMes[idx] += r.utilidad;
    });
    values = ventasMes.map((v, i) => (v > 0 ? utilMes[i] / v : 0));
  } else {
    base.forEach(r => {
      const idx = r.mes - 1;
      if (idx < 0 || idx > 11) return;
      values[idx] += metric.get(r);
    });
  }

  const ctxM = canvasMensual.getContext("2d");
  if (charts.mensual) charts.mensual.destroy();

  const typeMensual = (quick.chart === "pie" || quick.chart === "doughnut") ? "bar" : quick.chart;

  charts.mensual = new Chart(ctxM, {
    type: typeMensual,
    data: {
      labels: labelsMes,
      datasets: [{ label: metric.title, data: values }]
    },
    options: {
      responsive: true,
      scales: (typeMensual === "pie" || typeMensual === "doughnut") ? {} : {
        y: {
          ticks: {
            callback: v => (quick.metric === "margen" ? (v*100).toFixed(1) + "%" : v.toLocaleString("es-MX"))
          }
        }
      }
    }
  });

  // Por almacén
  const porAlm = {};
  data.forEach(r => {
    const a = r.almacen || "(sin)";
    if (!porAlm[a]) porAlm[a] = { ventas: 0 };
    porAlm[a].ventas += r.subtotal;
  });

  const labsAlm = Object.keys(porAlm);
  const ventasAlm = labsAlm.map(a => porAlm[a].ventas);
  const ventasM2Alm = labsAlm.map(a => {
    const m2 = M2_POR_ALMACEN[a] || 0;
    return m2 > 0 ? porAlm[a].ventas / m2 : 0;
  });

  const ctxA = canvasAlm.getContext("2d");
  if (charts.almacen) charts.almacen.destroy();
  charts.almacen = new Chart(ctxA, {
    type: "bar",
    data: {
      labels: labsAlm,
      datasets: [
        { label: "Ventas", data: ventasAlm },
        { label: "Ventas por m²", data: ventasM2Alm }
      ]
    },
    options: {
      responsive: true,
      scales: { y: { ticks: { callback: v => v.toLocaleString("es-MX") } } }
    }
  });
}

function actualizarTop(data) {
  if (!tablaTopClientes || !tablaTopVendedores) return;

  // Clientes
  const clientes = {};
  data.forEach(r => {
    const c = r.cliente || "(sin cliente)";
    if (!clientes[c]) clientes[c] = { ventas: 0, util: 0, mAcum: 0, n: 0 };
    clientes[c].ventas += r.subtotal;
    if (r.incluirUtilidad) {
      clientes[c].util += r.utilidad;
      clientes[c].mAcum += r.subtotal > 0 ? r.utilidad / r.subtotal : 0;
      clientes[c].n++;
    }
  });

  const listaC = Object.entries(clientes)
    .map(([nombre, d]) => ({
      nombre,
      ventas: d.ventas,
      utilidad: d.util,
      margen: d.n ? d.mAcum / d.n : 0
    }))
    .sort((a, b) => b.ventas - a.ventas)
    .slice(0, 5);

  tablaTopClientes.innerHTML = "";
  listaC.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${r.nombre}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right">${formatCurrency(r.utilidad)}</td>
      <td class="text-right">${formatPercent(r.margen)}</td>
    `;
    tablaTopClientes.appendChild(tr);
  });

  // Vendedores
  const vend = {};
  const opsPorVend = {};
  data.forEach(r => {
    const v = r.vendedor || "(sin vendedor)";
    if (!vend[v]) vend[v] = { ventas: 0, util: 0 };
    vend[v].ventas += r.subtotal;
    if (r.incluirUtilidad) vend[v].util += r.utilidad;

    const key = r.origen + "|" + r.folio;
    if (!opsPorVend[v]) opsPorVend[v] = new Set();
    opsPorVend[v].add(key);
  });

  const listaV = Object.entries(vend)
    .map(([nombre, d]) => ({
      nombre,
      ventas: d.ventas,
      utilidad: d.util,
      ops: (opsPorVend[nombre] || new Set()).size
    }))
    .sort((a, b) => b.ventas - a.ventas)
    .slice(0, 5);

  tablaTopVendedores.innerHTML = "";
  listaV.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${r.nombre}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right">${formatCurrency(r.utilidad)}</td>
      <td class="text-right">${r.ops.toLocaleString("es-MX")}</td>
    `;
    tablaTopVendedores.appendChild(tr);
  });
}

// ==================== YOY ====================

function actualizarYoY() {
  if (!tablaYoYBody || !thYearPrev || !thYearCurrent) return;

  const datos = filtrarRecords(true);
  if (!datos.length) {
    tablaYoYBody.innerHTML = "<tr><td colspan='4' class='muted'>No hay datos.</td></tr>";
    return;
  }

  const years = unique(datos.map(r => r.anio)).sort((a, b) => a - b);
  if (years.length < 2) {
    tablaYoYBody.innerHTML = "<tr><td colspan='4' class='muted'>Se requiere al menos 2 años.</td></tr>";
    return;
  }

  const yCur = years[years.length - 1];
  const yPrev = years[years.length - 2];

  thYearPrev.textContent = yPrev;
  thYearCurrent.textContent = yCur;

  const dataPrev = datos.filter(r => r.anio === yPrev);
  const dataCur = datos.filter(r => r.anio === yCur);

  const ventasPrev = sumField(dataPrev, "subtotal");
  const ventasCur = sumField(dataCur, "subtotal");

  const utilPrev = sumField(dataPrev.filter(r => r.incluirUtilidad), "utilidad");
  const utilCur = sumField(dataCur.filter(r => r.incluirUtilidad), "utilidad");

  const mPrev = ventasPrev > 0 ? utilPrev / ventasPrev : 0;
  const mCur = ventasCur > 0 ? utilCur / ventasCur : 0;

  const credPrev = sumField(dataPrev.filter(r => r.esCredito), "subtotal");
  const credCur = sumField(dataCur.filter(r => r.esCredito), "subtotal");
  const pctCredPrev = ventasPrev > 0 ? credPrev / ventasPrev : 0;
  const pctCredCur = ventasCur > 0 ? credCur / ventasCur : 0;

  const utilNegPrevBase = dataPrev.filter(r => r.incluirUtilidad && r.subtotal > 0);
  const utilNegCurBase = dataCur.filter(r => r.incluirUtilidad && r.subtotal > 0);
  const pctNegPrev = utilNegPrevBase.length ? utilNegPrevBase.filter(r => r.utilidad < 0).length / utilNegPrevBase.length : 0;
  const pctNegCur = utilNegCurBase.length ? utilNegCurBase.filter(r => r.utilidad < 0).length / utilNegCurBase.length : 0;

  const filtros = getFiltros();
  const m2Valor = filtros.store && M2_POR_ALMACEN[filtros.store] ? M2_POR_ALMACEN[filtros.store] : M2_POR_ALMACEN["Todas"];

  const rows = [
    { id: "ventas", nombre: "Ventas (Subtotal)", prev: ventasPrev, cur: ventasCur, tipo: "money" },
    { id: "utilidad", nombre: "Utilidad Bruta", prev: utilPrev, cur: utilCur, tipo: "money" },
    { id: "margen", nombre: "Margen Bruto %", prev: mPrev, cur: mCur, tipo: "percent" },
    { id: "pct_credito", nombre: "% Ventas a Crédito", prev: pctCredPrev, cur: pctCredCur, tipo: "percent" },
    { id: "pct_util_neg", nombre: "% Ventas con Utilidad Negativa", prev: pctNegPrev, cur: pctNegCur, tipo: "percent" },
    { id: "m2", nombre: "m² disponibles", prev: m2Valor, cur: m2Valor, tipo: "plain" }
  ];

  tablaYoYBody.innerHTML = "";
  rows.forEach(r => {
    let crec;
    if (r.prev === 0 && r.cur > 0) crec = 1;
    else if (r.prev === 0 && r.cur === 0) crec = 0;
    else crec = (r.cur - r.prev) / (r.prev || 1);

    const crecStr = (crec * 100).toFixed(1) + "%";
    const cls = crec > 0 ? "pill pos" : crec < 0 ? "pill neg" : "pill neu";
    const icon = crec > 0 ? "▲" : crec < 0 ? "▼" : "●";

    function fmt(v) {
      if (r.tipo === "money") return formatCurrency(v);
      if (r.tipo === "percent") return formatPercent(v);
      if (r.tipo === "plain") return v.toLocaleString("es-MX");
      return v;
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.nombre}</td>
      <td class="text-right">${fmt(r.prev)}</td>
      <td class="text-right">${fmt(r.cur)}</td>
      <td class="text-right"><span class="${cls}">${icon} ${crecStr}</span></td>
    `;
    tr.addEventListener("click", () => abrirDetalleMetrica(r.id, yPrev, yCur));
    tablaYoYBody.appendChild(tr);
  });
}

function abrirDetalleMetrica(metricId, yPrev, yCur) {
  if (!modalBackdrop || !modalTableBody) return;

  const datos = filtrarRecords(true);
  const dataPrev = datos.filter(r => r.anio === yPrev);
  const dataCur = datos.filter(r => r.anio === yCur);

  modalCtx.metricId = metricId;
  modalCtx.yPrev = yPrev;
  modalCtx.yCur = yCur;
  modalCtx.rows = [];

  modalYearPrev.textContent = yPrev;
  modalYearCurrent.textContent = yCur;

  let titulo = "";
  let calc;
  let tipo = "money";

  if (metricId === "ventas") {
    titulo = "Detalle de Ventas (Subtotal) por categoría";
    calc = arr => sumField(arr, "subtotal");
  } else if (metricId === "utilidad") {
    titulo = "Detalle de Utilidad Bruta por categoría";
    calc = arr => sumField(arr.filter(r => r.incluirUtilidad), "utilidad");
  } else if (metricId === "margen") {
    titulo = "Detalle de Margen Bruto % por categoría";
    tipo = "percent";
    calc = arr => {
      const v = sumField(arr, "subtotal");
      const u = sumField(arr.filter(r => r.incluirUtilidad), "utilidad");
      return v > 0 ? u / v : 0;
    };
  } else if (metricId === "pct_credito") {
    titulo = "Detalle % Ventas a Crédito por categoría";
    tipo = "percent";
    calc = arr => {
      const vTot = sumField(arr, "subtotal");
      const vCred = sumField(arr.filter(r => r.esCredito), "subtotal");
      return vTot > 0 ? vCred / vTot : 0;
    };
  } else if (metricId === "pct_util_neg") {
    titulo = "% Ventas con Utilidad Negativa por categoría";
    tipo = "percent";
    calc = arr => {
      const base = arr.filter(r => r.incluirUtilidad && r.subtotal > 0);
      if (!base.length) return 0;
      const neg = base.filter(r => r.utilidad < 0).length;
      return neg / base.length;
    };
  } else if (metricId === "m2") {
    titulo = "Detalle de m² por almacén";
    modalTitle.textContent = titulo;
    modalSub.textContent = "m² configurados";

    modalDimTitle.textContent = "Almacén";
    modalTableBody.innerHTML = "";

    const filas = [];
    Object.keys(M2_POR_ALMACEN).forEach(k => {
      if (k === "Todas") return;
      filas.push({ dim: k, prev: M2_POR_ALMACEN[k], cur: M2_POR_ALMACEN[k] });
    });

    modalCtx.rows = filas.map(f => ({ dimension: f.dim, prev: f.prev, cur: f.cur, growth: 0 }));

    filas.forEach(f => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${f.dim}</td>
        <td class="text-right">${f.prev.toLocaleString("es-MX")}</td>
        <td class="text-right">${f.cur.toLocaleString("es-MX")}</td>
        <td class="text-right"><span class="pill neu">● 0.0%</span></td>
      `;
      modalTableBody.appendChild(tr);
    });

    modalBackdrop.classList.add("active");
    return;
  }

  modalTitle.textContent = titulo;
  const filtros = getFiltros();
  modalSub.textContent = `Filtros: almacén=${filtros.store || "Todos"}, tipo=${filtros.tipo}, categoría=${filtros.categoria || "Todas"}`;

  modalDimTitle.textContent = "Categoría";

  const cats = unique(datos.map(r => r.categoria || "(Sin categoría)"));

  const filas = cats.map(cat => {
    const aPrev = dataPrev.filter(r => (r.categoria || "(Sin categoría)") === cat);
    const aCur = dataCur.filter(r => (r.categoria || "(Sin categoría)") === cat);
    return { dim: cat, prev: calc(aPrev), cur: calc(aCur) };
  });

  modalTableBody.innerHTML = "";

  const sorted = filas.sort((a, b) => b.cur - a.cur);

  modalCtx.rows = sorted.map(f => {
    let crec;
    if (f.prev === 0 && f.cur > 0) crec = 1;
    else if (f.prev === 0 && f.cur === 0) crec = 0;
    else crec = (f.cur - f.prev) / (f.prev || 1);
    return { dimension: f.dim, prev: f.prev, cur: f.cur, growth: crec };
  });

  sorted.forEach(f => {
    let crec;
    if (f.prev === 0 && f.cur > 0) crec = 1;
    else if (f.prev === 0 && f.cur === 0) crec = 0;
    else crec = (f.cur - f.prev) / (f.prev || 1);

    const crecStr = (crec * 100).toFixed(1) + "%";
    const cls = crec > 0 ? "pill pos" : crec < 0 ? "pill neg" : "pill neu";
    const icon = crec > 0 ? "▲" : crec < 0 ? "▼" : "●";

    function fmt(v) {
      if (tipo === "money") return formatCurrency(v);
      if (tipo === "percent") return formatPercent(v);
      return v.toLocaleString("es-MX");
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="linkish">${f.dim}</td>
      <td class="text-right">${fmt(f.prev)}</td>
      <td class="text-right">${fmt(f.cur)}</td>
      <td class="text-right"><span class="${cls}">${icon} ${crecStr}</span></td>
    `;

    tr.addEventListener("click", () => {
      // ir a detalle filtrado por esa categoría
      if (filterCategory) {
        filterCategory.value = categoriasDisponibles.includes(f.dim) ? f.dim : "all";
      }
      switchTab("tab-detalle");
      renderDetalle();
      modalBackdrop.classList.remove("active");
    });

    modalTableBody.appendChild(tr);
  });

  modalBackdrop.classList.add("active");
}

if (modalClose && modalBackdrop) {
  modalClose.addEventListener("click", () => modalBackdrop.classList.remove("active"));
  modalBackdrop.addEventListener("click", e => { if (e.target === modalBackdrop) modalBackdrop.classList.remove("active"); });
}

// Exportaciones modal YoY
if (modalExportXlsx) modalExportXlsx.addEventListener("click", () => {
  if (!modalCtx.rows?.length) return;
  const out = modalCtx.rows.map(r => ({
    Dimension: r.dimension,
    [modalYearPrev.textContent]: r.prev,
    [modalYearCurrent.textContent]: r.cur,
    "Crecimiento %": (r.growth * 100).toFixed(1) + "%"
  }));
  exportJsonToXlsx(out, `Detalle_${modalCtx.metricId}_${modalCtx.yPrev}_${modalCtx.yCur}.xlsx`, "Detalle");
});

if (modalExportPdf) modalExportPdf.addEventListener("click", () => {
  if (!modalCtx.rows?.length) return;
  const headers = [modalDimTitle.textContent, modalYearPrev.textContent, modalYearCurrent.textContent, "Crecimiento %"];
  const rows = modalCtx.rows.map(r => [
    r.dimension,
    r.prev,
    r.cur,
    (r.growth * 100).toFixed(1) + "%"
  ]);
  exportTableToPdf(modalTitle.textContent, headers, rows, `Detalle_${modalCtx.metricId}_${modalCtx.yPrev}_${modalCtx.yCur}.pdf`);
});

if (modalGoDetalle) modalGoDetalle.addEventListener("click", () => {
  switchTab("tab-detalle");
  modalBackdrop.classList.remove("active");
});

// ==================== CRÉDITO VS CONTADO ====================

function actualizarCredito(data) {
  const canvas = $("chart-credito");
  if (!canvas || !tablaCreditoBody) return;

  const filtros = getFiltros();
  let yearBase = filtros.year;
  if (!yearBase) yearBase = Math.max(...yearsDisponibles);

  const base = data.filter(r => r.anio === yearBase);

  const ventasCont = new Array(12).fill(0);
  const ventasCred = new Array(12).fill(0);

  base.forEach(r => {
    const idx = r.mes - 1;
    if (idx < 0 || idx > 11) return;
    if (r.tipoFactura === "credito") ventasCred[idx] += r.subtotal;
    else ventasCont[idx] += r.subtotal;
  });

  const labelsMes = monthLabels();
  const ctx = canvas.getContext("2d");

  if (charts.credito) charts.credito.destroy();
  charts.credito = new Chart(ctx, {
    type: "bar",
    data: {
      labels: labelsMes,
      datasets: [
        { label: "Contado", data: ventasCont, stack: "ventas" },
        { label: "Crédito", data: ventasCred, stack: "ventas" }
      ]
    },
    options: {
      responsive: true,
      scales: {
        x: { stacked: true },
        y: { stacked: true, ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });

  tablaCreditoBody.innerHTML = "";
  labelsMes.forEach((mesLabel, idx) => {
    const cont = ventasCont[idx];
    const cred = ventasCred[idx];
    const total = cont + cred;
    if (total === 0) return;
    const pctCred = total > 0 ? cred / total : 0;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${mesLabel} ${yearBase}</td>
      <td class="text-right">${formatCurrency(cont)}</td>
      <td class="text-right">${formatCurrency(cred)}</td>
      <td class="text-right">${formatCurrency(total)}</td>
      <td class="text-right">${formatPercent(pctCred)}</td>
    `;
    tablaCreditoBody.appendChild(tr);
  });
}

// ==================== CATEGORÍAS ====================

function actualizarCategorias(data) {
  const canvas = $("chart-categorias");
  if (!canvas || !tablaCategoriasBody) return;

  const grupos = {};
  data.forEach(r => {
    const cat = r.categoria || "(Sin categoría)";
    if (!grupos[cat]) grupos[cat] = { ventas: 0, utilidad: 0, ventasCred: 0, ventasTot: 0 };
    grupos[cat].ventas += r.subtotal;
    if (r.incluirUtilidad) grupos[cat].utilidad += r.utilidad;
    if (r.esCredito) grupos[cat].ventasCred += r.subtotal;
    grupos[cat].ventasTot += r.subtotal;
  });

  const lista = Object.entries(grupos).map(([cat, d]) => {
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    const pctCred = d.ventasTot > 0 ? d.ventasCred / d.ventasTot : 0;
    return { cat, ventas: d.ventas, utilidad: d.utilidad, margen, pctCred };
  });

  const top = lista.slice().sort((a,b) => b.ventas - a.ventas).slice(0, 12);

  const labels = top.map(x => x.cat);
  const valores = top.map(x => x.ventas);

  const ctx = canvas.getContext("2d");
  if (charts.categorias) charts.categorias.destroy();
  charts.categorias = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ label: "Ventas", data: valores }] },
    options: {
      responsive: true,
      indexAxis: "y",
      onClick: (evt, elements) => {
        if (!elements?.length) return;
        const i = elements[0].index;
        const cat = labels[i];
        if (filterCategory) filterCategory.value = categoriasDisponibles.includes(cat) ? cat : "all";
        switchTab("tab-detalle");
        renderDetalle();
      },
      scales: {
        x: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });

  tablaCategoriasBody.innerHTML = "";
  lista.sort((a,b) => b.ventas - a.ventas).forEach(r => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="linkish">${r.cat}</td>
      <td class="text-right">${formatCurrency(r.ventas)}</td>
      <td class="text-right">${formatCurrency(r.utilidad)}</td>
      <td class="text-right">${formatPercent(r.margen)}</td>
      <td class="text-right">${formatPercent(r.pctCred)}</td>
    `;
    tr.addEventListener("click", () => {
      if (filterCategory) filterCategory.value = categoriasDisponibles.includes(r.cat) ? r.cat : "all";
      switchTab("tab-detalle");
      renderDetalle();
    });
    tablaCategoriasBody.appendChild(tr);
  });
}

// ==================== DETALLE (CONFIGURABLE + CLICK FACTURA) ====================

function cargarDetalleLayout() {
  detalleColsLayout = [...DETALLE_DEFAULT_COLS];
  try {
    const raw = localStorage.getItem(LS_KEYS.detailLayout);
    if (!raw) return;
    const arr = JSON.parse(raw);
    const filtrado = arr.filter(c => DETALLE_ALL_COLS.includes(c));
    if (filtrado.length === 7) detalleColsLayout = filtrado;
  } catch {}
}

function initDetalleUI() {
  cargarDetalleLayout();
  detalleFiltros = {};
  detalleSort = { col: null, asc: true };

  if (searchGlobalInput) {
    searchGlobalInput.addEventListener("input", debounce(() => {
      detalleBusqueda = (searchGlobalInput.value || "").toLowerCase();
      renderDetalle();
    }, 200));
  }

  renderDetalleHeaders();
  renderDetalleChips();
}

function renderDetalleChips() {
  if (!detalleChipsContainer) return;
  detalleChipsContainer.innerHTML = "";

  const disponibles = DETALLE_ALL_COLS.filter(c => !detalleColsLayout.includes(c));
  disponibles.forEach(col => {
    const chip = document.createElement("button");
    chip.className = "chip";
    chip.textContent = col;
    chip.draggable = true;
    chip.dataset.col = col;

    chip.addEventListener("dragstart", () => { detalleDragCol = col; });

    // click rápido: reemplaza la primera columna
    chip.addEventListener("click", () => {
      detalleColsLayout[0] = col;
      if (cfg.saveView) saveConfigs();
      renderDetalleHeaders();
      renderDetalleChips();
      renderDetalle();
    });

    detalleChipsContainer.appendChild(chip);
  });
}

function renderDetalleHeaders() {
  if (!detalleHeaderRow || !detalleFilterRow) return;

  detalleHeaderRow.innerHTML = "";
  detalleFilterRow.innerHTML = "";
  detalleFiltros = {};

  detalleColsLayout.forEach(col => {
    // header
    const th = document.createElement("th");
    th.classList.add("sortable", "th-drop-target");
    th.dataset.col = col;
    th.textContent = col;

    const sortSpan = document.createElement("span");
    sortSpan.textContent = "⇵";
    th.appendChild(sortSpan);

    th.addEventListener("click", () => sortDetalle(col));

    th.addEventListener("dragover", e => { e.preventDefault(); th.classList.add("th-drop-target-over"); });
    th.addEventListener("dragleave", () => th.classList.remove("th-drop-target-over"));
    th.addEventListener("drop", e => {
      e.preventDefault();
      th.classList.remove("th-drop-target-over");
      if (!detalleDragCol) return;
      const idx = detalleColsLayout.indexOf(col);
      if (idx === -1) return;
      detalleColsLayout[idx] = detalleDragCol;
      detalleDragCol = null;
      if (cfg.saveView) saveConfigs();
      renderDetalleHeaders();
      renderDetalleChips();
      renderDetalle();
    });

    detalleHeaderRow.appendChild(th);

    // filtros por columna
    const thF = document.createElement("th");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "filter-input";
    inp.placeholder = "Filtrar...";
    inp.addEventListener("input", debounce(() => {
      detalleFiltros[col] = (inp.value || "").toLowerCase();
      renderDetalle();
    }, 150));
    thF.appendChild(inp);
    detalleFilterRow.appendChild(thF);
  });
}

function getValorDetalle(r, col) {
  switch (col) {
    case "Año": return r.anio;
    case "Fecha": return r.fecha ? r.fecha.toISOString().slice(0,10) : "";
    case "Hora": return r.hora || "";
    case "Almacén": return r.almacen || "";
    case "Factura/Nota": return r.folio || "";
    case "Cliente": return r.cliente || "";
    case "Categoría": return r.categoria || "";
    case "Producto": return r.producto || "";
    case "Tipo factura": return r.tipoFactura === "credito" ? "Crédito" : "Contado";
    case "Subtotal": return r.subtotal;
    case "Costo": return r.costo;
    case "Descuento": return r.descuentoMonto || 0;
    case "Utilidad": return r.utilidad;
    case "Margen %": return r.subtotal > 0 ? r.utilidad / r.subtotal : 0;
    case "Marca": return r.marca || "";
    case "Vendedor": return r.vendedor || "";
    default: return "";
  }
}

function sortDetalle(col) {
  if (detalleSort.col === col) detalleSort.asc = !detalleSort.asc;
  else { detalleSort.col = col; detalleSort.asc = true; }
  renderDetalle();
}

function renderDetalle() {
  if (!detalleTableBody) return;
  if (!records.length) { detalleTableBody.innerHTML = ""; return; }

  const base = filtrarRecords(false);
  let arr = base.slice();

  // filtro: solo márgenes negativos (config rápida)
  if (quick.negOnly) {
    arr = arr.filter(r => (r.subtotal > 0 ? (r.utilidad / r.subtotal) : 0) < 0);
  }

  // filtros por columna
  arr = arr.filter(r => {
    for (const col in detalleFiltros) {
      const text = detalleFiltros[col];
      if (!text) continue;
      const v = getValorDetalle(r, col);
      const s = (typeof v === "number") ? v.toString() : (v || "").toString().toLowerCase();
      if (!s.includes(text)) return false;
    }
    return true;
  });

  // filtro global
  if (detalleBusqueda) {
    arr = arr.filter(r => {
      const campos = [r.cliente, r.vendedor, r.folio, r.almacen, r.categoria, r.producto];
      return campos.some(c => c && c.toString().toLowerCase().includes(detalleBusqueda));
    });
  }

  // orden
  if (detalleSort.col) {
    const col = detalleSort.col;
    const asc = detalleSort.asc;
    arr.sort((a, b) => {
      const va = getValorDetalle(a, col);
      const vb = getValorDetalle(b, col);
      if (typeof va === "number" && typeof vb === "number") return asc ? va - vb : vb - va;
      const sa = (va || "").toString();
      const sb = (vb || "").toString();
      return asc ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }

  // performance limit
  const max = cfg.maxDetalleRows || 5000;
  const view = arr.slice(0, max);

  detalleTableBody.innerHTML = "";
  view.forEach(r => {
    const tr = document.createElement("tr");

    detalleColsLayout.forEach(col => {
      const td = document.createElement("td");
      let v = getValorDetalle(r, col);

      if (col === "Factura/Nota") {
        td.classList.add("linkish");
        td.textContent = v || "";
        td.addEventListener("click", () => abrirDocumento(r.origen, r.folio));
      }
      else if (col === "Subtotal" || col === "Costo" || col === "Utilidad" || col === "Descuento") {
        td.classList.add("text-right");
        const num = toNumber(v);
        td.textContent = formatCurrency(num);
        if (num < 0) td.classList.add("neg");
      }
      else if (col === "Margen %") {
        td.classList.add("text-right");
        const num = typeof v === "number" ? v : toNumber(v);
        td.textContent = formatPercent(num);
        if (num < 0) td.classList.add("neg");
      }
      else {
        td.textContent = v || "";
      }

      tr.appendChild(td);
    });

    detalleTableBody.appendChild(tr);
  });
}

// ==================== MODAL FACTURA / NOTA ====================

function abrirDocumento(origen, folio) {
  if (!docBackdrop || !docBody) return;

  const datos = filtrarRecords(false);
  const lines = datos.filter(r => r.origen === origen && r.folio === folio);

  if (!lines.length) return;

  const head = lines[0];
  const meta = {
    fecha: head.fecha ? head.fecha.toISOString().slice(0,10) : "",
    hora: head.hora || "",
    cliente: head.cliente || "",
    vendedor: head.vendedor || "",
    almacen: head.almacen || "",
    categoria: head.categoria || "",
    origen,
    folio
  };

  docCtx.origen = origen;
  docCtx.folio = folio;
  docCtx.meta = meta;

  docTitle.textContent = `${origen === "factura" ? "Factura" : "Nota"}: ${folio}`;
  docSub.textContent = `Cliente: ${meta.cliente} | Fecha: ${meta.fecha} ${meta.hora} | Almacén: ${meta.almacen} | Categoría: ${meta.categoria}`;

  // Construir filas del modal
  docBody.innerHTML = "";
  docCtx.rows = lines.map(r => {
    const qty = toNumber(r.cantidad);
    const precio = toNumber(r.subtotal); // asumimos subtotal por línea sin IVA
    return {
      Fecha: r.fecha ? r.fecha.toISOString().slice(0,10) : "",
      Hora: r.hora || "",
      Cliente: r.cliente || "",
      Descripción: r.producto || "",
      Cantidad: qty,
      Costo: toNumber(r.costo),
      "Precio (sin IVA)": precio,
      Vendedor: r.vendedor || ""
    };
  });

  docCtx.rows.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row["Fecha"]}</td>
      <td>${row["Hora"]}</td>
      <td>${row["Cliente"]}</td>
      <td>${row["Descripción"]}</td>
      <td class="text-right">${row["Cantidad"].toLocaleString("es-MX")}</td>
      <td class="text-right">${formatCurrency(row["Costo"])}</td>
      <td class="text-right">${formatCurrency(row["Precio (sin IVA)"])}</td>
      <td>${row["Vendedor"]}</td>
    `;
    docBody.appendChild(tr);
  });

  docBackdrop.classList.add("active");
}

if (docClose && docBackdrop) {
  docClose.addEventListener("click", () => docBackdrop.classList.remove("active"));
  docBackdrop.addEventListener("click", e => { if (e.target === docBackdrop) docBackdrop.classList.remove("active"); });
}

if (docExportXlsx) docExportXlsx.addEventListener("click", () => {
  if (!docCtx.rows?.length) return;
  exportJsonToXlsx(docCtx.rows, `Documento_${docCtx.origen}_${docCtx.folio}.xlsx`, "Lineas");
});

if (docExportPdf) docExportPdf.addEventListener("click", () => {
  if (!docCtx.rows?.length) return;
  const headers = ["Fecha","Hora","Cliente","Descripción","Cantidad","Costo","Precio (sin IVA)","Vendedor"];
  const rows = docCtx.rows.map(r => [
    r["Fecha"], r["Hora"], r["Cliente"], r["Descripción"],
    r["Cantidad"], r["Costo"], r["Precio (sin IVA)"], r["Vendedor"]
  ]);
  exportTableToPdf(`${docTitle.textContent}`, headers, rows, `Documento_${docCtx.origen}_${docCtx.folio}.pdf`);
});

if (docGoDetalle) docGoDetalle.addEventListener("click", () => {
  // Ir al tab Detalle filtrando por folio (en búsqueda global)
  switchTab("tab-detalle");
  if (searchGlobalInput) {
    searchGlobalInput.value = docCtx.folio;
    detalleBusqueda = docCtx.folio.toLowerCase();
  }
  renderDetalle();
  docBackdrop.classList.remove("active");
});

// ==================== EXPLORADOR ====================

let explorerCfg = { dim: "almacen", met: "ventas", type: "bar", top: 10 };

function initExplorerUI() {
  if (expDim) expDim.addEventListener("change", () => { explorerCfg.dim = expDim.value; actualizarExplorador(); });
  if (expMet) expMet.addEventListener("change", () => { explorerCfg.met = expMet.value; actualizarExplorador(); });
  if (expType) expType.addEventListener("change", () => { explorerCfg.type = expType.value; actualizarExplorador(); });
  if (expTop) expTop.addEventListener("change", () => { explorerCfg.top = parseInt(expTop.value,10); actualizarExplorador(); });
  if (expRefresh) expRefresh.addEventListener("click", actualizarExplorador);

  if (expExportPng) expExportPng.addEventListener("click", () => {
    chartToPng(charts.explorer, `Explorador_${explorerCfg.dim}_${explorerCfg.met}.png`);
  });
}

function dimValue(r, dim) {
  if (dim === "almacen") return r.almacen || "(sin)";
  if (dim === "categoria") return r.categoria || "(sin)";
  if (dim === "cliente") return r.cliente || "(sin cliente)";
  if (dim === "vendedor") return r.vendedor || "(sin vendedor)";
  if (dim === "marca") return r.marca || "(sin marca)";
  if (dim === "mes") return monthLabels()[Math.max(0, Math.min(11, (r.mes||1)-1))];
  return "(sin)";
}

function metricAgg(group, met) {
  if (met === "ventas") return group.ventas;
  if (met === "utilidad") return group.utilidad;
  if (met === "margen") return group.ventas > 0 ? group.utilidad / group.ventas : 0;
  if (met === "transacciones") return group.ops;
  if (met === "clientes") return group.clientes;
  return group.ventas;
}

function metricLabel(met) {
  if (met === "ventas") return "Ventas";
  if (met === "utilidad") return "Utilidad";
  if (met === "margen") return "Margen %";
  if (met === "transacciones") return "Transacciones";
  if (met === "clientes") return "Clientes únicos";
  return met;
}

function actualizarExplorador() {
  const canvas = $("chart-explorer");
  if (!canvas || !tablaExplorer) return;
  if (!records.length) return;

  // sincroniza UI
  if (expDim) expDim.value = explorerCfg.dim;
  if (expMet) expMet.value = explorerCfg.met;
  if (expType) expType.value = explorerCfg.type;
  if (expTop) expTop.value = String(explorerCfg.top);

  const data = filtrarRecords(false);

  const groups = {};
  const opsByDim = {};
  const clientsByDim = {};

  data.forEach(r => {
    const d = dimValue(r, explorerCfg.dim);
    if (!groups[d]) groups[d] = { ventas: 0, utilidad: 0, ops: 0, clientes: 0 };
    groups[d].ventas += r.subtotal;
    if (r.incluirUtilidad) groups[d].utilidad += r.utilidad;

    const opKey = r.origen + "|" + r.folio;
    if (!opsByDim[d]) opsByDim[d] = new Set();
    opsByDim[d].add(opKey);

    if (!clientsByDim[d]) clientsByDim[d] = new Set();
    if (r.cliente) clientsByDim[d].add(r.cliente);
  });

  Object.keys(groups).forEach(d => {
    groups[d].ops = (opsByDim[d] || new Set()).size;
    groups[d].clientes = (clientsByDim[d] || new Set()).size;
  });

  let list = Object.keys(groups).map(d => ({ d, val: metricAgg(groups[d], explorerCfg.met) }));

  // si es margen, conviene ordenar por valor absoluto? aquí por valor normal
  list.sort((a,b) => b.val - a.val);
  list = list.slice(0, explorerCfg.top);

  const labels = list.map(x => x.d);
  const values = list.map(x => x.val);

  // Tabla
  expThDim.textContent = explorerCfg.dim.toUpperCase();
  expThMet.textContent = metricLabel(explorerCfg.met);

  tablaExplorer.innerHTML = "";
  list.forEach(x => {
    const tr = document.createElement("tr");
    const isPct = explorerCfg.met === "margen";
    tr.innerHTML = `
      <td class="linkish">${x.d}</td>
      <td class="text-right">${isPct ? formatPercent(x.val) : formatCurrency(x.val)}</td>
    `;
    tr.addEventListener("click", () => drillToDetalle(explorerCfg.dim, x.d));
    tablaExplorer.appendChild(tr);
  });

  // Chart
  if (expTitle) expTitle.textContent = `Explorador: ${explorerCfg.dim} vs ${metricLabel(explorerCfg.met)}`;
  const ctx = canvas.getContext("2d");
  if (charts.explorer) charts.explorer.destroy();

  charts.explorer = new Chart(ctx, {
    type: explorerCfg.type,
    data: {
      labels,
      datasets: [{ label: metricLabel(explorerCfg.met), data: values }]
    },
    options: {
      responsive: true,
      onClick: (evt, elements) => {
        if (!elements?.length) return;
        const i = elements[0].index;
        const val = labels[i];
        drillToDetalle(explorerCfg.dim, val);
      },
      scales: (explorerCfg.type === "pie" || explorerCfg.type === "doughnut") ? {} : {
        y: {
          ticks: {
            callback: v => (explorerCfg.met === "margen" ? (v*100).toFixed(1) + "%" : v.toLocaleString("es-MX"))
          }
        }
      }
    }
  });
}

function drillToDetalle(dim, value) {
  // aplica filtros aproximados desde explorador
  if (dim === "almacen" && filterStore) filterStore.value = almacenesDisponibles.includes(value) ? value : "all";
  if (dim === "categoria" && filterCategory) filterCategory.value = categoriasDisponibles.includes(value) ? value : "all";

  // para dimensiones no existentes como dropdown, usa búsqueda global
  if (dim === "cliente" && searchGlobalInput) {
    switchTab("tab-detalle");
    searchGlobalInput.value = value;
    detalleBusqueda = value.toLowerCase();
    renderDetalle();
    return;
  }
  if (dim === "vendedor" && searchGlobalInput) {
    switchTab("tab-detalle");
    searchGlobalInput.value = value;
    detalleBusqueda = value.toLowerCase();
    renderDetalle();
    return;
  }
  if (dim === "marca" && searchGlobalInput) {
    switchTab("tab-detalle");
    searchGlobalInput.value = value;
    detalleBusqueda = value.toLowerCase();
    renderDetalle();
    return;
  }

  // mes: manda a detalle con búsqueda del nombre del mes (simple)
  if (dim === "mes" && searchGlobalInput) {
    switchTab("tab-detalle");
    searchGlobalInput.value = value;
    detalleBusqueda = value.toLowerCase();
    renderDetalle();
    return;
  }

  switchTab("tab-detalle");
  renderDetalle();
}

// ==================== EXPORTS + BOTONES PNG ====================

if (btnExportChartMensual) btnExportChartMensual.addEventListener("click", () => chartToPng(charts.mensual, "Resumen_Mensual.png"));
if (btnExportChartAlmacen) btnExportChartAlmacen.addEventListener("click", () => chartToPng(charts.almacen, "Resumen_Almacen.png"));
if (btnExportChartCredito) btnExportChartCredito.addEventListener("click", () => chartToPng(charts.credito, "Credito_vs_Contado.png"));
if (btnExportChartCategorias) btnExportChartCategorias.addEventListener("click", () => chartToPng(charts.categorias, "Categorias_Top.png"));

// ==================== CONFIG TAB ====================

function initConfigTab() {
  if (cfgMaxRows) cfgMaxRows.value = String(cfg.maxDetalleRows || 5000);
  if (cfgSaveView) cfgSaveView.checked = !!cfg.saveView;

  if (cfgMaxRows) cfgMaxRows.addEventListener("change", () => {
    cfg.maxDetalleRows = parseInt(cfgMaxRows.value, 10) || 5000;
    saveConfigs();
    renderDetalle();
  });

  if (cfgSaveView) cfgSaveView.addEventListener("change", () => {
    cfg.saveView = !!cfgSaveView.checked;
    saveConfigs();
  });

  if (cfgResetView) cfgResetView.addEventListener("click", () => {
    // reset a defaults (sin borrar archivo cargado)
    quick = { metric: "ventas", chart: "bar", mode: "normal", negOnly: false };
    cfg.maxDetalleRows = 5000;
    detalleColsLayout = [...DETALLE_DEFAULT_COLS];
    applyQuickToUI();
    if (cfgMaxRows) cfgMaxRows.value = "5000";
    renderDetalleHeaders();
    renderDetalleChips();
    saveConfigs();
    actualizarTodo();
  });

  if (cfgClearStorage) cfgClearStorage.addEventListener("click", () => {
    clearSaved();
    // también resetea en memoria
    quick = { metric: "ventas", chart: "bar", mode: "normal", negOnly: false };
    cfg.maxDetalleRows = 5000;
    detalleColsLayout = [...DETALLE_DEFAULT_COLS];
    applyQuickToUI();
    if (cfgMaxRows) cfgMaxRows.value = "5000";
    renderDetalleHeaders();
    renderDetalleChips();
    actualizarTodo();
  });

  if (cfgExportDetalleXlsx) cfgExportDetalleXlsx.addEventListener("click", () => {
    if (!records.length) return;
    const data = filtrarRecords(false);
    const out = data.map(r => ({
      Año: r.anio,
      Fecha: r.fecha ? r.fecha.toISOString().slice(0,10) : "",
      Hora: r.hora || "",
      Almacén: r.almacen,
      "Factura/Nota": r.folio,
      Cliente: r.cliente,
      Categoría: r.categoria,
      Producto: r.producto,
      "Tipo factura": r.tipoFactura === "credito" ? "Crédito" : "Contado",
      Subtotal: r.subtotal,
      Costo: r.costo,
      Descuento: r.descuentoMonto || 0,
      Utilidad: r.utilidad,
      "Margen %": r.subtotal > 0 ? (r.utilidad / r.subtotal) : 0,
      Marca: r.marca,
      Vendedor: r.vendedor
    }));
    exportJsonToXlsx(out, "Detalle_Filtrado.xlsx", "Detalle");
  });

  if (cfgExportResumenPdf) cfgExportResumenPdf.addEventListener("click", () => {
    if (!records.length) return;

    const data = filtrarRecords(false);
    const ventas = sumField(data, "subtotal");
    const util = sumField(data.filter(r => r.incluirUtilidad), "utilidad");
    const margen = ventas > 0 ? util / ventas : 0;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });

    doc.setFontSize(16);
    doc.text("Resumen de Ventas – El Cedro", 40, 40);

    doc.setFontSize(11);
    const f = getFiltros();
    doc.text(`Filtros: Año=${f.year || "Todos"} | Almacén=${f.store || "Todos"} | Tipo=${f.tipo} | Categoría=${f.categoria || "Todas"}`, 40, 60);

    doc.setFontSize(12);
    doc.text(`Ventas: ${formatCurrency(ventas)}`, 40, 90);
    doc.text(`Utilidad: ${formatCurrency(util)}`, 40, 110);
    doc.text(`Margen: ${formatPercent(margen)}`, 40, 130);

    // intenta pegar imágenes de charts (si existen)
    let y = 160;
    try {
      if (charts.mensual) {
        const img = charts.mensual.toBase64Image();
        doc.addImage(img, "PNG", 40, y, 520, 200);
        y += 220;
      }
      if (charts.almacen) {
        const img2 = charts.almacen.toBase64Image();
        doc.addImage(img2, "PNG", 40, y, 520, 200);
        y += 220;
      }
    } catch {}

    doc.save("Resumen.pdf");
  });
}

// ==================== TABS ====================

function switchTab(tabId) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));

  const btn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
  const panel = $(tabId);
  if (btn) btn.classList.add("active");
  if (panel) panel.classList.add("active");
}

Array.from(document.querySelectorAll(".tab-btn")).forEach(btn => {
  btn.addEventListener("click", () => {
    const tabId = btn.dataset.tab;
    switchTab(tabId);
  });
});

// ==================== INICIALIZACIÓN ====================

loadConfigs();
applyQuickToUI();
initQuickConfig();
initConfigTab();
initDetalleUI();
initExplorerUI();
renderDetalle();
