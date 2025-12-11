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

// Mapeos de columnas
const RENAME_FACTURAS = {
  "No_fac": "Factura",
  "Falta_fac": "Fecha",
  "Descuento": "Descuento ($)",
  "Subt_fac": "Sub. Factura",
  "Total_fac": "Total Factura",
  "Iva": "IVA",
  "Cve_cte": "ID Cliente",
  "Cse_prod": "Categoria",
  "Cve_prod": "Clave",
  "Cant_surt": "Pz.",
  "Dcto1": "Descuento (%)",
  "Cve_age": "ID Vendedor",
  "Desc_prod": "Articulo",
  "Cost_prom": "Costo Prom.",
  "Lugar": "Almacen2",
  "Hora_fac": "Hora",
  "Des_tial": "Marca",
  "Cto_ent": "Costo Ent.",
  "Nom_cte": "Cliente",
  "Nom_age": "Vendedor"
};

const RENAME_NOTAS = {
  "No_fac": "Nota",
  "Cve_suc": "Albaranes",
  "Falta_fac": "Fecha",
  "Lugar": "Almacen",
  "Descuento": "Descuento",
  "Cve_prod": "Clave",
  "Desc_prod": "Articulo",
  "Cant_surt": "Pz.",
  "Subt_prod": "Sub. Total",
  "Iva_prod": "IVA",
  "Dcto1": "Descuento (%)",
  "Hravta": "Hora",
  "Cse_prod": "Categoria",
  "Total_fac": "Total Nota",
  "Nom_age": "Vendedor",
  "Des_tial": "Marca",
  "Cto_ent": "Costo Entrada",
  "Nom_fac": "Cliente",
  "Cte_fac": "ID Cliente"
};

// Todas las columnas posibles para la tabla de detalle
const DETALLE_COLS_ALL = [
  "Año",
  "Fecha",
  "Hora",
  "Almacen",
  "Factura/Nota",
  "Cliente",
  "Categoria",
  "Producto",
  "Tipo factura",
  "Subtotal",
  "Costo",
  "Utilidad",
  "Margen %",
  "Pz.",
  "Marca",
  "Vendedor"
];

// Número máximo de columnas visibles en detalle
const DETALLE_MAX_COLS = 7;

// ==================== UTILIDADES ====================

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
  return v.toLocaleString("es-MX", {
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

// ==================== ESTADO Y DOM ====================

let records = [];
let yearsDisponibles = [];
let almacenesDisponibles = [];
let categoriasDisponibles = [];

let charts = { mensual: null, almacen: null };
let expChart = null;

// configuración persistente
let config = {
  mainMetric: "ventas",
  chartType: "bar",
  tableMode: "compact",
  hideNegativeMargin: false,
  maxFilasDetalle: 5000,
  resaltarNegativos: true,
  guardarVista: true,
  expDimension: "mes",
  expMetrica: "subtotal",
  expChartType: "bar",
  expExcluirNegativos: true
};

let detalleSort = { col: null, asc: true };
let detalleFiltros = {};
let detalleBusqueda = "";
let detalleColsActivas = ["Año", "Almacen", "Categoria", "Subtotal", "Margen %", "Utilidad", "Vendedor"];

let ultimaFacturaFiltro = null;

// DOM
const fileInput = document.getElementById("file-input");
const fileNameSpan = document.getElementById("file-name");
const errorDiv = document.getElementById("error");

const filterYear = document.getElementById("filter-year");
const filterStore = document.getElementById("filter-store");
const filterType = document.getElementById("filter-type");
const filterCategory = document.getElementById("filter-category");

const kpiVentas = document.getElementById("kpi-ventas");
const kpiVentasSub = document.getElementById("kpi-ventas-sub");
const kpiUtilidad = document.getElementById("kpi-utilidad");
const kpiMargen = document.getElementById("kpi-margen");
const kpiMargenSub = document.getElementById("kpi-margen-sub");
const kpiM2 = document.getElementById("kpi-m2");
const kpiM2Sub = document.getElementById("kpi-m2-sub");
const kpiTrans = document.getElementById("kpi-trans");
const kpiClientes = document.getElementById("kpi-clientes");

const tablaTopClientes = document.getElementById("tabla-top-clientes");
const tablaTopVendedores = document.getElementById("tabla-top-vendedores");

const thYearPrev = document.getElementById("th-year-prev");
const thYearCurrent = document.getElementById("th-year-current");
const tablaYoYBody = document.getElementById("tabla-yoy");

const detalleHeaderRow = document.getElementById("detalle-header-row");
const detalleFilterRow = document.getElementById("detalle-filter-row");
const detalleTableBody = document.getElementById("tabla-detalle");
const searchGlobalInput = document.getElementById("search-global");
const detalleChipContainer = document.getElementById("detalle-chip-container");

const modalBackdrop = document.getElementById("modal-backdrop");
const modalClose = document.getElementById("modal-close");
const modalTitle = document.getElementById("modal-title");
const modalSub = document.getElementById("modal-sub");
const modalYearPrev = document.getElementById("modal-year-prev");
const modalYearCurrent = document.getElementById("modal-year-current");
const modalTableBody = document.querySelector("#modal-table tbody");

// Modal factura
const modalFacturaBackdrop = document.getElementById("modal-factura-backdrop");
const modalFacturaClose = document.getElementById("modal-factura-close");
const modalFacturaTitle = document.getElementById("modal-factura-title");
const modalFacturaSub = document.getElementById("modal-factura-sub");
const modalFacturaTableBody = document.querySelector("#modal-factura-table tbody");
const modalFacturaExportXls = document.getElementById("modal-factura-export-xls");
const modalFacturaExportPdf = document.getElementById("modal-factura-export-pdf");
const modalFacturaVerDetalle = document.getElementById("modal-factura-ver-detalle");

// Config rápida
const qcMainMetric = document.getElementById("qc-main-metric");
const qcChartType = document.getElementById("qc-chart-type");
const qcTableMode = document.getElementById("qc-table-mode");
const qcHideNegativeMargin = document.getElementById("qc-hide-negative-margin");

// Explorador
const expDimensionSel = document.getElementById("exp-dimension");
const expMetricaSel = document.getElementById("exp-metrica");
const expChartTypeSel = document.getElementById("exp-chart-type");
const expExcluirNegSel = document.getElementById("exp-excluir-negativos");
const expChartTitle = document.getElementById("exp-chart-title");

// Config tab
const confMaxFilas = document.getElementById("conf-max-filas");
const confResaltarNegativos = document.getElementById("conf-resaltar-negativos");
const confGuardarVista = document.getElementById("conf-guardar-vista");
const confReset = document.getElementById("conf-reset");

// ==================== LOCALSTORAGE ====================

function loadConfig() {
  try {
    const raw = localStorage.getItem("cedroDashboardConfig");
    if (!raw) return;
    const saved = JSON.parse(raw);
    config = { ...config, ...saved };
  } catch (e) {
    console.warn("No se pudo leer config:", e);
  }
}

function saveConfig() {
  if (!config.guardarVista) return;
  try {
    localStorage.setItem("cedroDashboardConfig", JSON.stringify(config));
  } catch (e) {
    console.warn("No se pudo guardar config:", e);
  }
}

function applyConfigToUI() {
  if (qcMainMetric) qcMainMetric.value = config.mainMetric;
  if (qcChartType) qcChartType.value = config.chartType;
  if (qcTableMode) qcTableMode.value = config.tableMode;
  if (qcHideNegativeMargin) qcHideNegativeMargin.checked = config.hideNegativeMargin;

  if (expDimensionSel) expDimensionSel.value = config.expDimension;
  if (expMetricaSel) expMetricaSel.value = config.expMetrica;
  if (expChartTypeSel) expChartTypeSel.value = config.expChartType;
  if (expExcluirNegSel) expExcluirNegSel.checked = config.expExcluirNegativos;

  if (confMaxFilas) confMaxFilas.value = config.maxFilasDetalle;
  if (confResaltarNegativos) confResaltarNegativos.checked = config.resaltarNegativos;
  if (confGuardarVista) confGuardarVista.checked = config.guardarVista;

  document.body.classList.toggle("table-comfortable", config.tableMode === "comfortable");
}

loadConfig();
applyConfigToUI();

// ==================== LECTURA DE ARCHIVO ====================

if (fileInput) {
  fileInput.addEventListener("change", handleFile);
}

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

      // Encabezados para localizar columna BB (Factura del día / Rango / Base)
      const headerRows = XLSX.utils.sheet_to_json(sheetFact, {
        header: 1,
        range: 0,
        raw: true
      });
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

        const subtotal = toNumber(row["Sub. Factura"]);
        const costo = toNumber(row["Costo Ent."]);
        const descuentoMonto = toNumber(row["Descuento ($)"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad =
          !(flagTexto === "rango" || flagTexto === "base");

        let esCredito = false;
        if (descuentoMonto === 0 && flagTexto !== "base") esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (
          row["Almacen"] ||
          row["Almacen2"] ||
          ""
        )
          .toString()
          .trim()
          .toUpperCase();

        const categoria = (row["Categoria"] || "").toString().trim();
        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Factura"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();
        const articulo = (row["Articulo"] || "").toString().trim();
        const hora = (row["Hora"] || "").toString().trim();
        const pz = toNumber(row["Pz."]);
        const precioUnit = pz > 0 ? subtotal / pz : 0;

        records.push({
          origen: "factura",
          anio,
          mes,
          fecha,
          hora,
          almacen,
          categoria,
          cliente,
          vendedor,
          marca,
          folio,
          articulo,
          pz,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura,
          precioUnit
        });
      }

      // NOTAS (solo REM)
      for (let i = 0; i < rawNotas.length; i++) {
        const row = notasRen[i];
        const albaran = (row["Albaranes"] || "")
          .toString()
          .trim()
          .toLowerCase();
        if (albaran !== "rem") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;
        const anio = fecha.getFullYear();
        const mes = fecha.getMonth() + 1;

        const subtotal = toNumber(row["Sub. Total"]);
        const costo = toNumber(row["Costo Entrada"]);
        const descuentoMonto = toNumber(row["Descuento"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;

        const incluirUtilidad = true;

        let esCredito = false;
        if (descuentoMonto === 0) esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (row["Almacen"] || "")
          .toString()
          .trim()
          .toUpperCase();
        const categoria = (row["Categoria"] || "").toString().trim();
        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Nota"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();
        const articulo = (row["Articulo"] || "").toString().trim();
        const hora = (row["Hora"] || "").toString().trim();
        const pz = toNumber(row["Pz."]);
        const precioUnit = pz > 0 ? subtotal / pz : 0;

        records.push({
          origen: "nota",
          anio,
          mes,
          fecha,
          hora,
          almacen,
          categoria,
          cliente,
          vendedor,
          marca,
          folio,
          articulo,
          pz,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura,
          precioUnit
        });
      }

      if (!records.length) {
        throw new Error("No se generaron registros válidos.");
      }

      yearsDisponibles = unique(records.map(r => r.anio)).sort((a, b) => a - b);
      almacenesDisponibles = unique(records.map(r => r.almacen || ""));
      categoriasDisponibles = unique(
        records.map(r => r.categoria || "").filter(x => x)
      );

      poblarFiltros();
      initDetalleHeaders();
      initDetalleChips();
      actualizarTodo();
    } catch (err) {
      console.error(err);
      if (errorDiv) {
        errorDiv.textContent = "Error al procesar el archivo: " + err.message;
      }
    }
  };
  reader.onerror = function () {
    if (errorDiv) errorDiv.textContent = "No se pudo leer el archivo.";
  };
  reader.readAsArrayBuffer(file);
}

// ==================== FILTROS ====================

function poblarFiltros() {
  if (!filterYear || !filterStore || !filterType || !filterCategory) return;

  filterYear.innerHTML = "<option value='all'>Todos los años</option>";
  yearsDisponibles.forEach(y => {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    filterYear.appendChild(opt);
  });

  filterStore.innerHTML = "<option value='all'>Todos los almacenes</option>";
  almacenesDisponibles.forEach(a => {
    if (!a) return;
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    filterStore.appendChild(opt);
  });

  filterCategory.innerHTML =
    "<option value='all'>Todas las categorías</option>";
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

function getFiltros() {
  return {
    year: filterYear && filterYear.value === "all"
      ? null
      : filterYear
      ? parseInt(filterYear.value, 10)
      : null,
    store: filterStore && filterStore.value === "all" ? null : filterStore ? filterStore.value : null,
    tipo: filterType ? filterType.value : "both",
    categoria:
      filterCategory && filterCategory.value === "all"
        ? null
        : filterCategory
        ? filterCategory.value
        : null
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

// ==================== KPIs / GRÁFICAS / TOP ====================

function actualizarTodo() {
  if (!records.length) return;
  const datos = filtrarRecords(false);
  actualizarKpis(datos);
  actualizarGraficas(datos);
  actualizarTop(datos);
  actualizarYoY();
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
  if (kpiMargen) kpiMargen.textContent = formatPercent(margen);
  if (kpiMargenSub) kpiMargenSub.textContent = `Basado en ${utilData.length} filas válidas`;

  let m2 = 0;
  if (filtros.store && M2_POR_ALMACEN[filtros.store]) {
    m2 = M2_POR_ALMACEN[filtros.store];
  } else {
    m2 = M2_POR_ALMACEN["Todas"];
  }
  const ventasM2 = m2 > 0 ? ventas / m2 : 0;
  if (kpiM2) kpiM2.textContent = formatCurrency(ventasM2) + "/m²";
  if (kpiM2Sub) kpiM2Sub.textContent = `${filtros.store || "Todas"} – ${m2.toLocaleString("es-MX")} m²`;

  const opsSet = new Set(data.map(r => r.origen + "|" + r.folio));
  if (kpiTrans) kpiTrans.textContent = opsSet.size.toLocaleString("es-MX");

  const clientesSet = new Set(
    data.map(r => r.cliente || "").filter(x => x !== "")
  );
  if (kpiClientes) kpiClientes.textContent = clientesSet.size.toLocaleString("es-MX");
}

function actualizarGraficas(data) {
  const canvasMensual = document.getElementById("chart-mensual");
  const canvasAlm = document.getElementById("chart-almacen");
  if (!canvasMensual || !canvasAlm) return;

  const ctxMensual = canvasMensual.getContext("2d");
  const ctxAlm = canvasAlm.getContext("2d");

  const filtros = getFiltros();
  let base = data;
  if (!filtros.year) {
    const maxYear = Math.max(...yearsDisponibles);
    base = data.filter(r => r.anio === maxYear);
  }

  const ventasMes = new Array(12).fill(0);
  base.forEach(r => {
    const idx = r.mes - 1;
    if (idx >= 0 && idx < 12) ventasMes[idx] += r.subtotal;
  });
  const labelsMes = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

  if (charts.mensual) charts.mensual.destroy();
  charts.mensual = new Chart(ctxMensual, {
    type: config.chartType === "line" ? "line" : "bar",
    data: {
      labels: labelsMes,
      datasets: [{ label: "Ventas", data: ventasMes }]
    },
    options: {
      responsive: true,
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });

  const porAlm = {};
  data.forEach(r => {
    const a = r.almacen || "(sin)";
    if (!porAlm[a]) porAlm[a] = { ventas: 0, m2Ventas: 0 };
    porAlm[a].ventas += r.subtotal;
  });

  const labsAlm = Object.keys(porAlm);
  const ventasAlm = labsAlm.map(a => porAlm[a].ventas);
  const ventasM2Alm = labsAlm.map(a => {
    const m2 = M2_POR_ALMACEN[a] || 0;
    return m2 > 0 ? porAlm[a].ventas / m2 : 0;
  });

  if (charts.almacen) charts.almacen.destroy();
  charts.almacen = new Chart(ctxAlm, {
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
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
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

// ==================== COMPARATIVO YOY ====================

function actualizarYoY() {
  if (!tablaYoYBody || !thYearPrev || !thYearCurrent) return;

  const datos = filtrarRecords(true);
  if (!datos.length) {
    tablaYoYBody.innerHTML =
      "<tr><td colspan='4' class='muted'>No hay datos.</td></tr>";
    return;
  }
  const years = unique(datos.map(r => r.anio)).sort((a, b) => a - b);
  if (years.length < 2) {
    tablaYoYBody.innerHTML =
      "<tr><td colspan='4' class='muted'>Se requiere al menos 2 años.</td></tr>";
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

  const utilPrev = sumField(
    dataPrev.filter(r => r.incluirUtilidad),
    "utilidad"
  );
  const utilCur = sumField(
    dataCur.filter(r => r.incluirUtilidad),
    "utilidad"
  );

  const mPrev = ventasPrev > 0 ? utilPrev / ventasPrev : 0;
  const mCur = ventasCur > 0 ? utilCur / ventasCur : 0;

  const credPrev = sumField(
    dataPrev.filter(r => r.esCredito),
    "subtotal"
  );
  const credCur = sumField(
    dataCur.filter(r => r.esCredito),
    "subtotal"
  );
  const pctCredPrev = ventasPrev > 0 ? credPrev / ventasPrev : 0;
  const pctCredCur = ventasCur > 0 ? credCur / ventasCur : 0;

  const utilNegPrevBase = dataPrev.filter(
    r => r.incluirUtilidad && r.subtotal > 0
  );
  const utilNegCurBase = dataCur.filter(
    r => r.incluirUtilidad && r.subtotal > 0
  );
  const pctNegPrev = utilNegPrevBase.length
    ? utilNegPrevBase.filter(r => r.utilidad < 0).length /
      utilNegPrevBase.length
    : 0;
  const pctNegCur = utilNegCurBase.length
    ? utilNegCurBase.filter(r => r.utilidad < 0).length /
      utilNegCurBase.length
    : 0;

  const filtros = getFiltros();
  const m2Valor =
    filtros.store && M2_POR_ALMACEN[filtros.store]
      ? M2_POR_ALMACEN[filtros.store]
      : M2_POR_ALMACEN["Todas"];

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
    tr.dataset.metricId = r.id;
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
  if (!modalBackdrop || !modalTitle || !modalSub || !modalYearPrev || !modalYearCurrent || !modalTableBody) return;

  const datos = filtrarRecords(true);
  const dataPrev = datos.filter(r => r.anio === yPrev);
  const dataCur = datos.filter(r => r.anio === yCur);

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
    const filtros = getFiltros();
    modalSub.textContent = filtros.store
      ? `Almacén ${filtros.store}`
      : "Todas las sucursales";
    modalTableBody.innerHTML = "";

    const filas = [];
    if (filtros.store && M2_POR_ALMACEN[filtros.store]) {
      filas.push({
        cat: filtros.store,
        prev: M2_POR_ALMACEN[filtros.store],
        cur: M2_POR_ALMACEN[filtros.store]
      });
    } else {
      Object.keys(M2_POR_ALMACEN).forEach(k => {
        if (k === "Todas") return;
        filas.push({ cat: k, prev: M2_POR_ALMACEN[k], cur: M2_POR_ALMACEN[k] });
      });
    }

    filas.forEach(f => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${f.cat}</td>
        <td class="text-right">${f.prev.toLocaleString("es-MX")}</td>
        <td class="text-right">${f.cur.toLocaleString("es-MX")}</td>
        <td class="text-right"><span class="pill neu">● 0.0%</span></td>
      `;
      modalTableBody.appendChild(tr);
    });

    modalBackdrop.classList.add("active");
    return;
  }

  const filtros = getFiltros();
  modalTitle.textContent = titulo;
  modalSub.textContent = `Filtros: almacén=${filtros.store || "Todos"}, tipo=${filtros.tipo}, categoría=${filtros.categoria || "Todas"}`;

  const cats = unique(datos.map(r => r.categoria || "(Sin categoría)"));
  const filas = cats.map(cat => {
    const aPrev = dataPrev.filter(r => (r.categoria || "(Sin categoría)") === cat);
    const aCur = dataCur.filter(r => (r.categoria || "(Sin categoría)") === cat);
    return { cat, prev: calc(aPrev), cur: calc(aCur) };
  });

  modalTableBody.innerHTML = "";
  filas
    .sort((a, b) => b.cur - a.cur)
    .forEach(f => {
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
        <td>${f.cat}</td>
        <td class="text-right">${fmt(f.prev)}</td>
        <td class="text-right">${fmt(f.cur)}</td>
        <td class="text-right"><span class="${cls}">${icon} ${crecStr}</span></td>
      `;
      modalTableBody.appendChild(tr);
    });

  modalBackdrop.classList.add("active");
}

if (modalClose && modalBackdrop) {
  modalClose.addEventListener("click", () =>
    modalBackdrop.classList.remove("active")
  );
  modalBackdrop.addEventListener("click", e => {
    if (e.target === modalBackdrop) modalBackdrop.classList.remove("active");
  });
}

// ==================== DETALLE ====================

function initDetalleHeaders() {
  if (!detalleHeaderRow || !detalleFilterRow) return;

  renderDetalleHeaders();
  detalleFiltros = {};

  if (searchGlobalInput) {
    searchGlobalInput.addEventListener(
      "input",
      debounce(() => {
        detalleBusqueda = searchGlobalInput.value.toLowerCase();
        renderDetalle();
      }, 200)
    );
  }
}

function renderDetalleHeaders() {
  detalleHeaderRow.innerHTML = "";
  detalleFilterRow.innerHTML = "";

  detalleColsActivas.forEach(col => {
    const th = document.createElement("th");
    th.classList.add("sortable", "th-drop-target");
    th.dataset.col = col;
    th.draggable = false;
    th.innerHTML = `${col} <span>⇅</span>`;
    th.addEventListener("click", () => sortDetalle(col));
    th.addEventListener("dragover", e => {
      e.preventDefault();
    });
    th.addEventListener("drop", e => {
      e.preventDefault();
      const colName = e.dataTransfer.getData("text/colName");
      if (!colName) return;
      // insertar en la posición de esta columna
      const idx = detalleColsActivas.indexOf(col);
      if (idx === -1) return;
      if (!detalleColsActivas.includes(colName)) {
        // sustituir
        detalleColsActivas[idx] = colName;
      } else {
        // reordenar
        const from = detalleColsActivas.indexOf(colName);
        detalleColsActivas.splice(from, 1);
        detalleColsActivas.splice(idx, 0, colName);
      }
      renderDetalleHeaders();
      renderDetalle();
      initDetalleChips();
      saveConfig();
    });
    detalleHeaderRow.appendChild(th);

    const thF = document.createElement("th");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "filter-input";
    inp.placeholder = "Filtrar...";
    inp.value = detalleFiltros[col] || "";
    inp.addEventListener(
      "input",
      debounce(() => {
        detalleFiltros[col] = inp.value.toLowerCase();
        renderDetalle();
      }, 150)
    );
    thF.appendChild(inp);
    detalleFilterRow.appendChild(thF);
  });
}

function initDetalleChips() {
  if (!detalleChipContainer) return;
  detalleChipContainer.innerHTML = "";

  DETALLE_COLS_ALL.forEach(col => {
    const chip = document.createElement("div");
    chip.className = "chip";
    chip.textContent = col;
    chip.dataset.col = col;
    if (detalleColsActivas.includes(col)) {
      chip.classList.add("active");
    }
    chip.draggable = true;
    chip.addEventListener("dragstart", e => {
      e.dataTransfer.setData("text/colName", col);
    });
    detalleChipContainer.appendChild(chip);
  });
}

function getValorDetalle(r, col) {
  switch (col) {
    case "Año":
      return r.anio;
    case "Fecha":
      return r.fecha.toISOString().slice(0, 10);
    case "Hora":
      return r.hora || "";
    case "Almacen":
      return r.almacen;
    case "Factura/Nota":
      return r.folio;
    case "Cliente":
      return r.cliente;
    case "Categoria":
      return r.categoria;
    case "Producto":
      return r.articulo;
    case "Tipo factura":
      return r.tipoFactura === "credito" ? "Crédito" : "Contado";
    case "Subtotal":
      return r.subtotal;
    case "Costo":
      return r.costo;
    case "Utilidad":
      return r.utilidad;
    case "Margen %":
      return r.subtotal > 0 ? r.utilidad / r.subtotal : 0;
    case "Pz.":
      return r.pz;
    case "Marca":
      return r.marca;
    case "Vendedor":
      return r.vendedor;
    default:
      return "";
  }
}

function sortDetalle(col) {
  if (detalleSort.col === col) detalleSort.asc = !detalleSort.asc;
  else {
    detalleSort.col = col;
    detalleSort.asc = true;
  }
  renderDetalle();
}

function renderDetalle() {
  if (!detalleTableBody) return;

  const base = filtrarRecords(false);
  let arr = base.slice();

  // filtros por columna
  arr = arr.filter(r => {
    for (const col of detalleColsActivas) {
      const text = (detalleFiltros[col] || "").trim();
      if (!text) continue;
      const v = getValorDetalle(r, col);
      const s =
        typeof v === "number"
          ? v.toString()
          : (v || "").toString().toLowerCase();
      if (!s.includes(text)) return false;
    }
    return true;
  });

  // filtro global
  if (detalleBusqueda) {
    arr = arr.filter(r => {
      const campos = [
        r.cliente,
        r.vendedor,
        r.folio,
        r.almacen,
        r.categoria,
        r.articulo
      ];
      return campos.some(
        c => c && c.toString().toLowerCase().includes(detalleBusqueda)
      );
    });
  }

  // ocultar márgenes negativos si aplica
  if (config.hideNegativeMargin) {
    arr = arr.filter(r => r.subtotal <= 0 || r.utilidad / r.subtotal >= 0);
  }

  // orden
  if (detalleSort.col) {
    const col = detalleSort.col;
    const asc = detalleSort.asc;
    arr.sort((a, b) => {
      const va = getValorDetalle(a, col);
      const vb = getValorDetalle(b, col);
      if (typeof va === "number" && typeof vb === "number") {
        return asc ? va - vb : vb - va;
      }
      const sa = (va || "").toString();
      const sb = (vb || "").toString();
      return asc ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }

  // limitar filas
  arr = arr.slice(0, config.maxFilasDetalle || 5000);

  detalleTableBody.innerHTML = "";
  arr.forEach(r => {
    const tr = document.createElement("tr");
    tr.classList.add("clickable-row");
    tr.dataset.origen = r.origen;
    tr.dataset.folio = r.folio;
    detalleColsActivas.forEach(col => {
      const td = document.createElement("td");
      let v = getValorDetalle(r, col);
      if (col === "Subtotal" || col === "Costo" || col === "Utilidad") {
        td.classList.add("text-right");
        const num = toNumber(v);
        td.textContent = formatCurrency(num);
        if (num < 0 && config.resaltarNegativos) td.classList.add("neg");
      } else if (col === "Margen %") {
        td.classList.add("text-right");
        td.textContent = formatPercent(v);
        if (v < 0 && config.resaltarNegativos) td.classList.add("neg");
      } else if (col === "Pz.") {
        td.classList.add("text-right");
        td.textContent = v || "";
      } else {
        td.textContent = v || "";
      }
      tr.appendChild(td);
    });
    tr.addEventListener("click", () => {
      abrirDetalleFactura(r.origen, r.folio);
    });
    detalleTableBody.appendChild(tr);
  });
}

// ==================== MODAL DETALLE FACTURA ====================

function abrirDetalleFactura(origen, folio) {
  if (!modalFacturaBackdrop) return;

  const items = records.filter(
    r => r.origen === origen && r.folio === folio
  );
  if (!items.length) return;

  const r0 = items[0];
  ultimaFacturaFiltro = { origen, folio, cliente: r0.cliente };

  modalFacturaTitle.textContent = `Detalle de ${origen === "factura" ? "factura" : "nota"} ${folio}`;
  const fechaTxt = r0.fecha.toISOString().slice(0, 10);
  modalFacturaSub.textContent =
    `Fecha: ${fechaTxt} | Cliente: ${r0.cliente || "(sin cliente)"} | Almacén: ${r0.almacen || ""} | Vendedor: ${r0.vendedor || ""}`;

  modalFacturaTableBody.innerHTML = "";
  items.forEach(it => {
    const tr = document.createElement("tr");
    const fechaLinea = it.fecha.toISOString().slice(0, 10);
    const precioUnit = it.precioUnit || (it.pz > 0 ? it.subtotal / it.pz : 0);
    tr.innerHTML = `
      <td>${fechaLinea}</td>
      <td>${it.hora || ""}</td>
      <td>${it.folio}</td>
      <td>${it.cliente || ""}</td>
      <td>${it.articulo || ""}</td>
      <td class="text-right">${(it.pz || 0).toLocaleString("es-MX")}</td>
      <td class="text-right">${formatCurrency(it.costo)}</td>
      <td class="text-right">${formatCurrency(precioUnit)}</td>
      <td>${it.vendedor || ""}</td>
    `;
    modalFacturaTableBody.appendChild(tr);
  });

  modalFacturaBackdrop.classList.add("active");
}

if (modalFacturaClose && modalFacturaBackdrop) {
  modalFacturaClose.addEventListener("click", () =>
    modalFacturaBackdrop.classList.remove("active")
  );
  modalFacturaBackdrop.addEventListener("click", e => {
    if (e.target === modalFacturaBackdrop) modalFacturaBackdrop.classList.remove("active");
  });
}

function exportTableToCsv(table, filename) {
  const rows = Array.from(table.querySelectorAll("tr"));
  const lines = rows.map(row => {
    const cells = Array.from(row.querySelectorAll("th,td"));
    return cells
      .map(c => {
        const text = c.textContent.replace(/"/g, '""');
        return `"${text}"`;
      })
      .join(",");
  });
  const csv = lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

if (modalFacturaExportXls) {
  modalFacturaExportXls.addEventListener("click", () => {
    const table = document.getElementById("modal-factura-table");
    if (!table) return;
    const name = ultimaFacturaFiltro
      ? `detalle_${ultimaFacturaFiltro.origen}_${ultimaFacturaFiltro.folio}.csv`
      : "detalle_factura.csv";
    exportTableToCsv(table, name);
  });
}

if (modalFacturaExportPdf) {
  modalFacturaExportPdf.addEventListener("click", () => {
    const table = document.getElementById("modal-factura-table");
    if (!table) return;
    const win = window.open("", "_blank");
    win.document.write("<html><head><title>Detalle factura</title></head><body>");
    win.document.write("<h3>Detalle de factura / nota</h3>");
    win.document.write(table.outerHTML);
    win.document.write("</body></html>");
    win.document.close();
    win.focus();
    win.print();
  });
}

if (modalFacturaVerDetalle) {
  modalFacturaVerDetalle.addEventListener("click", () => {
    if (!ultimaFacturaFiltro) return;
    // ir a tab Detalle y filtrar por factura
    const tabBtnDetalle = document.querySelector('.tab-btn[data-tab="tab-detalle"]');
    if (tabBtnDetalle) tabBtnDetalle.click();

    detalleFiltros["Factura/Nota"] = ultimaFacturaFiltro.folio.toString().toLowerCase();
    renderDetalleHeaders(); // reinyecta filtros
    renderDetalle();

    modalFacturaBackdrop.classList.remove("active");
  });
}

// ==================== TABS ====================

Array.from(document.querySelectorAll(".tab-btn")).forEach(btn => {
  btn.addEventListener("click", () => {
    const tabId = btn.dataset.tab;
    document
      .querySelectorAll(".tab-btn")
      .forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document
      .querySelectorAll(".tab-panel")
      .forEach(p => p.classList.remove("active"));
    const panel = document.getElementById(tabId);
    if (panel) panel.classList.add("active");
  });
});

// Inicializar tabla de detalle vacía
initDetalleHeaders();
initDetalleChips();
renderDetalle();

// ==================== CONFIG RÁPIDA EVENTOS ====================

if (qcMainMetric) {
  qcMainMetric.addEventListener("change", () => {
    config.mainMetric = qcMainMetric.value;
    saveConfig();
  });
}
if (qcChartType) {
  qcChartType.addEventListener("change", () => {
    config.chartType = qcChartType.value;
    saveConfig();
    actualizarTodo();
  });
}
if (qcTableMode) {
  qcTableMode.addEventListener("change", () => {
    config.tableMode = qcTableMode.value;
    document.body.classList.toggle("table-comfortable", config.tableMode === "comfortable");
    saveConfig();
  });
}
if (qcHideNegativeMargin) {
  qcHideNegativeMargin.addEventListener("change", () => {
    config.hideNegativeMargin = qcHideNegativeMargin.checked;
    saveConfig();
    renderDetalle();
  });
}

// CONFIG TAB

if (confMaxFilas) {
  confMaxFilas.addEventListener("change", () => {
    const v = parseInt(confMaxFilas.value, 10);
    if (!isNaN(v) && v > 0) {
      config.maxFilasDetalle = v;
      saveConfig();
      renderDetalle();
    }
  });
}

if (confResaltarNegativos) {
  confResaltarNegativos.addEventListener("change", () => {
    config.resaltarNegativos = confResaltarNegativos.checked;
    saveConfig();
    renderDetalle();
  });
}

if (confGuardarVista) {
  confGuardarVista.addEventListener("change", () => {
    config.guardarVista = confGuardarVista.checked;
    saveConfig();
  });
}

if (confReset) {
  confReset.addEventListener("click", () => {
    config = {
      mainMetric: "ventas",
      chartType: "bar",
      tableMode: "compact",
      hideNegativeMargin: false,
      maxFilasDetalle: 5000,
      resaltarNegativos: true,
      guardarVista: true,
      expDimension: "mes",
      expMetrica: "subtotal",
      expChartType: "bar",
      expExcluirNegativos: true
    };
    saveConfig();
    applyConfigToUI();
    actualizarTodo();
  });
}

// ==================== EXPLORADOR DE GRÁFICOS ====================

function actualizarExplorador() {
  const canvas = document.getElementById("exp-chart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  const data = filtrarRecords(false);
  if (!data.length) return;

  const dimension = config.expDimension;
  const metrica = config.expMetrica;
  const chartType = config.expChartType;
  const excluirNeg = config.expExcluirNegativos;

  const grupos = {};

  function keyAndLabel(r) {
    if (dimension === "mes") {
      const idx = r.mes - 1;
      const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
      return { key: r.mes, label: meses[idx] || r.mes };
    }
    if (dimension === "almacen") {
      const k = r.almacen || "(sin)";
      return { key: k, label: k };
    }
    if (dimension === "categoria") {
      const k = r.categoria || "(Sin categoría)";
      return { key: k, label: k };
    }
    if (dimension === "vendedor") {
      const k = r.vendedor || "(Sin vendedor)";
      return { key: k, label: k };
    }
    if (dimension === "cliente") {
      const k = r.cliente || "(Sin cliente)";
      return { key: k, label: k };
    }
    return { key: "otros", label: "Otros" };
  }

  const opsSet = {};

  data.forEach(r => {
    const { key, label } = keyAndLabel(r);
    if (!grupos[key]) {
      grupos[key] = {
        label,
        ventas: 0,
        utilidad: 0,
        ventasCred: 0,
        ventasTot: 0
      };
      opsSet[key] = new Set();
    }
    grupos[key].ventas += r.subtotal;
    if (r.incluirUtilidad) grupos[key].utilidad += r.utilidad;
    if (r.esCredito) grupos[key].ventasCred += r.subtotal;
    grupos[key].ventasTot += r.subtotal;

    const opKey = r.origen + "|" + r.folio;
    opsSet[key].add(opKey);
  });

  const items = Object.keys(grupos).map(k => {
    const g = grupos[k];
    let valor = 0;
    if (metrica === "subtotal") valor = g.ventas;
    else if (metrica === "utilidad") valor = g.utilidad;
    else if (metrica === "margen")
      valor = g.ventas > 0 ? g.utilidad / g.ventas : 0;
    else if (metrica === "ops")
      valor = opsSet[k].size;
    else if (metrica === "cred_pct")
      valor = g.ventasTot > 0 ? g.ventasCred / g.ventasTot : 0;
    return { key: k, label: g.label, valor };
  });

  let filtrados = items;
  if (excluirNeg && (metrica === "margen" || metrica === "cred_pct")) {
    filtrados = items.filter(i => i.valor >= 0);
  }

  filtrados.sort((a, b) => b.valor - a.valor);

  const labels = filtrados.map(i => i.label);
  const valores = filtrados.map(i => i.valor);

  if (expChart) expChart.destroy();
  expChart = new Chart(ctx, {
    type: chartType,
    data: {
      labels,
      datasets: [{
        label:
          metrica === "subtotal" ? "Ventas" :
          metrica === "utilidad" ? "Utilidad" :
          metrica === "margen" ? "Margen %" :
          metrica === "ops" ? "Operaciones" :
          "% Crédito",
        data: valores
      }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.parsed;
              if (metrica === "subtotal" || metrica === "utilidad") return formatCurrency(v);
              if (metrica === "margen" || metrica === "cred_pct") return formatPercent(v);
              return v.toLocaleString("es-MX");
            }
          }
        }
      },
      scales: chartType === "pie" || chartType === "doughnut" ? {} : {
        y: {
          ticks: {
            callback: v => {
              if (metrica === "subtotal" || metrica === "utilidad") return formatCurrency(v);
              if (metrica === "margen" || metrica === "cred_pct") return formatPercent(v);
              return v.toLocaleString("es-MX");
            }
          }
        }
      }
    }
  });

  if (expChartTitle) {
    const dimTxt = {
      mes: "Mes",
      almacen: "Almacén",
      categoria: "Categoría",
      vendedor: "Vendedor",
      cliente: "Cliente"
    }[dimension] || "Dimensión";

    const metTxt = {
      subtotal: "Ventas (Subtotal)",
      utilidad: "Utilidad bruta",
      margen: "Margen %",
      ops: "Número de operaciones",
      cred_pct: "% Ventas a crédito"
    }[metrica] || "Métrica";

    expChartTitle.textContent = `${metTxt} por ${dimTxt}`;
  }
}

// eventos explorador

if (expDimensionSel) {
  expDimensionSel.addEventListener("change", () => {
    config.expDimension = expDimensionSel.value;
    saveConfig();
    actualizarExplorador();
  });
}
if (expMetricaSel) {
  expMetricaSel.addEventListener("change", () => {
    config.expMetrica = expMetricaSel.value;
    saveConfig();
    actualizarExplorador();
  });
}
if (expChartTypeSel) {
  expChartTypeSel.addEventListener("change", () => {
    config.expChartType = expChartTypeSel.value;
    saveConfig();
    actualizarExplorador();
  });
}
if (expExcluirNegSel) {
  expExcluirNegSel.addEventListener("change", () => {
    config.expExcluirNegativos = expExcluirNegSel.checked;
    saveConfig();
    actualizarExplorador();
  });
}
