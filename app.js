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
  "Cse_prod": "ID Categoria",
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
  "Cse_prod": "ID Clase",
  "Total_fac": "Total Nota",
  "Nom_age": "Vendedor",
  "Des_tial": "Marca",
  "Cto_ent": "Costo Entrada",
  "Nom_fac": "Cliente",
  "Cte_fac": "ID Cliente"
};

const DETALLE_COLS = [
  "Año",
  "Fecha",
  "Almacen",
  "Factura/Nota",
  "Cliente",
  "Categoria",
  "Tipo factura",
  "Subtotal",
  "Costo",
  "Utilidad",
  "Margen %",
  "Vendedor"
];

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

let charts = {
  mensual: null,
  almacen: null,
  sucursal: null,
  creditoMix: null,
  creditoMensual: null,
  categorias: null,
  vendedores: null,
  clientes: null,
  explorer: null
};

let detalleSort = { col: null, asc: true };
let detalleFiltros = {};
let detalleBusqueda = "";

// DOM general
const fileInput = document.getElementById("file-input");
const fileNameSpan = document.getElementById("file-name");
const errorDiv = document.getElementById("error");

const filterYear = document.getElementById("filter-year");
const filterStore = document.getElementById("filter-store");
const filterType = document.getElementById("filter-type");
const filterCategory = document.getElementById("filter-category");

// KPIs resumen
const kpiVentas = document.getElementById("kpi-ventas");
const kpiVentasSub = document.getElementById("kpi-ventas-sub");
const kpiUtilidad = document.getElementById("kpi-utilidad");
const kpiMargen = document.getElementById("kpi-margen");
const kpiMargenSub = document.getElementById("kpi-margen-sub");
const kpiM2 = document.getElementById("kpi-m2");
const kpiM2Sub = document.getElementById("kpi-m2-sub");
const kpiTrans = document.getElementById("kpi-trans");
const kpiClientes = document.getElementById("kpi-clientes");

// Tablas resumen
const tablaTopClientes = document.getElementById("tabla-top-clientes");
const tablaTopVendedores = document.getElementById("tabla-top-vendedores");

// YoY
const thYearPrev = document.getElementById("th-year-prev");
const thYearCurrent = document.getElementById("th-year-current");
const tablaYoYBody = document.getElementById("tabla-yoy");

// Detalle
const detalleHeaderRow = document.getElementById("detalle-header-row");
const detalleFilterRow = document.getElementById("detalle-filter-row");
const detalleTableBody = document.getElementById("tabla-detalle");
const searchGlobalInput = document.getElementById("search-global");
const btnExportDetalle = document.getElementById("btn-export-detalle");

// Modal
const modalBackdrop = document.getElementById("modal-backdrop");
const modalClose = document.getElementById("modal-close");
const modalTitle = document.getElementById("modal-title");
const modalSub = document.getElementById("modal-sub");
const modalYearPrev = document.getElementById("modal-year-prev");
const modalYearCurrent = document.getElementById("modal-year-current");
const modalTableBody = document.querySelector("#modal-table tbody");

// Análisis por sucursal
const tablaSucursal = document.getElementById("tabla-sucursal");

// Crédito vs contado
const kpiContadoVentas = document.getElementById("kpi-contado-ventas");
const kpiContadoMargen = document.getElementById("kpi-contado-margen");
const kpiCreditoVentas = document.getElementById("kpi-credito-ventas");
const kpiCreditoMargen = document.getElementById("kpi-credito-margen");
const kpiCreditoPct = document.getElementById("kpi-credito-pct");

// Categorías
const tablaCategorias = document.getElementById("tabla-categorias");

// Vendedores
const tablaVendedoresDetalle = document.getElementById("tabla-vendedores-detalle");

// Clientes
const tablaClientesDetalle = document.getElementById("tabla-clientes-detalle");

// EXPLORADOR DE GRÁFICOS – DOM
const explorerDimSelect = document.getElementById("explorer-dimension");
const explorerMetricSelect = document.getElementById("explorer-metric");
const explorerTypeSelect = document.getElementById("explorer-type");
const explorerTopSelect = document.getElementById("explorer-top");
const explorerCanvas = document.getElementById("chart-explorer");
const explorerViewNameInput = document.getElementById("explorer-view-name");
const explorerSaveBtn = document.getElementById("explorer-save-view");
const explorerViewsSelect = document.getElementById("explorer-saved-views");

// localStorage key
const EXPLORER_VIEWS_KEY = "cedro_dashboard_explorer_views";

// ==================== EXPLORADOR – CONFIG ====================

const EXPLORER_DIMENSIONS = {
  anio: {
    label: "Año",
    keyFn: r => r.anio,
    labelFn: k => String(k)
  },
  mes: {
    label: "Mes",
    keyFn: r => r.mes,
    labelFn: k => {
      const idx = parseInt(k, 10) - 1;
      const nombres = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
      return nombres[idx] || `Mes ${k}`;
    }
  },
  almacen: {
    label: "Almacén",
    keyFn: r => r.almacen || "(Sin almacén)",
    labelFn: k => k
  },
  categoria: {
    label: "Categoría",
    keyFn: r => r.categoria || "(Sin categoría)",
    labelFn: k => k
  },
  vendedor: {
    label: "Vendedor",
    keyFn: r => r.vendedor || "(Sin vendedor)",
    labelFn: k => k
  },
  cliente: {
    label: "Cliente",
    keyFn: r => r.cliente || "(Sin cliente)",
    labelFn: k => k
  },
  tipo: {
    label: "Tipo de factura",
    keyFn: r => r.tipoFactura,
    labelFn: k => (k === "credito" ? "Crédito" : k === "contado" ? "Contado" : k)
  }
};

const EXPLORER_METRICS = {
  ventas: {
    label: "Ventas (Subtotal)",
    type: "money",
    calc: rows => sumField(rows, "subtotal")
  },
  utilidad: {
    label: "Utilidad bruta",
    type: "money",
    calc: rows => sumField(rows.filter(r => r.incluirUtilidad), "utilidad")
  },
  margen: {
    label: "Margen bruto %",
    type: "percent",
    calc: rows => {
      const v = sumField(rows, "subtotal");
      const u = sumField(rows.filter(r => r.incluirUtilidad), "utilidad");
      return v > 0 ? u / v : 0;
    }
  },
  operaciones: {
    label: "Transacciones únicas",
    type: "plain",
    calc: rows => {
      const set = new Set(rows.map(r => r.origen + "|" + r.folio));
      return set.size;
    }
  },
  clientes: {
    label: "Clientes únicos",
    type: "plain",
    calc: rows => {
      const set = new Set(rows.map(r => r.cliente || "").filter(x => x));
      return set.size;
    }
  }
};

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

        const categoria = (row["ID Categoria"] || "").toString().trim();
        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Factura"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();

        records.push({
          origen: "factura",
          anio,
          mes,
          fecha,
          almacen,
          categoria,
          cliente,
          vendedor,
          marca,
          folio,
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
        const categoria = (row["ID Clase"] || "").toString().trim();
        const cliente = (row["Cliente"] || "").toString().trim();
        const vendedor = (row["Vendedor"] || "").toString().trim();
        const folio = (row["Nota"] || "").toString().trim();
        const marca = (row["Marca"] || "").toString().trim();

        records.push({
          origen: "nota",
          anio,
          mes,
          fecha,
          almacen,
          categoria,
          cliente,
          vendedor,
          marca,
          folio,
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
      actualizarTodo();
      initExplorer();
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

// versión para sucursal (ignora almacén)
function filtrarRecordsSinAlmacen() {
  const f = getFiltros();
  return records.filter(r => {
    if (f.year && r.anio !== f.year) return false;
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
  actualizarAnalisisSucursales();
  actualizarAnalisisCredito(datos);
  actualizarAnalisisCategorias(datos);
  actualizarAnalisisVendedores(datos);
  actualizarAnalisisClientes(datos);
  renderExplorerChart();
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
    type: "bar",
    data: {
      labels: labelsMes,
      datasets: [{ label: "Ventas", data: ventasMes }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            title: items => `Mes: ${items[0].label}`,
            label: item => {
              const valor = item.parsed.y || 0;
              const total = ventasMes.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? (valor / total) * 100 : 0;
              return [
                `Ventas: ${formatCurrency(valor)}`,
                `Participación: ${pct.toFixed(1)}%`
              ];
            }
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const mes = idx + 1;
        abrirDetalleMensual(mes);
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });

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
      plugins: {
        tooltip: {
          callbacks: {
            title: items => `Almacén: ${items[0].label}`,
            label: item => {
              const datasetLabel = item.dataset.label || "";
              const valor = item.parsed.y || 0;
              if (datasetLabel.includes("m²")) {
                return `${datasetLabel}: ${formatCurrency(valor)}/m²`;
              }
              return `${datasetLabel}: ${formatCurrency(valor)}`;
            }
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const almacenSeleccionado = labsAlm[idx];
        if (filterStore) {
          filterStore.value = almacenSeleccionado;
          actualizarTodo();
        }
      },
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

    const thead = document.querySelector("#modal-table thead tr");
    if (thead) {
      thead.innerHTML = `
        <th>Almacén</th>
        <th class="text-right">m²</th>
        <th class="text-right">m²</th>
        <th class="text-right">Crecimiento %</th>
      `;
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

  const thead = document.querySelector("#modal-table thead tr");
  if (thead) {
    thead.innerHTML = `
      <th>Categoría</th>
      <th id="modal-year-prev" class="text-right">${yPrev}</th>
      <th id="modal-year-current" class="text-right">${yCur}</th>
      <th class="text-right">Crecimiento %</th>
    `;
  }

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

      if (metricId !== "m2") {
        tr.style.cursor = "pointer";
        tr.addEventListener("click", () => irADetalleConCategoria(f.cat));
      }

      modalTableBody.appendChild(tr);
    });

  modalBackdrop.classList.add("active");
}

function abrirDetalleMensual(mesSeleccionado) {
  if (!modalBackdrop || !modalTitle || !modalSub || !modalTableBody) return;

  const filtros = getFiltros();

  let yearBase = filtros.year;
  if (!yearBase) {
    yearBase = Math.max(...yearsDisponibles);
  }

  const datos = filtrarRecords(false).filter(r => {
    return r.anio === yearBase && r.mes === mesSeleccionado;
  });

  const mesNombre = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][mesSeleccionado - 1];

  modalTitle.textContent = `Detalle del mes: ${mesNombre} ${yearBase}`;
  modalSub.textContent = `Filtros: almacén=${filtros.store || "Todos"}, tipo=${filtros.tipo}, categoría=${filtros.categoria || "Todas"}`;

  const porCat = {};
  datos.forEach(r => {
    const cat = r.categoria || "(Sin categoría)";
    if (!porCat[cat]) {
      porCat[cat] = { ventas: 0, utilidad: 0 };
    }
    porCat[cat].ventas += r.subtotal;
    if (r.incluirUtilidad) {
      porCat[cat].utilidad += r.utilidad;
    }
  });

  const filas = Object.entries(porCat).map(([cat, d]) => {
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    return { cat, ventas: d.ventas, utilidad: d.utilidad, margen };
  }).sort((a, b) => b.ventas - a.ventas);

  const thead = document.querySelector("#modal-table thead tr");
  if (thead) {
    thead.innerHTML = `
      <th>Categoría</th>
      <th class="text-right">Ventas</th>
      <th class="text-right">Utilidad</th>
      <th class="text-right">Margen %</th>
    `;
  }

  modalTableBody.innerHTML = "";
  filas.forEach(f => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.cat}</td>
      <td class="text-right">${formatCurrency(f.ventas)}</td>
      <td class="text-right">${formatCurrency(f.utilidad)}</td>
      <td class="text-right">${formatPercent(f.margen)}</td>
    `;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => irADetalleConCategoria(f.cat));
    modalTableBody.appendChild(tr);
  });

  modalBackdrop.classList.add("active");
}

function irADetalleConCategoria(cat) {
  if (!cat || cat === "(Sin categoría)") return;
  if (filterCategory) {
    filterCategory.value = cat;
  }
  if (modalBackdrop) modalBackdrop.classList.remove("active");
  activarTab("tab-detalle");
  actualizarTodo();
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

  detalleHeaderRow.innerHTML = "";
  detalleFilterRow.innerHTML = "";
  detalleFiltros = {};

  DETALLE_COLS.forEach(col => {
    const th = document.createElement("th");
    th.textContent = col;
    th.classList.add("sortable");
    th.addEventListener("click", () => sortDetalle(col));
    detalleHeaderRow.appendChild(th);

    const thF = document.createElement("th");
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "filter-input";
    inp.placeholder = "Filtrar...";
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

  if (searchGlobalInput) {
    searchGlobalInput.addEventListener(
      "input",
      debounce(() => {
        detalleBusqueda = searchGlobalInput.value.toLowerCase();
        renderDetalle();
      }, 200)
    );
  }

  if (btnExportDetalle) {
    btnExportDetalle.addEventListener("click", exportDetalleToExcel);
  }
}

function getValorDetalle(r, col) {
  switch (col) {
    case "Año":
      return r.anio;
    case "Fecha":
      return r.fecha.toISOString().slice(0, 10);
    case "Almacen":
      return r.almacen;
    case "Factura/Nota":
      return r.folio;
    case "Cliente":
      return r.cliente;
    case "Categoria":
      return r.categoria;
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
    for (const col in detalleFiltros) {
      const text = detalleFiltros[col];
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
        r.categoria
      ];
      return campos.some(
        c => c && c.toString().toLowerCase().includes(detalleBusqueda)
      );
    });
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

  detalleTableBody.innerHTML = "";
  arr.forEach(r => {
    const tr = document.createElement("tr");
    DETALLE_COLS.forEach(col => {
      const td = document.createElement("td");
      let v = getValorDetalle(r, col);
      if (col === "Subtotal" || col === "Costo" || col === "Utilidad") {
        td.classList.add("text-right");
        const num = toNumber(v);
        td.textContent = formatCurrency(num);
        if (num < 0) td.classList.add("neg");
      } else if (col === "Margen %") {
        td.classList.add("text-right");
        td.textContent = formatPercent(v);
        if (v < 0) td.classList.add("neg");
      } else {
        td.textContent = v || "";
      }
      tr.appendChild(td);
    });
    detalleTableBody.appendChild(tr);
  });
}

// ==================== ANÁLISIS POR SUCURSAL ====================

function actualizarAnalisisSucursales() {
  const data = filtrarRecordsSinAlmacen();
  if (!data.length || !tablaSucursal) return;

  const porAlm = {};
  data.forEach(r => {
    const a = r.almacen || "(Sin almacén)";
    if (!porAlm[a]) {
      porAlm[a] = {
        ventas: 0,
        utilidad: 0,
        nUtil: 0,
        ops: new Set(),
        clientes: new Set()
      };
    }
    porAlm[a].ventas += r.subtotal;
    if (r.incluirUtilidad) {
      porAlm[a].utilidad += r.utilidad;
      porAlm[a].nUtil++;
    }
    porAlm[a].ops.add(r.origen + "|" + r.folio);
    if (r.cliente) porAlm[a].clientes.add(r.cliente);
  });

  const filas = Object.entries(porAlm).map(([alm, d]) => {
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    const m2 = M2_POR_ALMACEN[alm] || 0;
    const vM2 = m2 > 0 ? d.ventas / m2 : 0;
    return {
      almacén: alm,
      ventas: d.ventas,
      utilidad: d.utilidad,
      margen,
      ventasM2: vM2,
      trans: d.ops.size,
      clientes: d.clientes.size
    };
  }).sort((a, b) => b.ventas - a.ventas);

  tablaSucursal.innerHTML = "";
  filas.forEach(f => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.almacén}</td>
      <td class="text-right">${formatCurrency(f.ventas)}</td>
      <td class="text-right">${formatCurrency(f.utilidad)}</td>
      <td class="text-right">${formatPercent(f.margen)}</td>
      <td class="text-right">${formatCurrency(f.ventasM2)}/m²</td>
      <td class="text-right">${f.trans.toLocaleString("es-MX")}</td>
      <td class="text-right">${f.clientes.toLocaleString("es-MX")}</td>
    `;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => {
      if (filterStore) {
        filterStore.value = f.almacén;
        activarTab("tab-detalle");
        actualizarTodo();
      }
    });
    tablaSucursal.appendChild(tr);
  });

  const canvas = document.getElementById("chart-sucursal");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  const labels = filas.map(f => f.almacén);
  const datosVentas = filas.map(f => f.ventas);
  const datosVentasM2 = filas.map(f => f.ventasM2);

  if (charts.sucursal) charts.sucursal.destroy();
  charts.sucursal = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "Ventas", data: datosVentas },
        { label: "Ventas por m²", data: datosVentasM2 }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => {
              const label = item.dataset.label;
              const v = item.parsed.y || 0;
              if (label.includes("m²")) {
                return `${label}: ${formatCurrency(v)}/m²`;
              }
              return `${label}: ${formatCurrency(v)}`;
            }
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const alm = labels[idx];
        if (filterStore) {
          filterStore.value = alm;
          activarTab("tab-resumen");
          actualizarTodo();
        }
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });
}

// ==================== CRÉDITO VS CONTADO ====================

function actualizarAnalisisCredito(data) {
  if (!data.length) return;

  const cont = data.filter(r => r.tipoFactura === "contado");
  const cred = data.filter(r => r.tipoFactura === "credito");

  const vCont = sumField(cont, "subtotal");
  const vCred = sumField(cred, "subtotal");
  const vTot = vCont + vCred;

  const uCont = sumField(cont.filter(r => r.incluirUtilidad), "utilidad");
  const uCred = sumField(cred.filter(r => r.incluirUtilidad), "utilidad");

  const mCont = vCont > 0 ? uCont / vCont : 0;
  const mCred = vCred > 0 ? uCred / vCred : 0;
  const pctCred = vTot > 0 ? vCred / vTot : 0;

  if (kpiContadoVentas) kpiContadoVentas.textContent = formatCurrency(vCont);
  if (kpiContadoMargen) kpiContadoMargen.textContent = `Margen: ${formatPercent(mCont)}`;
  if (kpiCreditoVentas) kpiCreditoVentas.textContent = formatCurrency(vCred);
  if (kpiCreditoMargen) kpiCreditoMargen.textContent = `Margen: ${formatPercent(mCred)}`;
  if (kpiCreditoPct) kpiCreditoPct.textContent = formatPercent(pctCred);

  const canvasMix = document.getElementById("chart-credito-mix");
  const canvasMensual = document.getElementById("chart-credito-mensual");
  if (!canvasMix || !canvasMensual) return;

  const ctxMix = canvasMix.getContext("2d");
  const ctxMensual = canvasMensual.getContext("2d");

  if (charts.creditoMix) charts.creditoMix.destroy();
  charts.creditoMix = new Chart(ctxMix, {
    type: "doughnut",
    data: {
      labels: ["Contado", "Crédito"],
      datasets: [{
        data: [vCont, vCred]
      }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => {
              const v = item.parsed || 0;
              const pct = vTot > 0 ? (v / vTot) * 100 : 0;
              return `${item.label}: ${formatCurrency(v)} (${pct.toFixed(1)}%)`;
            }
          }
        }
      }
    }
  });

  const labelsMes = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
  const contMes = new Array(12).fill(0);
  const credMes = new Array(12).fill(0);

  cont.forEach(r => {
    const i = r.mes - 1;
    if (i >= 0 && i < 12) contMes[i] += r.subtotal;
  });
  cred.forEach(r => {
    const i = r.mes - 1;
    if (i >= 0 && i < 12) credMes[i] += r.subtotal;
  });

  if (charts.creditoMensual) charts.creditoMensual.destroy();
  charts.creditoMensual = new Chart(ctxMensual, {
    type: "line",
    data: {
      labels: labelsMes,
      datasets: [
        { label: "Contado", data: contMes },
        { label: "Crédito", data: credMes }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => `${item.dataset.label}: ${formatCurrency(item.parsed.y || 0)}`
          }
        }
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });
}

// ==================== CATEGORÍAS ====================

function actualizarAnalisisCategorias(data) {
  if (!data.length || !tablaCategorias) return;

  const porCat = {};
  data.forEach(r => {
    const c = r.categoria || "(Sin categoría)";
    if (!porCat[c]) porCat[c] = { ventas: 0, utilidad: 0 };
    porCat[c].ventas += r.subtotal;
    if (r.incluirUtilidad) porCat[c].utilidad += r.utilidad;
  });

  const totalVentas = Object.values(porCat).reduce((acc, d) => acc + d.ventas, 0);

  const filas = Object.entries(porCat).map(([cat, d]) => {
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    const pct = totalVentas > 0 ? d.ventas / totalVentas : 0;
    return { cat, ventas: d.ventas, utilidad: d.utilidad, margen, pct };
  }).sort((a, b) => b.ventas - a.ventas);

  tablaCategorias.innerHTML = "";
  filas.forEach(f => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.cat}</td>
      <td class="text-right">${formatCurrency(f.ventas)}</td>
      <td class="text-right">${formatCurrency(f.utilidad)}</td>
      <td class="text-right">${formatPercent(f.margen)}</td>
      <td class="text-right">${formatPercent(f.pct)}</td>
    `;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => irADetalleConCategoria(f.cat));
    tablaCategorias.appendChild(tr);
  });

  const canvas = document.getElementById("chart-categorias");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  const topN = filas.slice(0, 10);
  const labels = topN.map(f => f.cat);
  const dataVentas = topN.map(f => f.ventas);

  if (charts.categorias) charts.categorias.destroy();
  charts.categorias = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Ventas", data: dataVentas }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => `Ventas: ${formatCurrency(item.parsed.y || 0)}`
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const cat = labels[idx];
        irADetalleConCategoria(cat);
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });
}

// ==================== VENDEDORES ====================

function actualizarAnalisisVendedores(data) {
  if (!data.length || !tablaVendedoresDetalle) return;

  const vend = {};
  const opsPorVend = {};
  data.forEach(r => {
    const v = r.vendedor || "(Sin vendedor)";
    if (!vend[v]) vend[v] = { ventas: 0, utilidad: 0 };
    vend[v].ventas += r.subtotal;
    if (r.incluirUtilidad) vend[v].utilidad += r.utilidad;

    const key = r.origen + "|" + r.folio;
    if (!opsPorVend[v]) opsPorVend[v] = new Set();
    opsPorVend[v].add(key);
  });

  const filas = Object.entries(vend).map(([nombre, d]) => {
    const ops = opsPorVend[nombre] ? opsPorVend[nombre].size : 0;
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    return { nombre, ventas: d.ventas, utilidad: d.utilidad, margen, ops };
  }).sort((a, b) => b.ventas - a.ventas);

  tablaVendedoresDetalle.innerHTML = "";
  filas.forEach(f => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.nombre}</td>
      <td class="text-right">${formatCurrency(f.ventas)}</td>
      <td class="text-right">${formatCurrency(f.utilidad)}</td>
      <td class="text-right">${formatPercent(f.margen)}</td>
      <td class="text-right">${f.ops.toLocaleString("es-MX")}</td>
    `;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => {
      // filtrar detalle por vendedor
      detalleFiltros["Vendedor"] = (f.nombre || "").toLowerCase();
      if (searchGlobalInput) searchGlobalInput.value = "";
      activarTab("tab-detalle");
      renderDetalle();
    });
    tablaVendedoresDetalle.appendChild(tr);
  });

  const canvas = document.getElementById("chart-vendedores");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  const topN = filas.slice(0, 10);
  const labels = topN.map(f => f.nombre);
  const dataVentas = topN.map(f => f.ventas);

  if (charts.vendedores) charts.vendedores.destroy();
  charts.vendedores = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Ventas", data: dataVentas }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => `Ventas: ${formatCurrency(item.parsed.y || 0)}`
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const nombre = labels[idx];
        detalleFiltros["Vendedor"] = (nombre || "").toLowerCase();
        activarTab("tab-detalle");
        renderDetalle();
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });
}

// ==================== CLIENTES ====================

function actualizarAnalisisClientes(data) {
  if (!data.length || !tablaClientesDetalle) return;

  const clientes = {};
  const opsPorCliente = {};
  data.forEach(r => {
    const c = r.cliente || "(Sin cliente)";
    if (!clientes[c]) clientes[c] = { ventas: 0, utilidad: 0 };
    clientes[c].ventas += r.subtotal;
    if (r.incluirUtilidad) clientes[c].utilidad += r.utilidad;

    const key = r.origen + "|" + r.folio;
    if (!opsPorCliente[c]) opsPorCliente[c] = new Set();
    opsPorCliente[c].add(key);
  });

  const filas = Object.entries(clientes).map(([nombre, d]) => {
    const ops = opsPorCliente[nombre] ? opsPorCliente[nombre].size : 0;
    const margen = d.ventas > 0 ? d.utilidad / d.ventas : 0;
    return { nombre, ventas: d.ventas, utilidad: d.utilidad, margen, ops };
  }).sort((a, b) => b.ventas - a.ventas);

  tablaClientesDetalle.innerHTML = "";
  filas.forEach(f => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${f.nombre}</td>
      <td class="text-right">${formatCurrency(f.ventas)}</td>
      <td class="text-right">${formatCurrency(f.utilidad)}</td>
      <td class="text-right">${formatPercent(f.margen)}</td>
      <td class="text-right">${f.ops.toLocaleString("es-MX")}</td>
    `;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => {
      detalleFiltros["Cliente"] = (f.nombre || "").toLowerCase();
      activarTab("tab-detalle");
      renderDetalle();
    });
    tablaClientesDetalle.appendChild(tr);
  });

  const canvas = document.getElementById("chart-clientes");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  const topN = filas.slice(0, 10);
  const labels = topN.map(f => f.nombre);
  const dataVentas = topN.map(f => f.ventas);

  if (charts.clientes) charts.clientes.destroy();
  charts.clientes = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Ventas", data: dataVentas }]
    },
    options: {
      responsive: true,
      plugins: {
        tooltip: {
          callbacks: {
            label: item => `Ventas: ${formatCurrency(item.parsed.y || 0)}`
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const nombre = labels[idx];
        detalleFiltros["Cliente"] = (nombre || "").toLowerCase();
        activarTab("tab-detalle");
        renderDetalle();
      },
      scales: {
        y: { ticks: { callback: v => v.toLocaleString("es-MX") } }
      }
    }
  });
}

// ==================== EXPLORADOR – LÓGICA ====================

function renderExplorerChart() {
  if (!explorerCanvas || !explorerDimSelect || !explorerMetricSelect || !explorerTypeSelect || !explorerTopSelect) return;
  if (!records.length) return;

  const dimKey = explorerDimSelect.value;
  const metKey = explorerMetricSelect.value;
  const chartType = explorerTypeSelect.value;
  const topN = parseInt(explorerTopSelect.value, 10) || 0;

  const dimCfg = EXPLORER_DIMENSIONS[dimKey];
  const metCfg = EXPLORER_METRICS[metKey];
  if (!dimCfg || !metCfg) return;

  const datos = filtrarRecords(false);
  const grupos = {};

  datos.forEach(r => {
    let k = dimCfg.keyFn(r);
    if (k === undefined || k === null || k === "") k = "(Sin valor)";
    if (!grupos[k]) grupos[k] = [];
    grupos[k].push(r);
  });

  const entries = Object.entries(grupos).map(([k, rows]) => ({
    key: k,
    label: dimCfg.labelFn ? dimCfg.labelFn(k) : k,
    value: metCfg.calc(rows)
  }));

  entries.sort((a, b) => b.value - a.value);
  let recs = entries;
  if (topN > 0 && entries.length > topN) {
    recs = entries.slice(0, topN);
  }

  const labels = recs.map(e => e.label);
  const data = recs.map(e => e.value);

  if (charts.explorer) charts.explorer.destroy();
  const ctx = explorerCanvas.getContext("2d");

  const baseType = chartType === "bar-horizontal" ? "bar" : chartType;

  charts.explorer = new Chart(ctx, {
    type: baseType,
    data: {
      labels,
      datasets: [
        {
          label: metCfg.label,
          data
        }
      ]
    },
    options: {
      responsive: true,
      indexAxis: chartType === "bar-horizontal" ? "y" : "x",
      plugins: {
        tooltip: {
          callbacks: {
            label: item => {
              const v = item.parsed.y ?? item.parsed;
              if (metCfg.type === "money") return `${metCfg.label}: ${formatCurrency(v)}`;
              if (metCfg.type === "percent") return `${metCfg.label}: ${(v * 100).toFixed(1)}%`;
              return `${metCfg.label}: ${v.toLocaleString("es-MX")}`;
            }
          }
        }
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        const rec = recs[idx];
        manejarClickExplorer(dimKey, rec.key);
      },
      scales: baseType === "pie" || baseType === "doughnut"
        ? {}
        : {
            y: {
              ticks: {
                callback: v => {
                  if (metCfg.type === "money") return v.toLocaleString("es-MX");
                  return v;
                }
              }
            }
          }
    }
  });

  const titulo = document.getElementById("explorer-title");
  if (titulo) {
    titulo.textContent = `${metCfg.label} por ${dimCfg.label}`;
  }
}

function manejarClickExplorer(dimKey, rawKey) {
  if (dimKey === "almacen" && filterStore) {
    filterStore.value = rawKey;
    activarTab("tab-resumen");
    actualizarTodo();
  } else if (dimKey === "categoria" && filterCategory) {
    if (rawKey === "(Sin categoría)") return;
    filterCategory.value = rawKey;
    activarTab("tab-detalle");
    actualizarTodo();
  } else if (dimKey === "anio" && filterYear) {
    filterYear.value = String(rawKey);
    activarTab("tab-resumen");
    actualizarTodo();
  } else if (dimKey === "tipo" && filterType) {
    filterType.value = rawKey;
    activarTab("tab-resumen");
    actualizarTodo();
  }
}

// ----- Vistas en localStorage -----

function cargarVistasExplorerDesdeStorage() {
  let views = [];
  try {
    const raw = localStorage.getItem(EXPLORER_VIEWS_KEY);
    if (raw) views = JSON.parse(raw) || [];
  } catch (e) {
    views = [];
  }
  actualizarSelectVistasExplorer(views);
}

function actualizarSelectVistasExplorer(views) {
  if (!explorerViewsSelect) return;
  explorerViewsSelect.innerHTML = "<option value=''>Vistas guardadas...</option>";
  views.forEach((v, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = v.name;
    explorerViewsSelect.appendChild(opt);
  });
}

function guardarVistaExplorer() {
  if (!explorerDimSelect || !explorerMetricSelect || !explorerTypeSelect || !explorerTopSelect || !explorerViewNameInput) return;

  const nombre = explorerViewNameInput.value.trim();
  if (!nombre) {
    alert("Escribe un nombre para la vista.");
    return;
  }

  let views = [];
  try {
    const raw = localStorage.getItem(EXPLORER_VIEWS_KEY);
    if (raw) views = JSON.parse(raw) || [];
  } catch (e) {
    views = [];
  }

  const nueva = {
    name: nombre,
    dim: explorerDimSelect.value,
    metric: explorerMetricSelect.value,
    type: explorerTypeSelect.value,
    top: explorerTopSelect.value
  };

  views.push(nueva);
  localStorage.setItem(EXPLORER_VIEWS_KEY, JSON.stringify(views));
  explorerViewNameInput.value = "";
  actualizarSelectVistasExplorer(views);
}

function aplicarVistaExplorerPorIndice(idx) {
  if (!explorerDimSelect || !explorerMetricSelect || !explorerTypeSelect || !explorerTopSelect) return;

  let views = [];
  try {
    const raw = localStorage.getItem(EXPLORER_VIEWS_KEY);
    if (raw) views = JSON.parse(raw) || [];
  } catch (e) {
    views = [];
  }

  const v = views[Number(idx)];
  if (!v) return;

  explorerDimSelect.value = v.dim;
  explorerMetricSelect.value = v.metric;
  explorerTypeSelect.value = v.type;
  explorerTopSelect.value = v.top;
  renderExplorerChart();
}

function initExplorer() {
  if (!explorerDimSelect || !explorerMetricSelect || !explorerTypeSelect || !explorerTopSelect || !explorerCanvas) return;

  explorerDimSelect.innerHTML = "";
  Object.entries(EXPLORER_DIMENSIONS).forEach(([key, cfg]) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = cfg.label;
    explorerDimSelect.appendChild(opt);
  });

  explorerMetricSelect.innerHTML = "";
  Object.entries(EXPLORER_METRICS).forEach(([key, cfg]) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = cfg.label;
    explorerMetricSelect.appendChild(opt);
  });

  explorerTypeSelect.innerHTML = `
    <option value="bar">Barras</option>
    <option value="bar-horizontal">Barras horizontales</option>
    <option value="line">Línea</option>
    <option value="pie">Pastel</option>
    <option value="doughnut">Dona</option>
  `;

  explorerTopSelect.innerHTML = `
    <option value="0">Todos</option>
    <option value="5">Top 5</option>
    <option value="10">Top 10</option>
    <option value="20">Top 20</option>
  `;

  explorerDimSelect.value = "categoria";
  explorerMetricSelect.value = "ventas";
  explorerTypeSelect.value = "bar";
  explorerTopSelect.value = "10";

  explorerDimSelect.addEventListener("change", renderExplorerChart);
  explorerMetricSelect.addEventListener("change", renderExplorerChart);
  explorerTypeSelect.addEventListener("change", renderExplorerChart);
  explorerTopSelect.addEventListener("change", renderExplorerChart);

  if (explorerSaveBtn) {
    explorerSaveBtn.addEventListener("click", guardarVistaExplorer);
  }

  if (explorerViewsSelect) {
    explorerViewsSelect.addEventListener("change", () => {
      if (!explorerViewsSelect.value) return;
      aplicarVistaExplorerPorIndice(explorerViewsSelect.value);
    });
  }

  cargarVistasExplorerDesdeStorage();
  renderExplorerChart();
}

// ==================== EXPORTAR A EXCEL ====================

function getDetalleFiltradoArray() {
  const base = filtrarRecords(false);
  let arr = base.slice();

  arr = arr.filter(r => {
    for (const col in detalleFiltros) {
      const text = detalleFiltros[col];
      if (!text) continue;
      const v = getValorDetalle(r, col);
      const s = typeof v === "number" ? v.toString() : (v || "").toString().toLowerCase();
      if (!s.includes(text)) return false;
    }
    return true;
  });

  if (detalleBusqueda) {
    arr = arr.filter(r => {
      const campos = [r.cliente, r.vendedor, r.folio, r.almacen, r.categoria];
      return campos.some(c => c && c.toString().toLowerCase().includes(detalleBusqueda));
    });
  }

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

  const header = DETALLE_COLS.slice();
  const data = [header];

  arr.forEach(r => {
    const row = DETALLE_COLS.map(col => {
      const v = getValorDetalle(r, col);
      if (col === "Subtotal" || col === "Costo" || col === "Utilidad") {
        return toNumber(v);
      }
      if (col === "Margen %") {
        return (r.subtotal > 0 ? r.utilidad / r.subtotal : 0);
      }
      return v || "";
    });
    data.push(row);
  });

  return data;
}

function exportDetalleToExcel() {
  if (!records.length) return;
  const aoa = getDetalleFiltradoArray();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Detalle");

  const filtros = getFiltros();
  const yearTxt = filtros.year || "todos";
  const almTxt = filtros.store || "todas";

  const fileName = `detalle_ventas_${yearTxt}_${almTxt}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

// ==================== DESCARGAR GRÁFICAS PNG / PDF ====================

function descargarGraficaPNG(key) {
  const chart = charts[key];
  if (!chart) return;
  const link = document.createElement("a");
  link.href = chart.toBase64Image("image/png", 1.0);
  link.download = `grafica_${key}.png`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function descargarGraficaPDF(key, titulo) {
  const chart = charts[key];
  if (!chart) return;

  const imgData = chart.toBase64Image("image/png", 1.0);
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: "landscape",
    unit: "pt",
    format: "a4"
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 40;
  const imgWidth = pageWidth - margin * 2;
  const imgHeight = imgWidth * 0.5;

  doc.setFontSize(14);
  doc.text(titulo || "Gráfica", margin, margin - 10);
  doc.addImage(imgData, "PNG", margin, margin, imgWidth, imgHeight);
  doc.save(`grafica_${key}.pdf`);
}

// ==================== TABS ====================

function activarTab(tabId) {
  document.querySelectorAll(".tab-btn").forEach(b => {
    b.classList.toggle("active", b.dataset.tab === tabId);
  });
  document.querySelectorAll(".tab-panel").forEach(p => {
    p.classList.toggle("active", p.id === tabId);
  });
}

Array.from(document.querySelectorAll(".tab-btn")).forEach(btn => {
  btn.addEventListener("click", () => {
    const tabId = btn.dataset.tab;
    activarTab(tabId);
  });
});

// Inicializar tabla de detalle vacía
initDetalleHeaders();
renderDetalle();
