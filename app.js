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

// ==================== ESTADO ====================

let records = [];
let yearsDisponibles = [];
let almacenesDisponibles = [];
let categoriasDisponibles = [];

let charts = { mensual: null, almacen: null };

let detalleSort = { col: null, asc: true };
let detalleFiltros = {};
let detalleBusqueda = "";

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

const modalBackdrop = document.getElementById("modal-backdrop");
const modalClose = document.getElementById("modal-close");
const modalTitle = document.getElementById("modal-title");
const modalSub = document.getElementById("modal-sub");
const modalYearPrev = document.getElementById("modal-year-prev");
const modalYearCurrent = document.getElementById("modal-year-current");
const modalTableBody = document.querySelector("#modal-table tbody");

// ==================== LECTURA DEL ARCHIVO ====================

fileInput.addEventListener("change", handleFile);

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

      // encabezados de facturas para localizar columna BB
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

      // -------- FACTURAS ----------
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

      // -------- NOTAS (REM) ----------
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
    year: filterYear.value === "all" ? null : parseInt(filterYear.value, 10),
    store: filterStore.value === "all" ? null : filterStore.value,
    tipo: filterType.value,
    categoria: filterCategory.value === "all" ? null : filterCategory.value
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

filterYear.addEventListener("change", actualizarTodo);
filterStore.addEventListener("change", actualizarTodo);
filterType.addEventListener("change", actualizarTodo);
filterCategory.addEventListener("change", actualizarTodo);

// ==================== KPIs, GRÁFICAS, TOPS ====================

function actualizarTodo() {
  if (!records.length) return;
  const datos = filtrarRecords(false);
  actualizarKpis(datos);
  actualizarGraficas(datos);
  actualizarTop(datos);
  actualizarYoY();
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

  kpiVentas.textContent = formatCurrency(ventas);
  kpiVentasSub.textContent = `Periodo: ${yearTxt}, almacén: ${almTxt}`;
  kpiUtilidad.textContent = formatCurrency(utilidad);
  kpiMargen.textContent = formatPercent(margen);
  kpiMargenSub.textContent = `Basado en ${utilData.length} filas válidas`;

  let m2 = 0;
  if (filtros.store && M2_POR_ALMACEN[filtros.store]) {
    m2 = M2_POR_ALMACEN[filtros.store];
  } else {
    m2 = M2_POR_ALMACEN["Todas"];
  }
  const ventasM2 = m2 > 0 ? ventas / m2 : 0;
  kpiM2.textContent = formatCurrency(ventasM2) + "/m²";
  kpiM2Sub.textContent = `${filtros.store || "Todas"} – ${m2.toLocaleString(
    "es-MX"
  )} m²`;

  const opsSet = new Set(data.map(r => r.origen + "|" + r.folio));
  kpiTrans.textContent = opsSet.size.toLocaleString("es-MX");

  const clientesSet = new Set(
    data.map(r => r.cliente || "").filter(x => x !== "")
  );
  kpiClientes.textContent = clientesSet.size.toLocaleString("es-MX");
}

function actualizarGraficas(data) {
  const ctxMensual = document
    .getElementById("chart-mensual")
    .getContext("2d");
  const ctxAlm = document.getElementById("chart-almacen").getContext("2d");

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
  const labelsMes = [
    "Ene",
    "Feb",
    "Mar",
    "Abr",
    "May",
    "Jun",
    "Jul",
    "Ago",
    "Sep",
    "Oct",
    "Nov",
    "Dic"
  ];

  if (charts.mensual) charts.mensual.destroy();
  charts.mensual = new Chart(ctxMensual, {
    type: "bar",
    data: {
      labels: labelsMes,
      datasets: [{ label: "Ventas", data: ventasMes }]
    },
    options: {
      responsive: true,
      scales: {
        y: {
          ticks: {
            callback: v => v.toLocaleString("es-MX")
          }
        }
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
      scales: {
        y: {
          ticks: {
            callback: v => v.toLocaleString("es-MX")
          }
        }
      }
    }
  });
}

function actualizarTop(data) {
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
    calc = arr =>
      sumField(
        arr.filter(r => r.incluirUtilidad),
        "utilidad"
      );
  } else if (metricId === "margen") {
    titulo = "Detalle de Margen Bruto % por categoría";
    tipo = "percent";
    calc = arr => {
      const v = sumField(arr, "subtotal");
      const u = sumField(
        arr.filter(r => r.incluirUtilidad),
        "utilidad"
      );
      return v > 0 ? u / v : 0;
    };
  } else if (metricId === "pct_credito") {
    titulo = "Detalle % Ventas a Crédito por categoría";
    tipo = "percent";
    calc = arr => {
      const vTot = sumField(arr, "subtotal");
      const vCred = sumField(
        arr.filter(r => r.esCredito),
        "subtotal"
      );
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

    modalTitle.textContent = titulo;
    modalBackdrop.classList.add("active");
    return;
  }

  const filtros = getFiltros();
  modalTitle.textContent = titulo;
  modalSub.textContent = `Filtros: almacén=${filtros.store ||
    "Todos"}, tipo=${filtros.tipo}, categoría=${filtros.categoria || "Todas"}`;

  const cats = unique(datos.map(r => r.categoria || "(Sin categoría)"));
  const filas = cats.map(cat => {
    const aPrev = dataPrev.filter(
      r => (r.categoria || "(Sin categoría)") === cat
    );
    const aCur = dataCur.filter(
      r => (r.categoria || "(Sin categoría)") === cat
    );
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

modalClose.addEventListener("click", () =>
  modalBackdrop.classList.remove("active")
);
modalBackdrop.addEventListener("click", e => {
  if (e.target === modalBackdrop) modalBackdrop.classList.remove("active");
});

// ==================== DETALLE ====================

function initDetalleHeaders() {
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

  searchGlobalInput.addEventListener(
    "input",
    debounce(() => {
      detalleBusqueda = searchGlobalInput.value.toLowerCase();
      renderDetalle();
    }, 200)
  );
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
      if (
        col === "Subtotal" ||
        col === "Costo" ||
        col === "Utilidad"
      ) {
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
    document.getElementById(tabId).classList.add("active");
  });
});

// init vacío (sin datos todavía)
initDetalleHeaders();
renderDetalle();
