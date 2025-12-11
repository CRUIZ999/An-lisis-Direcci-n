// =======================================================
//  CONFIGURACIÓN FIJA
// =======================================================

// m² por almacén
const M2_POR_ALMACEN = {
  "GENERAL": 1538,
  "EXPRESS": 369,
  "SAN AGUST": 870,
  "ADELITAS": 348,
  "H ILUSTRES": 100,
  "Todas": 3225
};

// Mapear columnas originales -> nombres comunes (FACTURAS)
const RENAME_FACTURAS = {
  "No_fac": "Factura",
  "Falta_fac": "Fecha",
  "Status_fac": "Status_fac",
  "Descuento": "Descuento ($)",
  "Subt_fac": "Sub. Factura",
  "Total_fac": "Total Factura",
  "Retiva_fac": "Retiva_fac",
  "Retisrfac": "Retisrfac",
  "Impto1": "Impto1",
  "Impto2": "Impto2",
  "Timpto1": "Timpto1",
  "Timpto2": "Timpto2",
  "Otrosimptos": "Otrosimptos",
  "Iva": "IVA",
  "Saldo_fac": "Saldo_fac",
  "Iva_prod": "Iva_prod",
  "Ieps_prod": "Ieps_prod",
  "Cve_factu": "Cve_factu",
  "Cve_cte": "ID Cliente",
  "Cse_prod": "ID Categoria",
  "Cve_prod": "Clave",
  "New_med": "New_med",
  "Valor_prod": "Valor_prod",
  "Cant_surt": "Pz.",
  "Cve_mon": "Cve_mon",
  "Tip_cam": "Tip_cam",
  "Unidad": "Unidad",
  "No_ped": "No_ped",
  "No_rem": "No_rem",
  "Cve_suc": "Almacen",
  "Lote": "Lote",
  "Dcto1": "Descuento (%)",
  "Dcto2": "Dcto2",
  "Cve_age": "ID Vendedor",
  "Nom_fac": "Nom_fac",
  "Cve_entre": "Cve_entre",
  "Desc_prod": "Articulo",
  "Part_fac": "Part_fac",
  "Ref_lote": "Ref_lote",
  "Cost_prom": "Costo Prom.",
  "Lugar": "Almacen2",
  "Hora_fac": "Hora",
  "Cve_age2": "Cve_age2",
  "Lista_med": "Lista_med",
  "Contrarec": "Contrarec",
  "Num_fac": "Num_fac",
  "Ind_med": "Ind_med",
  "Desc_med": "Desc_med",
  "Costoprom2": "Costoprom2",
  "Des_tial": "Marca",
  "Cto_ent": "Costo Ent.",
  "Nom_cte": "Cliente",
  "Nom_age": "Vendedor",
  "Nombre_ext001": "Nombre_ext001",
  "Nombre_ext002": "F1",
  "Nombre_ext003": "F2"
};

// Mapear columnas originales -> nombres comunes (NOTAS)
const RENAME_NOTAS = {
  "No_fac": "Nota",
  "Cve_suc": "Albaranes",
  "Falta_fac": "Fecha",
  "Lugar": "Almacen",
  "Status_fac": "Status_fac",
  "Status_c": "Status_c",
  "Cte_fac": "ID Cliente",
  "Nom_fac": "Cliente",
  "Rfc_cte": "Rfc_cte",
  "Descuento": "Descuento",
  "Cve_mon": "Cve_mon",
  "Factura": "Factura",
  "Cve_prod": "Clave",
  "New_med": "New_med",
  "Desc_prod": "Articulo",
  "Cant_surt": "Pz.",
  "Valor_prod": "Valor_prod",
  "Descu_prod": "Descuento2",
  "Part_nta": "Part_nta",
  "Subt_prod": "Sub. Total",
  "Iva_prod": "IVA",
  "Iva_ieps": "Iva_ieps",
  "Ieps_prod": "Ieps_prod",
  "Factor": "Factor",
  "Unidad": "Unidad",
  "Descue": "Descue",
  "Descue2": "Descue2",
  "Descue3": "Descue3",
  "Descue4": "Descue4",
  "Dcto1": "Descuento (%)",
  "Dcto2": "Dcto2",
  "Lote": "Lote",
  "Ref_lote": "Ref_lote",
  "Hravta": "Hora",
  "Foliosep": "Foliosep",
  "Separado": "Separado",
  "Cse_prod": "ID Clase",
  "Des_cse": "Clase",
  "Sub_cse": "Sub_cse",
  "Des_sub": "Des_sub",
  "Sub_subcse": "Sub_subcse",
  "Tip_cam": "Tip_cam",
  "Retisrvtad": "Retisrvtad",
  "Retiva_pro": "Retiva_pro",
  "Impto1": "Impto1",
  "Impto2": "Impto2",
  "Timpto1": "Timpto1",
  "Timpto2": "Timpto2",
  "Otrosimptos": "Otrosimptos",
  "Dessubsub": "Dessubsub",
  "Cve_factu": "Cve_factu",
  "Dato_1": "Dato_1",
  "Dato_2": "Dato_2",
  "Dato_3": "Dato_3",
  "Dato_4": "Dato_4",
  "Desc_med": "Desc_med",
  "Total_fac": "Total Nota",
  "Nom_age": "Vendedor",
  "Des_tial": "Marca",
  "Costoprom2": "Costoprom2",
  "Cto_ent": "Costo Entrada"
};

// Campos disponibles para la pestaña Detalle
const DISPONIBLES_DETALLE = [
  "Año", "Fecha", "Hora", "Almacen", "Factura/Nota", "Cliente", "ID Cliente",
  "Categoria", "Clave", "Tipo factura", "Subtotal", "Costo", "Descuento ($)",
  "Descuento (%)", "Utilidad", "Margen %", "Marca", "Vendedor"
];

const DEFAULT_DETALLE_COLS = [
  "Año", "Almacen", "Categoria", "Subtotal", "Costo", "Margen %", "Utilidad"
];

// -------------------------------------------------------
// Normalización para mapear columnas
// -------------------------------------------------------
function normalizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(/_/g, "");
}
function buildNormalizedMap(mapping) {
  const out = {};
  for (const k in mapping) {
    out[normalizeKey(k)] = mapping[k];
  }
  return out;
}
const NORM_FACT = buildNormalizedMap(RENAME_FACTURAS);
const NORM_NOTAS = buildNormalizedMap(RENAME_NOTAS);

// =======================================================
//  ESTADO GLOBAL Y DOM
// =======================================================

let records = [];
let yearsDisponibles = [];
let categoriasDisponibles = [];
let almacenesDisponibles = [];

let charts = { mensual: null, almacen: null };

let detalleState = {
  columnasSeleccionadas: [...DEFAULT_DETALLE_COLS],
  filtrosColumna: {},
  filtroGlobal: ""
};
let detalleSort = { col: null, asc: true };

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

const tablaTopClientes = document.querySelector("#tabla-top-clientes");
const tablaTopVendedores = document.querySelector("#tabla-top-vendedores");

const thYearPrev = document.getElementById("th-year-prev");
const thYearCurrent = document.getElementById("th-year-current");
const tablaYoYBody = document.getElementById("tabla-yoy");

const chipsContainer = document.getElementById("chips-container");
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

// =======================================================
//  UTILIDADES
// =======================================================

function toNumber(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return value;
  const str = value.toString().replace(/\s/g, "").replace(/,/g, "");
  const n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}

function formatCurrency(value) {
  return value.toLocaleString("es-MX", {
    style: "currency",
    currency: "MXN",
    maximumFractionDigits: 0
  });
}

function formatPercent(dec) {
  if (!isFinite(dec)) return "0.0%";
  return (dec * 100).toFixed(1) + "%";
}

function parseFecha(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + value * 86400000);
  }
  const str = value.toString().trim();
  const d1 = new Date(str);
  if (!isNaN(d1.getTime())) return d1;

  const parts = str.split(/[\/\-]/);
  if (parts.length === 3) {
    const [p1, p2, p3] = parts.map(p => parseInt(p, 10));
    if (p1 > 1900) return new Date(p1, p2 - 1, p3);
    if (p3 > 1900) return new Date(p3, p2 - 1, p1);
  }
  return null;
}

function sumField(arr, field) {
  return arr.reduce((acc, r) => acc + (r[field] || 0), 0);
}

function unique(arr) {
  return Array.from(new Set(arr));
}

function renameRow(row, normalizedMap) {
  const nuevo = {};
  for (const key in row) {
    const dest = normalizedMap[normalizeKey(key)];
    if (dest) nuevo[dest] = row[key];
  }
  return nuevo;
}

function debounce(fn, delay) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), delay);
  };
}

// =======================================================
//  LECTURA DEL ARCHIVO
// =======================================================

fileInput.addEventListener("change", handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  fileNameSpan.textContent = file.name;
  errorDiv.textContent = "";
  records = [];
  yearsDisponibles = [];
  categoriasDisponibles = [];
  almacenesDisponibles = [];

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });

      const nombreFacturas =
        wb.SheetNames.find(n => n.toLowerCase().includes("factura")) ||
        wb.SheetNames[0];
      const nombreNotas =
        wb.SheetNames.find(n => n.toLowerCase().includes("nota")) ||
        wb.SheetNames[1];

      const sheetFacturas = wb.Sheets[nombreFacturas];
      const sheetNotas = wb.Sheets[nombreNotas];

      if (!sheetFacturas || !sheetNotas) {
        throw new Error(
          "No se encontraron las hojas 'facturas' y 'notas' en el archivo."
        );
      }

      // Detectar encabezado de columna BB para "Factura del dia / Rango / Base"
      const headerRowsFact = XLSX.utils.sheet_to_json(sheetFacturas, {
        header: 1,
        range: 0,
        raw: true
      });
      const headerFact = headerRowsFact[0] || [];
      const idxBB = XLSX.utils.decode_col("BB");
      const filtroColKey = headerFact[idxBB];

      const rawFacturas = XLSX.utils.sheet_to_json(sheetFacturas, {
        defval: null
      });
      const rawNotas = XLSX.utils.sheet_to_json(sheetNotas, {
        defval: null
      });

      const facturasRen = rawFacturas.map(row => renameRow(row, NORM_FACT));
      const notasRen = rawNotas.map(row => renameRow(row, NORM_NOTAS));

      // ---------------- FACTURAS ----------------
      for (let i = 0; i < rawFacturas.length; i++) {
        const rowOrig = rawFacturas[i];
        const row = facturasRen[i];

        const flagTexto = filtroColKey
          ? (rowOrig[filtroColKey] ?? "").toString().trim().toLowerCase()
          : "";

        // excluir "Factura del dia"
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

        // Para utilidad/margen excluir RANGO y BASE
        const incluirUtilidad =
          !(flagTexto === "rango" || flagTexto === "base");

        // Crédito = descuento $ 0 y que no sea "base"
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

        records.push({
          origen: "factura",
          anio,
          mes,
          fecha,
          almacen,
          categoria,
          cliente,
          clienteId: (row["ID Cliente"] || "").toString().trim(),
          vendedor,
          folio,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura,
          marca: (row["Marca"] || "").toString().trim()
        });
      }

      // ---------------- NOTAS (solo REM) ----------------
      for (let i = 0; i < rawNotas.length; i++) {
        const row = notasRen[i];
        const tipoAlbaran = (row["Albaranes"] || "")
          .toString()
          .trim()
          .toLowerCase();
        if (tipoAlbaran !== "rem") continue;

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

        records.push({
          origen: "nota",
          anio,
          mes,
          fecha,
          almacen,
          categoria,
          cliente,
          clienteId: (row["ID Cliente"] || "").toString().trim(),
          vendedor,
          folio,
          subtotal,
          costo,
          utilidad,
          incluirUtilidad,
          descuentoMonto,
          descuentoPct,
          esCredito,
          tipoFactura,
          marca: (row["Marca"] || "").toString().trim()
        });
      }

      if (!records.length) {
        throw new Error(
          "No se generaron registros válidos. Revisa que el archivo tenga datos."
        );
      }

      yearsDisponibles = unique(records.map(r => r.anio)).sort((a, b) => a - b);
      categoriasDisponibles = unique(
        records.map(r => r.categoria || "").filter(x => x)
      );
      almacenesDisponibles = unique(
        records.map(r => r.almacen || "").filter(x => x)
      );

      poblarFiltros();
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

// =======================================================
//  FILTROS Y ACTUALIZACIÓN
// =======================================================

function poblarFiltros() {
  // Año
  filterYear.innerHTML = "<option value='all'>Todos los años</option>";
  yearsDisponibles.forEach(y => {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    filterYear.appendChild(opt);
  });

  // Almacén
  filterStore.innerHTML = "<option value='all'>Todos los almacenes</option>";
  almacenesDisponibles.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    filterStore.appendChild(opt);
  });

  // Tipo
  filterType.value = "both";

  // Categoría
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

function filtrarRecords(ignorarYearEnYoY) {
  const f = getFiltros();
  return records.filter(r => {
    if (!ignorarYearEnYoY && f.year && r.anio !== f.year) return false;
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

function actualizarTodo() {
  if (!records.length) return;
  const filtrados = filtrarRecords(false);
  actualizarKpis(filtridos);
}

