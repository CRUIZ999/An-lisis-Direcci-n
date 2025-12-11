// =======================================================
//  CONFIGURACIÓN Y MAPEOS
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

// Mapear columnas originales -> nombre común (FACTURAS)
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

// Mapear columnas originales -> nombre común (NOTAS)
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

// Campos que puedes usar en la pestaña Detalle
const DISPONIBLES_DETALLE = [
  "Año","Fecha","Hora","Almacen","Factura/Nota","Cliente","ID Cliente","Categoria",
  "Producto","Clave","Tipo factura","Subtotal","Costo","Descuento ($)","Descuento (%)",
  "Utilidad","Margen %","Marca","Vendedor"
];

const DEFAULT_DETALLE_COLS = [
  "Año","Almacen","Categoria","Subtotal","Costo","Margen %","Utilidad"
];

// Normalizar nombres de columnas para mapear aunque traigan espacios/rutas raras
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
//  ESTADO GLOBAL
// =======================================================
let records = [];               // dataset unificado
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

const tablaTopClientes = document.querySelector("#tabla-top-clientes tbody");
const tablaTopVendedores = document.querySelector("#tabla-top-vendedores tbody");

const thYearPrev = document.getElementById("th-year-prev");
const thYearCurrent = document.getElementById("th-year-current");
const tablaYoYBody = document.querySelector("#tabla-yoy tbody");

const chipsContainer = document.getElementById("chips-container");
const detalleHeaderRow = document.getElementById("detalle-header-row");
const detalleFilterRow = document.getElementById("detalle-filter-row");
const detalleTableBody = document.querySelector("#tabla-detalle tbody");
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
    const [p1,p2,p3] = parts.map(p=>parseInt(p,10));
    if (p1>1900) return new Date(p1, p2-1, p3);
    if (p3>1900) return new Date(p3, p2-1, p1);
  }
  return null;
}
function sumField(arr, field) {
  return arr.reduce((acc,r)=>acc+(r[field]||0),0);
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
    t = setTimeout(()=>fn(...args), delay);
  };
}

// =======================================================
//  LECTURA DE ARCHIVO
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
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });

      const nombreFacturas = wb.SheetNames.find(n=>n.toLowerCase().includes("factura")) || wb.SheetNames[0];
      const nombreNotas = wb.SheetNames.find(n=>n.toLowerCase().includes("nota")) || wb.SheetNames[1];

      const sheetFacturas = wb.Sheets[nombreFacturas];
      const sheetNotas = wb.Sheets[nombreNotas];

      if (!sheetFacturas || !sheetNotas) {
        throw new Error("No se encontraron las hojas 'facturas' y 'notas'.");
      }

      // leer encabezados para identificar la columna BB (Factura del día / Rango / Base)
      const headerRowsFact = XLSX.utils.sheet_to_json(sheetFacturas, { header: 1, range: 0, raw: true });
      const headerFact = headerRowsFact[0] || [];
      const idxBB = XLSX.utils.decode_col("BB");
      const filtroColKey = headerFact[idxBB];

      const rawFacturas = XLSX.utils.sheet_to_json(sheetFacturas, { defval: null });
      const rawNotas = XLSX.utils.sheet_to_json(sheetNotas, { defval: null });

      const facturasRen = rawFacturas.map(row => renameRow(row, NORM_FACT));
      const notasRen = rawNotas.map(row => renameRow(row, NORM_NOTAS));

      // ---------------- FACTURAS ----------------
      for (let i=0;i<rawFacturas.length;i++) {
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
        const mes = fecha.getMonth()+1;

        const subtotal = toNumber(row["Sub. Factura"]);
        const costo = toNumber(row["Costo Ent."]);
        const descuentoMonto = toNumber(row["Descuento ($)"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;
        const incluirUtilidad = !(flagTexto === "rango" || flagTexto === "base");

        let esCredito = false;
        if (descuentoMonto === 0 && flagTexto !== "base") esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (row["Almacen"] || row["Almacen2"] || "").toString().trim().toUpperCase();
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
      for (let i=0;i<rawNotas.length;i++) {
        const row = notasRen[i];
        const tipoAlbaran = (row["Albaranes"] || "").toString().trim().toLowerCase();
        if (tipoAlbaran !== "rem") continue;

        const fecha = parseFecha(row["Fecha"]);
        if (!fecha) continue;
        const anio = fecha.getFullYear();
        const mes = fecha.getMonth()+1;

        const subtotal = toNumber(row["Sub. Total"]);
        const costo = toNumber(row["Costo Entrada"]);
        const descuentoMonto = toNumber(row["Descuento"]);
        const descuentoPct = toNumber(row["Descuento (%)"]);
        const utilidad = subtotal - costo;
        const incluirUtilidad = true;

        let esCredito = false;
        if (descuentoMonto === 0) esCredito = true;
        const tipoFactura = esCredito ? "credito" : "contado";

        const almacen = (row["Almacen"] || "").toString().trim().toUpperCase();
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
        throw new Error("No se generaron registros válidos. Revisa las columnas y hojas.");
      }

      yearsDisponibles = unique(records.map(r=>r.anio)).sort((a,b)=>a-b);
      categoriasDisponibles = unique(records.map(r=>r.categoria || "")).filter(x=>x);
      almacenesDisponibles = unique(records.map(r=>r.almacen || "")).filter(x=>x);

      poblarFiltros();
      actualizarTodo();
    } catch (err) {
      console.error(err);
      errorDiv.textContent = "Error al procesar el archivo: " + err.message;
    }
  };
  reader.onerror = ()=> errorDiv.textContent = "No se pudo leer el archivo.";
  reader.readAsArrayBuffer(file);
}

// =======================================================
//  FILTROS Y ACTUALIZACIÓN
// =======================================================
function poblarFiltros() {
  filterYear.innerHTML = "<option value='all'>Todos los años</option>";
  yearsDisponibles.forEach(y=>{
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    filterYear.appendChild(opt);
  });

  filterStore.innerHTML = "<option value='all'>Todos los almacenes</option>";
  almacenesDisponibles.forEach(a=>{
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    filterStore.appendChild(opt);
  });

  filterType.value = "both";

  filterCategory.innerHTML = "<option value='all'>Todas las categorías</option>";
  categoriasDisponibles.forEach(c=>{
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
    year: filterYear.value === "all" ? null : parseInt(filterYear.value,10),
    store: filterStore.value === "all" ? null : filterStore.value,
    tipo: filterType.value,
    categoria: filterCategory.value === "all" ? null : filterCategory.value
  };
}

function filtrarRecords(ignorarYearEnYoY=false) {
  const f = getFiltros();
  return records.filter(r=>{
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
  actualizarKpis(filtrados);
  actualizarGraficas(filtrados);
  actualizarTop(filtrados);
  actualizarYoY();
  renderDetalle();
}

// =======================================================
//  KPIs
// =======================================================
function actualizarKpis(data) {
  const ventas = sumField(data,"subtotal");
  const utilData = data.filter(r=>r.incluirUtilidad);
  const utilidad = sumField(utilData,"utilidad");
  const margen = ventas>0 ? utilidad/ventas : 0;

  const filtros = getFiltros();
  const yearTxt = filtros.year || "todos los años";
  const almTxt = filtros.store || "todas las sucursales";

  kpiVentas.textContent = formatCurrency(ventas);
  kpiVentasSub.textContent = `Periodo: ${yearTxt}, almacén: ${almTxt}`;
  kpiUtilidad.textContent = formatCurrency(utilidad);
  kpiMargen.textContent = formatPercent(margen);
  kpiMargenSub.textContent = `Basado en ${utilData.length} filas válidas`;

  let m2 = 0;
  if (filtros.store && M2_POR_ALMACEN[filtros.store]) m2 = M2_POR_ALMACEN[filtros.store];
  else m2 = M2_POR_ALMACEN["Todas"];

  const ventasPorM2 = m2>0 ? ventas/m2 : 0;
  kpiM2.textContent = formatCurrency(ventasPorM2) + "/m²";
  kpiM2Sub.textContent = `${filtros.store || "Todas"} – ${m2.toLocaleString("es-MX")} m²`;

  const setOps = new Set(data.map(r=>r.origen + "|" + r.folio));
  kpiTrans.textContent = setOps.size.toLocaleString("es-MX");

  const setCli = new Set(data.map(r=>r.cliente || "").filter(x=>x));
  kpiClientes.textContent = setCli.size.toLocaleString("es-MX");
}

// =======================================================
//  GRÁFICAS
// =======================================================
function actualizarGraficas(data) {
  const ctxMensual = document.getElementById("chart-mensual").getContext("2d");
  const ctxAlmacen = document.getElementById("chart-almacen").getContext("2d");

  const filtros = getFiltros();
  let datosMensual = data;
  if (!filtros.year) {
    const maxYear = Math.max(...yearsDisponibles);
    datosMensual = data.filter(r=>r.anio===maxYear);
  }
  const ventasMes = new Array(12).fill(0);
  datosMensual.forEach(r=>{
    const idx = r.mes-1;
    if (idx>=0 && idx<12) ventasMes[idx]+=r.subtotal;
  });
  const labelsMes = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

  if (charts.mensual) charts.mensual.destroy();
  charts.mensual = new Chart(ctxMensual,{
    type:"bar",
    data:{
      labels:labelsMes,
      datasets:[{ label:"Ventas", data:ventasMes, borderWidth:1 }]
    },
    options:{
      responsive:true,
      scales:{ y:{ ticks:{ callback:v=>v.toLocaleString("es-MX") } } }
    }
  });

  const porAlm = {};
  data.forEach(r=>{
    const a = r.almacen || "SIN ALMACEN";
    if (!porAlm[a]) porAlm[a]={ventas:0,util:0};
    porAlm[a].ventas += r.subtotal;
    porAlm[a].util += r.utilidad;
  });
  const labelsAlm = Object.keys(porAlm);
  const datosVentas = labelsAlm.map(a=>porAlm[a].ventas);
  const datosVentasM2 = labelsAlm.map(a=>{
    const m2 = M2_POR_ALMACEN[a] || 0;
    return m2>0 ? porAlm[a].ventas/m2 : 0;
  });

  if (charts.almacen) charts.almacen.destroy();
  charts.almacen = new Chart(ctxAlmacen,{
    type:"bar",
    data:{
      labels:labelsAlm,
      datasets:[
        { label:"Ventas", data:datosVentas, borderWidth:1 },
        { label:"Ventas por m²", data:datosVentasM2, borderWidth:1 }
      ]
    },
    options:{
      responsive:true,
      scales:{ y:{ ticks:{ callback:v=>v.toLocaleString("es-MX") } } }
    }
  });
}

// =======================================================
//  TOP CLIENTES Y VENDEDORES
// =======================================================
function actualizarTop(data) {
  const porCliente = {};
  data.forEach(r=>{
    const cli = r.cliente || "(Sin cliente)";
    if (!porCliente[cli]) porCliente[cli]={ventas:0,util:0,margenAcum:0,count:0};
    porCliente[cli].ventas += r.subtotal;
    if (r.incluirUtilidad) {
      porCliente[cli].util += r.utilidad;
      porCliente[cli].margenAcum += (r.subtotal>0 ? r.utilidad/r.subtotal : 0);
      porCliente[cli].count++;
    }
  });
  const listaC = Object.entries(porCliente).map(([nombre,d])=>({
    nombre,
    ventas:d.ventas,
    utilidad:d.util,
    margen:d.count? d.margenAcum/d.count : 0
  })).sort((a,b)=>b.ventas-a.ventas).slice(0,5);

  tablaTopClientes.innerHTML="";
  listaC.forEach((row,idx)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${idx+1}</td>
      <td>${row.nombre}</td>
      <td class="text-right">${formatCurrency(row.ventas)}</td>
      <td class="text-right">${formatCurrency(row.utilidad)}</td>
      <td class="text-right">${formatPercent(row.margen)}</td>
    `;
    tablaTopClientes.appendChild(tr);
  });

  const porVend = {};
  const opsPorV = {};
  data.forEach(r=>{
    const v = r.vendedor || "(Sin vendedor)";
    if (!porVend[v]) porVend[v]={ventas:0,util:0};
    porVend[v].ventas += r.subtotal;
    if (r.incluirUtilidad) porVend[v].util += r.utilidad;

    const opKey = r.origen + "|" + r.folio;
    if (!opsPorV[v]) opsPorV[v] = new Set();
    opsPorV[v].add(opKey);
  });
  const listaV = Object.entries(porVend).map(([nombre,d])=>({
    nombre,
    ventas:d.ventas,
    utilidad:d.util,
    ops:(opsPorV[nombre]||new Set()).size
  })).sort((a,b)=>b.ventas-a.ventas).slice(0,5);

  tablaTopVendedores.innerHTML="";
  listaV.forEach((row,idx)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${idx+1}</td>
      <td>${row.nombre}</td>
      <td class="text-right">${formatCurrency(row.ventas)}</td>
      <td class="text-right">${formatCurrency(row.utilidad)}</td>
      <td class="text-right">${row.ops.toLocaleString("es-MX")}</td>
    `;
    tablaTopVendedores.appendChild(tr);
  });
}

// =======================================================
//  COMPARATIVO YOY
// =======================================================
function actualizarYoY() {
  const datos = filtrarRecords(true);
  if (!datos.length) {
    tablaYoYBody.innerHTML="<tr><td colspan='4' class='muted'>No hay datos para el comparativo.</td></tr>";
    return;
  }
  const years = unique(datos.map(r=>r.anio)).sort((a,b)=>a-b);
  if (years.length<2) {
    tablaYoYBody.innerHTML="<tr><td colspan='4' class='muted'>Se requiere al menos 2 años.</td></tr>";
    return;
  }
  const yCur = years[years.length-1];
  const yPrev = years[years.length-2];

  thYearPrev.textContent = yPrev;
  thYearCurrent.textContent = yCur;

  const dataPrev = datos.filter(r=>r.anio===yPrev);
  const dataCur = datos.filter(r=>r.anio===yCur);

  const ventasPrev = sumField(dataPrev,"subtotal");
  const ventasCur = sumField(dataCur,"subtotal");

  const utilPrev = sumField(dataPrev.filter(r=>r.incluirUtilidad),"utilidad");
  const utilCur = sumField(dataCur.filter(r=>r.incluirUtilidad),"utilidad");

  const margenPrev = ventasPrev>0? utilPrev/ventasPrev : 0;
  const margenCur = ventasCur>0? utilCur/ventasCur : 0;

  const credPrev = sumField(dataPrev.filter(r=>r.esCredito),"subtotal");
  const credCur = sumField(dataCur.filter(r=>r.esCredito),"subtotal");
  const pctCredPrev = ventasPrev>0? credPrev/ventasPrev : 0;
  const pctCredCur = ventasCur>0? credCur/ventasCur : 0;

  const utilNegPrevArr = dataPrev.filter(r=>r.incluirUtilidad && r.subtotal>0);
  const utilNegCurArr = dataCur.filter(r=>r.incluirUtilidad && r.subtotal>0);
  const pctNegPrev = utilNegPrevArr.length ? utilNegPrevArr.filter(r=>r.utilidad<0).length/utilNegPrevArr.length : 0;
  const pctNegCur = utilNegCurArr.length ? utilNegCurArr.filter(r=>r.utilidad<0).length/utilNegCurArr.length : 0;

  const filtros = getFiltros();
  let m2Valor = filtros.store && M2_POR_ALMACEN[filtros.store]
    ? M2_POR_ALMACEN[filtros.store]
    : M2_POR_ALMACEN["Todas"];

  const filas = [
    { id:"ventas", nombre:"Ventas (Subtotal)", prev:ventasPrev, cur:ventasCur, formato:"money" },
    { id:"utilidad", nombre:"Utilidad Bruta", prev:utilPrev, cur:utilCur, formato:"money" },
    { id:"margen", nombre:"Margen Bruto %", prev:margenPrev, cur:margenCur, formato:"percent" },
    { id:"pct_credito", nombre:"% Ventas a Crédito", prev:pctCredPrev, cur:pctCredCur, formato:"percent" },
    { id:"pct_util_neg", nombre:"% Ventas con Utilidad Negativa", prev:pctNegPrev, cur:pctNegCur, formato:"percent" },
    { id:"m2", nombre:"m² disponibles", prev:m2Valor, cur:m2Valor, formato:"plain" }
  ];

  tablaYoYBody.innerHTML="";
  filas.forEach(f=>{
    const crec = (f.prev===0 && f.cur>0) ? 1 :
                 (f.prev===0 && f.cur===0) ? 0 :
                 (f.cur - f.prev)/(f.prev || 1);
    const crecStr = (crec*100).toFixed(1) + "%";

    function fmt(v) {
      if (f.formato==="money") return formatCurrency(v);
      if (f.formato==="percent") return formatPercent(v);
      if (f.formato==="plain") return v.toLocaleString("es-MX");
      return v.toString();
    }

    const cls = crec>0 ? "pill pos" : (crec<0 ? "pill neg" : "pill neu");
    const icon = crec>0 ? "▲" : (crec<0 ? "▼" : "●");

    const tr = document.createElement("tr");
    tr.dataset.metricId = f.id;
    tr.innerHTML = `
      <td>${f.nombre}</td>
      <td class="text-right">${fmt(f.prev)}</td>
      <td class="text-right">${fmt(f.cur)}</td>
      <td class="text-right"><span class="${cls}">${icon} ${crecStr}</span></td>
    `;
    tablaYoYBody.appendChild(tr);
  });

  Array.from(tablaYoYBody.querySelectorAll("tr")).forEach(tr=>{
    tr.addEventListener("click",()=> abrirDetalleMetrica(tr.dataset.metricId,yPrev,yCur));
  });
}

// =======================================================
//  MODAL DE DETALLE POR MÉTRICA
// =======================================================
function abrirDetalleMetrica(metricId, yPrev, yCur) {
  const datos = filtrarRecords(true);
  const dataPrev = datos.filter(r=>r.anio===yPrev);
  const dataCur = datos.filter(r=>r.anio===yCur);

  modalYearPrev.textContent = yPrev;
  modalYearCurrent.textContent = yCur;

  let titulo = "";
  let calc = null;
  let formato = "money";

  if (metricId==="ventas") {
    titulo = "Detalle de Ventas (Subtotal) por categoría";
    calc = arr => sumField(arr,"subtotal");
  } else if (metricId==="utilidad") {
    titulo = "Detalle de Utilidad Bruta por categoría";
    calc = arr => sumField(arr.filter(r=>r.incluirUtilidad),"utilidad");
  } else if (metricId==="margen") {
    titulo = "Detalle de Margen Bruto % por categoría";
    calc = arr => {
      const v=sumField(arr,"subtotal");
      const u=sumField(arr.filter(r=>r.incluirUtilidad),"utilidad");
      return v>0? u/v : 0;
    };
    formato = "percent";
  } else if (metricId==="pct_credito") {
    titulo = "Detalle de % Ventas a Crédito por categoría";
    calc = arr => {
      const vTot=sumField(arr,"subtotal");
      const vCred=sumField(arr.filter(r=>r.esCredito),"subtotal");
      return vTot>0? vCred/vTot : 0;
    };
    formato = "percent";
  } else if (metricId==="pct_util_neg") {
    titulo = "Detalle de % Ventas con Utilidad Negativa por categoría";
    calc = arr => {
      const utilArr = arr.filter(r=>r.incluirUtilidad && r.subtotal>0);
      if (!utilArr.length) return 0;
      const neg=utilArr.filter(r=>r.utilidad<0).length;
      return neg/utilArr.length;
    };
    formato = "percent";
  } else if (metricId==="m2") {
    titulo = "Detalle de m² por almacén";
    const filtros=getFiltros();
    modalSub.textContent = filtros.store ? `Almacén ${filtros.store}` : "Todas las sucursales";
    modalTitle.textContent = titulo;
    modalTableBody.innerHTML="";

    const filas=[];
    if (filtros.store && M2_POR_ALMACEN[filtros.store]) {
      filas.push({cat:filtros.store, prev:M2_POR_ALMACEN[filtros.store], cur:M2_POR_ALMACEN[filtros.store]});
    } else {
      Object.keys(M2_POR_ALMACEN).forEach(k=>{
        if (k==="Todas") return;
        filas.push({cat:k, prev:M2_POR_ALMACEN[k], cur:M2_POR_ALMACEN[k]});
      });
    }
    filas.forEach(f=>{
      const tr=document.createElement("tr");
      tr.innerHTML=`
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

  modalTitle.textContent = titulo;
  const filtros = getFiltros();
  modalSub.textContent = `Filtros: almacén=${filtros.store||"Todos"}, tipo=${filtros.tipo}, categoría=${filtros.categoria||"Todas"}`;

  const cats = unique(datos.map(r=>r.categoria || "(Sin categoría)"));
  const filas = cats.map(cat=>{
    const arrPrev=dataPrev.filter(r=>(r.categoria || "(Sin categoría)")===cat);
    const arrCur=dataCur.filter(r=>(r.categoria || "(Sin categoría)")===cat);
    return {cat, prev:calc(arrPrev), cur:calc(arrCur)};
  }).sort((a,b)=>b.cur-a.cur);

  modalTableBody.innerHTML="";
  filas.forEach(f=>{
    const crec=(f.prev===0 && f.cur>0)?1:
               (f.prev===0 && f.cur===0)?0:
               (f.cur-f.prev)/(f.prev||1);
    const crecStr=(crec*100).toFixed(1)+"%";
    const cls = crec>0 ? "pill pos" : (crec<0 ? "pill neg" : "pill neu");
    const icon = crec>0 ? "▲" : (crec<0 ? "▼" : "●");

    function fmt(v){
      if (formato==="money") return formatCurrency(v);
      if (formato==="percent") return formatPercent(v);
      return v.toLocaleString("es-MX");
    }

    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${f.cat}</td>
      <td class="text-right">${fmt(f.prev)}</td>
      <td class="text-right">${fmt(f.cur)}</td>
      <td class="text-right"><span class="${cls}">${icon} ${crecStr}</span></td>
    `;
    modalTableBody.appendChild(tr);
  });

  modalBackdrop.classList.add("active");
}

modalClose.addEventListener("click",()=> modalBackdrop.classList.remove("active"));
modalBackdrop.addEventListener("click",e=>{
  if (e.target===modalBackdrop) modalBackdrop.classList.remove("active");
});

// =======================================================
//  DETALLE DE DATOS (TABLA PERSONALIZABLE)
// =======================================================
function initChips() {
  chipsContainer.innerHTML = "";
  DISPONIBLES_DETALLE.forEach(nombre=>{
    const chip=document.createElement("button");
    chip.type="button";
    chip.className="chip";
    chip.textContent=nombre;
    chip.dataset.colName=nombre;
    chip.addEventListener("click",()=> toggleColumnaDetalle(nombre));
    chipsContainer.appendChild(chip);
  });
}
function toggleColumnaDetalle(nombre) {
  const idx=detalleState.columnasSeleccionadas.indexOf(nombre);
  if (idx>=0) {
    if (detalleState.columnasSeleccionadas.length>1)
      detalleState.columnasSeleccionadas.splice(idx,1);
  } else {
    if (detalleState.columnasSeleccionadas.length>=7)
      detalleState.columnasSeleccionadas.shift();
    detalleState.columnasSeleccionadas.push(nombre);
  }
  renderDetalle();
}

searchGlobalInput.addEventListener("input", debounce(()=>{
  detalleState.filtroGlobal = searchGlobalInput.value.toLowerCase();
  renderDetalleBody();
}, 180));

function renderDetalle() {
  Array.from(chipsContainer.children).forEach(chip=>{
    const name=chip.dataset.colName;
    chip.classList.toggle("selected", detalleState.columnasSeleccionadas.includes(name));
  });

  detalleHeaderRow.innerHTML="";
  detalleFilterRow.innerHTML="";
  detalleState.filtrosColumna = {};

  detalleState.columnasSeleccionadas.forEach(col=>{
    const th=document.createElement("th");
    th.className="sortable";
    th.dataset.colName=col;
    th.innerHTML=`${col} <span>↕</span>`;
    th.addEventListener("click",()=> sortDetalleBy(col));
    detalleHeaderRow.appendChild(th);

    const thFilt=document.createElement("th");
    const inp=document.createElement("input");
    inp.className="filter-input";
    inp.placeholder="Filtrar...";
    inp.dataset.colName=col;
    inp.addEventListener("input", debounce(()=>{
      detalleState.filtrosColumna[col] = inp.value.toLowerCase();
      renderDetalleBody();
    },150));
    thFilt.appendChild(inp);
    detalleFilterRow.appendChild(thFilt);
  });

  renderDetalleBody();
}
function sortDetalleBy(col) {
  if (detalleSort.col===col) detalleSort.asc=!detalleSort.asc;
  else { detalleSort.col=col; detalleSort.asc=true; }
  renderDetalleBody();
}

function getCampoDetalle(rec,col) {
  switch(col){
    case "Año": return rec.anio;
    case "Fecha": return rec.fecha.toISOString().slice(0,10);
    case "Hora": return "";
    case "Almacen": return rec.almacen;
    case "Factura/Nota": return rec.folio;
    case "Cliente": return rec.cliente;
    case "ID Cliente": return rec.clienteId;
    case "Categoria": return rec.categoria;
    case "Producto": return "";
    case "Clave": return rec.clave || "";
    case "Tipo factura": return rec.tipoFactura==="credito" ? "Crédito" : "Contado";
    case "Subtotal": return rec.subtotal;
    case "Costo": return rec.costo;
    case "Descuento ($)": return rec.descuentoMonto;
    case "Descuento (%)": return rec.descuentoPct;
    case "Utilidad": return rec.utilidad;
    case "Margen %": return rec.subtotal>0 ? rec.utilidad/rec.subtotal : 0;
    case "Marca": return rec.marca || "";
    case "Vendedor": return rec.vendedor;
  }
  return "";
}

function renderDetalleBody() {
  const datos = filtrarRecords(false);
  let arr = datos.slice();

  arr = arr.filter(rec=>{
    for (const [col,val] of Object.entries(detalleState.filtrosColumna)) {
      if (!val) continue;
      const campo = getCampoDetalle(rec,col);
      const str = (campo===null || campo===undefined) ? "" :
        (typeof campo==="number" ? campo.toString() : campo.toString().toLowerCase());
      if (!str.includes(val)) return false;
    }
    return true;
  });

  if (detalleState.filtroGlobal) {
    const txt=detalleState.filtroGlobal;
    arr = arr.filter(rec=>{
      const campos=[rec.cliente,rec.vendedor,rec.folio,rec.almacen,rec.categoria];
      return campos.some(c=>c && c.toString().toLowerCase().includes(txt));
    });
  }

  if (detalleSort.col) {
    const col=detalleSort.col;
    const asc=detalleSort.asc;
    arr.sort((a,b)=>{
      const va=getCampoDetalle(a,col);
      const vb=getCampoDetalle(b,col);
      if (typeof va==="number" && typeof vb==="number")
        return asc ? va-vb : vb-va;
      const sa=(va||"").toString();
      const sb=(vb||"").toString();
      return asc ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }

  detalleTableBody.innerHTML="";
  arr.forEach(rec=>{
    const tr=document.createElement("tr");
    detalleState.columnasSeleccionadas.forEach(col=>{
      const td=document.createElement("td");
      let v=getCampoDetalle(rec,col);
      if (["Subtotal","Costo","Descuento ($)","Utilidad"].includes(col)) {
        td.className="text-right";
        td.textContent = formatCurrency(toNumber(v));
      } else if (col==="Margen %") {
        td.className="text-right";
        td.textContent = formatPercent(v);
      } else td.textContent = v ?? "";
      tr.appendChild(td);
    });
    detalleTableBody.appendChild(tr);
  });
}

// =======================================================
//  TABS & INIT
// =======================================================
Array.from(document.querySelectorAll(".tab-btn")).forEach(btn=>{
  btn.addEventListener("click",()=>{
    const tabId=btn.dataset.tab;
    document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
    document.getElementById(tabId).classList.add("active");
  });
});

initChips();
renderDetalle();
