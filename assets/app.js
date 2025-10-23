// --- Config ---
const LS_KEY_DATA = "pt_inv_dataset_v1";
const LS_KEY_OBS_PREFIX = "pt_obs:"; // obs por código
const DICC_URL = "assets/diccionario.csv"; // linea;codigo;producto (;)

// === KPI: Umbrales de DÍAS DE PISO (configurable) ===
// Definición típica en PT: rojo (<=3), naranja (<=7), amarillo (<=14), verde (>14)
// === KPI: Umbrales de DÍAS DE PISO ===
const KPI_DIAS = Object.freeze({
  rojo:     3,
  naranja:  7,
  amarillo: 14
});

const fmtNum = (x) => (x == null ? "–" : (Number.isInteger(x) ? x.toString() : x.toFixed(2)));

function chipDias(val) {
  if (val == null || !Number.isFinite(val)) return `<span class="kpi na">–</span>`;
  if (val <= KPI_DIAS.rojo)     return `<span class="kpi bad">${fmtNum(val)}</span>`;
  if (val <= KPI_DIAS.naranja)  return `<span class="kpi warn">${fmtNum(val)}</span>`;
  if (val <= KPI_DIAS.amarillo) return `<span class="kpi mid">${fmtNum(val)}</span>`;
  return `<span class="kpi ok">${fmtNum(val)}</span>`;
}


// --- Carga dinámica de XLSX (lazy) con fallbacks ---
async function ensureXLSX() {
  if (window.XLSX) return;
  const tryLoad = (src) =>
    new Promise((res, rej) => {
      const s = document.createElement("script");
      s.defer = true; s.src = src;
      s.onload = () => res(true);
      s.onerror = () => rej(new Error("fail " + src));
      document.head.appendChild(s);
    });
  try {
    await tryLoad("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js");
  } catch {
    try {
      await tryLoad("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.20.2/xlsx.full.min.js");
    } catch {
      // último recurso: archivo local (opcional si lo agregas en tu repo)
      await tryLoad("./assets/vendor/xlsx.full.min.js");
    }
  }
}

// Reemplaza COL_SYNONYMS por esto:
const COL_SYNONYMS = {
  codigo:  ["codigo","código","code","cod","clave","pro"], // ← añade "pro"
  invtot:  ["invtot","inv. total","inv total","inventario total","total"],
  invlle:  ["invlle","inv lleno","lleno"],
  invvac:  ["invvac","inv vacio","inv vacío","vacio","vacío"],
  // OJO: quitamos "pro" de venpro para no confundirla con el código
  venpro:  ["venpro","venta prom","venta promedio","ventas prom","ventas promedio","promedio"],
  diatotc: ["diatotc","dias piso","días piso","dias de piso","días de piso","cobertura","dias cobertura"],
};


const byAlias = (obj) => {
  const map = {};
  for (const [k, v] of Object.entries(obj)) map[String(k).trim().toLowerCase()] = v;
  return map;
};
function pick(row, key) {
  const r = byAlias(row);
  for (const alias of COL_SYNONYMS[key] || []) if (alias in r) return r[alias];
  return undefined;
}

function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function toNum(v) {
  if (v == null) return NaN;
  const s0 = String(v).trim();
  const s1 = s0.replace(/[\s,]/g, "");                  // quita separadores de miles (espacio/coma)
  const s2 = s1.replace(/(\d)\.(?=\d{3}(\D|$))/g, "$1"); // quita puntos de miles tipo 12.345
  const s  = s2.replace(",", ".");                      // por si quedó decimal con coma
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}

// --- Diccionario (linea;codigo;producto) ---
let DICC = { byCode: new Map(), byLinea: new Map() };

async function loadDiccionario() {
  return new Promise((resolve, reject) => {
    Papa.parse(DICC_URL, {
      download: true, delimiter: ",", header: true, encoding: "UTF-8",
      complete: (res) => {
        const rows = res.data.filter(r => r.linea && r.codigo && r.producto);
        DICC.byCode = new Map(); DICC.byLinea = new Map();
        for (const r of rows) {
          const linea = r.linea.trim();
          const codigo = String(r.codigo).trim();
          const producto = (r.producto || "").trim().replace(/^\s*\d{1,6}\s+/, "");
          DICC.byCode.set(codigo, { linea, producto });
          if (!DICC.byLinea.has(linea)) DICC.byLinea.set(linea, []);
          DICC.byLinea.get(linea).push({ codigo, producto });
        }
        for (const [L, arr] of DICC.byLinea.entries()) {
          arr.sort((a,b) => (parseInt(a.codigo)||0) - (parseInt(b.codigo)||0));
        }
        // badge
        const meta = document.getElementById("fileMeta");
        meta.textContent = `Dicc: ${DICC.byCode.size} entradas`;
        resolve();
      },
      error: (err) => {
        const meta = document.getElementById("fileMeta");
        meta.textContent = "⚠ No se pudo cargar el diccionario.csv";
        reject(err);
      }
    });
  });
}

// --- XLSX → objetos (detecta encabezados aunque haya títulos arriba) ---
async function parseXlsx(file) {
  await ensureXLSX(); // <- asegura la librería
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });
  const firstSheet = wb.SheetNames[0];
  const sheet = wb.Sheets[firstSheet];

  // Matriz cruda
  const rows2D = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const norm = (s) => String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/\p{Diacritic}/gu,"")
    .replace(/\s+/g," ").trim();

  const wants = {
    codigo:   ["codigo","código","code","cod","clave","pro"],    // ← "pro"
    producto: ["producto","descripcion","descripción","nombre","concepto"], // ← "concepto"
    invtot:   ["inv. total","inv total","invtot","inventario total","total"],
    venpro:   ["venta prom","venta promedio","venpro","promedio","ventas prom","ventas promedio"],
    diatotc:  ["dias de piso","días de piso","dias piso","cobertura","diatotc"]
  };

  let headerRowIdx = -1, headerMap = null;
  for (let i = 0; i < Math.min(rows2D.length, 100); i++) {
    const row = rows2D[i];
    const hasCod = row.some(c => wants.codigo.includes(norm(c)));
    const hasProd = row.some(c => wants.producto.includes(norm(c)));
    if (hasCod && hasProd) {
      headerRowIdx = i;
      headerMap = {};
      row.forEach((cell, idx) => {
        const n = norm(cell);
        for (const [k, arr] of Object.entries(wants)) {
          if (arr.includes(n) && headerMap[k] === undefined) headerMap[k] = idx;
        }
      });
      break;
    }
  }
  if (headerRowIdx === -1) {
    return { rows: [], sheetName: firstSheet, fileName: file.name };
  }

  const dataRows = [];
  for (let r = headerRowIdx + 1; r < rows2D.length; r++) {
    const row = rows2D[r];
    if (!row || row.every(c => String(c).trim() === "")) continue;

    const obj = {};
    if (headerMap.codigo   !== undefined) obj["codigo"]       = row[headerMap.codigo] ?? "";
    if (headerMap.producto !== undefined) obj["producto"]     = row[headerMap.producto] ?? "";
    if (headerMap.invtot   !== undefined) obj["inv. total"]   = row[headerMap.invtot] ?? "";
    if (headerMap.venpro   !== undefined) obj["venta prom"]   = row[headerMap.venpro] ?? "";
    if (headerMap.diatotc  !== undefined) obj["dias de piso"] = row[headerMap.diatotc] ?? ""; // ← simple y correcto
    // extra: si existe 'concepto', guárdalo crudo para derivar nombre
    if (headerMap.producto === undefined && headerMap.producto === undefined && headerMap["producto"] === undefined) {}
    if (headerMap["producto"] === undefined && headerMap["concepto"] !== undefined) {
      obj["concepto"] = row[headerMap["concepto"]] ?? "";
    } else if (headerMap["concepto"] !== undefined) {
      obj["concepto"] = row[headerMap["concepto"]] ?? "";
    }

    dataRows.push(obj);
  }
  return { rows: dataRows, sheetName: firstSheet, fileName: file.name };
}

// --- Post-proceso / badge ---
function afterParseAndRender(raw, norm) {
  const dataset = { meta: { fileName: raw.fileName, sheetName: raw.sheetName }, rows: norm };
  saveDataset(dataset.meta, dataset.rows);
  renderTables(dataset, document.getElementById("toggleCatalogo").checked);

  const meta = document.getElementById("fileMeta");
  const diccCount = DICC.byCode?.size ?? 0;
  meta.textContent = `Dicc: ${diccCount} entradas | Filas Excel: ${dataset.rows.length}`;

  if (!dataset.rows.length) {
    alert("No se detectaron filas de datos. Revisa que la hoja tenga una fila de encabezados con 'Código' y 'Producto'.");
  }
}

// --- Normalización / derivación ---
// --- Normalización / derivación: SOLO códigos presentes en el diccionario ---
function normalizeRows(rows) {
  const out = [];
  for (const row of rows) {
    // Código (obligatorio)
    const codigo_raw = pick(row, "codigo");
    if (!codigo_raw) continue;
    const codigo = String(codigo_raw).trim();

    // Si el código NO está en el diccionario, lo ignoramos
    if (!DICC.byCode.has(codigo)) {
      continue;
    }

    // Datos numéricos del Excel (opcionalmente usados)
    const invtot_raw  = pick(row, "invtot");
    const invlle_raw  = pick(row, "invlle");
    const invvac_raw  = pick(row, "invvac");
    const venpro_raw  = pick(row, "venpro");
    const diatotc_raw = pick(row, "diatotc");

    const invlle = toNum(invlle_raw);
    const invvac = toNum(invvac_raw);

    let invtot = toNum(invtot_raw);
    if (!Number.isFinite(invtot)) {
      const suma = (Number.isFinite(invlle) ? invlle : 0) + (Number.isFinite(invvac) ? invvac : 0);
      invtot = Number.isFinite(suma) ? suma : NaN;
    }

    let venprom = toNum(venpro_raw);
    if (!Number.isFinite(venprom)) venprom = NaN;

    let dias = toNum(diatotc_raw);
    if (!Number.isFinite(dias) && Number.isFinite(invtot) && Number.isFinite(venprom) && venprom > 0) {
      dias = invtot / venprom;
    }

    // Línea y nombre SIEMPRE del diccionario
    const { linea, producto } = DICC.byCode.get(codigo);

    out.push({
      linea,
      codigo,
      producto,
      inv_total: Number.isFinite(invtot) ? invtot : null,
      venta_prom: Number.isFinite(venprom) ? venprom : null,
      dias_piso: Number.isFinite(dias) ? dias : null
    });
  }
  return out;
}

// --- Persistencia ---
function saveDataset(meta, rows) {
  localStorage.setItem(LS_KEY_DATA, JSON.stringify({ meta, rows, savedAt: new Date().toISOString() }));
}
function loadDataset() {
  const raw = localStorage.getItem(LS_KEY_DATA);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

// --- Observaciones por código ---
const getObs = (code) => localStorage.getItem(LS_KEY_OBS_PREFIX + code) || "";
const setObs = (code, val) => localStorage.setItem(LS_KEY_OBS_PREFIX + code, val);

// --- Helpers de UI ---


function chipDias(val) {
  if (val == null || !Number.isFinite(val)) return `<span class="kpi na">–</span>`;
  if (val <= KPI_DIAS.rojo)     return `<span class="kpi bad">${fmtNum(val)}</span>`;     // rojo
  if (val <= KPI_DIAS.naranja)  return `<span class="kpi warn">${fmtNum(val)}</span>`;    // naranja
  if (val <= KPI_DIAS.amarillo) return `<span class="kpi mid">${fmtNum(val)}</span>`;     // amarillo
  return `<span class="kpi ok">${fmtNum(val)}</span>`;                                     // verde
}

const sum = (arr, k) => arr.reduce((a,b)=> a + (Number.isFinite(b[k]) ? b[k] : 0), 0);
const avg = (arr, k) => {
  const vals = arr.map(r => r[k]).filter(Number.isFinite);
  return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
};

// --- Render ---
function renderTables(dataset, mostrarCatalogoCompleto=false) {
  const host = document.getElementById("tables");
  host.innerHTML = "";

  const byLineaData = new Map();
  for (const r of dataset.rows) {
    if (!byLineaData.has(r.linea)) byLineaData.set(r.linea, []);
    byLineaData.get(r.linea).push(r);
  }
  for (const arr of byLineaData.values()) {
    arr.sort((a,b) => (parseInt(a.codigo)||0) - (parseInt(b.codigo)||0));
  }

  const meta = document.getElementById("fileMeta");
  meta.textContent = `Dicc: ${DICC.byCode.size} entradas | Filas Excel: ${dataset.rows.length}`;

  const lineas = Array.from(DICC.byLinea.keys());
  if (byLineaData.has("SIN LÍNEA")) lineas.push("SIN LÍNEA");

  for (const linea of lineas) {
    const presentes = byLineaData.get(linea) || [];

    const cat = DICC.byLinea.get(linea) || [];
    const presentCodes = new Set(presentes.map(r => r.codigo));
    const faltantes = cat
      .filter(x => !presentCodes.has(x.codigo))
      .map(x => ({ linea, codigo: x.codigo, producto: x.producto, inv_total:null, venta_prom:null, dias_piso:null }));

    const rows = mostrarCatalogoCompleto ? [...presentes, ...faltantes] : presentes;
    if (!rows.length) continue;

    const card = document.createElement("div");
    card.className = "card linea";

    const h = document.createElement("h2");
    const chip = `<span class="chip"></span>`;
    h.innerHTML = `${linea} ${chip}`;
    card.appendChild(h);

    const tbl = document.createElement("table");
    const theadTopOffset = 80;
    tbl.innerHTML = `
      <thead style="top:${theadTopOffset}px">
        <tr>
          <th>Código</th>
          <th>Producto</th>
          <th class="right">Inv. Total</th>
          <th class="right">Venta prom</th>
          <th class="right">Días de piso</th>
          <th>Observaciones</th>
        </tr>
      </thead>
      <tbody></tbody>
      <tfoot>
        <tr>
          <td></td>
          <td><strong>Totales / Promedios</strong></td>
          <td class="right" id="sumInv"></td>
          <td class="right" id="avgVen"></td>
          <td class="right" id="avgDias"></td>
          <td></td>
        </tr>
      </tfoot>
    `;
    const tbody = tbl.querySelector("tbody");

    for (const r of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.codigo}</td>
        <td>${r.producto}</td>
        <td class="right">${fmtNum(r.inv_total)}</td>
        <td class="right">${fmtNum(r.venta_prom)}</td>
        <td class="right">${chipDias(r.dias_piso)}</td>
        <td class="obs"><textarea data-code="${r.codigo}" rows="1" placeholder="Notas..."></textarea></td>
      `;
      tbody.appendChild(tr);
      const ta = tr.querySelector("textarea");
      // después de crear el <textarea>:
      ta.setAttribute("maxlength", "220");   // límite de caracteres
      ta.style.maxWidth = "520px";           // (opcional) límite visual en pantallas anchas

      ta.value = getObs(r.codigo);
      ta.addEventListener("input", (e) => setObs(r.codigo, e.target.value));
    }

    tbl.querySelector("#sumInv").textContent = fmtNum(sum(presentes, "inv_total"));
    tbl.querySelector("#avgVen").textContent = fmtNum(avg(presentes, "venta_prom"));
    tbl.querySelector("#avgDias").textContent = fmtNum(avg(presentes, "dias_piso"));

    card.appendChild(tbl);
    host.appendChild(card);
  }

  if (!host.children.length) {
    const p = document.createElement("div");
    p.className = "card";
    p.innerHTML = `<div class="empty">Carga un Excel para ver datos.</div>`;
    host.appendChild(p);
  }
}

// --- Boot ---
(async function () {
  await loadDiccionario();

  const file = document.getElementById("file");
  const fileMeta = document.getElementById("fileMeta");
  const clearData = document.getElementById("clearData");
  const clearObs = document.getElementById("clearObs");
  const toggleCatalogo = document.getElementById("toggleCatalogo");

  const prev = loadDataset();
  if (prev?.rows?.length) {
    renderTables(prev, toggleCatalogo.checked);
    fileMeta.textContent = `Cargando de memoria: ${prev.meta?.fileName || "dataset"} — hoja: ${prev.meta?.sheetName || "única"}`;
  }

  toggleCatalogo.addEventListener("change", () => {
    const d = loadDataset();
    if (d?.rows?.length) renderTables(d, toggleCatalogo.checked);
  });

  file.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const raw = await parseXlsx(f);
      const norm = normalizeRows(raw.rows);
      afterParseAndRender(raw, norm);
    } catch (err) {
      console.error(err);
      alert("No pude leer la hoja: " + err.message);
    }
  });

  clearData.addEventListener("click", () => {
    localStorage.removeItem(LS_KEY_DATA);
    document.getElementById("tables").innerHTML = "";
    fileMeta.textContent = "";
  });

  clearObs.addEventListener("click", () => {
    for (const code of DICC.byCode.keys()) localStorage.removeItem(LS_KEY_OBS_PREFIX + code);
    const d = loadDataset();
    if (d?.rows?.length) renderTables(d, toggleCatalogo.checked);
  });
})();
