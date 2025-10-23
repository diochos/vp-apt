// --- Config ---
const LS_KEY_DATA = "pt_inv_dataset_v1";
const LS_KEY_OBS_PREFIX = "pt_obs:"; // obs por código
const DICC_URL = "assets/diccionario.enc";

// === KPI: Umbrales de DÍAS DE PISO (tus nuevos valores) ===
const KPI_DIAS = Object.freeze({ rojo:1, naranja:3, amarillo:5 }); // verde: >5

// === Parámetros para cálculos adicionales (tus nuevos valores) ===
const CFG = Object.freeze({
  targetDOS: 5,     // cobertura objetivo (días)
  excesoDOS: 10     // exceso si pasa de este umbral
});

// --- Decode XOR + Base64 usando keyB1 (string) ---
function decodeB64XorWithKey(b64, keyStr) {
  if (!keyStr) {
    throw new Error("No hay clave para diccionario (B1).");
  }
  const clean = b64.replace(/\s+/g, "");
  const bin = atob(clean);
  const out = new Uint8Array(bin.length);
  const klen = keyStr.length;
  for (let i = 0; i < bin.length; i++) {
    out[i] = bin.charCodeAt(i) ^ keyStr.charCodeAt(i % klen);
  }
  return new TextDecoder().decode(out);
}


const fmtNum = (x) =>
  (x == null || Number.isNaN(x)) ? "–" :
  (Math.abs(x) >= 1000 || !Number.isInteger(x) ? Number(x).toFixed(2) : String(x));

const today = new Date();
const fmtDate = (d) => !d ? "–" :
  d.toLocaleDateString("es-MX", { day:"2-digit", month:"short", year:"numeric" }).replace(".", "");

function fechaQuiebre(dias) {
  if (!Number.isFinite(dias)) return null;
  const dd = new Date(today); dd.setDate(dd.getDate() + Math.ceil(dias));
  return dd;
}

function chipDias(val) {
  if (val == null || !Number.isFinite(val)) return `<span class="kpi na">–</span>`;
  if (val <= KPI_DIAS.rojo)     return `<span class="kpi bad">${fmtNum(val)}</span>`;
  if (val <= KPI_DIAS.naranja)  return `<span class="kpi warn">${fmtNum(val)}</span>`;
  if (val <= KPI_DIAS.amarillo) return `<span class="kpi mid">${fmtNum(val)}</span>`;
  return `<span class="kpi ok">${fmtNum(val)}</span>`;
}

// --- Carga dinámica de XLSX ---
async function ensureXLSX() {
  if (window.XLSX) return;
  const tryLoad = (src) => new Promise((res, rej) => {
    const s = document.createElement("script");
    s.defer = true; s.src = src;
    s.onload = () => res(true);
    s.onerror = () => rej(new Error("fail " + src));
    document.head.appendChild(s);
  });
  try {
    await tryLoad("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js");
  } catch {
    try { await tryLoad("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.20.2/xlsx.full.min.js"); }
    catch { await tryLoad("./assets/vendor/xlsx.full.min.js"); }
  }
}

// --- Aliases de encabezados ---
const COL_SYNONYMS = {
  codigo:  ["codigo","código","code","cod","clave","pro"],
  invtot:  ["invtot","inv. total","inv total","inventario total","total"],
  invlle:  ["invlle","inv lleno","lleno"],
  invvac:  ["invvac","inv vacio","inv vacío","vacio","vacío"],
  venpro:  ["venpro","venta prom","venta promedio","ventas prom","ventas promedio","promedio"],
  diatotc: ["diatotc","dias piso","días piso","dias de piso","días de piso","cobertura","dias cobertura"],
  producto:["producto","descripcion","descripción","nombre","concepto"]
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
function escapeRegex(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function toNum(v) {
  if (v == null) return NaN;
  const s0 = String(v).trim();
  const s1 = s0.replace(/[\s,]/g, "");
  const s2 = s1.replace(/(\d)\.(?=\d{3}(\D|$))/g, "$1");
  const s  = s2.replace(",", ".");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}

// --- Diccionario (linea,codigo,producto) ---
let DICC = { byCode: new Map(), byLinea: new Map() };

// Carga diccionario.enc y lo descifra con la clave proveniente del Excel (B1)
async function loadDiccionarioWithKey(keyB1) {
  const res = await fetch(DICC_URL, { cache: "no-cache" });
  if (!res.ok) throw new Error("No pude descargar diccionario.enc");
  const b64 = await res.text();
  const csvText = decodeB64XorWithKey(b64, keyB1);

  return new Promise((resolve, reject) => {
    Papa.parse(csvText, {
      header: true,
      complete: ({data}) => {
        const rows = data.filter(r => r.linea && r.codigo && r.producto);
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
        resolve();
      },
      error: reject
    });
  });
}



// --- XLSX → objetos ---
async function parseXlsx(file) {
  await ensureXLSX();
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });
  const firstSheet = wb.SheetNames[0];
  const sheet = wb.Sheets[firstSheet];

  // Lee toda la hoja a 2D para detección de encabezados y para buscar B1 flexible
  const rows2D = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  // --- detectar encabezados como ya lo hacías ---
  const norm = (s) => String(s || "")
    .toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"")
    .replace(/\s+/g," ").trim();

  const wants = {
    codigo:   ["codigo","código","code","cod","clave","pro"],
    producto: ["producto","descripcion","descripción","nombre","concepto"],
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

  // === CLAVE ROBUSTA ===
  // 1) Si la fila 0 tiene algo en Col B (index 1), úsalo (B1 real)
  // 2) Si detectamos encabezado en otra fila, toma Col B de ESA fila
  // 3) Si todavía no, busca la primera celda NO vacía en Col B dentro de las primeras 10 filas
  // 4) Fallback: si existe sheet["B1"].v úsalo
  let keyB1 = "";
  if (rows2D[0] && rows2D[0][1] !== undefined && String(rows2D[0][1]).trim() !== "") {
    keyB1 = String(rows2D[0][1]).trim();
  } else if (headerRowIdx >= 0 && rows2D[headerRowIdx] && rows2D[headerRowIdx][1] !== undefined && String(rows2D[headerRowIdx][1]).trim() !== "") {
    keyB1 = String(rows2D[headerRowIdx][1]).trim();
  } else {
    for (let r = 0; r < Math.min(rows2D.length, 10); r++) {
      if (rows2D[r] && rows2D[r][1] !== undefined && String(rows2D[r][1]).trim() !== "") {
        keyB1 = String(rows2D[r][1]).trim();
        break;
      }
    }
    if (!keyB1 && sheet["B1"] && sheet["B1"].v != null) {
      keyB1 = String(sheet["B1"].v).trim();
    }
  }

  if (!keyB1) {
    throw new Error("No se encontró clave en columna B (B1 o fila de encabezados).");
  }

  // --- si no hay encabezados, cortamos igual (pero devolvemos meta con keyB1) ---
  if (headerRowIdx === -1) {
    return { rows: [], sheetName: firstSheet, fileName: file.name, keyB1 };
  }

  // --- construir dataRows como ya lo hacías ---
  const dataRows = [];
  for (let r = headerRowIdx + 1; r < rows2D.length; r++) {
    const row = rows2D[r];
    if (!row || row.every(c => String(c).trim() === "")) continue;

    const obj = {};
    if (headerMap.codigo   !== undefined) obj["codigo"]       = row[headerMap.codigo] ?? "";
    if (headerMap.producto !== undefined) obj["producto"]     = row[headerMap.producto] ?? "";
    if (headerMap.invtot   !== undefined) obj["inv. total"]   = row[headerMap.invtot] ?? "";
    if (headerMap.venpro   !== undefined) obj["venta prom"]   = row[headerMap.venpro] ?? "";
    if (headerMap.diatotc  !== undefined) obj["dias de piso"] = row[headerMap.diatotc] ?? "";
    if (headerMap["concepto"] !== undefined) obj["concepto"]  = row[headerMap["concepto"]] ?? "";
    dataRows.push(obj);
  }

  return { rows: dataRows, sheetName: firstSheet, fileName: file.name, keyB1 };
}

// --- Normalización: SOLO códigos presentes en el diccionario ---
function normalizeRows(rows) {
  const out = [];
  for (const row of rows) {
    const codigo_raw = pick(row, "codigo");
    if (!codigo_raw) continue;
    const codigo = String(codigo_raw).trim();
    if (!DICC.byCode.has(codigo)) continue; // filtro duro

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

    const { linea, producto } = DICC.byCode.get(codigo);

    out.push({
      linea, codigo, producto,
      inv_total: Number.isFinite(invtot) ? invtot : null,
      venta_prom: Number.isFinite(venprom) ? venprom : null,
      dias_piso: Number.isFinite(dias) ? dias : null
    });
  }
  return out;
}

// --- Persistencia ---
function saveDataset(meta, rows) {
  // meta puede traer keyB1 si venía de parseXlsx
  localStorage.setItem(LS_KEY_DATA, JSON.stringify({ meta, rows, savedAt: new Date().toISOString() }));
}

function loadDataset() {
  const raw = localStorage.getItem(LS_KEY_DATA);
  if (!raw) return null; try { return JSON.parse(raw); } catch { return null; }
}

// --- Observaciones por código ---
const getObs = (code) => localStorage.getItem(LS_KEY_OBS_PREFIX + code) || "";
const setObs = (code, val) => localStorage.setItem(LS_KEY_OBS_PREFIX + code, val);

// --- Helpers KPIs / Acción Hoy ---
const sum = (arr, k) => arr.reduce((a,b)=> a + (Number.isFinite(b[k]) ? b[k] : 0), 0);
const avg = (arr, k) => {
  const vals = arr.map(r => r[k]).filter(Number.isFinite);
  return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : null;
};
function qtyReorden(inv, ven, targetDOS=CFG.targetDOS){
  if(!Number.isFinite(ven) || ven<=0) return 0;
  const need = targetDOS*ven - (inv||0);
  return Math.max(0, Math.round(need));
}
function kpiBuckets(d) {
  let rojo=0,naranja=0,amarillo=0,verde=0;
  for (const r of d) {
    const v = r.dias_piso;
    if (!Number.isFinite(v)) continue;
    if (v<=KPI_DIAS.rojo) rojo++;
    else if (v<=KPI_DIAS.naranja) naranja++;
    else if (v<=KPI_DIAS.amarillo) amarillo++;
    else verde++;
  }
  return {rojo,naranja,amarillo,verde};
}

// --- Render KPIs globales ---
function renderKPIs(rows) {
  const host = document.getElementById("kpis");
  if (!rows || !rows.length) { host.style.display = "none"; host.innerHTML = ""; return; }

  const activos = rows.filter(r => Number.isFinite(r.venta_prom) && r.venta_prom>0).length;
  const muertos = rows.filter(r => (!Number.isFinite(r.venta_prom) || r.venta_prom===0) && Number.isFinite(r.inv_total) && r.inv_total>0).length;
  const invSum = sum(rows, "inv_total");
  const venSum = sum(rows, "venta_prom");
  const cobertura = venSum>0 ? invSum/venSum : null;
  const {rojo,naranja,amarillo,verde} = kpiBuckets(rows);
  const exceso = rows.filter(r => Number.isFinite(r.dias_piso) && r.dias_piso>CFG.excesoDOS).length;

  host.style.display = "block";
  host.innerHTML = `
    <h2 style="margin-bottom:12px">KPIs globales</h2>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-title">SKUs activos</div><div class="kpi-value">${activos}</div><div class="kpi-sub">con venta &gt; 0</div></div>
      <div class="kpi-card"><div class="kpi-title">Muertos</div><div class="kpi-value">${muertos}</div><div class="kpi-sub">inv &gt; 0 y venta = 0</div></div>
      <div class="kpi-card"><div class="kpi-title">Cobertura promedio</div><div class="kpi-value">${fmtNum(cobertura)}</div><div class="kpi-sub">días ponderados</div></div>
      <div class="kpi-card"><div class="kpi-title">Críticos por color</div><div class="kpi-value">🔴 ${rojo} · 🟠 ${naranja} · 🟡 ${amarillo} · 🟢 ${verde}</div><div class="kpi-sub">umbrales 1/3/5/5+</div></div>
      <div class="kpi-card"><div class="kpi-title">Exceso</div><div class="kpi-value">${exceso}</div><div class="kpi-sub">&gt; ${CFG.excesoDOS} días</div></div>
    </div>
  `;
}

// --- Render principal (tablas por línea + Acción hoy) ---
function renderTables(dataset, mostrarCatalogoCompleto=false) {
  const host = document.getElementById("tables");
  host.innerHTML = "";

  // KPIs globales primero
  renderKPIs(dataset.rows);

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

    // Enriquecer con fecha de quiebre y sugerido
    const enriched = rows.map(r => {
      const dos = Number.isFinite(r.dias_piso) ? r.dias_piso :
                  (Number.isFinite(r.inv_total) && Number.isFinite(r.venta_prom) && r.venta_prom>0 ? r.inv_total/r.venta_prom : null);
      const fq = fechaQuiebre(dos);
      const sug = qtyReorden(r.inv_total, r.venta_prom, CFG.targetDOS);
      return { ...r, dias_calc: dos, fecha_quiebre: fq, sugerido: sug };
    });

    const card = document.createElement("div");
    card.className = "card linea";

    const h = document.createElement("h2");
    const chip = `<span class="chip">${enriched.length} productos</span>`;
    h.innerHTML = `${linea} ${chip}`;
    card.appendChild(h);

    // Leyenda de colores
    const legend = document.createElement("div");
    legend.className = "legend";
    legend.innerHTML = `
      <span>Semáforo días de piso:</span>
      <span class="dot" style="background:var(--kpi-red)"></span><span>≤ ${KPI_DIAS.rojo}</span>
      <span class="dot" style="background:var(--kpi-orange)"></span><span>≤ ${KPI_DIAS.naranja}</span>
      <span class="dot" style="background:var(--kpi-yellow)"></span><span>≤ ${KPI_DIAS.amarillo}</span>
      <span class="dot" style="background:var(--kpi-green)"></span><span>&gt; ${KPI_DIAS.amarillo}</span>
    `;
    card.appendChild(legend);

    const tbl = document.createElement("table");
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Código</th>
          <th>Producto</th>
          <th class="right">Inv. Total</th>
          <th class="right">Venta prom</th>
          <th class="right">Días de piso</th>
          <th class="right">Fecha de quiebre</th>
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
          <td class="right" id="nextOut"></td>
          <td></td>
        </tr>
      </tfoot>
    `;
    const tbody = tbl.querySelector("tbody");

    for (const r of enriched) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.codigo}</td>
        <td>${r.producto}</td>
        <td class="right">${fmtNum(r.inv_total)}</td>
        <td class="right">${fmtNum(r.venta_prom)}</td>
        <td class="right">${chipDias(r.dias_piso)}</td>
        <td class="right">${fmtDate(r.fecha_quiebre)}</td>
        <td class="obs"><textarea data-code="${r.codigo}" rows="1" placeholder="Notas..."></textarea></td>
      `;
      tbody.appendChild(tr);
      const ta = tr.querySelector("textarea");
      ta.setAttribute("maxlength","220");
      ta.style.maxWidth = "520px";
      ta.value = getObs(r.codigo);
      ta.addEventListener("input", (e) => setObs(r.codigo, e.target.value));
    }

    const presentesSolo = enriched.filter(r => r.inv_total!=null || r.venta_prom!=null || r.dias_piso!=null);
    tbl.querySelector("#sumInv").textContent = fmtNum(sum(presentesSolo, "inv_total"));
    tbl.querySelector("#avgVen").textContent = fmtNum(avg(presentesSolo, "venta_prom"));
    tbl.querySelector("#avgDias").textContent = fmtNum(avg(presentesSolo, "dias_piso"));

    // Fecha de quiebre más próxima (entre presentes)
    const proximas = presentesSolo.map(r => r.fecha_quiebre).filter(Boolean).sort((a,b)=>a-b);
    tbl.querySelector("#nextOut").textContent = fmtDate(proximas[0] || null);

    card.appendChild(tbl);

    // === ACCIÓN HOY (debajo de cada tabla) ===
    const panel = document.createElement("div");
    panel.className = "todo";
    panel.innerHTML = `<h3>Acción hoy — ${linea}</h3>`;

    // helpers bonitos
    const toK = (n)=> Number.isFinite(n) ? n.toLocaleString('es-MX') : "–";
    const mkItem = (prod, extraLeft, extraRightHTML="") => {
      const div = document.createElement("div");
      div.className = "todo-item";
      div.innerHTML = `
        <span class="badge prod">${prod}</span>
        <span class="spacer"></span>
        ${extraLeft}
        ${extraRightHTML}
      `;
      return div;
    };

    // Tooltips de secciones
    const tipReab = `
      <div class="hint">?
        <div class="tip">
          <b>Reabastecer:</b> SKUs por debajo del objetivo de cobertura (<b>targetDOS = ${CFG.targetDOS} días</b>).<br/>
          Cantidad sugerida: <code>sugerido = max(0, targetDOS × venta_prom − inv_total)</code>.
        </div>
      </div>`;
    const tipRevis = `
      <div class="hint">?
        <div class="tip">
          <b>Revisar:</b> SKUs críticos por baja cobertura (colores rojo/naranja, es decir <b>días de piso ≤ ${KPI_DIAS.naranja}</b>) y con <b>venta &gt; 0</b>.<br/>
          En paréntesis se muestran los días de piso actuales.
        </div>
      </div>`;
    const tipTrans = `
      <div class="hint">?
        <div class="tip">
          <b>Traspaso:</b> SKUs muertos (hay inventario pero <b>venta_prom = 0</b>).<br/>
          En paréntesis se muestran las piezas que podrían moverse/traspasarse.
        </div>
      </div>`;

    // 1) Reabastecer (top 6 por sugerido)
    const reab = enriched
      .map(r => ({...r, sugerido: qtyReorden(r.inv_total, r.venta_prom)}))
      .filter(r => r.sugerido>0)
      .sort((a,b)=> b.sugerido - a.sugerido)
      .slice(0,6);

    const s1 = document.createElement("div");
    s1.className = "todo-section";
    s1.innerHTML = `<div class="todo-title">🔧 Reabastecer ${tipReab}</div>`;
    const l1 = document.createElement("div"); l1.className = "todo-list";
    if (reab.length) {
      for (const r of reab) {
        const left = `<span class="badge qty">${toK(r.sugerido)} pzas</span>`;
        const right = `<span class="badge dos">${chipDias(r.dias_piso)}</span>`;
        l1.appendChild(mkItem(r.producto, left, right));
      }
    } else {
      const p = document.createElement("div"); p.className="small"; p.textContent="Sin reabastecimientos urgentes.";
      s1.appendChild(p);
    }
    s1.appendChild(l1);
    panel.appendChild(s1);

    // 2) Revisar (críticos con venta > 0, por menor DOS)
    const revis = enriched
      .filter(r => Number.isFinite(r.venta_prom) && r.venta_prom>0 &&
                   Number.isFinite(r.dias_piso) && (r.dias_piso<=KPI_DIAS.naranja))
      .sort((a,b)=> (a.dias_piso||Infinity) - (b.dias_piso||Infinity))
      .slice(0,6);

    const s2 = document.createElement("div");
    s2.className = "todo-section";
    s2.innerHTML = `<div class="todo-title">⚠️ Revisar ${tipRevis}</div>`;
    const l2 = document.createElement("div"); l2.className = "todo-list";
    if (revis.length) {
      for (const r of revis) {
        const left = `<span class="badge">Días Piso ${fmtNum(r.dias_piso)} d</span>`;
        const right = `<span class="badge dos">${chipDias(r.dias_piso)}</span>`;
        l2.appendChild(mkItem(r.producto, left, right));
      }
    } else {
      const p = document.createElement("div"); p.className="small"; p.textContent="Sin críticos (rojo/naranja).";
      s2.appendChild(p);
    }
    s2.appendChild(l2);
    panel.appendChild(s2);

    // 3) Traspaso (mudertos)
    const trans = enriched
      .filter(r => (r.inv_total||0) > 0 && (!Number.isFinite(r.venta_prom) || r.venta_prom===0))
      .sort((a,b)=> (b.inv_total||0) - (a.inv_total||0))
      .slice(0,6);

    const s3 = document.createElement("div");
    s3.className = "todo-section";
    s3.innerHTML = `<div class="todo-title">🚚 Traspaso ${tipTrans}</div>`;
    const l3 = document.createElement("div"); l3.className = "todo-list";
    if (trans.length) {
      for (const r of trans) {
        const left = `<span class="badge qty">${toK(r.inv_total)} pzas</span>`;
        l3.appendChild(mkItem(r.producto, left));
      }
    } else {
      const p = document.createElement("div"); p.className="small"; p.textContent="Sin candidatos a traspaso.";
      s3.appendChild(p);
    }
    s3.appendChild(l3);
    panel.appendChild(s3);

    card.appendChild(panel);

    host.appendChild(card);
  }

  if (!host.children.length) {
    const p = document.createElement("div");
    p.className = "card";
    p.innerHTML = `<div class="empty">Carga un Excel para ver datos.</div>`;
    host.appendChild(p);
  }
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
    alert("No se detectaron filas de datos. Revisa que la hoja tenga encabezados con 'Código' y 'Producto'.");
  }
}

// --- Boot ---

(async function () {
  // 1) NO cargues el diccionario aquí sin clave
  // await loadDiccionarioWithKey();  // <-- ELIMINAR

  const file = document.getElementById("file");
  const fileMeta = document.getElementById("fileMeta");
  const clearData = document.getElementById("clearData");
  const clearObs = document.getElementById("clearObs");
  const toggleCatalogo = document.getElementById("toggleCatalogo");

  // 2) Si hay dataset guardado, primero carga el diccionario con la key guardada
  const prev = loadDataset();
  if (prev?.rows?.length) {
    try {
      if (!prev.meta?.keyB1) throw new Error("No hay keyB1 guardada. Sube un Excel una vez.");
      await loadDiccionarioWithKey(prev.meta.keyB1);
      renderTables(prev, toggleCatalogo.checked);
      fileMeta.textContent =
        `Cargando de memoria: ${prev.meta?.fileName || "dataset"} — hoja: ${prev.meta?.sheetName || "única"}`;
    } catch (e) {
      console.warn(e.message);
      fileMeta.textContent = "Sube un Excel para leer clave B y el diccionario.";
    }
  }

  toggleCatalogo.addEventListener("change", () => {
    const d = loadDataset();
    if (d?.rows?.length) renderTables(d, toggleCatalogo.checked);
  });

  // Al subir archivo:
  file.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const raw = await parseXlsx(f);              // trae raw.keyB1 (clave)
      await loadDiccionarioWithKey(raw.keyB1);     // descifra diccionario con esa clave
      const norm = normalizeRows(raw.rows);
      const meta = { fileName: raw.fileName, sheetName: raw.sheetName, keyB1: raw.keyB1 };
      saveDataset(meta, norm);
      renderTables({ meta, rows: norm }, toggleCatalogo.checked);
      fileMeta.textContent = `Dicc: ${DICC.byCode.size} entradas | Filas Excel: ${norm.length}`;
    } catch (err) {
      console.error(err);
      alert("Error al preparar datos: " + err.message);
    }
  });

  clearData.addEventListener("click", () => {
    localStorage.removeItem(LS_KEY_DATA);
    document.getElementById("tables").innerHTML = "";
    document.getElementById("kpis").style.display = "none";
    fileMeta.textContent = "";
  });

  clearObs.addEventListener("click", () => {
    for (const code of DICC.byCode.keys()) localStorage.removeItem(LS_KEY_OBS_PREFIX + code);
    const d = loadDataset();
    if (d?.rows?.length) renderTables(d, toggleCatalogo.checked);
  });
})();
