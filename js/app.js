// Relevamiento de Campo — Carga de Fuego (PWA). Desarrollado por Ignacio Burd · Alonger SRL.
// Interfaz y formato tomados de la página de relevamiento V2.2.0 (lo importa Carga de Fuego desde la V2.2.0).
// Cambios: funciona sin señal (todo queda en el celular), cola de envío a SharePoint al llegar al WiFi,
// login de Microsoft como Remitos de Campo y equipos tipo contados (se calculan en la app de escritorio).
/* global JSZip */
import { CONFIG, MODO_DEMO } from './config.js';
import { CFG, GENERAL, UNITS, EQ_TYPES, SYS_TYPES, PRESETS, COMP_TYPES, EXT_STATUS, SERVICE_LINES, DEFAULT_TRACTOR, UNIT_TYPES } from './data.js';
import { iniciarAuth, ingresar, salir, cuentaActual } from './auth.js';
import { borrador, fotos, cola, enviados, persistir } from './almacen.js';
import { enviarRelevamiento } from './nube.js';

const TABS = [["datos", "Datos"], ["sectores", "Sectores"], ["equipos", "Equipos"], ["inventario", "Inventario"], ["extintores", "Extintores"], ["cierre", "Cierre"], ["envios", "Envíos"]];

let S = null;              // relevamiento en curso
let tab = "datos";
let filterSector = "";
const blobs = new Map();   // nombre de foto -> Blob local
const urls = new Map();    // nombre de foto -> object URL
let pendingConfirm = "";
let notice = null;
let queue = [];            // relevamientos terminados esperando envío
let sentList = [];         // historial de enviados
let sending = "";          // survey_id que se está subiendo
let sendProgress = "";
let unitLine = "";         // línea elegida en "Equipos tipo"


const $ = (s, r = document) => r.querySelector(s);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const num = v => { const n = parseFloat(String(v ?? "").replace(",", ".")); return isFinite(n) ? n : 0; };
const fmt = n => n.toLocaleString("es-AR", { maximumFractionDigits: 2 });
const slug = v => String(v || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 28) || "foto";
const today = () => { const d = new Date(); return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`; };
const newId = () => { const d = new Date(); const p = n => String(n).padStart(2, "0"); return `REL-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${Math.random().toString(36).slice(2, 6).toUpperCase()}`; };

function blankSurvey(installation) {
  const s = { format: "alonger-cfu-relevamiento", version: 1, survey_id: newId(), created_at: new Date().toISOString(),
    general: { tipo_instalacion: installation, fecha_estudio: today() }, sectors: [], items: [], equipment: [], units: [], extinguishers: [], systems: [], photos: [], notes: "" };
  if (installation) applyTemplate(s, installation);
  return s;
}
function applyTemplate(s, name) {
  const t = CFG.templates[name] || { sectors: [], items: [] };
  s.general.tipo_instalacion = name;
  s.sectors = t.sectors.map(([n, a]) => ({ name: n, activity: a, present: null, largo: "", ancho: "", area: "", notes: "", photos: [] }));
  s.items = t.items.map(i => ({ item: i, sector: "", present: "", quantity: "", capacity: "", unit: "", origin: "Confirmado en campo", notes: "", photos: [] }));
}
function applyPreset(name) {
  const p = PRESETS[name](); const s = blankSurvey(p.general.tipo_instalacion);
  Object.assign(s.general, p.general);
  s.sectors = p.sectors.map(([n, a]) => ({ name: n, activity: a, present: null, largo: "", ancho: "", area: "", notes: "", photos: [] }));
  s.equipment = p.equipment;
  s.items = p.items.map(i => ({ item: i, sector: "", present: "", quantity: "", capacity: "", unit: "", origin: "Confirmado en campo", notes: "", photos: [] }));
  s.systems = p.systems; s.preset = name;
  dropDraftPhotos(); S = s; blobs.clear(); urls.clear();
}
function sectorArea(sec) { const manual = num(sec.area); return manual > 0 ? manual : num(sec.largo) * num(sec.ancho); }


/* ---------- almacenamiento en el celular (IndexedDB) ---------- */
let saveTimer = null;
function touch() {
  S.updated_at = new Date().toISOString();
  $("#saveState").textContent = "Guardando…";
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try { await borrador.guardar({ survey: S }); $("#saveState").textContent = "Guardado en este celular"; }
    catch { $("#saveState").textContent = "No se pudo guardar: liberá espacio en el celular"; }
  }, 400);
}
// Fotos del borrador anterior que nunca se terminó (al precargar o empezar de nuevo).
function dropDraftPhotos() { if (S && !queue.some(q => q.survey_id === S.survey_id)) fotos.borrarTodas(S.survey_id); }

/* ---------- fotos ---------- */
function photoUrl(name) {
  if (urls.has(name)) return urls.get(name);
  if (blobs.has(name)) { const u = URL.createObjectURL(blobs.get(name)); urls.set(name, u); return u; }
  return "";
}
async function shrink(file) {
  try {
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas"); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    return await new Promise(r => c.toBlob(b => r(b || file), "image/jpeg", 0.82));
  } catch { return file; }
}
function photoName(prefix) {
  let k = 1, name;
  do { name = `${prefix}_${String(k).padStart(2, "0")}.jpg`; k++; } while (S.photos.some(p => p.file === name));
  return name;
}
async function addPhotos(kind, index, files) {
  const target = S[kind][index];
  for (const f of files) {
    let prefix;
    if (kind === "sectors") prefix = `S${String(index + 1).padStart(2, "0")}_${slug(target.name)}`;
    else if (kind === "items") { const si = S.sectors.findIndex(s => s.name === target.sector); prefix = `${si >= 0 ? "S" + String(si + 1).padStart(2, "0") : "SX"}_${slug(target.item)}`; }
    else if (kind === "systems") prefix = `SF${String(index + 1).padStart(2, "0")}_${slug(target.equipment || target.type)}`;
    else if (kind === "equipment") prefix = `EQ${String(index + 1).padStart(2, "0")}_${slug(target.name)}`;
    else prefix = `EXT_${String(index + 1).padStart(2, "0")}_${slug(target.agent)}`;
    const name = photoName(prefix);
    const blob = await shrink(f);
    blobs.set(name, blob);
    await fotos.guardar(S.survey_id, name, blob);
    target.photos.push(name);
    S.photos.push({ file: name, sector: kind === "sectors" ? target.name : (target.sector || ""), item: kind === "items" ? target.item : kind === "equipment" ? target.name : (kind === "extinguishers" ? "Extintor " + (index + 1) : "") });
  }
  touch(); render();
}
function removePhoto(kind, index, name) {
  const t = S[kind][index];
  t.photos = t.photos.filter(p => p !== name);
  S.photos = S.photos.filter(p => p.file !== name);
  blobs.delete(name); fotos.borrar(S.survey_id, name);
  touch(); render();
}
function photosHtml(kind, i, list) {
  return `<div class="photos">${list.map(p => `<div class="thumb"><img alt="${esc(p)}" src="${esc(photoUrl(p))}"><button type="button" data-act="delphoto" data-kind="${kind}" data-i="${i}" data-name="${esc(p)}" aria-label="Quitar foto">✕</button></div>`).join("")}
    <label class="addphoto">Sacar foto<input type="file" accept="image/*" capture="environment" data-photo="${kind}" data-i="${i}" aria-label="Sacar foto con la cámara"></label>
    <label class="addphoto">Desde galería<input type="file" accept="image/*" multiple data-photo="${kind}" data-i="${i}" aria-label="Elegir fotos de la galería"></label></div>`;
}

/* ---------- validaciones de cierre ---------- */
function checks() {
  const out = [];
  const g = S.general;
  for (const [k, l] of [["establecimiento", "Establecimiento"], ["fecha_estudio", "Fecha visita"], ["responsable_medicion", "Responsable del relevamiento"]])
    if (!String(g[k] || "").trim()) out.push(["bad", `Falta ${l} (pestaña Datos).`]);
  S.sectors.forEach(s => {
    if (s.present === null) out.push(["warn", `Sector "${s.name}": marcar si está presente o no aplica.`]);
    else if (s.present && sectorArea(s) <= 0) out.push(["bad", `Sector "${s.name}": falta la superficie (largo × ancho).`]);
    else if (s.present && !s.photos.length) out.push(["warn", `Sector "${s.name}": sin foto general.`]);
  });
  S.items.forEach(it => {
    if (it.present === "Sí" && !String(it.quantity).trim()) out.push(["bad", `"${it.item}": presente pero sin cantidad.`]);
    if (it.present === "Sí" && !it.sector) out.push(["warn", `"${it.item}": indicar en qué sector está.`]);
    if (it.present === "Sí" && !it.photos.length) out.push(["warn", `"${it.item}": sin foto de placa o etiqueta.`]);
  });
  S.equipment.forEach(eq => {
    const label = eq.name || "Equipo sin nombre";
    if (!num(eq.units)) out.push(["bad", `Equipo "${label}": falta la cantidad de equipos iguales.`]);
    if (!eq.components.length) out.push(["bad", `Equipo "${label}": no tiene componentes cargados.`]);
    eq.components.forEach(c => { if (!String(c.capacity).trim()) out.push(["warn", `Equipo "${label}": "${c.item || "componente"}" sin capacidad.`]); });
    if (!eq.sector) out.push(["warn", `Equipo "${label}": indicar el sector.`]);
    if (!eq.photos.length) out.push(["warn", `Equipo "${label}": sin foto de placa.`]);
  });
  S.systems.forEach(y => {
    const label = `${y.type || "Sistema"}${y.equipment ? " en " + y.equipment : ""}`;
    if (y.status === "Pendiente de verificar") out.push(["warn", `${label}: quedó pendiente de verificar.`]);
    if (y.status !== "Fuera de servicio" && (!String(y.capacity).trim() || !String(y.recharge).trim())) out.push(["warn", `${label}: falta capacidad o fecha de recarga.`]);
    if (!y.photos.length) out.push(["warn", `${label}: sin foto de placa.`]);
  });
  (S.units || []).forEach(u => {
    if (!u.unit) out.push(["bad", "Hay un equipo tipo sin elegir (pestaña Equipos)."]);
    else if (!num(u.count)) out.push(["bad", `Equipo tipo "${u.unit}": falta la cantidad.`]);
    else if (!u.sector) out.push(["warn", `Equipo tipo "${u.unit}": sin sector (la app de escritorio arma uno con la huella).`]);
  });
  if (!S.extinguishers.length) out.push(["warn", "No hay extintores relevados. Si el set no tiene, dejarlo anotado en Observaciones."]);
  if (!S.general.tipo_instalacion) out.push(["bad", "Falta elegir el tipo de instalación (pestaña Datos)."]);
  const pending = S.items.filter(it => !it.present).length;
  if (pending) out.push(["warn", `${pending} ítem(s) del inventario sin marcar como Sí o No.`]);
  return out;
}

/* ---------- render ---------- */
function renderTabs() {
  const counts = { sectores: S.sectors.filter(s => s.present).length, inventario: S.items.filter(i => i.present === "Sí").length, equipos: S.equipment.length + (S.units || []).length, envios: queue.length, extintores: S.extinguishers.reduce((a, e) => a + Math.max(1, Math.round(num(e.qty) || 1)), 0) };
  $("#tabs").innerHTML = TABS.map(([k, l]) => `<button role="tab" type="button" aria-selected="${k === tab}" data-tab="${k}">${l}${counts[k] ? `<span class="n">${counts[k]}</span>` : ""}</button>`).join("");
}
function render() {
  renderTabs();
  const m = $("#main");
  if (tab === "datos") m.innerHTML = viewDatos();
  if (tab === "sectores") m.innerHTML = viewSectores();
  if (tab === "equipos") m.innerHTML = viewEquipos();
  if (tab === "inventario") m.innerHTML = viewInventario();
  if (tab === "extintores") m.innerHTML = viewExtintores();
  if (tab === "cierre") m.innerHTML = viewCierre();
  if (tab === "envios") m.innerHTML = viewEnvios();
  renderUser();
}
function sel(path, value, options, attrs = "") {
  return `<select data-k="${path}" ${attrs}>${options.map(o => `<option${o === value ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
}
function viewDatos() {
  const names = Object.keys(CFG.templates);
  const g = S.general;
  return `
  <h2>Datos del relevamiento</h2>
  <p class="hint">Código <span class="code">${esc(S.survey_id)}</span>. Si en la oficina prepararon el proyecto, importalo y se cargan los mismos sectores e ítems.</p>
  <section class="card">
    <label class="f" for="inst">Tipo de instalación
      <select id="inst" data-act="inst">${g.tipo_instalacion ? "" : `<option selected disabled value="">Elegí el tipo de instalación…</option>`}${names.map(n => `<option${n === g.tipo_instalacion ? " selected" : ""}>${esc(n)}</option>`).join("")}</select></label>
    ${pendingConfirm === "inst" ? `<div class="confirm">Se reemplazan los sectores e ítems cargados. <button class="btn small primary" type="button" data-act="inst-yes">Reemplazar</button><button class="btn small ghost" type="button" data-act="inst-no">Cancelar</button></div>` : ""}
    <label class="f" for="preset">Relevamiento precargado
      <select id="preset"><option value="">Ninguno</option>${Object.keys(PRESETS).map(n => `<option${n === S.preset ? " selected" : ""}>${esc(n)}</option>`).join("")}</select></label>
    ${pendingConfirm === "preset" ? `<div class="confirm">Se reemplaza todo lo cargado en este borrador. <button class="btn small primary" type="button" data-act="preset-yes">Precargar</button><button class="btn small ghost" type="button" data-act="cancel">Cancelar</button></div>` : ""}
    <div class="row"><label class="btn small" for="plan">Importar proyecto preparado (.cfu.json)</label><input id="plan" type="file" accept=".json,application/json" hidden></div>
  </section>
  <section class="card"><div class="grid two">
    ${GENERAL.map(([k, l]) => `<label class="f" for="g_${k}">${l}<input id="g_${k}" data-k="general.${k}" value="${esc(g[k] || "")}"${k === "fecha_estudio" ? ' placeholder="dd/mm/aaaa"' : ""}${k === "hora_inicio" ? ' placeholder="hh:mm"' : ""}></label>`).join("")}
  </div></section>`;
}
function viewSectores() {
  const total = S.sectors.filter(s => s.present).reduce((a, s) => a + sectorArea(s), 0);
  return `
  <h2>Sectores</h2>
  <p class="hint">Medí largo y ancho de cada sector; la superficie se calcula sola. Si el sector no tiene forma rectangular, cargá la superficie directa. Superficie relevada: <b class="area">${fmt(total)} m²</b>.</p>
  ${S.sectors.map((s, i) => `
  <section class="card${s.present === false ? " off" : ""}">
    <div class="card-head"><span class="code">S${String(i + 1).padStart(2, "0")}</span>
      <div class="seg"><button type="button" class="${s.present === true ? "on-yes" : ""}" data-act="spres" data-i="${i}" data-v="1">Presente</button><button type="button" class="${s.present === false ? "on-na" : ""}" data-act="spres" data-i="${i}" data-v="0">No aplica</button></div></div>
    <label class="f" for="sn${i}">Sector<input id="sn${i}" data-k="sectors.${i}.name" value="${esc(s.name)}"></label>
    ${s.present === false ? "" : `
    <div class="grid">
      <label class="f" for="sl${i}">Largo (m)<input id="sl${i}" inputmode="decimal" data-k="sectors.${i}.largo" value="${esc(s.largo)}"></label>
      <label class="f" for="sa${i}">Ancho (m)<input id="sa${i}" inputmode="decimal" data-k="sectors.${i}.ancho" value="${esc(s.ancho)}"></label>
      <label class="f" for="sx${i}">Superficie directa (m²)<input id="sx${i}" inputmode="decimal" data-k="sectors.${i}.area" value="${esc(s.area)}" placeholder="opcional"></label>
    </div>
    <div>Superficie: <b class="area" id="area${i}">${fmt(sectorArea(s))} m²</b></div>
    <label class="f" for="sd${i}">Actividad / equipos<input id="sd${i}" data-k="sectors.${i}.activity" value="${esc(s.activity)}"></label>
    <label class="f" for="so${i}">Observaciones<textarea id="so${i}" data-k="sectors.${i}.notes">${esc(s.notes)}</textarea></label>
    ${photosHtml("sectors", i, s.photos)}`}
  </section>`).join("")}
  <div class="row"><button class="btn" type="button" data-act="addsector">+ Agregar sector</button></div>`;
}
function eqSummary(eq) {
  const n = num(eq.units);
  if (!n) return "Indicá cuántos equipos iguales hay.";
  const parts = eq.components.filter(c => c.item).map(c => {
    const q = num(c.quantity) || 1, cap = num(c.capacity);
    return cap ? `${c.item}: ${fmt(n * q)} × ${fmt(cap)} ${c.unit || ""} = ${fmt(n * q * cap)} ${c.unit || ""}`.trim() : `${c.item}: ${fmt(n * q)} ${c.unit || "u."}`;
  });
  return `${fmt(n)} equipo(s). ` + (parts.join(" · ") || "Sin componentes.");
}
function viewEquipos() {
  const sectors = S.sectors.filter(s => s.present !== false).map(s => s.name);
  return `${viewUnits(sectors)}
  <h2>Equipos relevados (componentes)</h2>
  <p class="hint">Relevá un equipo una sola vez (por ejemplo un generador: aceite de cárter, refrigerante, neumáticos) e indicá cuántos iguales hay. Al exportar se multiplica por la cantidad.</p>
  <datalist id="eqtypes">${EQ_TYPES.map(t => `<option value="${esc(t)}">`).join("")}</datalist>
  <datalist id="comptypes">${COMP_TYPES.map(t => `<option value="${esc(t)}">`).join("")}</datalist>
  ${S.equipment.map((eq, i) => `
  <section class="card">
    <div class="card-head"><span class="code">EQ${String(i + 1).padStart(2, "0")}</span><button class="btn small ghost" type="button" data-act="deleq" data-i="${i}">Quitar equipo</button></div>
    <div class="grid">
      <label class="f" for="eqn${i}">Equipo<input id="eqn${i}" list="eqtypes" data-k="equipment.${i}.name" value="${esc(eq.name)}" placeholder="ej. Generador diésel"></label>
      <label class="f" for="eqm${i}">Marca / modelo<input id="eqm${i}" data-k="equipment.${i}.model" value="${esc(eq.model)}" placeholder="opcional"></label>
      <label class="f" for="equ${i}">Equipos iguales<input id="equ${i}" inputmode="decimal" data-k="equipment.${i}.units" value="${esc(eq.units)}" placeholder="ej. 10"></label>
      <label class="f">Sector${sel(`equipment.${i}.sector`, eq.sector, ["", ...sectors])}</label>
    </div>
    <label class="f" for="eqo${i}">Observaciones<input id="eqo${i}" data-k="equipment.${i}.notes" value="${esc(eq.notes || "")}"></label>
    <p class="card-title" style="font-size:.85rem">Componentes de UN equipo</p>
    <p class="hint">Cantidad por equipo: litros de aceite, unidades de neumáticos, metros de cable. Capacidad c/u solo si son varios recipientes iguales (ej. 2 tanques × 300 L).</p>
    ${eq.components.map((c, j) => `
    <div class="grid" style="border-top:1px solid var(--line);padding-top:8px">
      <label class="f" for="ec${i}_${j}">Componente<input id="ec${i}_${j}" list="comptypes" data-k="equipment.${i}.components.${j}.item" value="${esc(c.item)}"></label>
      <label class="f" for="eq${i}_${j}">Cantidad por equipo<input id="eq${i}_${j}" inputmode="decimal" data-k="equipment.${i}.components.${j}.quantity" value="${esc(c.quantity)}" placeholder="ej. 600 L o 16 u."></label>
      <label class="f" for="ek${i}_${j}">Capacidad c/u (opcional)<input id="ek${i}_${j}" inputmode="decimal" data-k="equipment.${i}.components.${j}.capacity" value="${esc(c.capacity)}"></label>
      <label class="f">Unidad${sel(`equipment.${i}.components.${j}.unit`, c.unit, ["", ...UNITS])}</label>
      <label class="f">Origen${sel(`equipment.${i}.components.${j}.origin`, c.origin, CFG.origins)}</label>
      <div class="row" style="align-self:end"><button class="btn small ghost" type="button" data-act="delcomp" data-i="${i}" data-j="${j}">Quitar</button></div>
    </div>`).join("")}
    <div class="row"><button class="btn small" type="button" data-act="addcomp" data-i="${i}">+ Componente</button></div>
    <div class="msg" id="eqtot${i}">${esc(eqSummary(eq))}</div>
    ${photosHtml("equipment", i, eq.photos)}
  </section>`).join("") || `<p class="hint">Todavía no hay equipos. Agregá uno para relevarlo una vez y multiplicarlo.</p>`}
  <div class="row"><button class="btn" type="button" data-act="addeq">+ Agregar equipo</button></div>`;
}
function viewInventario() {
  const sectors = S.sectors.filter(s => s.present !== false).map(s => s.name);
  const list = S.items.map((it, i) => [it, i]).filter(([it]) => !filterSector || it.sector === filterSector || (filterSector === "-" && !it.sector));
  return `
  <h2>Inventario</h2>
  <p class="hint">Para cada elemento: si está, en qué sector, cuántos y de qué capacidad. Sacá foto de la placa o etiqueta (capacidad de aceite, IBC, FDS).</p>
  <div class="chips">${[["", "Todos"], ["-", "Sin sector"], ...sectors.map(s => [s, s])].map(([v, l]) => `<button class="chip" type="button" aria-pressed="${filterSector === v}" data-act="filter" data-v="${esc(v)}">${esc(l)}</button>`).join("")}</div>
  ${list.map(([it, i]) => `
  <section class="card${it.present === "No" ? " off" : ""}">
    <div class="card-head"><p class="card-title">${esc(it.item)}</p>
      <div class="seg"><button type="button" class="${it.present === "Sí" ? "on-yes" : ""}" data-act="ipres" data-i="${i}" data-v="Sí">Sí</button><button type="button" class="${it.present === "No" ? "on-no" : ""}" data-act="ipres" data-i="${i}" data-v="No">No</button></div></div>
    ${it.present === "No" ? "" : `
    <div class="grid">
      <label class="f">Sector${sel(`items.${i}.sector`, it.sector, ["", ...sectors])}</label>
      <label class="f" for="iq${i}">Cantidad<input id="iq${i}" inputmode="decimal" data-k="items.${i}.quantity" value="${esc(it.quantity)}"></label>
      <label class="f" for="ic${i}">Capacidad c/u<input id="ic${i}" data-k="items.${i}.capacity" value="${esc(it.capacity)}" placeholder="ej. 1500"></label>
      <label class="f">Unidad${sel(`items.${i}.unit`, it.unit, ["", ...UNITS])}</label>
    </div>
    <div class="grid two">
      <label class="f">Origen del dato${sel(`items.${i}.origin`, it.origin, CFG.origins)}</label>
      <label class="f" for="in${i}">Observaciones<input id="in${i}" data-k="items.${i}.notes" value="${esc(it.notes)}" placeholder="producto, tag, estado…"></label>
    </div>
    ${photosHtml("items", i, it.photos)}`}
  </section>`).join("") || `<p class="hint">No hay ítems en este filtro.</p>`}
  <div class="row"><input id="newitem" placeholder="Otro elemento (ej. Tanque de lubricante)" style="flex:1 1 220px"><button class="btn" type="button" data-act="additem">+ Agregar</button></div>`;
}
function viewExtintores() {
  const sectors = S.sectors.filter(s => s.present !== false).map(s => s.name);
  const agents = Object.keys(CFG.agents);
  return `
  <h2>Extintores existentes</h2>
  <p class="hint">Si hay varios extintores iguales (mismo agente, capacidad, sector y estado), cargalos una vez e indicá la cantidad. La ubicación exacta se dibuja después en el layout; acá alcanza con el sector.</p>
  ${S.extinguishers.map((e, i) => `
  <section class="card">
    <div class="card-head"><span class="code">E-${String(i + 1).padStart(2, "0")}${num(e.qty) > 1 ? ` · ${fmt(num(e.qty))} equipos` : ""}</span><button class="btn small ghost" type="button" data-act="delext" data-i="${i}">Quitar</button></div>
    <div class="grid">
      <label class="f" for="eq_n${i}">Cantidad iguales<input id="eq_n${i}" inputmode="numeric" data-k="extinguishers.${i}.qty" value="${esc(e.qty || "1")}"></label>
      <label class="f">Agente${sel(`extinguishers.${i}.agent`, e.agent, agents, 'data-rerender="1"')}</label>
      <label class="f">Capacidad${sel(`extinguishers.${i}.capacity`, e.capacity, CFG.agents[e.agent] || [])}</label>
      <label class="f">Sector${sel(`extinguishers.${i}.sector`, e.sector, ["", ...sectors])}</label>
      <label class="f">Estado${sel(`extinguishers.${i}.status`, e.status, EXT_STATUS)}</label>
      <label class="f" for="ev${i}">Vencimiento<input id="ev${i}" data-k="extinguishers.${i}.expiry" value="${esc(e.expiry)}" placeholder="mm/aaaa"></label>
      <label class="f" for="ec${i}">Condición<input id="ec${i}" data-k="extinguishers.${i}.condition" value="${esc(e.condition)}" placeholder="carga, precinto, señalización"></label>
    </div>
    ${photosHtml("extinguishers", i, e.photos)}
  </section>`).join("") || `<p class="hint">Todavía no hay extintores cargados.</p>`}
  <div class="row"><button class="btn" type="button" data-act="addext">+ Agregar extintor</button></div>
  <h2>Sistemas fijos de supresión</h2>
  <p class="hint">Sistemas instalados en los equipos (por ejemplo AFEX en los generadores). Sacá foto de la placa del tanque y del cartucho.</p>
  <datalist id="systypes">${SYS_TYPES.map(t => `<option value="${esc(t)}">`).join("")}</datalist>
  ${S.systems.map((y, i) => `
  <section class="card">
    <div class="card-head"><span class="code">SF-${String(i + 1).padStart(2, "0")}</span><button class="btn small ghost" type="button" data-act="delsys" data-i="${i}">Quitar</button></div>
    <div class="grid">
      <label class="f" for="yt${i}">Sistema<input id="yt${i}" list="systypes" data-k="systems.${i}.type" value="${esc(y.type)}"></label>
      <label class="f" for="ye${i}">Equipo protegido<input id="ye${i}" data-k="systems.${i}.equipment" value="${esc(y.equipment)}"></label>
      <label class="f">Sector${sel(`systems.${i}.sector`, y.sector, ["", ...sectors])}</label>
      <label class="f">Estado${sel(`systems.${i}.status`, y.status, ["Existente", "Pendiente de verificar", "Fuera de servicio"])}</label>
      <label class="f" for="ya${i}">Agente<input id="ya${i}" data-k="systems.${i}.agent" value="${esc(y.agent)}"></label>
      <label class="f" for="yc${i}">Capacidad<input id="yc${i}" data-k="systems.${i}.capacity" value="${esc(y.capacity)}" placeholder="ej. 5 gal"></label>
      <label class="f" for="ys${i}">Modelo / serie<input id="ys${i}" data-k="systems.${i}.serial" value="${esc(y.serial)}"></label>
      <label class="f" for="yr${i}">Última recarga<input id="yr${i}" data-k="systems.${i}.recharge" value="${esc(y.recharge)}" placeholder="mm/aaaa"></label>
    </div>
    <label class="f" for="yn${i}">Observaciones<input id="yn${i}" data-k="systems.${i}.notes" value="${esc(y.notes)}"></label>
    ${photosHtml("systems", i, y.photos)}
  </section>`).join("") || `<p class="hint">Sin sistemas fijos cargados.</p>`}
  <div class="row"><button class="btn" type="button" data-act="addsys">+ Agregar sistema fijo</button></div>`;
}
function viewUnits(sectors) {
  const lines = SERVICE_LINES.filter(l => UNIT_TYPES.some(u => u.line === l));
  if (!unitLine) unitLine = lineForInstallation(S.general.tipo_instalacion) || lines[0];
  const options = UNIT_TYPES.filter(u => u.line === unitLine);
  return `
  <h2>Equipos tipo (contar)</h2>
  <p class="hint">Para un set: elegí el equipo y cargá cuántos hay (por ejemplo 22 bombeadores). La carga de fuego se calcula en la app de escritorio con los valores de cada unidad; el tractor va aparte. Si un equipo es distinto, relevalo abajo por componentes.</p>
  <section class="card">
    <div class="grid two">
      <label class="f" for="unitline">Línea de servicio<select id="unitline">${lines.map(l => `<option${l === unitLine ? " selected" : ""}>${esc(l)}</option>`).join("")}</select></label>
      <label class="f" for="unitpick">Equipo<select id="unitpick">${options.map(u => `<option value="${esc(u.name)}">${esc(u.name)} · ${u.largo}×${u.ancho} m</option>`).join("")}</select></label>
    </div>
    <div class="row"><button class="btn" type="button" data-act="addunit">+ Agregar equipo tipo</button></div>
  </section>
  ${(S.units || []).map((u, i) => {
    const t = UNIT_TYPES.find(x => x.name === u.unit);
    return `
  <section class="card">
    <div class="card-head"><p class="card-title">${esc(u.unit)}</p><button class="btn small ghost" type="button" data-act="delunit" data-i="${i}">Quitar</button></div>
    <div class="grid">
      <label class="f" for="uc${i}">Cantidad<input id="uc${i}" inputmode="numeric" data-k="units.${i}.count" value="${esc(u.count)}" placeholder="ej. 22"></label>
      <label class="f">Sector${sel(`units.${i}.sector`, u.sector, ["", ...sectors])}</label>
      ${t && t.towed ? `<label class="f">Con tractor<span class="row"><input type="checkbox" data-tractor="${i}" ${u.tractor ? "checked" : ""} style="width:auto"> ${esc(DEFAULT_TRACTOR)}</span></label>` : ""}
    </div>
    <label class="f" for="un${i}">Observaciones<input id="un${i}" data-k="units.${i}.notes" value="${esc(u.notes || "")}" placeholder="marca, modelo, diferencias con la unidad tipo…"></label>
    ${t ? `<p class="hint">Unidad tipo: ${fmt(t.kg_madera)} kg de madera equivalente por equipo, huella ${t.largo}×${t.ancho} m${t.towed && u.tractor ? " + tractor 5×4 m" : ""}.</p>` : ""}
  </section>`; }).join("")}`;
}
function lineForInstallation(inst) {
  return ({ "Fractura hidráulica": "Fractura hidráulica (diésel)", "Fractura eléctrica (e-frac)": "Fractura eléctrica (e-frac)",
    "Wireline / Slickline": "Wireline / Slickline", "Equipo de perforación": "Perforación", "Coiled Tubing": "Coiled Tubing",
    "Cementación": "Cementación", "Pulling / Workover": "Pulling / Workover" })[inst] || "";
}

function viewCierre() {
  const c = checks();
  const bad = c.filter(x => x[0] === "bad").length;
  const present = S.sectors.filter(s => s.present);
  return `
  <h2>Cierre del relevamiento</h2>
  <div class="stats">
    <div class="stat"><b>${present.length}</b><span>Sectores presentes</span></div>
    <div class="stat"><b>${fmt(present.reduce((a, s) => a + sectorArea(s), 0))}</b><span>m² relevados</span></div>
    <div class="stat"><b>${(S.units || []).reduce((a, u) => a + num(u.count), 0) + S.equipment.reduce((a, e) => a + num(e.units), 0)}</b><span>Equipos</span></div>
    <div class="stat"><b>${S.photos.length}</b><span>Fotos</span></div>
  </div>
  <section class="card">
    <div class="card-head"><p class="card-title">Controles</p><span class="pill ${bad ? "bad" : c.length ? "warn" : "ok"}">${bad ? bad + " a corregir" : c.length ? "Con avisos" : "Completo"}</span></div>
    ${c.length ? `<ul class="checks">${c.map(([k, t]) => `<li><span class="pill ${k}">${k === "bad" ? "Falta" : "Aviso"}</span><span>${esc(t)}</span></li>`).join("")}</ul>` : `<p class="hint">No faltan datos.</p>`}
  </section>
  <section class="card">
    <label class="f" for="gn">Observaciones generales<textarea id="gn" data-k="notes">${esc(S.notes)}</textarea></label>
  </section>
  ${notice ? `<div class="msg ${notice[0]}">${esc(notice[1])}</div>` : ""}
  <section class="card">
    <p class="card-title">Terminar</p>
    <p class="hint">El relevamiento queda guardado en el celular listo para enviar, y se empieza uno nuevo. Se envía solo cuando haya WiFi y sesión iniciada (pestaña Envíos).</p>
    ${pendingConfirm === "finish" ? `<div class="confirm">${bad ? `Hay ${bad} dato(s) a corregir. ` : ""}¿Terminar y dejar listo para enviar? <button class="btn small primary" type="button" data-act="finish-yes">Terminar</button><button class="btn small ghost" type="button" data-act="cancel">Seguir editando</button></div>`
      : `<div class="row"><button class="btn primary" type="button" data-act="finish">Terminar y dejar para enviar</button><button class="btn" type="button" data-act="zip">Descargar .zip (respaldo)</button></div>`}
  </section>
  <section class="card">
    ${pendingConfirm === "new" ? `<div class="confirm">Se borra este borrador y sus fotos del celular. <button class="btn small primary" type="button" data-act="new-yes">Descartar y empezar de nuevo</button><button class="btn small ghost" type="button" data-act="cancel">Cancelar</button></div>` : `<div class="row"><button class="btn ghost" type="button" data-act="new">Descartar borrador</button></div>`}
  </section>`;
}

function viewEnvios() {
  const user = cuentaActual();
  return `
  <h2>Envíos a SharePoint</h2>
  <p class="hint">Con WiFi: iniciá sesión una vez con tu cuenta de Alonger y tocá Enviar. Las fotos suben una por una; si se corta, sigue desde donde quedó. En la PC: Carga de Fuego → Archivo → Importar relevamiento desde SharePoint.</p>
  <section class="card">
    <div class="card-head"><p class="card-title">Cuenta</p><span class="net ${navigator.onLine ? "on" : "off"}">${navigator.onLine ? "Con conexión" : "Sin conexión"}</span></div>
    ${MODO_DEMO ? `<p class="hint">Modo demostración: no se envía nada.</p>` : user ? `<div class="row"><span>${esc(user.name || user.username)}</span><button class="btn small ghost" type="button" data-act="logout">Cerrar sesión</button></div>`
      : `<div class="row"><button class="btn primary" type="button" data-act="login" ${navigator.onLine ? "" : "disabled"}>Iniciar sesión (Microsoft 365)</button></div>${navigator.onLine ? "" : `<p class="hint">Necesita WiFi la primera vez.</p>`}`}
  </section>
  ${notice ? `<div class="msg ${notice[0]}">${esc(notice[1])}</div>` : ""}
  <section class="card">
    <div class="card-head"><p class="card-title">Pendientes de envío</p><span class="pill ${queue.length ? "warn" : "ok"}">${queue.length}</span></div>
    ${queue.length ? queue.map(q => `
    <div class="queue-item">
      <div><b>${esc(q.establecimiento || "Sin establecimiento")}</b> <span class="code">${esc(q.survey_id)}</span></div>
      <span class="hint">${esc(q.cliente || "")} · ${esc(q.fecha || "")} · ${q.fotos.length} fotos${q.subidas && q.subidas.length ? ` · ${q.subidas.length} ya subidas` : ""}</span>
      ${sending === q.survey_id ? `<div class="progress"><i style="width:${sendProgress || "0%"}"></i></div>` : ""}
      ${q.error ? `<span class="pill bad">${esc(q.error)}</span>` : ""}
      ${pendingConfirm === "qdel:" + q.survey_id ? `<div class="confirm">Se borra del celular sin enviar. <button class="btn small primary" type="button" data-act="qdel-yes" data-id="${esc(q.survey_id)}">Borrar</button><button class="btn small ghost" type="button" data-act="cancel">Cancelar</button></div>`
        : `<div class="row"><button class="btn small" type="button" data-act="qzip" data-id="${esc(q.survey_id)}">Descargar .zip</button><button class="btn small ghost" type="button" data-act="qdel" data-id="${esc(q.survey_id)}">Borrar</button></div>`}
    </div>`).join("") : `<p class="hint">No hay relevamientos esperando envío.</p>`}
    <div class="row"><button class="btn primary" type="button" data-act="send" ${queue.length && !sending && !MODO_DEMO ? "" : "disabled"}>${sending ? "Enviando…" : "Enviar ahora"}</button></div>
  </section>
  <section class="card">
    <p class="card-title">Enviados</p>
    ${sentList.length ? sentList.slice(0, 30).map(e => `<div class="queue-item"><div><b>${esc(e.establecimiento || "Sin establecimiento")}</b> <span class="code">${esc(e.survey_id)}</span></div><span class="hint">${esc(e.fecha || "")} · enviado ${esc(new Date(e.enviado).toLocaleString("es-AR"))}${e.webUrl ? ` · <a href="${esc(e.webUrl)}" target="_blank" rel="noopener">ver en SharePoint</a>` : ""}</span></div>`).join("") : `<p class="hint">Todavía no se envió nada desde este celular.</p>`}
  </section>`;
}

function renderUser() {
  const n = $("#net");
  if (n) { n.textContent = navigator.onLine ? "WiFi/datos" : "Sin señal"; n.className = "net " + (navigator.onLine ? "on" : "off"); }
  const q = $("#qcount"); if (q) q.textContent = queue.length ? `${queue.length} por enviar` : "";
}

/* ---------- eventos ---------- */
function setPath(path, value) {
  const parts = path.split("."); let o = S;
  for (let i = 0; i < parts.length - 1; i++) o = o[parts[i]];
  o[parts[parts.length - 1]] = value;
}
document.addEventListener("input", e => {
  const k = e.target.dataset.k; if (!k) return;
  const rn = k.match(/^sectors\.(\d+)\.name$/);
  if (rn) {
    const old = S.sectors[+rn[1]].name, nu = e.target.value;
    for (const list of [S.items, S.equipment, S.units, S.extinguishers, S.systems]) for (const x of list || []) if (x.sector === old) x.sector = nu;
  }
  setPath(k, e.target.value); touch();
  const m = k.match(/^sectors\.(\d+)\.(largo|ancho|area)$/);
  if (m) { const sec = S.sectors[+m[1]]; if (sec.present === null && sectorArea(sec) > 0) sec.present = true; const el = $("#area" + m[1]); if (el) el.textContent = fmt(sectorArea(sec)) + " m²"; }
  const q = k.match(/^equipment\.(\d+)\./);
  if (q) { const el = $("#eqtot" + q[1]); if (el) el.textContent = eqSummary(S.equipment[+q[1]]); }
});
document.addEventListener("change", async e => {
  const t = e.target;
  if (t.dataset.k && t.tagName === "SELECT") { setPath(t.dataset.k, t.value); touch(); if (t.dataset.rerender) { const i = t.dataset.k.split(".")[1]; S.extinguishers[i].capacity = (CFG.agents[t.value] || [""])[0] || ""; } render(); return; }
  if (t.dataset.photo && t.files && t.files.length) { const files = [...t.files]; t.value = ""; try { await addPhotos(t.dataset.photo, +t.dataset.i, files); } catch (err) { $("#saveState").textContent = "No se pudo agregar la foto"; } return; }
  if (t.id === "inst") { if (S.sectors.some(s => s.present !== null) || S.items.some(i => i.present)) { S._pendingInst = t.value; pendingConfirm = "inst"; render(); } else { applyTemplate(S, t.value); touch(); render(); } return; }
  if (t.id === "plan" && t.files[0]) { importPlan(t.files[0]); }
  if (t.id === "unitline") { unitLine = t.value; render(); return; }
  if (t.dataset.tractor !== undefined) { S.units[+t.dataset.tractor].tractor = t.checked; touch(); render(); return; }
  if (t.id === "preset" && t.value) { S._pendingPreset = t.value; if (S.sectors.some(x => x.present !== null) || S.equipment.length || S.photos.length) { pendingConfirm = "preset"; render(); } else { applyPreset(t.value); touch(); render(); } }
});
document.addEventListener("click", async e => {
  const b = e.target.closest("[data-tab],[data-act]"); if (!b || b.tagName === "SELECT") return;
  if (b.dataset.tab) { tab = b.dataset.tab; pendingConfirm = ""; notice = null; render(); scrollTo(0, 0); return; }
  const a = b.dataset.act, i = +b.dataset.i;
  if (a === "spres") { S.sectors[i].present = b.dataset.v === "1"; touch(); render(); }
  if (a === "ipres") { S.items[i].present = S.items[i].present === b.dataset.v ? "" : b.dataset.v; touch(); render(); }
  if (a === "delphoto") removePhoto(b.dataset.kind, i, b.dataset.name);
  if (a === "filter") { filterSector = b.dataset.v; render(); }
  if (a === "addsector") { S.sectors.push({ name: "Sector " + (S.sectors.length + 1), activity: "", present: true, largo: "", ancho: "", area: "", notes: "", photos: [] }); touch(); render(); }
  if (a === "additem") { const v = $("#newitem").value.trim(); if (v) { S.items.push({ item: v, sector: filterSector && filterSector !== "-" ? filterSector : "", present: "Sí", quantity: "", capacity: "", unit: "", origin: "Confirmado en campo", notes: "", photos: [] }); touch(); render(); } }
  if (a === "addext") { const ag = Object.keys(CFG.agents)[0]; S.extinguishers.push({ qty: "1", agent: ag, capacity: CFG.agents[ag][2] || CFG.agents[ag][0], sector: "", status: "Existente", expiry: "", condition: "", notes: "", photos: [] }); touch(); render(); }
  if (a === "addeq") { S.equipment.push({ name: "", model: "", units: "", sector: "", notes: "", photos: [], components: COMP_TYPES.slice(0, 2).concat(["Neumáticos"]).map(c => ({ item: c, quantity: "", capacity: "", unit: c === "Neumáticos" ? "unidad" : "L", origin: "Confirmado en campo" })) }); touch(); render(); }
  if (a === "deleq") { S.equipment.splice(i, 1); touch(); render(); }
  if (a === "addcomp") { S.equipment[i].components.push({ item: "", quantity: "", capacity: "", unit: "", origin: "Confirmado en campo" }); touch(); render(); }
  if (a === "delcomp") { S.equipment[i].components.splice(+b.dataset.j, 1); touch(); render(); }
  if (a === "delext") { S.extinguishers.splice(i, 1); touch(); render(); }
  if (a === "inst-yes") { applyTemplate(S, S._pendingInst); delete S._pendingInst; pendingConfirm = ""; touch(); render(); }
  if (a === "inst-no") { delete S._pendingInst; pendingConfirm = ""; render(); }
  if (a === "new") { pendingConfirm = "new"; render(); }
  if (a === "new-yes") { await resetDraft(); }
  if (a === "cancel") { pendingConfirm = ""; render(); }
  if (a === "preset-yes") { applyPreset(S._pendingPreset); pendingConfirm = ""; touch(); render(); }
  if (a === "addsys") { S.systems.push({ type: SYS_TYPES[0], equipment: "", sector: "", agent: "", capacity: "", serial: "", recharge: "", status: "Existente", notes: "", photos: [] }); touch(); render(); }
  if (a === "delsys") { S.systems.splice(i, 1); touch(); render(); }
  if (a === "zip") exportZip(b);
  if (a === "finish") { pendingConfirm = "finish"; render(); }
  if (a === "finish-yes") await finishSurvey();
  if (a === "send") sendQueue(true);
  if (a === "login") ingresar();
  if (a === "logout") salir();
  if (a === "qzip") exportQueuedZip(b.dataset.id, b);
  if (a === "qdel") { pendingConfirm = "qdel:" + b.dataset.id; render(); }
  if (a === "qdel-yes") { await discardQueued(b.dataset.id); }
  if (a === "addunit") { const v = $("#unitpick").value; if (v) { const u = UNIT_TYPES.find(x => x.name === v); S.units.push({ unit: v, count: "", tractor: !!(u && u.towed), sector: filterSector && filterSector !== "-" ? filterSector : "", notes: "" }); touch(); render(); } }
  if (a === "delunit") { S.units.splice(i, 1); touch(); render(); }
});

/* ---------- importar proyecto preparado en la oficina ---------- */
function importPlan(file) {
  const r = new FileReader();
  r.onload = () => {
    try {
      const p = JSON.parse(r.result);
      if (p.format !== "alonger-carga-fuego") throw new Error("No es un proyecto de Carga de Fuego.");
      const g = p.general || {};
      const inst = g.tipo_instalacion && CFG.templates[g.tipo_instalacion] ? g.tipo_instalacion : S.general.tipo_instalacion;
      const s = blankSurvey(inst);
      for (const [k] of GENERAL) if (g[k]) s.general[k] = String(g[k]);
      s.general.razon_social = g.razon_social || "";
      s.project_id = p.project_id || "";
      if ((p.sectors || []).length) s.sectors = p.sectors.map(x => ({ name: x.name, activity: x.activity || "", present: null, largo: "", ancho: "", area: num(x.area) > 0 ? String(x.area) : "", notes: "", photos: [] }));
      const items = (p.field_guide && p.field_guide.items) || [];
      if (items.length) s.items = items.map(x => ({ item: x.item, sector: x.sector || "", present: "", quantity: x.quantity || "", capacity: x.capacity || "", unit: x.unit || "", origin: "Confirmado en campo", notes: x.notes || "", photos: [] }));
      S = s; touch(); notice = null; render();
      $("#saveState").textContent = "Proyecto importado";
    } catch (err) { alertInline(err.message || "No se pudo leer el archivo."); }
  };
  r.readAsText(file);
}
function alertInline(text) { notice = ["bad", text]; tab = "cierre"; render(); }

/* ---------- exportar ---------- */
function exportPayload() {
  const out = JSON.parse(JSON.stringify(S));
  delete out._pendingInst;
  out.exported_at = new Date().toISOString();
  out.sectors = out.sectors.map(s => ({ ...s, area: sectorArea(s) || "" }));
  // Extintores iguales: se cargan una vez con cantidad y se exportan uno por equipo.
  out.extinguishers = (out.extinguishers || []).flatMap(e => Array.from({ length: Math.max(1, Math.min(200, Math.round(num(e.qty) || 1))) }, () => ({ ...e, qty: "1" })));
  // Equipos iguales: se relevan una vez y se exportan multiplicados como ítems del inventario.
  for (const eq of out.equipment || []) {
    const n = num(eq.units) || 1;
    const name = [eq.name || "Equipo", eq.model].filter(Boolean).join(" ");
    for (const c of eq.components || []) {
      if (!String(c.item || "").trim()) continue;
      if (!String(c.quantity || "").trim() && !String(c.capacity || "").trim()) continue; // componente sin relevar
      const q = num(c.quantity) || 1;
      out.items.push({ item: `${name} - ${c.item}`, sector: eq.sector || "", present: "Sí", quantity: String(n * q), capacity: String(c.capacity || ""),
        unit: c.unit || "", per_unit_quantity: String(c.quantity || ""), origin: c.origin || "Confirmado en campo", notes: `${fmt(n)} equipo(s) iguales × ${fmt(q)} por equipo` + (eq.notes ? `. ${eq.notes}` : ""), photos: [...(eq.photos || [])] });
    }
  }
  return out;
}
async function photoBlob(name, id = S.survey_id) {
  if (id === S.survey_id && blobs.has(name)) return blobs.get(name);
  try { return await fotos.leer(id, name) || null; } catch { return null; }
}
function zipName(g) { return `Relevamiento CFU - ${slug(g.establecimiento || "estudio")} (${String(g.fecha_estudio || "").replace(/\//g, "-")}).zip`; }
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
async function buildZip(payload, photoNames, id) {
  const zip = new JSZip();
  zip.file("relevamiento.json", JSON.stringify(payload, null, 2));
  let missing = 0;
  for (const n of photoNames) { const b = await photoBlob(n, id); if (b) zip.file("fotos/" + n, b); else missing++; }
  return { blob: await zip.generateAsync({ type: "blob" }), missing };
}
async function exportZip(btn) {
  btn.disabled = true; btn.textContent = "Armando…";
  try {
    const { blob, missing } = await buildZip(exportPayload(), S.photos.map(p => p.file), S.survey_id);
    download(blob, zipName(S.general));
    notice = ["ok", "Archivo descargado (queda en Descargas del celular)." + (missing ? ` ${missing} foto(s) no estaban disponibles.` : "")];
  } catch (err) { notice = ["bad", "No se pudo armar el .zip: " + (err && err.message || "error")]; }
  render();
}
async function exportQueuedZip(id, btn) {
  const q = queue.find(x => x.survey_id === id); if (!q) return;
  btn.disabled = true; btn.textContent = "Armando…";
  try { const { blob } = await buildZip(q.payload, q.fotos, id); download(blob, zipName(q.payload.general || {})); }
  catch (err) { notice = ["bad", "No se pudo armar el .zip."]; }
  render();
}

/* ---------- terminar, cola y envío ---------- */
async function finishSurvey() {
  const payload = exportPayload();
  const entry = { survey_id: S.survey_id, payload, fotos: S.photos.map(p => p.file), subidas: [],
    establecimiento: S.general.establecimiento || "", cliente: S.general.cliente || "", fecha: S.general.fecha_estudio || "",
    terminado: new Date().toISOString() };
  await cola.guardar(entry);
  queue = await cola.todos();
  const inst = S.general.tipo_instalacion, resp = S.general.responsable_medicion, mat = S.general.matricula_medicion;
  S = blankSurvey(inst); S.general.responsable_medicion = resp || ""; S.general.matricula_medicion = mat || "";
  blobs.clear(); urls.clear(); pendingConfirm = "";
  await borrador.guardar({ survey: S });
  notice = ["ok", `Relevamiento ${entry.survey_id} listo para enviar. Se envía al tener WiFi.`];
  tab = "envios"; render(); scrollTo(0, 0);
  sendQueue(false);
}
async function discardQueued(id) {
  await cola.quitar(id); await fotos.borrarTodas(id);
  queue = await cola.todos(); pendingConfirm = ""; render();
}
async function sendQueue(manual) {
  if (sending || MODO_DEMO || !queue.length) return;
  if (!navigator.onLine) { if (manual) { notice = ["bad", "Sin conexión. Se envía solo al conectarse al WiFi."]; render(); } return; }
  if (!cuentaActual()) { if (manual) { notice = ["bad", "Primero iniciá sesión con tu cuenta de Alonger."]; render(); } return; }
  for (const q of [...queue]) {
    sending = q.survey_id; sendProgress = "0%"; q.error = ""; render();
    try {
      const r = await enviarRelevamiento(q, n => photoBlob(n, q.survey_id),
        (done, total) => { sendProgress = Math.round(100 * done / total) + "%"; if (tab === "envios") { const bar = document.querySelector(".progress i"); if (bar) bar.style.width = sendProgress; } },
        e => cola.guardar(e));
      await enviados.guardar({ survey_id: q.survey_id, establecimiento: q.establecimiento, fecha: q.fecha, enviado: new Date().toISOString(), webUrl: r.webUrl });
      await cola.quitar(q.survey_id); await fotos.borrarTodas(q.survey_id);
      notice = ["ok", `Enviado: ${q.establecimiento || q.survey_id}.`];
    } catch (err) {
      q.error = err.red ? "Sin conexión: se reintenta solo" : (err.message || "Error al enviar");
      await cola.guardar(q);
      notice = ["bad", `No se pudo enviar ${q.establecimiento || q.survey_id}: ${q.error}`];
      if (err.red || (err.status && err.status < 500 && err.status !== 429)) break;
    }
  }
  sending = ""; sendProgress = "";
  queue = await cola.todos(); sentList = (await enviados.todos()).sort((a, b) => String(b.enviado).localeCompare(String(a.enviado)));
  render();
}
async function resetDraft() {
  await fotos.borrarTodas(S.survey_id);
  blobs.clear(); urls.clear();
  S = blankSurvey(S.general.tipo_instalacion);
  pendingConfirm = ""; notice = null; tab = "datos"; touch(); render();
}

/* ---------- inicio ---------- */
(async function start() {
  try {
    const saved = await borrador.leer();
    if (saved && saved.survey && saved.survey.format === "alonger-cfu-relevamiento") {
      S = saved.survey; S.equipment = S.equipment || []; S.systems = S.systems || []; S.units = S.units || [];
      for (const p of S.photos) { const b = await fotos.leer(S.survey_id, p.file); if (b) blobs.set(p.file, b); }
      $("#saveState").textContent = "Borrador recuperado";
    } else S = blankSurvey("");
    queue = await cola.todos();
    sentList = (await enviados.todos()).sort((a, b) => String(b.enviado).localeCompare(String(a.enviado)));
    persistir();
  } catch {
    S = S || blankSurvey("");
    $("#saveState").textContent = "Sin almacenamiento local: descargá el .zip antes de cerrar";
  }
  render();
  if (!MODO_DEMO) {
    try { await iniciarAuth(); } catch { /* sin señal: el login se hace después */ }
    render();
    sendQueue(false);
  }
  addEventListener("online", () => { render(); sendQueue(false); });
  addEventListener("offline", () => render());
  setInterval(() => { if (document.visibilityState === "visible") sendQueue(false); }, 90000);
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
