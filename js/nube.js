// Subida a SharePoint por Microsoft Graph, con la sesión del técnico.
// Cada relevamiento va a <biblioteca de documentos>/Relevamientos Carga de Fuego/<ID> - <establecimiento>/
//   fotos/<archivo>.jpg   (una por una: si se corta el WiFi, se retoma donde quedó)
//   relevamiento.json     (AL FINAL: la app de escritorio solo lista carpetas que lo tienen)
import { CONFIG } from './config.js';
import { obtenerToken } from './auth.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const esperar = ms => new Promise(r => setTimeout(r, ms));
let sitio = null;

async function g(ruta, { method = 'GET', body, contentType, reintentos = 3 } = {}) {
  const headers = { Authorization: 'Bearer ' + (await obtenerToken()) };
  if (contentType) headers['Content-Type'] = contentType;
  let r;
  try {
    r = await fetch(ruta.startsWith('http') ? ruta : GRAPH + ruta, { method, headers, body });
  } catch (e) {
    throw Object.assign(new Error('Sin conexión con SharePoint.'), { red: true });
  }
  if ([429, 500, 502, 503, 504].includes(r.status) && reintentos > 0) {
    await esperar(1000 * (+r.headers.get('Retry-After') || 3));
    return g(ruta, { method, body, contentType, reintentos: reintentos - 1 });
  }
  if (!r.ok) {
    let msg = r.statusText;
    try { msg = (await r.json()).error?.message || msg; } catch { /* sin cuerpo */ }
    throw Object.assign(new Error(`SharePoint respondió ${r.status}: ${msg}`), { status: r.status });
  }
  return r.status === 204 ? null : r.json();
}

async function sitioId() {
  if (!sitio) sitio = (await g(`/sites/${CONFIG.sitio.host}:${CONFIG.sitio.ruta}`)).id;
  return sitio;
}

const ruta = partes => partes.map(p => encodeURIComponent(p)).join('/');

// Nombre de carpeta válido en SharePoint (sin " * : < > ? / \ | # %).
export function carpetaDe(e) {
  const est = String(e.establecimiento || 'Sin establecimiento').replace(/["*:<>?/\\|#%~&{}]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  return `${e.survey_id} - ${est}`.replace(/[. ]+$/, '');
}

async function subir(partes, cuerpo, tipo) {
  const sid = await sitioId();
  return g(`/sites/${sid}/drive/root:/${ruta(partes)}:/content`,
    { method: 'PUT', body: cuerpo, contentType: tipo });
}

// Sube un relevamiento de la cola. leerFoto(nombre) -> Blob. alAvanzar(hechas, total).
// Marca en e.subidas las fotos que ya subieron, para retomar si se corta.
export async function enviarRelevamiento(e, leerFoto, alAvanzar, guardarProgreso) {
  const carpeta = [CONFIG.carpeta, carpetaDe(e)];
  e.subidas = e.subidas || [];
  const total = e.fotos.length + 1;
  for (const nombre of e.fotos) {
    if (e.subidas.includes(nombre)) continue;
    const blob = await leerFoto(nombre);
    if (blob) await subir([...carpeta, 'fotos', nombre], blob, 'image/jpeg');
    e.subidas.push(nombre);
    await guardarProgreso(e);
    alAvanzar(e.subidas.length, total);
  }
  const json = new Blob([JSON.stringify(e.payload, null, 2)], { type: 'application/json' });
  const item = await subir([...carpeta, 'relevamiento.json'], json, 'application/json');
  alAvanzar(total, total);
  return { webUrl: item?.webUrl || '', carpeta: carpeta.join('/') };
}
