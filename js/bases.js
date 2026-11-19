// Bases de Alonger desde Dataverse (solo lectura), igual que la app de escritorio
// (alonger_cfu/dataverse_sources.py): clientes, operadoras, yacimientos y personal.
// Se leen con WiFi y quedan guardadas en el celular para usarlas sin señal.
import { CONFIG } from './config.js';
import { obtenerToken } from './auth.js';
import { kv } from './almacen.js';

const FORMATEADO = '@OData.Community.Display.V1.FormattedValue';
const CLAVE = 'bases';

const texto = v => {
  const s = v == null ? '' : String(v).trim();
  return ['nan', 'none', 'nat'].includes(s.toLowerCase()) ? '' : s;
};
export const normalizar = v => texto(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
const activo = r => (r.statecode === 0 || r.statecode == null) && r.cr1d0_activo !== false;
const orden = (a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' });

export function armarNombres(filas, campo) {
  const vistos = new Set(), out = [];
  for (const r of filas) {
    const v = texto(r[campo]), k = normalizar(v);
    if (v && k && !vistos.has(k) && activo(r)) { vistos.add(k); out.push(v); }
  }
  return out.sort(orden);
}

// Dataverse guarda "Apellido Nombres": 'Burd José Ignacio' -> 'José Ignacio Burd' (como la app de escritorio).
export function reordenarNombre(completo) {
  const p = texto(completo).split(/\s+/).filter(Boolean);
  return p.length < 2 ? p.join(' ') : [...p.slice(1), p[0]].join(' ');
}

export function armarPersonal(filas) {
  const vistos = new Set(), out = [];
  for (const r of filas) {
    const completo = texto(r.cr1d0_nombrecompleto);
    if (!completo || !activo(r)) continue;
    const nombre = reordenarNombre(completo), k = normalizar(nombre);
    if (vistos.has(k)) continue;
    vistos.add(k);
    out.push({ nombre, matricula: texto(r.cr1d0_numerodematricula), sector: texto(r['_cr1d0_sector_value' + FORMATEADO]) });
  }
  return out.sort((a, b) => orden(a.nombre, b.nombre));
}

async function leerTabla(token, tabla, campos) {
  let url = `${CONFIG.dataverse.url}/api/data/v9.2/${tabla}?$select=${campos.join(',')}`;
  const filas = [];
  while (url) {
    let r;
    try {
      r = await fetch(url, { headers: {
        Authorization: 'Bearer ' + token, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=500' } });
    } catch {
      throw Object.assign(new Error('Sin conexión con Dataverse.'), { red: true });
    }
    if (!r.ok) {
      let msg = r.statusText;
      try { msg = (await r.json()).error?.message || msg; } catch { /* sin cuerpo */ }
      throw new Error(`Dataverse respondió ${r.status} (${tabla}): ${msg}`);
    }
    const j = await r.json();
    filas.push(...(j.value || []));
    url = j['@odata.nextLink'] || '';
  }
  return filas;
}

// Lee las 4 tablas y guarda la copia local. Puede redirigir al login si falta el permiso.
export async function actualizarBases(interactivo = true) {
  let token;
  try { token = await obtenerToken([CONFIG.dataverse.scope], interactivo); }
  catch (e) { throw interactivo ? e : new Error('Falta autorizar Dataverse: tocá "Actualizar bases".'); }
  const [clientes, operadoras, yacimientos, personal] = await Promise.all([
    leerTabla(token, 'cr1d0_clienteempresas', ['cr1d0_clienteempresanombre', 'cr1d0_activo', 'statecode']),
    leerTabla(token, 'cr1d0_generaloperadoras', ['cr1d0_nombreoperadora', 'cr1d0_activo', 'statecode']),
    leerTabla(token, 'cr1d0_generalyacimientos', ['cr1d0_nombreyacimiento', 'cr1d0_activo', 'statecode']),
    leerTabla(token, 'cr1d0_generalpersonals', ['cr1d0_nombrecompleto', 'cr1d0_numerodematricula', '_cr1d0_sector_value', 'cr1d0_activo', 'statecode']),
  ]);
  const bases = {
    clientes: armarNombres(clientes, 'cr1d0_clienteempresanombre'),
    operadoras: armarNombres(operadoras, 'cr1d0_nombreoperadora'),
    yacimientos: armarNombres(yacimientos, 'cr1d0_nombreyacimiento'),
    personal: armarPersonal(personal),
    leido: new Date().toISOString(),
  };
  if (!bases.clientes.length) throw new Error('Dataverse no devolvió clientes.');
  await kv.guardar(CLAVE, bases);
  return bases;
}

export const basesGuardadas = () => kv.leer(CLAVE).catch(() => null);

// Vale la pena releer si nunca se leyó o pasó más de un día.
export const basesViejas = b => !b || !b.leido || Date.now() - Date.parse(b.leido) > 24 * 3600 * 1000;
