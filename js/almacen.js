// Guardado en el celular (IndexedDB). Funciona sin señal.
// - kv/"draft": el relevamiento que se está cargando.
// - fotos: las fotos de cada relevamiento, con clave "<ID>/<archivo>" (no se pisan entre relevamientos).
// - cola: relevamientos terminados esperando WiFi para subir a SharePoint.
// - enviados: historial de lo que ya se subió (sin fotos).

const DB = 'relevamiento-campo';
let db = null;

export function abrir() {
  if (db) return Promise.resolve(db);
  return new Promise((ok, mal) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      d.createObjectStore('kv');
      d.createObjectStore('fotos');
      d.createObjectStore('cola', { keyPath: 'survey_id' });
      d.createObjectStore('enviados', { keyPath: 'survey_id' });
    };
    r.onsuccess = () => { db = r.result; ok(db); };
    r.onerror = () => mal(r.error);
  });
}

async function tx(almacen, modo, fn) {
  const d = await abrir();
  return new Promise((ok, mal) => {
    const t = d.transaction(almacen, modo);
    const pedido = fn(t.objectStore(almacen));
    t.oncomplete = () => ok(pedido?.result);
    t.onerror = () => mal(t.error);
    t.onabort = () => mal(t.error);
  });
}

export const borrador = {
  guardar: v => tx('kv', 'readwrite', s => s.put(v, 'draft')),
  leer: () => tx('kv', 'readonly', s => s.get('draft')),
};

const claveFoto = (id, nombre) => `${id}/${nombre}`;
export const fotos = {
  guardar: (id, nombre, blob) => tx('fotos', 'readwrite', s => s.put(blob, claveFoto(id, nombre))),
  leer: (id, nombre) => tx('fotos', 'readonly', s => s.get(claveFoto(id, nombre))),
  borrar: (id, nombre) => tx('fotos', 'readwrite', s => s.delete(claveFoto(id, nombre))),
  async borrarTodas(id) {
    const d = await abrir();
    return new Promise(ok => {
      const t = d.transaction('fotos', 'readwrite');
      t.objectStore('fotos').delete(IDBKeyRange.bound(`${id}/`, `${id}/￿`));
      t.oncomplete = () => ok(); t.onerror = () => ok();
    });
  },
};

export const cola = {
  guardar: e => tx('cola', 'readwrite', s => s.put(e)),
  todos: () => tx('cola', 'readonly', s => s.getAll()),
  leer: id => tx('cola', 'readonly', s => s.get(id)),
  quitar: id => tx('cola', 'readwrite', s => s.delete(id)),
};

export const enviados = {
  guardar: e => tx('enviados', 'readwrite', s => s.put(e)),
  todos: () => tx('enviados', 'readonly', s => s.getAll()),
};

// Pide al navegador que no borre los datos si el celular se queda sin espacio.
export async function persistir() {
  try { return navigator.storage && navigator.storage.persist ? await navigator.storage.persist() : false; }
  catch { return false; }
}
