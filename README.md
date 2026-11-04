# Relevamiento de Campo — Alonger SRL

**Versión 1.0.0** · Desarrollado por Ignacio Burd

App web para relevar **carga de fuego** desde el celular **sin señal** (pensada para el Apex sin datos).
Mismo esquema que Remitos de Campo: código en GitHub Pages, datos en SharePoint, login de Microsoft 365.

```
Celular sin señal → sectores, equipos tipo (22 bombeadores…), inventario, extintores, fotos → todo queda en el celular
WiFi en la base → Enviar → SharePoint (sitio RemitosCampo / Relevamientos Carga de Fuego / <ID> - <establecimiento>)
PC → Carga de Fuego V2.3.0 → Archivo → Importar relevamiento desde SharePoint → proyecto con sectores, materiales y fotos
```

- **Los datos nunca van a GitHub.** Relevamientos y fotos quedan en SharePoint, detrás del login de Alonger.
- Sin servidor. Login con MSAL (PKCE) desde el navegador, con el mismo registro de Entra que Remitos de Campo.
- **Sin señal**: la app abre desde el celular (service worker), guarda el borrador y las fotos en IndexedDB.
  "Terminar" deja el relevamiento en la cola y se puede empezar otro. Al volver el WiFi se envía solo.
- Las fotos suben una por una; si se corta, sigue desde donde quedó. `relevamiento.json` se sube al final:
  la app de escritorio solo muestra los relevamientos completos.
- Respaldo: "Descargar .zip" (mismo formato que importa Carga de Fuego desde la V2.2.0).

## Puesta en marcha (una sola vez)

1. **Entra** (https://entra.microsoft.com) → Registros de aplicaciones → **Remitos de Campo** → Autenticación →
   plataforma *Aplicación de página única* → agregar la URI `https://iburd84.github.io/relevamiento-campo/`.
   Los permisos (User.Read, Sites.ReadWrite.All, Files.ReadWrite.All) son los que ya tiene Remitos.
2. **Publicar**: repo público `iburd84/relevamiento-campo` → Settings → Pages → rama `main`, carpeta `/`.
3. **En el Apex, con WiFi**: abrir la dirección en Chrome → menú → **Agregar a pantalla principal** →
   pestaña Envíos → **Iniciar sesión**. Desde ahí se usa sin señal.

La carpeta en SharePoint (`Relevamientos Carga de Fuego` en la biblioteca de documentos del sitio RemitosCampo)
se crea sola con el primer envío. Para probar la interfaz sin enviar nada: `?demo=1`.

## Uso en campo
Datos → tipo de instalación (o importar el proyecto preparado en la oficina) → Sectores (largo × ancho, foto) →
Equipos (equipos tipo: elegir y poner cantidad, con o sin tractor; o equipos por componentes) → Inventario →
Extintores → Cierre → **Terminar y dejar para enviar**. En la base, con WiFi: pestaña **Envíos** → Enviar ahora.

## Archivos
`index.html`, `styles.css`, `js/app.js` (interfaz), `js/data.js` (plantillas y unidades tipo, generado desde
la app de escritorio), `js/almacen.js` (IndexedDB), `js/nube.js` (Graph/SharePoint), `js/auth.js` (MSAL),
`js/config.js`, `sw.js`, `manifest.webmanifest`, `vendor/` (MSAL 4.30.0, JSZip 3.10.1).
Al cambiar la versión: actualizar `CONFIG.version`, `VERSION` en `sw.js` y el pie de `index.html`.
