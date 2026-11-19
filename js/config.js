// Configuración de la app. Estos datos NO son secretos (tenant y client ID de una app SPA son públicos).
// Los relevamientos y las fotos viven en SharePoint (sitio de Remitos de Campo), nunca en este repo.

export const CONFIG = {
  version: '1.1.0',

  // Microsoft Entra: el mismo registro SPA de Remitos de Campo ("Remitos de Campo", tipo SPA).
  // Hay que agregarle la URI de redirección https://iburd84.github.io/relevamiento-campo/
  tenantId: 'ad9b3575-dfa1-4374-83ee-797ea46ee8da',
  clientId: '6accec4b-39ef-40c6-9bd8-f861b007bb38',
  scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],

  // Dataverse (solo lectura): listas de clientes, operadoras, yacimientos y personal.
  // El registro necesita el permiso delegado "Dynamics CRM → user_impersonation".
  dataverse: {
    url: 'https://orgb73a9f30.crm2.dynamics.com',
    scope: 'https://orgb73a9f30.crm2.dynamics.com/user_impersonation',
  },

  // Sitio de SharePoint y carpeta dentro de su biblioteca de documentos.
  // La app de escritorio Carga de Fuego (V2.3.0) lee de esta misma carpeta.
  sitio: { host: 'alongercomar.sharepoint.com', ruta: '/sites/RemitosCampo' },
  carpeta: 'Relevamientos Carga de Fuego',
};

// Con ?demo=1 la app no envía nada: sirve para probar la interfaz.
export const MODO_DEMO = new URLSearchParams(location.search).has('demo');
