// Login con la cuenta corporativa (Microsoft 365) desde el navegador, sin servidor.
// MSAL con PKCE (igual que Remitos de Campo). Solo hace falta para ENVIAR: en campo se trabaja sin sesión.
/* global msal */
import { CONFIG } from './config.js';

let app = null;
let cuenta = null;

function redirectUri() {
  return location.origin + location.pathname.replace(/index\.html$/, '');
}

export async function iniciarAuth() {
  app = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: redirectUri(),
    },
    cache: { cacheLocation: 'localStorage' },
  });
  await app.initialize();
  const r = await app.handleRedirectPromise();
  cuenta = r?.account || app.getAllAccounts()[0] || null;
  if (cuenta) app.setActiveAccount(cuenta);
  return cuenta;
}

export function ingresar() {
  // Se pide también el permiso de Dataverse en el mismo login (para las listas de bases).
  return app.loginRedirect({ scopes: CONFIG.scopes, extraScopesToConsent: [CONFIG.dataverse.scope], prompt: 'select_account' });
}

export function salir() {
  return app.logoutRedirect({ account: cuenta });
}

export async function obtenerToken(scopes = CONFIG.scopes, interactivo = true) {
  if (!cuenta) throw new Error('No hay sesión iniciada.');
  try {
    const r = await app.acquireTokenSilent({ scopes, account: cuenta });
    return r.accessToken;
  } catch (e) {
    if (interactivo && e instanceof msal.InteractionRequiredAuthError && navigator.onLine) {
      await app.acquireTokenRedirect({ scopes, account: cuenta });
    }
    throw e;
  }
}

export const cuentaActual = () => cuenta;
