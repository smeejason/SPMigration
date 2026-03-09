import type { Configuration, PopupRequest } from '@azure/msal-browser'

// ─── Runtime config ──────────────────────────────────────────────────────────
// Values are read from window.__APP_CONFIG__ (injected at runtime for
// deployments) or fall back to Vite env vars for local dev.
// Set VITE_CLIENT_ID and VITE_TENANT_ID in a .env.local file during development.

function getConfig(): { clientId: string; tenantId: string } {
  const win = window as Window & { __APP_CONFIG__?: { clientId: string; tenantId: string } }
  if (win.__APP_CONFIG__) {
    return win.__APP_CONFIG__
  }
  const clientId = import.meta.env.VITE_CLIENT_ID as string | undefined
  const tenantId = import.meta.env.VITE_TENANT_ID as string | undefined
  if (!clientId || !tenantId) {
    console.warn(
      '[MSAL] VITE_CLIENT_ID / VITE_TENANT_ID not set. ' +
      'Create a .env.local file with these values for local development.'
    )
  }
  return {
    clientId: clientId ?? '',
    tenantId: tenantId ?? 'common',
  }
}

const { clientId, tenantId } = getConfig()

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
}

export const loginRequest: PopupRequest = {
  scopes: [
    'User.Read',
    'People.Read',
    'Sites.ReadWrite.All',
    'Sites.Manage.All',
  ],
}
