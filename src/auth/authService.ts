import {
  PublicClientApplication,
  type AccountInfo,
  InteractionRequiredAuthError,
} from '@azure/msal-browser'
import { msalConfig, loginRequest } from './msalConfig'
import type { AppUser } from '../types'

// ─── MSAL instance ────────────────────────────────────────────────────────────

const msalInstance = new PublicClientApplication(msalConfig)

let _initialized = false

export async function initAuth(): Promise<void> {
  if (_initialized) return
  await msalInstance.initialize()
  // Handle redirect response (if using redirect flow instead of popup)
  await msalInstance.handleRedirectPromise()
  _initialized = true
}

// ─── Account helpers ──────────────────────────────────────────────────────────

function getAccount(): AccountInfo | null {
  const accounts = msalInstance.getAllAccounts()
  return accounts.length > 0 ? accounts[0] : null
}

export function isAuthenticated(): boolean {
  return getAccount() !== null
}

export function getCurrentUser(): AppUser | null {
  const account = getAccount()
  if (!account) return null
  return {
    id: account.homeAccountId,
    displayName: account.name ?? '',
    mail: account.username,
    userPrincipalName: account.username,
  }
}

// ─── Sign in / out ────────────────────────────────────────────────────────────

export async function signIn(): Promise<AppUser> {
  try {
    const result = await msalInstance.loginPopup(loginRequest)
    const account = result.account
    return {
      id: account.homeAccountId,
      displayName: account.name ?? '',
      mail: account.username,
      userPrincipalName: account.username,
    }
  } catch (err) {
    // COOP browser policy can block popup detection even when login succeeded.
    // If the account was cached by MSAL, treat it as a successful sign-in.
    const account = getAccount()
    if (account) {
      return {
        id: account.homeAccountId,
        displayName: account.name ?? '',
        mail: account.username,
        userPrincipalName: account.username,
      }
    }
    // Fall back to redirect if popup was blocked entirely
    if ((err as Error).message?.includes('popup')) {
      await msalInstance.loginRedirect(loginRequest)
      return Promise.reject(new Error('Redirecting for login…'))
    }
    throw err
  }
}

export async function signOut(): Promise<void> {
  const account = getAccount()
  await msalInstance.logoutPopup({ account: account ?? undefined })
}

// ─── Token acquisition ────────────────────────────────────────────────────────

export async function getToken(scopes: string[] = loginRequest.scopes ?? []): Promise<string> {
  const account = getAccount()
  if (!account) throw new Error('Not authenticated')

  try {
    const result = await msalInstance.acquireTokenSilent({ scopes, account })
    return result.accessToken
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup({ scopes, account })
      return result.accessToken
    }
    throw err
  }
}
