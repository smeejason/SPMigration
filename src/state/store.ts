import type { AppState } from '../types'

// ─── Initial state ────────────────────────────────────────────────────────────

const initialState: AppState = {
  auth: { user: null, isAuthenticated: false },
  projects: [],
  currentProject: null,
  treeData: null,
  mappings: [],
  sites: [],
  pendingSiteCreations: [],
  ui: {
    activeView: 'login',
    loading: false,
    error: null,
  },
}

// ─── Store ────────────────────────────────────────────────────────────────────

type Listener = (state: AppState) => void

class Store {
  private state: AppState = { ...initialState }
  private listeners: Set<Listener> = new Set()

  getState(): AppState {
    return this.state
  }

  setState(patch: DeepPartial<AppState>): void {
    this.state = deepMerge(this.state, patch) as AppState
    this.notify()
  }

  subscribe(listener: Listener): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  private notify(): void {
    this.listeners.forEach((l) => l(this.state))
  }
}

export const store = new Store()

// ─── Convenience selectors ────────────────────────────────────────────────────

export function getState(): AppState {
  return store.getState()
}

export function setState(patch: DeepPartial<AppState>): void {
  store.setState(patch)
}

export function subscribe(listener: Listener): () => void {
  return store.subscribe(listener)
}

// ─── Deep merge util ──────────────────────────────────────────────────────────

type DeepPartial<T> = T extends object
  ? { [K in keyof T]?: DeepPartial<T[K]> }
  : T

function deepMerge<T>(target: T, source: DeepPartial<T>): T {
  if (Array.isArray(source)) return source as unknown as T
  if (typeof source !== 'object' || source === null) return source as unknown as T
  const result = { ...target } as Record<string, unknown>
  for (const key of Object.keys(source as object)) {
    const srcVal = (source as Record<string, unknown>)[key]
    const tgtVal = (target as Record<string, unknown>)[key]
    if (typeof srcVal === 'object' && srcVal !== null && !Array.isArray(srcVal)) {
      result[key] = deepMerge(tgtVal, srcVal as DeepPartial<typeof tgtVal>)
    } else {
      result[key] = srcVal
    }
  }
  return result as T
}
