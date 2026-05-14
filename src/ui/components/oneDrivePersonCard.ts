import { checkUserDriveAccess, grantUserDriveAccess, revokeUserDriveAccess, getUserDrive, getUserById } from '../../graph/graphClient'
import { getState, setState } from '../../state/store'
import { getCurrentUser } from '../../auth/authService'
import type { MigrationMapping, OneDriveAccessStatus } from '../../types'

export interface PersonCardOptions {
  mapping: MigrationMapping
  migrationAccount: string
  container: HTMLElement
  onAccessChanged: (newStatus: OneDriveAccessStatus, mappingId: string) => Promise<void>
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

export function accessStatusBadge(s: OneDriveAccessStatus | undefined): string {
  if (!s || s === 'unknown')               return `<span class="badge badge-neutral">— Not checked</span>`
  if (s === 'accessible' || s === 'granted') return `<span class="badge status-ready">✓ Has Access</span>`
  if (s === 'revoked')                     return `<span class="badge badge-revoked">↩ Revoked</span>`
  if (s === 'no-drive')                    return `<span class="badge status-error">✗ No Drive</span>`
  if (s === 'no-access')                   return `<span class="badge status-error">✗ No Access</span>`
  return `<span class="badge status-error">✗ Error</span>`
}

export function renderPersonCard(opts: PersonCardOptions): void {
  const { mapping, container, onAccessChanged } = opts
  // Fall back to the currently signed-in user when no explicit migration account is configured
  const migrationAccount = opts.migrationAccount || getCurrentUser()?.userPrincipalName || ''
  const userId = mapping.targetSite?.id ?? ''
  const displayName = mapping.resolvedDisplayName ?? mapping.targetSite?.displayName ?? '—'
  const webUrl = mapping.targetSite?.webUrl ?? ''
  const access = mapping.accessStatus

  container.innerHTML = `
    <div class="person-card">
      <div class="person-card-header">
        <div class="person-card-avatar">${escHtml(displayName.slice(0, 2).toUpperCase())}</div>
        <div class="person-card-name">${escHtml(displayName)}</div>
      </div>
      <div class="person-card-body">
        <div class="person-card-row">
          <span class="person-card-label">UPN</span>
          <span id="pc-upn" class="person-card-value">⏳ Loading…</span>
        </div>
        <div class="person-card-row">
          <span class="person-card-label">OneDrive URL</span>
          <span id="pc-url" class="person-card-value">
            ${webUrl ? `<a href="${escHtml(webUrl)}" target="_blank" rel="noopener" class="person-card-link">${escHtml(webUrl)}</a>` : '—'}
          </span>
        </div>
        <div class="person-card-row person-card-access-row">
          <span class="person-card-label">Access</span>
          <span id="pc-access" class="person-card-value">${accessStatusBadge(access)}</span>
        </div>
        <div id="pc-access-actions" class="person-card-actions"></div>
        <div id="pc-access-error" class="person-card-error" style="display:none"></div>
      </div>
    </div>
  `

  if (userId) {
    getUserById(userId).then(user => {
      const el = container.querySelector<HTMLElement>('#pc-upn')
      if (el) el.textContent = user?.userPrincipalName ?? user?.mail ?? '—'
    }).catch(() => {
      const el = container.querySelector<HTMLElement>('#pc-upn')
      if (el) el.textContent = '—'
    })

    renderAccessActions(container, userId, mapping.id, mapping.accessStatus, migrationAccount, onAccessChanged)

    // Always check live — migrationAccount only needed for the grant/revoke buttons, not the check itself
    checkAndRefreshAccess(container, userId, mapping.id, migrationAccount, onAccessChanged)
  }
}

function renderAccessActions(
  container: HTMLElement,
  userId: string,
  mappingId: string,
  currentStatus: OneDriveAccessStatus | undefined,
  migrationAccount: string,
  onAccessChanged: (newStatus: OneDriveAccessStatus, mappingId: string) => Promise<void>,
): void {
  const actionsEl = container.querySelector<HTMLElement>('#pc-access-actions')
  if (!actionsEl || !migrationAccount) return

  const s = currentStatus
  if (s === 'no-access' || s === 'revoked' || !s || s === 'unknown') {
    actionsEl.innerHTML = `<button class="btn btn-sm btn-warning" id="pc-btn-grant">Grant Access</button>`
    actionsEl.querySelector('#pc-btn-grant')?.addEventListener('click', () =>
      handleAccessAction(container, userId, mappingId, 'grant', migrationAccount, onAccessChanged))
  } else if (s === 'accessible' || s === 'granted') {
    actionsEl.innerHTML = `<button class="btn btn-sm btn-ghost" id="pc-btn-revoke">Revoke Access</button>`
    actionsEl.querySelector('#pc-btn-revoke')?.addEventListener('click', () =>
      handleAccessAction(container, userId, mappingId, 'revoke', migrationAccount, onAccessChanged))
  }
}

async function handleAccessAction(
  container: HTMLElement,
  userId: string,
  mappingId: string,
  action: 'grant' | 'revoke',
  migrationAccount: string,
  onAccessChanged: (newStatus: OneDriveAccessStatus, mappingId: string) => Promise<void>,
): Promise<void> {
  const errorEl = container.querySelector<HTMLElement>('#pc-access-error')
  const accessEl = container.querySelector<HTMLElement>('#pc-access')
  const actionsEl = container.querySelector<HTMLElement>('#pc-access-actions')
  if (!actionsEl) return

  const btn = actionsEl.querySelector('button') as HTMLButtonElement | null
  if (btn) { btn.disabled = true; btn.textContent = action === 'grant' ? 'Granting…' : 'Revoking…' }
  if (errorEl) errorEl.style.display = 'none'

  try {
    if (action === 'grant') {
      await grantUserDriveAccess(userId, migrationAccount)
      const newStatus: OneDriveAccessStatus = 'granted'
      if (accessEl) accessEl.innerHTML = accessStatusBadge(newStatus)
      syncMappingAccessStatus(userId, mappingId, newStatus)
      renderAccessActions(container, userId, mappingId, newStatus, migrationAccount, onAccessChanged)
      await onAccessChanged(newStatus, mappingId)
    } else {
      await revokeUserDriveAccess(userId, migrationAccount)
      const newStatus: OneDriveAccessStatus = 'revoked'
      if (accessEl) accessEl.innerHTML = accessStatusBadge(newStatus)
      syncMappingAccessStatus(userId, mappingId, newStatus)
      renderAccessActions(container, userId, mappingId, newStatus, migrationAccount, onAccessChanged)
      await onAccessChanged(newStatus, mappingId)
    }
  } catch (err) {
    const msg = (err as Error)?.message ?? String(err)
    if (errorEl) { errorEl.textContent = msg; errorEl.style.display = '' }
    if (btn) { btn.disabled = false; btn.textContent = '⚠ Failed — retry' }
  }
}

async function checkAndRefreshAccess(
  container: HTMLElement,
  userId: string,
  mappingId: string,
  migrationAccount: string,
  onAccessChanged: (newStatus: OneDriveAccessStatus, mappingId: string) => Promise<void>,
): Promise<void> {
  const accessEl = container.querySelector<HTMLElement>('#pc-access')
  if (accessEl) accessEl.innerHTML = `<span class="badge badge-neutral">⏳ Checking…</span>`

  try {
    const access = await checkUserDriveAccess(userId)

    let freshWebUrl: string | undefined
    if (access === 'accessible') {
      try {
        const drive = await getUserDrive(userId)
        if (drive?.webUrl) freshWebUrl = drive.webUrl
      } catch { /* non-fatal */ }
    }

    if (freshWebUrl) {
      const urlEl = container.querySelector<HTMLElement>('#pc-url')
      if (urlEl) urlEl.innerHTML = `<a href="${escHtml(freshWebUrl)}" target="_blank" rel="noopener" class="person-card-link">${escHtml(freshWebUrl)}</a>`
    }

    if (accessEl) accessEl.innerHTML = accessStatusBadge(access)
    syncMappingAccessStatus(userId, mappingId, access, freshWebUrl)
    renderAccessActions(container, userId, mappingId, access, migrationAccount, onAccessChanged)
  } catch {
    if (accessEl) accessEl.innerHTML = `<span class="badge status-error">✗ Could not check</span>`
  }
}

function syncMappingAccessStatus(
  userId: string,
  mappingId: string,
  newStatus: OneDriveAccessStatus,
  freshWebUrl?: string,
): void {
  setState({
    mappings: getState().mappings.map(m => {
      if (m.id !== mappingId && m.targetSite?.id !== userId) return m
      const updates: Partial<MigrationMapping> = { accessStatus: newStatus }
      if (freshWebUrl && m.targetSite) updates.targetSite = { ...m.targetSite, webUrl: freshWebUrl }
      return { ...m, ...updates }
    }),
  })
}
