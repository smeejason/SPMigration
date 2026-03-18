import { getState, setState } from '../../state/store'
import { checkUserDriveAccess, grantUserDriveAccess, saveMappingsFile } from '../../graph/graphClient'
import { getSpConfig } from '../../graph/projectService'
import type { MigrationMapping, OneDriveAccessStatus } from '../../types'

export function renderSummaryPanel(container: HTMLElement): void {
  const state = getState()
  if (state.currentProject?.type === 'OneDrive') {
    renderOneDriveSummary(container, state.mappings)
  } else {
    renderSharePointSummary(container, state.mappings, !!state.treeData)
  }
}

// ─── OneDrive Summary ─────────────────────────────────────────────────────────

function renderOneDriveSummary(container: HTMLElement, mappings: MigrationMapping[]): void {
  // Only show users that have actually been mapped (auto-matched or manually assigned)
  const odMappings = mappings.filter(m => m.targetSite !== null)

  if (odMappings.length === 0) {
    container.innerHTML = `
      <div class="summary-panel">
        <div class="summary-empty">
          <p>No OneDrive mappings yet. Run <strong>Auto Map</strong> to match users to their OneDrives.</p>
        </div>
      </div>`
    injectSummaryStyles()
    return
  }

  const matched      = odMappings.filter(m => m.matchStatus === 'matched')
  const hasAccess    = odMappings.filter(m => m.accessStatus === 'accessible' || m.accessStatus === 'granted')
  const noAccess     = odMappings.filter(m => m.accessStatus === 'no-access' || m.accessStatus === 'no-drive' || m.accessStatus === 'error')
  const unresolved   = odMappings.filter(m => m.matchStatus !== 'matched' && m.matchStatus !== 'pending' && m.matchStatus !== undefined)
  const totalSize    = odMappings.reduce((s, m) => s + m.sourceNode.sizeBytes, 0)
  const totalFiles   = odMappings.reduce((s, m) => s + m.sourceNode.fileCount, 0)

  container.innerHTML = `
    <div class="od-summary-panel">

      <!-- ── Stats ── -->
      <div class="summary-stats-row">
        <div class="stat-card">
          <div class="stat-value">${odMappings.length}</div>
          <div class="stat-label">Total Users</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${matched.length}</div>
          <div class="stat-label">Matched</div>
        </div>
        <div class="stat-card stat-card--success">
          <div class="stat-value" id="stat-val-has-access">${hasAccess.length}</div>
          <div class="stat-label">Has Access</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${formatBytes(totalSize)}</div>
          <div class="stat-label">Total Data</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${totalFiles.toLocaleString()}</div>
          <div class="stat-label">Files</div>
        </div>
        ${noAccess.length > 0 ? `
        <div class="stat-card stat-card--danger">
          <div class="stat-value">${noAccess.length}</div>
          <div class="stat-label">No Access</div>
        </div>` : ''}
        ${unresolved.length > 0 ? `
        <div class="stat-card stat-card--danger">
          <div class="stat-value">${unresolved.length}</div>
          <div class="stat-label">Unresolved</div>
        </div>` : ''}
      </div>

      <!-- ── Action Bar ── -->
      <div class="od-action-bar">
        <div class="od-btn-group">
          <div class="od-btn-cluster">
            <button id="btn-check-perms"  class="btn btn-secondary od-action-btn">🔍 Check Permissions</button>
            <button id="btn-grant-access" class="btn btn-secondary od-action-btn">🔑 Grant Drive Access</button>
          </div>
          <div class="od-btn-cluster od-btn-cluster--right">
            <button id="btn-export-csv"   class="btn btn-primary">Export as CSV</button>
            <button id="btn-export-json"  class="btn btn-ghost">Export as JSON</button>
          </div>
        </div>

        <!-- Check Permissions progress -->
        <div id="check-perms-progress" class="od-progress-section" style="display:none">
          <div class="od-progress-header">
            <span class="od-progress-title">Checking Permissions</span>
            <span id="check-perms-count" class="od-progress-count">0 / 0</span>
          </div>
          <div class="od-progress-bar-wrap">
            <div id="check-perms-bar" class="od-progress-bar" style="width:0%"></div>
          </div>
          <div class="od-progress-stats">
            <span id="check-perms-accessible" class="pstat pstat--ok">✓ Has Access: 0</span>
            <span id="check-perms-noaccess"   class="pstat pstat--err">✗ No Access: 0</span>
            <span id="check-perms-error"      class="pstat pstat--warn">⚠ Error: 0</span>
          </div>
        </div>

        <!-- Grant Drive Access progress -->
        <div id="grant-access-progress" class="od-progress-section" style="display:none">
          <div class="od-progress-header">
            <span class="od-progress-title">Granting Drive Access</span>
            <span id="grant-access-count" class="od-progress-count">0 / 0</span>
          </div>
          <div class="od-progress-bar-wrap">
            <div id="grant-access-bar" class="od-progress-bar od-progress-bar--grant" style="width:0%"></div>
          </div>
          <div class="od-progress-stats">
            <span id="grant-access-granted" class="pstat pstat--ok">✓ Granted: 0</span>
            <span id="grant-access-skipped" class="pstat pstat--info">— Already accessible: 0</span>
            <span id="grant-access-error"   class="pstat pstat--err">✗ Failed: 0</span>
          </div>
        </div>
      </div>

      <!-- ── Table ── -->
      <div class="summary-table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Source Path</th>
              <th>User</th>
              <th>Destination OneDrive URL</th>
              <th>Folder Path</th>
              <th>Match</th>
              <th>Has Access</th>
            </tr>
          </thead>
          <tbody>
            ${odMappings.map(odRowHtml).join('')}
          </tbody>
        </table>
      </div>

    </div>`

  injectSummaryStyles()

  container.querySelector('#btn-export-csv')?.addEventListener('click', () => exportOneDriveCsv(getState().mappings))
  container.querySelector('#btn-export-json')?.addEventListener('click', () => exportOneDriveJson(getState().mappings))
  container.querySelector('#btn-check-perms')?.addEventListener('click', () => runCheckPermissions(container))
  container.querySelector('#btn-grant-access')?.addEventListener('click', () => runGrantAccess(container))
}

// ─── Row / badge helpers ───────────────────────────────────────────────────────

function odRowHtml(m: MigrationMapping): string {
  const siteUrl    = siteUrlFromDriveUrl(m.targetSite?.webUrl ?? '')
  const folderPath = m.targetFolderPath || '/'
  const user       = m.resolvedDisplayName ?? m.targetSite?.displayName ?? '—'
  return `
    <tr>
      <td class="path-cell path-cell--wrap">${escHtml(m.sourceNode.originalPath)}</td>
      <td class="od-user-cell">${escHtml(user)}</td>
      <td class="path-cell path-cell--wrap">${siteUrl
        ? `<a href="${escHtml(siteUrl)}" target="_blank" rel="noopener">${escHtml(siteUrl)}</a>`
        : '—'}</td>
      <td class="path-cell">${escHtml(folderPath)}</td>
      <td>${odMatchBadge(m)}</td>
      <td data-access-for="${escHtml(m.id)}">${odAccessBadge(m)}</td>
    </tr>`
}

function odMatchBadge(m: MigrationMapping): string {
  if (m.matchStatus === undefined)    return `<span class="badge badge-manual">✏ Manual</span>`
  if (m.matchStatus === 'matched')    return `<span class="badge status-ready">✓ Matched</span>`
  if (m.matchStatus === 'not-found')  return `<span class="badge status-error">✗ Not Found</span>`
  if (m.matchStatus === 'ambiguous')  return `<span class="badge status-warning">? Ambiguous</span>`
  if (m.matchStatus === 'error')      return `<span class="badge status-error">✗ Error</span>`
  return `<span class="badge status-pending">⏳ Pending</span>`
}

function odAccessBadge(m: MigrationMapping): string {
  const s = m.accessStatus
  if (!s || s === 'unknown')              return `<span class="badge badge-neutral">— Not checked</span>`
  if (s === 'accessible' || s === 'granted') return `<span class="badge status-ready">✓ Has Access</span>`
  if (s === 'no-drive')                   return `<span class="badge status-error">✗ No Drive</span>`
  if (s === 'no-access')                  return `<span class="badge status-error">✗ No Access</span>`
  return `<span class="badge status-error">✗ Error</span>`
}

// ─── Check Permissions ────────────────────────────────────────────────────────

async function runCheckPermissions(container: HTMLElement): Promise<void> {
  const matchedMappings = getState().mappings.filter(m => m.targetSite?.id)
  if (matchedMappings.length === 0) return

  const btnCheck = container.querySelector<HTMLButtonElement>('#btn-check-perms')!
  const btnGrant = container.querySelector<HTMLButtonElement>('#btn-grant-access')!
  btnCheck.disabled = true
  btnGrant.disabled = true
  btnCheck.textContent = '⏳ Checking…'

  const section     = container.querySelector<HTMLElement>('#check-perms-progress')!
  const bar         = container.querySelector<HTMLElement>('#check-perms-bar')!
  const countEl     = container.querySelector<HTMLElement>('#check-perms-count')!
  const accessibleEl = container.querySelector<HTMLElement>('#check-perms-accessible')!
  const noaccessEl  = container.querySelector<HTMLElement>('#check-perms-noaccess')!
  const errorEl     = container.querySelector<HTMLElement>('#check-perms-error')!
  const statVal     = container.querySelector<HTMLElement>('#stat-val-has-access')

  section.style.display = ''
  bar.style.width = '0%'

  const total = matchedMappings.length
  let done = 0, accessibleCount = 0, noaccessCount = 0, errorCount = 0

  countEl.textContent = `0 / ${total}`

  for (const mapping of matchedMappings) {
    const userId = mapping.targetSite!.id
    let newStatus: OneDriveAccessStatus = 'error'
    try {
      const result = await checkUserDriveAccess(userId)
      newStatus = result
      if (result === 'accessible')                             accessibleCount++
      else if (result === 'no-access' || result === 'no-drive') noaccessCount++
      else                                                      errorCount++
    } catch {
      newStatus = 'error'
      errorCount++
    }

    setState({ mappings: getState().mappings.map(m =>
      m.id === mapping.id ? { ...m, accessStatus: newStatus } : m
    )})

    const cell = container.querySelector(`[data-access-for="${CSS.escape(mapping.id)}"]`) as HTMLElement | null
    if (cell) cell.innerHTML = odAccessBadge({ ...mapping, accessStatus: newStatus })

    done++
    const pct = Math.round((done / total) * 100)
    bar.style.width = `${pct}%`
    countEl.textContent = `${done} / ${total}`
    accessibleEl.textContent = `✓ Has Access: ${accessibleCount}`
    noaccessEl.textContent   = `✗ No Access: ${noaccessCount}`
    errorEl.textContent      = `⚠ Error: ${errorCount}`
    if (statVal) statVal.textContent = String(accessibleCount)

    await new Promise(r => setTimeout(r, 0))
  }

  btnCheck.disabled = false
  btnGrant.disabled = false
  btnCheck.textContent = '🔍 Check Permissions'

  await persistMappings()
}

// ─── Grant Drive Access ───────────────────────────────────────────────────────

async function runGrantAccess(container: HTMLElement): Promise<void> {
  const state = getState()
  const migrationAccount = state.currentProject?.projectData?.autoMapSettings?.migrationAccount
  if (!migrationAccount) {
    alert('No migration account configured. Set it in the Auto Map settings first.')
    return
  }

  const matchedMappings = state.mappings.filter(m => m.targetSite?.id)
  if (matchedMappings.length === 0) return

  const btnCheck = container.querySelector<HTMLButtonElement>('#btn-check-perms')!
  const btnGrant = container.querySelector<HTMLButtonElement>('#btn-grant-access')!
  btnCheck.disabled = true
  btnGrant.disabled = true
  btnGrant.textContent = '⏳ Granting…'

  const section   = container.querySelector<HTMLElement>('#grant-access-progress')!
  const bar       = container.querySelector<HTMLElement>('#grant-access-bar')!
  const countEl   = container.querySelector<HTMLElement>('#grant-access-count')!
  const grantedEl = container.querySelector<HTMLElement>('#grant-access-granted')!
  const skippedEl = container.querySelector<HTMLElement>('#grant-access-skipped')!
  const errorEl   = container.querySelector<HTMLElement>('#grant-access-error')!
  const statVal   = container.querySelector<HTMLElement>('#stat-val-has-access')

  section.style.display = ''
  bar.style.width = '0%'

  const total = matchedMappings.length
  let done = 0, grantedCount = 0, skippedCount = 0, errorCount = 0

  countEl.textContent = `0 / ${total}`

  for (const mapping of matchedMappings) {
    const userId = mapping.targetSite!.id
    let newStatus: OneDriveAccessStatus = 'error'
    try {
      const access = await checkUserDriveAccess(userId)
      if (access === 'accessible') {
        newStatus = 'accessible'
        skippedCount++
      } else if (access === 'no-access') {
        await grantUserDriveAccess(userId, migrationAccount)
        newStatus = 'granted'
        grantedCount++
      } else {
        newStatus = access  // 'no-drive' | 'error'
        errorCount++
      }
    } catch {
      newStatus = 'error'
      errorCount++
    }

    setState({ mappings: getState().mappings.map(m =>
      m.id === mapping.id ? { ...m, accessStatus: newStatus } : m
    )})

    const cell = container.querySelector(`[data-access-for="${CSS.escape(mapping.id)}"]`) as HTMLElement | null
    if (cell) cell.innerHTML = odAccessBadge({ ...mapping, accessStatus: newStatus })

    done++
    const pct = Math.round((done / total) * 100)
    bar.style.width = `${pct}%`
    countEl.textContent = `${done} / ${total}`
    grantedEl.textContent = `✓ Granted: ${grantedCount}`
    skippedEl.textContent = `— Already accessible: ${skippedCount}`
    errorEl.textContent   = `✗ Failed: ${errorCount}`

    const newHasAccess = getState().mappings.filter(m => m.accessStatus === 'accessible' || m.accessStatus === 'granted').length
    if (statVal) statVal.textContent = String(newHasAccess)

    await new Promise(r => setTimeout(r, 0))
  }

  btnCheck.disabled = false
  btnGrant.disabled = false
  btnGrant.textContent = '🔑 Grant Drive Access'

  await persistMappings()
}

async function persistMappings(): Promise<void> {
  try {
    const state = getState()
    const project = state.currentProject!
    const { siteId } = getSpConfig()
    await saveMappingsFile(siteId, project.title, project.id, state.mappings)
  } catch (err) {
    console.warn('[Summary] Failed to persist mappings:', err)
  }
}

// ─── OneDrive exports ─────────────────────────────────────────────────────────

function exportOneDriveCsv(mappings: MigrationMapping[]): void {
  const headers = ['Source Path', 'User', 'Destination OneDrive URL', 'Folder Path', 'Match Status', 'Access Status']
  const rows = mappings.map(m => [
    m.sourceNode.originalPath,
    m.resolvedDisplayName ?? m.targetSite?.displayName ?? '',
    siteUrlFromDriveUrl(m.targetSite?.webUrl ?? ''),
    m.targetFolderPath || '/',
    m.matchStatus ?? '',
    m.accessStatus ?? '',
  ])
  const csv = [headers, ...rows].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n')
  downloadFile(csv, 'onedrive-migration-plan.csv', 'text/csv')
}

function exportOneDriveJson(mappings: MigrationMapping[]): void {
  const tasks = mappings.map(m => ({
    SourcePath:             m.sourceNode.originalPath,
    TargetPath:             siteUrlFromDriveUrl(m.targetSite?.webUrl ?? ''),
    TargetList:             'Documents',
    TargetListRelativePath: m.targetFolderPath || '',
  }))
  downloadFile(JSON.stringify({ Tasks: tasks }, null, 2), 'onedrive-migration-plan.json', 'application/json')
}

// ─── SharePoint Summary ───────────────────────────────────────────────────────

function renderSharePointSummary(container: HTMLElement, mappings: MigrationMapping[], hasTree: boolean): void {
  if (!hasTree) {
    container.innerHTML = `<div class="summary-empty"><p>No data loaded. Start by uploading a TreeSize report.</p></div>`
    injectSummaryStyles()
    return
  }

  const ready      = mappings.filter(m => m.status === 'ready')
  const unmapped   = mappings.filter(m => m.status === 'pending')
  const totalSize  = ready.reduce((s, m) => s + m.sourceNode.sizeBytes, 0)
  const totalFiles = ready.reduce((s, m) => s + m.sourceNode.fileCount, 0)
  const uniqueSites = new Set(ready.map(m => m.targetSite?.id).filter(Boolean)).size

  container.innerHTML = `
    <div class="summary-panel">
      <div class="summary-stats-row">
        <div class="stat-card">
          <div class="stat-value">${ready.length}</div>
          <div class="stat-label">Mappings ready</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${formatBytes(totalSize)}</div>
          <div class="stat-label">Total data mapped</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${totalFiles.toLocaleString()}</div>
          <div class="stat-label">Files</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${uniqueSites}</div>
          <div class="stat-label">Target sites</div>
        </div>
      </div>

      ${unmapped.length > 0
        ? `<div class="summary-warning">⚠ ${unmapped.length} mapping${unmapped.length !== 1 ? 's' : ''} not yet assigned to a SharePoint target.</div>`
        : ''}

      <div class="summary-export-row">
        <button id="btn-export-csv"  class="btn btn-primary">Export as CSV</button>
        <button id="btn-export-json" class="btn btn-ghost">Export as JSON</button>
      </div>

      <div class="summary-table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Source Path</th>
              <th>Size</th>
              <th>Files</th>
              <th>Target Site</th>
              <th>Destination List</th>
              <th>Folder Path</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${mappings.length === 0
              ? `<tr><td colspan="7" class="table-empty">No mappings defined yet.</td></tr>`
              : mappings.map(spRowHtml).join('')}
          </tbody>
        </table>
      </div>
    </div>`

  injectSummaryStyles()
  container.querySelector('#btn-export-csv')?.addEventListener('click',  () => exportSpCsv(mappings))
  container.querySelector('#btn-export-json')?.addEventListener('click', () => exportSpJson(mappings))
}

function spRowHtml(m: MigrationMapping): string {
  const statusClass = m.status === 'ready' ? 'status-ready' : 'status-pending'
  const statusLabel = m.status === 'ready' ? '✅ Ready' : '⏳ Pending'
  return `
    <tr>
      <td class="path-cell" title="${escHtml(m.sourceNode.originalPath)}">${escHtml(m.sourceNode.originalPath)}</td>
      <td>${formatBytes(m.sourceNode.sizeBytes)}</td>
      <td>${m.sourceNode.fileCount.toLocaleString()}</td>
      <td>${m.targetSite  ? escHtml(m.targetSite.displayName)  : '—'}</td>
      <td>${m.targetDrive ? escHtml(m.targetDrive.name)        : '—'}</td>
      <td class="path-cell">${m.targetFolderPath ? escHtml(m.targetFolderPath) : '/'}</td>
      <td><span class="badge ${statusClass}">${statusLabel}</span></td>
    </tr>`
}

function exportSpCsv(mappings: MigrationMapping[]): void {
  const headers = ['Source Path', 'Size (Bytes)', 'File Count', 'Target Site', 'Destination List', 'Folder Path', 'Status']
  const rows = mappings.map(m => [
    m.sourceNode.originalPath,
    m.sourceNode.sizeBytes,
    m.sourceNode.fileCount,
    m.targetSite?.displayName ?? '',
    m.targetDrive?.name ?? '',
    m.targetFolderPath || '/',
    m.status,
  ])
  const csv = [headers, ...rows].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n')
  downloadFile(csv, 'migration-plan.csv', 'text/csv')
}

function exportSpJson(mappings: MigrationMapping[]): void {
  const data = mappings.map(m => ({
    sourcePath:        m.sourceNode.originalPath,
    sourceName:        m.sourceNode.name,
    sizeBytes:         m.sourceNode.sizeBytes,
    fileCount:         m.sourceNode.fileCount,
    targetSite:        m.targetSite?.webUrl  ?? null,
    targetSiteId:      m.targetSite?.id      ?? null,
    destinationList:   m.targetDrive?.name   ?? null,
    destinationListId: m.targetDrive?.id     ?? null,
    folderPath:        m.targetFolderPath || '/',
    status:            m.status,
  }))
  downloadFile(JSON.stringify(data, null, 2), 'migration-plan.json', 'application/json')
}

// ─── Shared helpers ───────────────────────────────────────────────────────────

function downloadFile(content: string, filename: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function formatBytes(bytes: number): string {
  if (!bytes) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function siteUrlFromDriveUrl(url: string): string {
  return url.replace(/\/Documents\/?$/i, '')
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectSummaryStyles(): void {
  if (document.getElementById('summary-styles')) return
  const style = document.createElement('style')
  style.id = 'summary-styles'
  style.textContent = `
    /* ── Shared ── */
    .summary-panel  { padding: 24px; }
    .summary-empty  { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .summary-warning { background: #fff4ce; color: #7d5900; padding: 10px 14px; border-radius: 4px;
      font-size: 0.88rem; margin-bottom: 16px; }

    /* ── Stat cards ── */
    .summary-stats-row { display: flex; gap: 16px; flex-wrap: wrap; }
    .stat-card { background: white; border: 1px solid var(--color-border); border-radius: 8px;
      padding: 16px 20px; flex: 1; min-width: 110px; }
    .stat-card--danger  { border-top: 3px solid var(--color-danger, #a4262c); }
    .stat-card--success { border-top: 3px solid #107c10; }
    .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--color-primary); }
    .stat-card--danger  .stat-value { color: var(--color-danger, #a4262c); }
    .stat-card--success .stat-value { color: #107c10; }
    .stat-label { font-size: 0.78rem; color: var(--color-text-muted); margin-top: 4px; }

    /* ── OneDrive summary layout ── */
    .od-summary-panel { padding: 24px; display: flex; flex-direction: column; gap: 20px; }

    /* ── Action bar ── */
    .od-action-bar { background: var(--color-surface, #fff); border: 1px solid var(--color-border);
      border-radius: 8px; padding: 16px 20px; display: flex; flex-direction: column; gap: 14px; }
    .od-btn-group  { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
    .od-btn-cluster { display: flex; gap: 8px; flex-wrap: wrap; }
    .od-btn-cluster--right { margin-left: auto; }
    .od-action-btn { display: inline-flex; align-items: center; gap: 6px; }

    /* ── Progress sections ── */
    .od-progress-section { display: flex; flex-direction: column; gap: 8px;
      padding-top: 14px; border-top: 1px solid var(--color-border); }
    .od-progress-header { display: flex; align-items: center; justify-content: space-between; }
    .od-progress-title  { font-size: 0.85rem; font-weight: 600; color: var(--color-text); }
    .od-progress-count  { font-size: 0.82rem; color: var(--color-text-muted);
      font-variant-numeric: tabular-nums; background: var(--color-surface-alt);
      padding: 2px 8px; border-radius: 10px; border: 1px solid var(--color-border); }
    .od-progress-bar-wrap { height: 10px; background: var(--color-border);
      border-radius: 5px; overflow: hidden; }
    .od-progress-bar { height: 100%; background: var(--color-primary, #0078d4);
      border-radius: 5px; transition: width 0.2s ease; min-width: 0; }
    .od-progress-bar--grant { background: #107c10; }
    .od-progress-stats { display: flex; gap: 16px; flex-wrap: wrap; }
    .pstat       { font-size: 0.78rem; font-weight: 600; }
    .pstat--ok   { color: #107c10; }
    .pstat--err  { color: #a4262c; }
    .pstat--warn { color: #7a5900; }
    .pstat--info { color: var(--color-text-muted); }

    /* ── Table ── */
    .summary-table-wrap { overflow-x: auto; border: 1px solid var(--color-border);
      border-radius: 6px; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
    .summary-table th { background: var(--color-surface-alt); padding: 10px 12px; text-align: left;
      font-weight: 600; border-bottom: 1px solid var(--color-border); white-space: nowrap; }
    .summary-table td { padding: 9px 12px; border-bottom: 1px solid var(--color-border);
      vertical-align: middle; }
    .summary-table tr:last-child td { border-bottom: none; }
    .summary-table a { color: var(--color-primary); text-decoration: none; }
    .summary-table a:hover { text-decoration: underline; }
    .path-cell { max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
      font-family: monospace; font-size: 0.8rem; }
    .path-cell--wrap { white-space: normal; overflow: visible; text-overflow: unset; word-break: break-all; }
    .od-user-cell { font-size: 0.85rem; white-space: nowrap; }
    .table-empty { text-align: center; color: var(--color-text-muted); padding: 24px !important; }

    /* ── Badges ── */
    .badge         { padding: 2px 8px; border-radius: 10px; font-size: 0.75rem; font-weight: 600;
      white-space: nowrap; }
    .status-ready   { background: #dff6dd; color: #107c10; }
    .status-pending { background: #fff4ce; color: #7d5900; }
    .status-warning { background: #fff4ce; color: #7d5900; }
    .status-error   { background: #fde7e9; color: var(--color-danger, #a4262c); }
    .badge-neutral  { background: var(--color-surface-alt, #f5f5f5); color: var(--color-text-muted);
      border: 1px solid var(--color-border); }
    .badge-manual   { background: #e8f4fd; color: #0078d4; }

    /* ── SharePoint export row ── */
    .summary-export-row { display: flex; gap: 10px; margin-bottom: 16px; }
  `
  document.head.appendChild(style)
}
