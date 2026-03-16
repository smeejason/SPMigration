import { getState } from '../../state/store'
import type { MigrationMapping } from '../../types'

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
  const odMappings = mappings.filter((m) => m.matchStatus !== undefined)

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

  const matched = odMappings.filter((m) => m.matchStatus === 'matched')
  const ready = matched.filter((m) => m.accessStatus === 'accessible' || m.accessStatus === 'granted')
  const unresolved = odMappings.filter((m) => m.matchStatus !== 'matched' && m.matchStatus !== 'pending')
  const totalSize = odMappings.reduce((s, m) => s + m.sourceNode.sizeBytes, 0)
  const totalFiles = odMappings.reduce((s, m) => s + m.sourceNode.fileCount, 0)

  container.innerHTML = `
    <div class="summary-panel">
      <div class="summary-stats-row">
        <div class="stat-card">
          <div class="stat-value">${odMappings.length}</div>
          <div class="stat-label">Total users</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${matched.length}</div>
          <div class="stat-label">Matched</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${ready.length}</div>
          <div class="stat-label">Access confirmed</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${formatBytes(totalSize)}</div>
          <div class="stat-label">Total data</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${totalFiles.toLocaleString()}</div>
          <div class="stat-label">Files</div>
        </div>
        ${unresolved.length > 0 ? `
        <div class="stat-card stat-card--danger">
          <div class="stat-value">${unresolved.length}</div>
          <div class="stat-label">Unresolved</div>
        </div>` : ''}
      </div>

      <div class="summary-table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Source Path</th>
              <th>Destination Site URL</th>
              <th>Destination List</th>
              <th>Folder Path</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${odMappings.map(odRowHtml).join('')}
          </tbody>
        </table>
      </div>

      <div class="summary-export-row">
        <button id="btn-export-csv" class="btn btn-primary">Export as CSV</button>
        <button id="btn-export-json" class="btn btn-ghost">Export as JSON</button>
      </div>
    </div>`

  injectSummaryStyles()
  container.querySelector('#btn-export-csv')?.addEventListener('click', () => exportOneDriveCsv(odMappings))
  container.querySelector('#btn-export-json')?.addEventListener('click', () => exportOneDriveJson(odMappings))
}

function odRowHtml(m: MigrationMapping): string {
  const siteUrl = m.targetSite?.webUrl ?? ''
  const folderPath = m.targetFolderPath || '/'
  return `
    <tr>
      <td class="path-cell" title="${escHtml(m.sourceNode.originalPath)}">${escHtml(m.sourceNode.originalPath)}</td>
      <td class="path-cell">${siteUrl
        ? `<a href="${escHtml(siteUrl)}" target="_blank" rel="noopener" title="${escHtml(siteUrl)}">${escHtml(siteUrl)}</a>`
        : '—'}</td>
      <td>Documents</td>
      <td class="path-cell">${escHtml(folderPath)}</td>
      <td>${odStatusBadge(m)}</td>
    </tr>`
}

function odStatusBadge(m: MigrationMapping): string {
  if (m.matchStatus === 'matched') {
    if (m.accessStatus === 'accessible' || m.accessStatus === 'granted') {
      return `<span class="badge status-ready">✅ Ready</span>`
    }
    if (m.accessStatus === 'no-access' || m.accessStatus === 'no-drive' || m.accessStatus === 'error') {
      return `<span class="badge status-error">✗ No access</span>`
    }
    return `<span class="badge status-warning">⚠ Matched</span>`
  }
  if (m.matchStatus === 'not-found') return `<span class="badge status-error">✗ Not found</span>`
  if (m.matchStatus === 'ambiguous') return `<span class="badge status-warning">? Ambiguous</span>`
  if (m.matchStatus === 'error') return `<span class="badge status-error">✗ Error</span>`
  return `<span class="badge status-pending">⏳ Pending</span>`
}

function exportOneDriveCsv(mappings: MigrationMapping[]): void {
  const headers = ['Source Path', 'Destination Site URL', 'Destination List', 'Folder Path', 'Match Status', 'Access Status']
  const rows = mappings.map((m) => [
    m.sourceNode.originalPath,
    m.targetSite?.webUrl ?? '',
    'Documents',
    m.targetFolderPath || '/',
    m.matchStatus ?? '',
    m.accessStatus ?? '',
  ])
  const csv = [headers, ...rows].map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n')
  downloadFile(csv, 'onedrive-migration-plan.csv', 'text/csv')
}

function exportOneDriveJson(mappings: MigrationMapping[]): void {
  const data = mappings.map((m) => ({
    sourcePath: m.sourceNode.originalPath,
    sourceName: m.sourceNode.name,
    sizeBytes: m.sourceNode.sizeBytes,
    fileCount: m.sourceNode.fileCount,
    destinationSiteUrl: m.targetSite?.webUrl ?? null,
    destinationList: 'Documents',
    folderPath: m.targetFolderPath || '/',
    matchStatus: m.matchStatus,
    accessStatus: m.accessStatus,
  }))
  downloadFile(JSON.stringify(data, null, 2), 'onedrive-migration-plan.json', 'application/json')
}

// ─── SharePoint Summary ───────────────────────────────────────────────────────

function renderSharePointSummary(container: HTMLElement, mappings: MigrationMapping[], hasTree: boolean): void {
  if (!hasTree) {
    container.innerHTML = `<div class="summary-empty"><p>No data loaded. Start by uploading a TreeSize report.</p></div>`
    injectSummaryStyles()
    return
  }

  const ready = mappings.filter((m) => m.status === 'ready')
  const unmapped = mappings.filter((m) => m.status === 'pending')
  const totalSize = ready.reduce((s, m) => s + m.sourceNode.sizeBytes, 0)
  const totalFiles = ready.reduce((s, m) => s + m.sourceNode.fileCount, 0)
  const uniqueSites = new Set(ready.map((m) => m.targetSite?.id).filter(Boolean)).size

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

      <div class="summary-export-row">
        <button id="btn-export-csv" class="btn btn-primary">Export as CSV</button>
        <button id="btn-export-json" class="btn btn-ghost">Export as JSON</button>
      </div>
    </div>`

  injectSummaryStyles()
  container.querySelector('#btn-export-csv')?.addEventListener('click', () => exportSpCsv(mappings))
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
      <td>${m.targetSite ? escHtml(m.targetSite.displayName) : '—'}</td>
      <td>${m.targetDrive ? escHtml(m.targetDrive.name) : '—'}</td>
      <td class="path-cell">${m.targetFolderPath ? escHtml(m.targetFolderPath) : '/'}</td>
      <td><span class="badge ${statusClass}">${statusLabel}</span></td>
    </tr>`
}

function exportSpCsv(mappings: MigrationMapping[]): void {
  const headers = ['Source Path', 'Size (Bytes)', 'File Count', 'Target Site', 'Destination List', 'Folder Path', 'Status']
  const rows = mappings.map((m) => [
    m.sourceNode.originalPath,
    m.sourceNode.sizeBytes,
    m.sourceNode.fileCount,
    m.targetSite?.displayName ?? '',
    m.targetDrive?.name ?? '',
    m.targetFolderPath || '/',
    m.status,
  ])
  const csv = [headers, ...rows].map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n')
  downloadFile(csv, 'migration-plan.csv', 'text/csv')
}

function exportSpJson(mappings: MigrationMapping[]): void {
  const data = mappings.map((m) => ({
    sourcePath: m.sourceNode.originalPath,
    sourceName: m.sourceNode.name,
    sizeBytes: m.sourceNode.sizeBytes,
    fileCount: m.sourceNode.fileCount,
    targetSite: m.targetSite?.webUrl ?? null,
    targetSiteId: m.targetSite?.id ?? null,
    destinationList: m.targetDrive?.name ?? null,
    destinationListId: m.targetDrive?.id ?? null,
    folderPath: m.targetFolderPath || '/',
    status: m.status,
  }))
  downloadFile(JSON.stringify(data, null, 2), 'migration-plan.json', 'application/json')
}

// ─── Shared helpers ───────────────────────────────────────────────────────────

function downloadFile(content: string, filename: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
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

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectSummaryStyles(): void {
  if (document.getElementById('summary-styles')) return
  const style = document.createElement('style')
  style.id = 'summary-styles'
  style.textContent = `
    .summary-panel { padding: 24px; }
    .summary-empty { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .summary-stats-row { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
    .stat-card { background: white; border: 1px solid var(--color-border); border-radius: 8px;
      padding: 16px 20px; flex: 1; min-width: 120px; }
    .stat-card--danger { border-top: 3px solid var(--color-danger); }
    .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--color-primary); }
    .stat-card--danger .stat-value { color: var(--color-danger); }
    .stat-label { font-size: 0.78rem; color: var(--color-text-muted); margin-top: 4px; }
    .summary-warning { background: #fff4ce; color: #7d5900; padding: 10px 14px; border-radius: 4px;
      font-size: 0.88rem; margin-bottom: 16px; }
    .summary-table-wrap { overflow-x: auto; border: 1px solid var(--color-border);
      border-radius: 6px; margin-bottom: 20px; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
    .summary-table th { background: var(--color-surface-alt); padding: 10px 12px; text-align: left;
      font-weight: 600; border-bottom: 1px solid var(--color-border); white-space: nowrap; }
    .summary-table td { padding: 9px 12px; border-bottom: 1px solid var(--color-border); }
    .summary-table tr:last-child td { border-bottom: none; }
    .summary-table a { color: var(--color-primary); text-decoration: none; }
    .summary-table a:hover { text-decoration: underline; }
    .path-cell { max-width: 240px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
      font-family: monospace; font-size: 0.8rem; }
    .badge { padding: 2px 8px; border-radius: 10px; font-size: 0.75rem; font-weight: 600; }
    .status-ready { background: #dff6dd; color: #107c10; }
    .status-pending { background: #fff4ce; color: #7d5900; }
    .status-warning { background: #fff4ce; color: #7d5900; }
    .status-error { background: #fde7e9; color: var(--color-danger, #a4262c); }
    .table-empty { text-align: center; color: var(--color-text-muted); padding: 24px !important; }
    .summary-export-row { display: flex; gap: 10px; }
  `
  document.head.appendChild(style)
}
