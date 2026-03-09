import { getState } from '../../state/store'
import type { MigrationMapping } from '../../types'

export function renderSummaryPanel(container: HTMLElement): void {
  const { mappings, treeData } = getState()

  if (!treeData) {
    container.innerHTML = `<div class="summary-empty"><p>No data loaded. Start by uploading a TreeSize report.</p></div>`
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
              <th>Library</th>
              <th>Path</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${mappings.length === 0
              ? `<tr><td colspan="7" class="table-empty">No mappings defined yet.</td></tr>`
              : mappings.map(rowHtml).join('')}
          </tbody>
        </table>
      </div>

      <div class="summary-export-row">
        <button id="btn-export-csv" class="btn btn-primary">Export as CSV</button>
        <button id="btn-export-json" class="btn btn-ghost">Export as JSON</button>
      </div>
    </div>
  `
  injectSummaryStyles()

  container.querySelector('#btn-export-csv')?.addEventListener('click', () => exportCsv(mappings))
  container.querySelector('#btn-export-json')?.addEventListener('click', () => exportJson(mappings))
}

function rowHtml(m: MigrationMapping): string {
  const statusClass = m.status === 'ready' ? 'status-ready' : 'status-pending'
  const statusLabel = m.status === 'ready' ? '✅ Ready' : '⏳ Pending'
  return `
    <tr>
      <td class="path-cell" title="${escHtml(m.sourceNode.path)}">${escHtml(m.sourceNode.name)}</td>
      <td>${formatBytes(m.sourceNode.sizeBytes)}</td>
      <td>${m.sourceNode.fileCount.toLocaleString()}</td>
      <td>${m.targetSite ? escHtml(m.targetSite.displayName) : '—'}</td>
      <td>${m.targetDrive ? escHtml(m.targetDrive.name) : '—'}</td>
      <td class="path-cell">${m.targetFolderPath ? escHtml(m.targetFolderPath) : '/'}</td>
      <td><span class="badge ${statusClass}">${statusLabel}</span></td>
    </tr>
  `
}

function exportCsv(mappings: MigrationMapping[]): void {
  const headers = ['Source Path', 'Size (Bytes)', 'File Count', 'Target Site', 'Target Library', 'Target Path', 'Status']
  const rows = mappings.map((m) => [
    m.sourceNode.path,
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

function exportJson(mappings: MigrationMapping[]): void {
  const data = mappings.map((m) => ({
    sourcePath: m.sourceNode.path,
    sourceName: m.sourceNode.name,
    sizeBytes: m.sourceNode.sizeBytes,
    fileCount: m.sourceNode.fileCount,
    targetSite: m.targetSite?.webUrl ?? null,
    targetSiteId: m.targetSite?.id ?? null,
    targetLibrary: m.targetDrive?.name ?? null,
    targetLibraryId: m.targetDrive?.id ?? null,
    targetFolderPath: m.targetFolderPath || '/',
    status: m.status,
  }))
  downloadFile(JSON.stringify(data, null, 2), 'migration-plan.json', 'application/json')
}

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
    .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--color-primary); }
    .stat-label { font-size: 0.78rem; color: var(--color-text-muted); margin-top: 4px; }
    .summary-warning { background: #fff4ce; color: #7d5900; padding: 10px 14px; border-radius: 4px;
      font-size: 0.88rem; margin-bottom: 16px; }
    .summary-table-wrap { overflow-x: auto; border: 1px solid var(--color-border); border-radius: 6px; margin-bottom: 20px; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
    .summary-table th { background: var(--color-surface-alt); padding: 10px 12px; text-align: left;
      font-weight: 600; border-bottom: 1px solid var(--color-border); white-space: nowrap; }
    .summary-table td { padding: 9px 12px; border-bottom: 1px solid var(--color-border); }
    .summary-table tr:last-child td { border-bottom: none; }
    .path-cell { max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
      font-family: monospace; font-size: 0.8rem; }
    .badge { padding: 2px 8px; border-radius: 10px; font-size: 0.75rem; font-weight: 600; }
    .status-ready { background: #dff6dd; color: #107c10; }
    .status-pending { background: #fff4ce; color: #7d5900; }
    .table-empty { text-align: center; color: var(--color-text-muted); padding: 24px !important; }
    .summary-export-row { display: flex; gap: 10px; }
  `
  document.head.appendChild(style)
}
