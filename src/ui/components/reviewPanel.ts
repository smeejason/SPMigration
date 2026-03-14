import { getState, setState } from '../../state/store'
import { getSpConfig } from '../../graph/projectService'
import { downloadDriveItem } from '../../graph/graphClient'
import { buildReviewTree } from '../../parsers/migrationResultParser'
import type { MigrationResultItem, MigrationResultSummary, ReviewData, ReviewNode } from '../../types'

// ─── Entry point ──────────────────────────────────────────────────────────────

export async function renderReviewPanel(container: HTMLElement): Promise<void> {
  injectReviewStyles()
  const state = getState()
  const project = state.currentProject
  if (!project) return

  const resultUploads = project.projectData.resultUploads ?? []

  if (resultUploads.length === 0) {
    container.innerHTML = `
      <div class="review-panel">
        <div class="review-empty">
          <div class="review-empty-icon">📦</div>
          <p class="review-empty-title">No migration results uploaded yet</p>
          <p class="review-empty-desc">Upload SPMT result ZIP files on the <strong>Upload</strong> tab. All uploaded ZIPs will be combined here into a single review tree.</p>
        </div>
      </div>`
    return
  }

  // Use cached data if available
  if (state.reviewData) {
    renderWithData(container, state.reviewData)
    return
  }

  // Show loading spinner while fetching
  container.innerHTML = `
    <div class="review-panel">
      <div class="review-loading">
        <span class="spinner"></span>
        Loading migration results…
      </div>
    </div>`

  try {
    const { siteId } = getSpConfig()
    const downloads = await Promise.all(
      resultUploads.map((u) => downloadDriveItem(siteId, u.summaryItemId))
    ) as MigrationResultSummary[]

    const combinedItems: MigrationResultItem[] = downloads.flatMap((d) => d.items ?? [])

    const tree = buildReviewTree(combinedItems)

    const totals = {
      migrated: downloads.reduce((s, d) => s + d.migratedCount, 0),
      failed: downloads.reduce((s, d) => s + d.failedCount, 0),
      skipped: downloads.reduce((s, d) => s + d.skippedCount, 0),
      partial: downloads.reduce((s, d) => s + d.partialCount, 0),
      total: downloads.reduce((s, d) => s + d.totalCount, 0),
      failedRecycleBin: combinedItems.filter((i) => i.status === 'Failed' && i.isRecycleBin).length,
      skippedRecycleBin: combinedItems.filter((i) => i.status === 'Skipped' && i.isRecycleBin).length,
    }

    const reviewData: ReviewData = { tree, items: combinedItems, totals }
    setState({ reviewData })
    renderWithData(container, reviewData)
  } catch (err) {
    container.innerHTML = `
      <div class="review-panel">
        <div class="upload-status upload-status--error" style="margin:24px">
          Failed to load migration results: ${escHtml((err as Error).message)}
        </div>
      </div>`
  }
}

// ─── Main render ──────────────────────────────────────────────────────────────

// Module-level state for filtering (reset on each call to renderWithData)
let _statusFilter: string = 'all'
let _hideRecycleBin = false
let _allItems: MigrationResultItem[] = []
let _treeRoot: ReviewNode | null = null

function renderWithData(container: HTMLElement, data: ReviewData): void {
  _statusFilter = 'all'
  _hideRecycleBin = false
  _allItems = data.items
  _treeRoot = data.tree

  const { totals } = data
  const pct = (n: number) => totals.total > 0 ? ` (${Math.round(n / totals.total * 100)}%)` : ''

  container.innerHTML = `
    <div class="review-panel">

      <div class="review-stats-bar">
        <div class="rstat-card">
          <div class="rstat-label">TOTAL ITEMS</div>
          <div class="rstat-value rstat-blue">${totals.total.toLocaleString()}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">MIGRATED</div>
          <div class="rstat-value rstat-green">${totals.migrated.toLocaleString()}${pct(totals.migrated)}</div>
        </div>
        <div class="rstat-card rstat-card--danger">
          <div class="rstat-label">FAILED</div>
          <div class="rstat-value rstat-red">${totals.failed.toLocaleString()}</div>
          <div class="rstat-sub">${(totals.failed - totals.failedRecycleBin).toLocaleString()} excl. recycle bin</div>
        </div>
        <div class="rstat-card rstat-card--skipped">
          <div class="rstat-label">SKIPPED</div>
          <div class="rstat-value rstat-amber">${totals.skipped.toLocaleString()}</div>
          <div class="rstat-sub">${(totals.skipped - totals.skippedRecycleBin).toLocaleString()} excl. recycle bin</div>
        </div>
        ${totals.partial > 0 ? `
        <div class="rstat-card">
          <div class="rstat-label">PARTIAL</div>
          <div class="rstat-value rstat-amber">${totals.partial.toLocaleString()}</div>
        </div>` : ''}
      </div>

      <div class="review-filter-bar">
        <div class="review-search-wrap">
          <input type="text" id="review-search" class="form-input review-search-input" placeholder="Search by path or name…" />
        </div>
        <div class="review-pill-group">
          <button class="review-pill review-pill--active" data-filter="all">All</button>
          <button class="review-pill review-pill--migrated" data-filter="Migrated">✓ Migrated</button>
          <button class="review-pill review-pill--failed" data-filter="Failed">✗ Failed</button>
          <button class="review-pill review-pill--skipped" data-filter="Skipped">⊘ Skipped</button>
          ${totals.partial > 0 ? '<button class="review-pill review-pill--partial" data-filter="Partial">◐ Partial</button>' : ''}
        </div>
        <label class="review-rb-label">
          <input type="checkbox" id="review-hide-rb" /> Hide Recycle Bin
        </label>
      </div>

      <div class="review-col-header">
        <span class="rch-name">PATH</span>
        <span class="rch-stat rch-migrated">MIGRATED</span>
        <span class="rch-stat rch-failed">FAILED</span>
        <span class="rch-stat rch-skipped">SKIPPED</span>
        <span class="rch-stat">TOTAL</span>
      </div>

      <ul class="review-tree" id="review-tree"></ul>

    </div>`

  const treeEl = container.querySelector('#review-tree') as HTMLElement
  renderTreeNodes(treeRoot(), treeEl, new Set())

  setupReviewFilters(container)
}

function treeRoot(): ReviewNode[] {
  if (!_treeRoot) return []
  return _treeRoot.path === '' ? _treeRoot.children : [_treeRoot]
}

// ─── Tree rendering ───────────────────────────────────────────────────────────

function renderTreeNodes(nodes: ReviewNode[], container: HTMLElement, expandedPaths: Set<string>): void {
  container.innerHTML = ''
  for (const node of nodes) {
    container.appendChild(createReviewNodeEl(node, expandedPaths))
  }
}

function createReviewNodeEl(node: ReviewNode, expandedPaths: Set<string>): HTMLLIElement {
  const li = document.createElement('li')
  li.className = 'review-node'

  const hasChildren = node.children.length > 0
  const isExpanded = expandedPaths.has(node.path)

  const row = document.createElement('div')
  row.className = `review-row${node.failedCount > 0 ? ' review-row--has-failed' : ''}`
  row.dataset.path = node.path

  const toggle = document.createElement('span')
  toggle.className = 'review-toggle'
  toggle.textContent = hasChildren ? (isExpanded ? '▼' : '▶') : ''

  const icon = document.createElement('span')
  icon.className = 'review-icon'
  icon.textContent = hasChildren ? '📁' : '📄'

  const name = document.createElement('span')
  name.className = 'review-name'
  name.textContent = node.name
  name.title = node.path

  const colMigrated = document.createElement('span')
  colMigrated.className = 'rstat rstat-migrated'
  colMigrated.textContent = node.migratedCount > 0 ? `✓ ${node.migratedCount.toLocaleString()}` : '—'

  const colFailed = document.createElement('span')
  colFailed.className = 'rstat rstat-failed'
  colFailed.textContent = node.failedCount > 0 ? `✗ ${node.failedCount.toLocaleString()}` : '—'

  const colSkipped = document.createElement('span')
  colSkipped.className = 'rstat rstat-skipped'
  colSkipped.textContent = node.skippedCount > 0 ? `⊘ ${node.skippedCount.toLocaleString()}` : '—'

  const colTotal = document.createElement('span')
  colTotal.className = 'rstat rstat-total'
  colTotal.textContent = node.totalCount.toLocaleString()

  row.appendChild(toggle)
  row.appendChild(icon)
  row.appendChild(name)
  row.appendChild(colMigrated)
  row.appendChild(colFailed)
  row.appendChild(colSkipped)
  row.appendChild(colTotal)

  li.appendChild(row)

  // Child list (rendered immediately if expanded)
  if (hasChildren) {
    const childList = document.createElement('ul')
    childList.className = 'review-children'
    childList.style.display = isExpanded ? '' : 'none'
    if (isExpanded) {
      for (const child of node.children) {
        childList.appendChild(createReviewNodeEl(child, expandedPaths))
      }
    }
    li.appendChild(childList)

    row.addEventListener('click', (e) => {
      e.stopPropagation()
      const open = childList.style.display !== 'none'
      if (open) {
        childList.style.display = 'none'
        toggle.textContent = '▶'
        icon.textContent = '📁'
        expandedPaths.delete(node.path)
      } else {
        if (childList.children.length === 0) {
          for (const child of node.children) {
            childList.appendChild(createReviewNodeEl(child, expandedPaths))
          }
        }
        childList.style.display = ''
        toggle.textContent = '▼'
        icon.textContent = '📂'
        expandedPaths.add(node.path)
      }
    })
  } else {
    // File node — click to show detail panel
    row.style.cursor = 'pointer'
    row.addEventListener('click', (e) => {
      e.stopPropagation()
      toggleItemDetail(li, node.path)
    })
  }

  return li
}

function toggleItemDetail(li: HTMLLIElement, path: string): void {
  const existing = li.querySelector('.review-detail-panel')
  if (existing) { existing.remove(); return }

  const items = _allItems.filter((i) => i.sourcePath === path)
  if (items.length === 0) return

  const item = items[0]
  const panel = document.createElement('div')
  panel.className = 'review-detail-panel'
  panel.innerHTML = `
    <dl class="review-detail-grid">
      <dt>Status</dt><dd>${statusBadgeHtml(item.status, item.isRecycleBin)}</dd>
      <dt>Result Category</dt><dd>${escHtml(item.resultCategory || '—')}</dd>
      ${item.message ? `<dt>Message</dt><dd class="review-detail-message">${escHtml(item.message)}</dd>` : ''}
      ${item.errorCode ? `<dt>Error Code</dt><dd class="review-detail-error">${escHtml(item.errorCode)}</dd>` : ''}
      <dt>File Size</dt><dd>${item.fileSizeBytes > 0 ? formatBytes(item.fileSizeBytes) : '—'}</dd>
      <dt>Source</dt><dd class="review-detail-path">${escHtml(item.source)}</dd>
      <dt>Destination</dt><dd class="review-detail-path${item.status !== 'Migrated' ? ' review-detail-path--muted' : ''}">${escHtml(item.destination)}</dd>
    </dl>`
  li.appendChild(panel)
}

function statusBadgeHtml(status: string, isRecycleBin: boolean): string {
  if (isRecycleBin) return `<span class="rbadge rbadge--rb">🗑️ Recycle Bin (${escHtml(status)})</span>`
  if (status === 'Migrated') return `<span class="rbadge rbadge--migrated">✓ Migrated</span>`
  if (status === 'Failed') return `<span class="rbadge rbadge--failed">✗ Failed</span>`
  if (status === 'Skipped') return `<span class="rbadge rbadge--skipped">⊘ Skipped</span>`
  if (status === 'Partial') return `<span class="rbadge rbadge--partial">◐ Partial</span>`
  return `<span class="rbadge">${escHtml(status)}</span>`
}

// ─── Filters ──────────────────────────────────────────────────────────────────

function setupReviewFilters(container: HTMLElement): void {
  const expandedPaths = new Set<string>()

  const rebuildTree = (): void => {
    const treeEl = container.querySelector('#review-tree') as HTMLElement
    const search = (container.querySelector('#review-search') as HTMLInputElement)?.value.trim().toLowerCase() ?? ''
    const nodes = treeRoot()
    const filtered = filterNodes(nodes, _statusFilter, _hideRecycleBin, search)
    renderTreeNodes(filtered, treeEl, expandedPaths)
  }

  // Status pills
  container.querySelector('.review-pill-group')?.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('.review-pill')
    if (!btn) return
    container.querySelectorAll('.review-pill').forEach((p) => p.classList.remove('review-pill--active'))
    btn.classList.add('review-pill--active')
    _statusFilter = btn.dataset.filter ?? 'all'
    rebuildTree()
  })

  // Hide recycle bin checkbox
  container.querySelector('#review-hide-rb')?.addEventListener('change', (e) => {
    _hideRecycleBin = (e.target as HTMLInputElement).checked
    rebuildTree()
  })

  // Search
  container.querySelector('#review-search')?.addEventListener('input', () => rebuildTree())
}

function filterNodes(
  nodes: ReviewNode[],
  statusFilter: string,
  hideRb: boolean,
  search: string
): ReviewNode[] {
  const result: ReviewNode[] = []
  for (const node of nodes) {
    const filteredChildren = filterNodes(node.children, statusFilter, hideRb, search)

    const nodeMatches = nodeMatchesFilters(node, statusFilter, hideRb, search)
    const hasMatchingChildren = filteredChildren.length > 0

    if (nodeMatches || hasMatchingChildren) {
      result.push({ ...node, children: filteredChildren })
    }
  }
  return result
}

function nodeMatchesFilters(node: ReviewNode, statusFilter: string, hideRb: boolean, search: string): boolean {
  // Search filter
  if (search && !node.path.toLowerCase().includes(search) && !node.name.toLowerCase().includes(search)) {
    return false
  }

  // Status filter — check if the subtree has matching counts
  if (statusFilter === 'Migrated' && node.migratedCount === 0) return false
  if (statusFilter === 'Failed' && node.failedCount === 0) return false
  if (statusFilter === 'Skipped' && node.skippedCount === 0) return false
  if (statusFilter === 'Partial' && node.partialCount === 0) return false

  // Hide recycle bin — only affects leaf nodes (files at that path)
  if (hideRb && node.children.length === 0) {
    const items = _allItems.filter((i) => i.sourcePath === node.path)
    if (items.length > 0 && items.every((i) => i.isRecycleBin)) return false
  }

  return true
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function escHtml(s: string): string {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectReviewStyles(): void {
  if (document.getElementById('review-styles')) return
  const style = document.createElement('style')
  style.id = 'review-styles'
  style.textContent = `
    .review-panel { padding: 0; display: flex; flex-direction: column; height: 100%; }

    /* Empty / loading states */
    .review-empty { display: flex; flex-direction: column; align-items: center; justify-content: center;
      padding: 80px 24px; text-align: center; }
    .review-empty-icon { font-size: 3rem; margin-bottom: 16px; }
    .review-empty-title { font-size: 1.1rem; font-weight: 600; color: var(--color-text); margin-bottom: 8px; }
    .review-empty-desc { font-size: 0.875rem; color: var(--color-text-muted); max-width: 420px; line-height: 1.5; }
    .review-loading { display: flex; align-items: center; gap: 10px; padding: 40px 24px;
      font-size: 0.9rem; color: var(--color-text-muted); }

    /* Stats bar */
    .review-stats-bar { display: flex; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface); flex-shrink: 0; overflow-x: auto; }
    .rstat-card { flex: 1; min-width: 110px; padding: 10px 14px;
      border-right: 1px solid var(--color-border); }
    .rstat-card:last-child { border-right: none; }
    .rstat-card--danger { border-left: 3px solid var(--color-danger); }
    .rstat-card--skipped { border-left: 3px solid #f3c00a; }
    .rstat-label { font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
    .rstat-value { font-size: 1.2rem; font-weight: 700; line-height: 1.1; }
    .rstat-sub { font-size: 0.68rem; color: var(--color-text-muted); margin-top: 2px; }
    .rstat-blue { color: var(--color-primary); }
    .rstat-green { color: #107c10; }
    .rstat-red { color: var(--color-danger, #a4262c); }
    .rstat-amber { color: #7d4200; }

    /* Filter bar */
    .review-filter-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 8px;
      padding: 10px 16px; background: var(--color-surface);
      border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .review-search-wrap { flex: 1; min-width: 200px; max-width: 320px; }
    .review-search-input { padding: 5px 10px; font-size: 0.85rem; }
    .review-pill-group { display: flex; gap: 4px; flex-wrap: wrap; }
    .review-pill { padding: 4px 12px; border-radius: 20px; border: 1px solid var(--color-border);
      background: white; font-size: 0.8rem; cursor: pointer; font-family: inherit;
      transition: background 0.12s, border-color 0.12s; }
    .review-pill:hover { background: var(--color-surface-alt); }
    .review-pill.review-pill--active { background: var(--color-primary); color: white;
      border-color: var(--color-primary); }
    .review-rb-label { display: flex; align-items: center; gap: 6px; font-size: 0.82rem;
      color: var(--color-text-muted); cursor: pointer; white-space: nowrap; }

    /* Column header */
    .review-col-header { display: flex; align-items: center; padding: 5px 16px 5px 0;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border);
      font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.05em; flex-shrink: 0; }
    .rch-name { flex: 1; padding-left: 64px; }
    .rch-stat { width: 80px; text-align: right; padding-right: 12px; flex-shrink: 0; }
    .rch-migrated { color: #107c10; }
    .rch-failed { color: var(--color-danger); }
    .rch-skipped { color: #605e5c; }

    /* Tree */
    .review-tree { flex: 1; overflow-y: auto; list-style: none; padding: 0; margin: 0; }
    .review-node { list-style: none; }
    .review-children { list-style: none; padding: 0; margin: 0 0 0 20px;
      border-left: 1px solid var(--color-border); }
    .review-row { display: flex; align-items: center; padding: 5px 8px; cursor: pointer;
      transition: background 0.1s; min-height: 32px; }
    .review-row:hover { background: #f0f6ff; }
    .review-row--has-failed { border-left: 3px solid var(--color-danger);
      background: rgba(209,52,56,0.04); }
    .review-row--has-failed:hover { background: rgba(209,52,56,0.09); }
    .review-toggle { width: 16px; font-size: 0.6rem; color: var(--color-text-muted); flex-shrink: 0; }
    .review-icon { width: 22px; text-align: center; flex-shrink: 0; margin-right: 4px; }
    .review-name { flex: 1; min-width: 0; font-size: 0.85rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .rstat { width: 80px; text-align: right; padding-right: 12px; font-size: 0.78rem;
      flex-shrink: 0; white-space: nowrap; }
    .rstat-migrated { color: #107c10; font-weight: 600; }
    .rstat-failed { color: var(--color-danger); font-weight: 600; }
    .rstat-skipped { color: var(--color-text-muted); }
    .rstat-total { color: var(--color-text-muted); }

    /* Item detail panel */
    .review-detail-panel { margin: 0 0 4px 42px; padding: 10px 14px;
      background: var(--color-surface-alt); border-left: 3px solid var(--color-border);
      border-radius: 0 4px 4px 0; font-size: 0.8rem; }
    .review-detail-grid { display: grid; grid-template-columns: 120px 1fr; gap: 4px 12px; }
    .review-detail-grid dt { color: var(--color-text-muted); font-weight: 600; align-self: start; padding-top: 1px; }
    .review-detail-grid dd { margin: 0; word-break: break-all; }
    .review-detail-path { font-family: 'Consolas', monospace; font-size: 0.75rem; color: var(--color-text); }
    .review-detail-path--muted { color: var(--color-text-muted); }
    .review-detail-message { color: var(--color-text); }
    .review-detail-error { color: var(--color-danger); font-family: 'Consolas', monospace; font-size: 0.75rem; }

    /* Status badges */
    .rbadge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.78rem; font-weight: 600; }
    .rbadge--migrated { background: rgba(16,124,16,0.12); color: #107c10; }
    .rbadge--failed { background: rgba(209,52,56,0.12); color: var(--color-danger); }
    .rbadge--skipped { background: var(--color-surface-alt); color: var(--color-text-muted); }
    .rbadge--partial { background: rgba(255,140,0,0.12); color: #7d4200; }
    .rbadge--rb { background: rgba(243,192,10,0.15); color: #7d5900; }

    /* Spinner (reuse from upload-styles if injected, define here for safety) */
    .review-panel .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid currentColor;
      border-top-color: transparent; border-radius: 50%; animation: review-spin 0.7s linear infinite; flex-shrink: 0; }
    @keyframes review-spin { to { transform: rotate(360deg); } }
  `
  document.head.appendChild(style)
}
