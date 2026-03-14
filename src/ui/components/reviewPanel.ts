import { getState, setState } from '../../state/store'
import { getSpConfig, updateProject } from '../../graph/projectService'
import { downloadDriveItem, getDefaultDriveWebUrl, getSharePointItemByPath } from '../../graph/graphClient'
import { buildReviewTree } from '../../parsers/migrationResultParser'
import type { MigrationResultItem, MigrationResultSummary, ReviewData, ReviewNode } from '../../types'
import type { SpDriveItemDetails } from '../../graph/graphClient'

// ─── Module-level state ───────────────────────────────────────────────────────

let _statusFilter = 'all'
let _hideRecycleBin = false
let _allItems: MigrationResultItem[] = []
let _treeRoot: ReviewNode | null = null
let _selectedNode: ReviewNode | null = null
let _expandedPaths = new Set<string>()
let _spFeedEnabled = false
let _driveWebUrl: string | null = null
let _treeEl: HTMLElement | null = null
let _rightPanel: HTMLElement | null = null

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

  if (state.reviewData) {
    renderWithData(container, state.reviewData)
    return
  }

  container.innerHTML = `
    <div class="review-panel">
      <div class="review-loading"><span class="spinner"></span> Loading migration results…</div>
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

function renderWithData(container: HTMLElement, data: ReviewData): void {
  const state = getState()
  _statusFilter = 'all'
  _hideRecycleBin = false
  _allItems = data.items
  _treeRoot = data.tree
  _selectedNode = null
  _expandedPaths = new Set()
  _spFeedEnabled = state.currentProject?.projectData.sharePointFeedEnabled ?? false
  _driveWebUrl = null

  const { totals } = data
  const pct = (n: number) => totals.total > 0 ? ` (${Math.round(n / totals.total * 100)}%)` : ''

  container.innerHTML = `
    <div class="review-panel">
      <div class="review-layout">

        <div class="review-left">
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
        </div>

        <div class="review-right" id="review-right">
          <div class="review-right-header">
            <label class="review-feed-toggle-label">
              <input type="checkbox" id="review-sp-feed-toggle" ${_spFeedEnabled ? 'checked' : ''} />
              <span>SharePoint Feed Enabled</span>
            </label>
          </div>

          <div class="review-item-panel" id="review-item-panel">
            <div class="review-item-placeholder">
              <span class="review-placeholder-arrow">←</span>
              <p>Select a file or folder</p>
            </div>
          </div>

          <div class="review-sp-section" id="review-sp-section" style="${_spFeedEnabled ? '' : 'display:none'}">
            <div class="review-sp-header">SharePoint Details</div>
            <div id="review-sp-content" class="review-sp-content">
              <div class="review-sp-placeholder">Select an item to load SharePoint details</div>
            </div>
          </div>
        </div>

      </div>
    </div>`

  _treeEl = container.querySelector('#review-tree') as HTMLElement
  _rightPanel = container.querySelector('#review-right') as HTMLElement

  renderTreeNodes(treeRootNodes(), _treeEl)
  setupReviewFilters(container)
  setupRightPanelControls(container)
}

function treeRootNodes(): ReviewNode[] {
  if (!_treeRoot) return []
  return _treeRoot.path === '' ? _treeRoot.children : [_treeRoot]
}

// ─── Tree rendering ───────────────────────────────────────────────────────────

function renderTreeNodes(nodes: ReviewNode[], container: HTMLElement): void {
  container.innerHTML = ''
  for (const node of nodes) {
    container.appendChild(createReviewNodeEl(node))
  }
}

function createReviewNodeEl(node: ReviewNode): HTMLLIElement {
  const li = document.createElement('li')
  li.className = 'review-node'

  const hasChildren = node.children.length > 0
  const isExpanded = _expandedPaths.has(node.path)
  const isSelected = _selectedNode?.path === node.path

  const row = document.createElement('div')
  row.className = [
    'review-row',
    node.failedCount > 0 ? 'review-row--has-failed' : '',
    isSelected ? 'review-row--selected' : '',
  ].filter(Boolean).join(' ')
  row.dataset.path = node.path

  const toggle = document.createElement('span')
  toggle.className = 'review-toggle'
  toggle.textContent = hasChildren ? (isExpanded ? '▼' : '▶') : ''

  const icon = document.createElement('span')
  icon.className = 'review-icon'
  icon.textContent = hasChildren ? (isExpanded ? '📂' : '📁') : '📄'

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

  row.append(toggle, icon, name, colMigrated, colFailed, colSkipped, colTotal)
  li.appendChild(row)

  if (hasChildren) {
    const childList = document.createElement('ul')
    childList.className = 'review-children'
    childList.style.display = isExpanded ? '' : 'none'
    if (isExpanded) {
      for (const child of node.children) childList.appendChild(createReviewNodeEl(child))
    }
    li.appendChild(childList)

    row.addEventListener('click', (e) => {
      e.stopPropagation()
      selectNode(node, row)
      const open = childList.style.display !== 'none'
      if (open) {
        childList.style.display = 'none'
        toggle.textContent = '▶'
        icon.textContent = '📁'
        _expandedPaths.delete(node.path)
      } else {
        if (childList.children.length === 0) {
          for (const child of node.children) childList.appendChild(createReviewNodeEl(child))
        }
        childList.style.display = ''
        toggle.textContent = '▼'
        icon.textContent = '📂'
        _expandedPaths.add(node.path)
      }
    })
  } else {
    row.addEventListener('click', (e) => { e.stopPropagation(); selectNode(node, row) })
  }

  return li
}

function selectNode(node: ReviewNode, rowEl: HTMLElement): void {
  _treeEl?.querySelectorAll('.review-row--selected').forEach((el) => el.classList.remove('review-row--selected'))
  rowEl.classList.add('review-row--selected')
  _selectedNode = node
  renderRightPanelContent()
}

// ─── Right panel content ──────────────────────────────────────────────────────

function renderRightPanelContent(): void {
  if (!_rightPanel) return

  const itemPanel = _rightPanel.querySelector('#review-item-panel') as HTMLElement
  if (!itemPanel) return

  if (!_selectedNode) {
    itemPanel.innerHTML = `<div class="review-item-placeholder"><span class="review-placeholder-arrow">←</span><p>Select a file or folder</p></div>`
    clearSpContent()
    return
  }

  const items = _allItems.filter((i) => i.sourcePath === _selectedNode!.path)
  const item = items[0] ?? null
  const isFolder = _selectedNode.children.length > 0

  itemPanel.innerHTML = isFolder || !item
    ? renderFolderCardHtml(_selectedNode)
    : renderFileCardHtml(_selectedNode, item)

  if (_spFeedEnabled) void loadSpFeed(item)
}

function clearSpContent(): void {
  const spContent = _rightPanel?.querySelector('#review-sp-content') as HTMLElement | null
  if (spContent) spContent.innerHTML = `<div class="review-sp-placeholder">Select an item to load SharePoint details</div>`
}

function renderFolderCardHtml(node: ReviewNode): string {
  const pct = (n: number) => node.totalCount > 0 ? ` (${Math.round(n / node.totalCount * 100)}%)` : ''
  return `
    <div class="review-item-card">
      <div class="review-item-title-row">
        <span class="review-item-type-icon">📁</span>
        <div class="review-item-title-text">
          <div class="review-item-name">${escHtml(node.name)}</div>
          <div class="review-item-path">${escHtml(node.path)}</div>
        </div>
      </div>
      <div class="review-item-counts">
        <div class="ric ric--migrated">
          <div class="ric-value">✓ ${node.migratedCount.toLocaleString()}</div>
          <div class="ric-label">Migrated${pct(node.migratedCount)}</div>
        </div>
        <div class="ric ric--failed">
          <div class="ric-value">✗ ${node.failedCount.toLocaleString()}</div>
          <div class="ric-label">Failed</div>
        </div>
        <div class="ric ric--skipped">
          <div class="ric-value">⊘ ${node.skippedCount.toLocaleString()}</div>
          <div class="ric-label">Skipped</div>
        </div>
        <div class="ric ric--total">
          <div class="ric-value">${node.totalCount.toLocaleString()}</div>
          <div class="ric-label">Total</div>
        </div>
      </div>
    </div>`
}

function renderFileCardHtml(node: ReviewNode, item: MigrationResultItem): string {
  const destMuted = item.status !== 'Migrated' ? ' review-detail-path--muted' : ''
  return `
    <div class="review-item-card">
      <div class="review-item-title-row">
        <span class="review-item-type-icon">📄</span>
        <div class="review-item-title-text">
          <div class="review-item-name">${escHtml(item.itemName || node.name)}</div>
          <div class="review-item-path">${escHtml(item.source)}</div>
        </div>
      </div>
      <dl class="review-detail-grid">
        <dt>Status</dt><dd>${statusBadgeHtml(item.status, item.isRecycleBin)}</dd>
        <dt>Result Category</dt><dd>${escHtml(item.resultCategory || '—')}</dd>
        ${item.message ? `<dt>Message</dt><dd class="review-detail-message">${escHtml(item.message)}</dd>` : ''}
        ${item.errorCode ? `<dt>Error Code</dt><dd class="review-detail-error">${escHtml(item.errorCode)}</dd>` : ''}
        <dt>File Size</dt><dd>${item.fileSizeBytes > 0 ? formatBytes(item.fileSizeBytes) : '—'}</dd>
        <dt>Source</dt><dd class="review-detail-path">${escHtml(item.source)}</dd>
        <dt>Destination</dt><dd class="review-detail-path${destMuted}">${escHtml(item.destination || '—')}</dd>
      </dl>
    </div>`
}

// ─── SharePoint live feed ─────────────────────────────────────────────────────

async function loadSpFeed(item: MigrationResultItem | null): Promise<void> {
  const spContent = _rightPanel?.querySelector('#review-sp-content') as HTMLElement | null
  if (!spContent) return

  if (!item?.destination) {
    spContent.innerHTML = `<div class="review-sp-placeholder">No destination URL available</div>`
    return
  }

  spContent.innerHTML = `<div class="review-sp-loading"><span class="spinner"></span> Loading from SharePoint…</div>`

  try {
    const { siteId } = getSpConfig()
    if (!_driveWebUrl) _driveWebUrl = await getDefaultDriveWebUrl(siteId)

    const relativePath = extractDriveRelativePath(item.destination, _driveWebUrl)
    if (!relativePath) {
      spContent.innerHTML = `<div class="review-sp-error">Could not resolve path from destination URL</div>`
      return
    }

    const details = await getSharePointItemByPath(siteId, relativePath)
    spContent.innerHTML = renderSpDetailsHtml(details)
  } catch (err) {
    spContent.innerHTML = `<div class="review-sp-error">Failed to load: ${escHtml((err as Error).message)}</div>`
  }
}

function extractDriveRelativePath(destination: string, driveWebUrl: string): string | null {
  try {
    const dest = decodeURIComponent(destination).replace(/\\/g, '/')
    const base = decodeURIComponent(driveWebUrl).replace(/\\/g, '/').replace(/\/$/, '')
    if (dest.toLowerCase().startsWith(base.toLowerCase())) {
      return dest.slice(base.length).replace(/^\//, '')
    }
    return null
  } catch {
    return null
  }
}

function renderSpDetailsHtml(details: SpDriveItemDetails): string {
  const title = details.listItem?.fields?.Title
  return `
    <dl class="review-detail-grid review-sp-detail-grid">
      <dt>Name</dt><dd>${escHtml(details.name)}</dd>
      ${title && title !== details.name ? `<dt>Title</dt><dd>${escHtml(title)}</dd>` : ''}
      <dt>Created By</dt><dd>${escHtml(details.createdBy?.user?.displayName ?? '—')}</dd>
      <dt>Created Date</dt><dd>${escHtml(formatDateTime(details.createdDateTime))}</dd>
      <dt>Modified By</dt><dd>${escHtml(details.lastModifiedBy?.user?.displayName ?? '—')}</dd>
      <dt>Modified Date</dt><dd>${escHtml(formatDateTime(details.lastModifiedDateTime))}</dd>
    </dl>`
}

// ─── Right panel controls ─────────────────────────────────────────────────────

function setupRightPanelControls(container: HTMLElement): void {
  container.querySelector('#review-sp-feed-toggle')?.addEventListener('change', async (e) => {
    const enabled = (e.target as HTMLInputElement).checked
    _spFeedEnabled = enabled

    const spSection = _rightPanel?.querySelector('#review-sp-section') as HTMLElement | null
    if (spSection) spSection.style.display = enabled ? '' : 'none'

    const project = getState().currentProject
    if (project) {
      const newProjectData = { ...project.projectData, sharePointFeedEnabled: enabled }
      try {
        await updateProject(project.id, { projectData: newProjectData })
        setState({ currentProject: { ...project, projectData: newProjectData } })
      } catch { /* non-critical */ }
    }

    if (enabled && _selectedNode) {
      const items = _allItems.filter((i) => i.sourcePath === _selectedNode!.path)
      void loadSpFeed(items[0] ?? null)
    } else if (!enabled) {
      clearSpContent()
    }
  })
}

// ─── Filters ──────────────────────────────────────────────────────────────────

function setupReviewFilters(container: HTMLElement): void {
  const rebuildTree = (): void => {
    if (!_treeEl) return
    const search = (container.querySelector('#review-search') as HTMLInputElement)?.value.trim().toLowerCase() ?? ''
    const filtered = filterNodes(treeRootNodes(), _statusFilter, _hideRecycleBin, search)
    renderTreeNodes(filtered, _treeEl)
  }

  container.querySelector('.review-pill-group')?.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('.review-pill')
    if (!btn) return
    container.querySelectorAll('.review-pill').forEach((p) => p.classList.remove('review-pill--active'))
    btn.classList.add('review-pill--active')
    _statusFilter = btn.dataset.filter ?? 'all'
    rebuildTree()
  })

  container.querySelector('#review-hide-rb')?.addEventListener('change', (e) => {
    _hideRecycleBin = (e.target as HTMLInputElement).checked
    rebuildTree()
  })

  container.querySelector('#review-search')?.addEventListener('input', () => rebuildTree())
}

function filterNodes(nodes: ReviewNode[], statusFilter: string, hideRb: boolean, search: string): ReviewNode[] {
  const result: ReviewNode[] = []
  for (const node of nodes) {
    const filteredChildren = filterNodes(node.children, statusFilter, hideRb, search)
    if (nodeMatchesFilters(node, statusFilter, hideRb, search) || filteredChildren.length > 0) {
      result.push({ ...node, children: filteredChildren })
    }
  }
  return result
}

function nodeMatchesFilters(node: ReviewNode, statusFilter: string, hideRb: boolean, search: string): boolean {
  if (search && !node.path.toLowerCase().includes(search) && !node.name.toLowerCase().includes(search)) return false
  if (statusFilter === 'Migrated' && node.migratedCount === 0) return false
  if (statusFilter === 'Failed' && node.failedCount === 0) return false
  if (statusFilter === 'Skipped' && node.skippedCount === 0) return false
  if (statusFilter === 'Partial' && node.partialCount === 0) return false
  if (hideRb && node.children.length === 0) {
    const items = _allItems.filter((i) => i.sourcePath === node.path)
    if (items.length > 0 && items.every((i) => i.isRecycleBin)) return false
  }
  return true
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function statusBadgeHtml(status: string, isRecycleBin: boolean): string {
  if (isRecycleBin) return `<span class="rbadge rbadge--rb">🗑️ Recycle Bin (${escHtml(status)})</span>`
  if (status === 'Migrated') return `<span class="rbadge rbadge--migrated">✓ Migrated</span>`
  if (status === 'Failed') return `<span class="rbadge rbadge--failed">✗ Failed</span>`
  if (status === 'Skipped') return `<span class="rbadge rbadge--skipped">⊘ Skipped</span>`
  if (status === 'Partial') return `<span class="rbadge rbadge--partial">◐ Partial</span>`
  return `<span class="rbadge">${escHtml(status)}</span>`
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function formatDateTime(iso: string): string {
  if (!iso) return '—'
  try {
    return new Date(iso).toLocaleString(undefined, {
      year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
    })
  } catch { return iso }
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
    .review-panel { padding: 0; display: flex; flex-direction: column;
      height: calc(100vh - 140px); overflow: hidden; }

    /* Empty / loading */
    .review-empty { display: flex; flex-direction: column; align-items: center; justify-content: center;
      padding: 80px 24px; text-align: center; }
    .review-empty-icon { font-size: 3rem; margin-bottom: 16px; }
    .review-empty-title { font-size: 1.1rem; font-weight: 600; margin-bottom: 8px; }
    .review-empty-desc { font-size: 0.875rem; color: var(--color-text-muted); max-width: 420px; line-height: 1.5; }
    .review-loading { display: flex; align-items: center; gap: 10px; padding: 40px 24px;
      font-size: 0.9rem; color: var(--color-text-muted); }

    /* Two-panel layout */
    .review-layout { flex: 1; display: grid; grid-template-columns: 2fr 1fr; overflow: hidden; min-height: 0; }
    .review-left { display: flex; flex-direction: column; overflow: hidden;
      border-right: 1px solid var(--color-border); min-height: 0; }
    .review-right { display: flex; flex-direction: column; overflow: hidden; min-width: 260px; min-height: 0; }

    /* Stats bar */
    .review-stats-bar { display: flex; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface); flex-shrink: 0; overflow-x: auto; }
    .rstat-card { flex: 1; min-width: 100px; padding: 10px 14px;
      border-right: 1px solid var(--color-border); }
    .rstat-card:last-child { border-right: none; }
    .rstat-card--danger { border-left: 3px solid var(--color-danger); }
    .rstat-card--skipped { border-left: 3px solid #f3c00a; }
    .rstat-label { font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
    .rstat-value { font-size: 1.1rem; font-weight: 700; line-height: 1.1; }
    .rstat-sub { font-size: 0.68rem; color: var(--color-text-muted); margin-top: 2px; }
    .rstat-blue { color: var(--color-primary); }
    .rstat-green { color: #107c10; }
    .rstat-red { color: var(--color-danger, #a4262c); }
    .rstat-amber { color: #7d4200; }

    /* Filter bar */
    .review-filter-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 8px;
      padding: 10px 16px; background: var(--color-surface);
      border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .review-search-wrap { flex: 1; min-width: 160px; max-width: 260px; }
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
    .review-tree { flex: 1; overflow-y: auto; list-style: none; padding: 0; margin: 0; min-height: 0; }
    .review-node { list-style: none; }
    .review-children { list-style: none; padding: 0; margin: 0 0 0 20px;
      border-left: 1px solid var(--color-border); }
    .review-row { display: flex; align-items: center; padding: 5px 8px; cursor: pointer;
      transition: background 0.1s; min-height: 32px; }
    .review-row:hover { background: #f0f6ff; }
    .review-row--selected { background: var(--color-primary-light, #deecf9) !important; }
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

    /* Right panel */
    .review-right-header { padding: 12px 16px; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface-alt); flex-shrink: 0; }
    .review-feed-toggle-label { display: flex; align-items: center; gap: 8px;
      font-size: 0.85rem; cursor: pointer; user-select: none; }

    /* Item panel */
    .review-item-panel { flex: 1; overflow-y: auto; border-bottom: 1px solid var(--color-border); min-height: 0; }
    .review-item-placeholder { display: flex; flex-direction: column; align-items: center;
      justify-content: center; height: 100%; padding: 32px 16px; text-align: center;
      color: var(--color-text-muted); font-size: 0.875rem; gap: 6px; }
    .review-placeholder-arrow { font-size: 1.5rem; }

    /* Item card */
    .review-item-card { padding: 16px; }
    .review-item-title-row { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 16px; }
    .review-item-type-icon { font-size: 1.8rem; flex-shrink: 0; margin-top: 2px; }
    .review-item-title-text { min-width: 0; }
    .review-item-name { font-size: 0.9rem; font-weight: 600; font-family: 'Consolas', monospace;
      word-break: break-all; margin-bottom: 3px; }
    .review-item-path { font-size: 0.72rem; color: var(--color-text-muted);
      font-family: 'Consolas', monospace; word-break: break-all; }

    /* Folder count cards */
    .review-item-counts { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .ric { padding: 10px 12px; border-radius: 6px; background: var(--color-surface-alt); text-align: center; }
    .ric-value { font-size: 1rem; font-weight: 700; margin-bottom: 2px; }
    .ric-label { font-size: 0.68rem; color: var(--color-text-muted); text-transform: uppercase; letter-spacing: 0.04em; }
    .ric--migrated .ric-value { color: #107c10; }
    .ric--failed .ric-value { color: var(--color-danger); }
    .ric--skipped .ric-value { color: #605e5c; }
    .ric--total .ric-value { color: var(--color-primary); }

    /* Detail grid */
    .review-detail-grid { display: grid; grid-template-columns: 110px 1fr; gap: 6px 12px; margin: 0; padding: 0; }
    .review-detail-grid dt { color: var(--color-text-muted); font-size: 0.78rem; font-weight: 600;
      align-self: start; padding-top: 2px; }
    .review-detail-grid dd { margin: 0; font-size: 0.82rem; word-break: break-all; }
    .review-detail-path { font-family: 'Consolas', monospace; font-size: 0.72rem; color: var(--color-text); }
    .review-detail-path--muted { color: var(--color-text-muted); }
    .review-detail-message { color: var(--color-text); }
    .review-detail-error { color: var(--color-danger); font-family: 'Consolas', monospace; font-size: 0.72rem; }

    /* Status badges */
    .rbadge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.78rem; font-weight: 600; }
    .rbadge--migrated { background: rgba(16,124,16,0.12); color: #107c10; }
    .rbadge--failed { background: rgba(209,52,56,0.12); color: var(--color-danger); }
    .rbadge--skipped { background: var(--color-surface-alt); color: var(--color-text-muted); }
    .rbadge--partial { background: rgba(255,140,0,0.12); color: #7d4200; }
    .rbadge--rb { background: rgba(243,192,10,0.15); color: #7d5900; }

    /* SharePoint feed */
    .review-sp-section { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-height: 0; }
    .review-sp-header { padding: 8px 16px; font-size: 0.72rem; font-weight: 700;
      text-transform: uppercase; letter-spacing: 0.05em; color: var(--color-text-muted);
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .review-sp-content { flex: 1; overflow-y: auto; min-height: 0; }
    .review-sp-placeholder { padding: 24px 16px; text-align: center;
      color: var(--color-text-muted); font-size: 0.85rem; }
    .review-sp-loading { display: flex; align-items: center; gap: 8px; padding: 16px;
      color: var(--color-text-muted); font-size: 0.85rem; }
    .review-sp-error { padding: 12px 16px; color: var(--color-danger); font-size: 0.82rem; }
    .review-sp-detail-grid { padding: 16px; }

    /* Spinner */
    .review-panel .spinner { display: inline-block; width: 14px; height: 14px;
      border: 2px solid currentColor; border-top-color: transparent; border-radius: 50%;
      animation: review-spin 0.7s linear infinite; flex-shrink: 0; }
    @keyframes review-spin { to { transform: rotate(360deg); } }
  `
  document.head.appendChild(style)
}
