import { getState, setState } from '../../state/store'
import { persistProjectMappings, getSpConfig, updateProject } from '../../graph/projectService'
import { renderPersonCard, accessStatusBadge } from './oneDrivePersonCard'
import { downloadDriveItem, resolveSharePointItemByUrl, resolveDriveItemRef, resolveOneDriveFolderByPath, listDriveItemsRecursive } from '../../graph/graphClient'
import { buildReviewTree } from '../../parsers/migrationResultParser'
import type { MigrationMapping, MigrationPhase, OneDriveAccessStatus, MigrationResultSummary, MigrationResultItem, ReviewData, ReviewNode } from '../../types'
import type { SpDriveItemDetails, DriveItemFlat } from '../../graph/graphClient'

// ─── Tree view module-level state (scoped per-open) ──────────────────────────

let _statusFilter = 'all'
let _hideRecycleBin = false
let _allItems: MigrationResultItem[] = []
let _selectedNode: ReviewNode | null = null
let _expandedPaths = new Set<string>()
let _spFeedEnabled = false
let _treeEl: HTMLElement | null = null
let _rightPanel: HTMLElement | null = null

// ─── Helpers ──────────────────────────────────────────────────────────────────

function escHtml(s: string): string {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function isOneDriveMapping(m: MigrationMapping): boolean {
  return m.matchStatus !== undefined || m.resolvedDisplayName !== undefined
}

// ─── Destination grouping ─────────────────────────────────────────────────────

interface DestGroup {
  key: string
  displayName: string
  isOneDrive: boolean
  webUrl: string
  mappings: MigrationMapping[]
}

function buildDestGroups(mappings: MigrationMapping[]): DestGroup[] {
  const map = new Map<string, DestGroup>()

  for (const m of mappings) {
    if (!m.targetSite && !m.plannedSite) continue
    const key = m.targetSite?.id ?? m.plannedSite?.alias ?? m.id
    const od = isOneDriveMapping(m)

    if (!map.has(key)) {
      map.set(key, {
        key,
        displayName: m.resolvedDisplayName ?? m.targetSite?.displayName ?? m.plannedSite?.displayName ?? key,
        isOneDrive: od,
        webUrl: m.targetSite?.webUrl ?? '',
        mappings: [],
      })
    }
    map.get(key)!.mappings.push(m)
  }

  return Array.from(map.values()).sort((a, b) => a.displayName.localeCompare(b.displayName))
}

// ─── Phase helpers ────────────────────────────────────────────────────────────

const PHASES: MigrationPhase[] = ['planning', 'migrated', 'testing', 'live']

function phaseBadgeHtml(phase: MigrationPhase | undefined): string {
  const p = phase ?? 'planning'
  if (p === 'live')     return `<span class="rev-phase-badge rev-phase--live">Live</span>`
  if (p === 'testing')  return `<span class="rev-phase-badge rev-phase--testing">Testing</span>`
  if (p === 'migrated') return `<span class="rev-phase-badge rev-phase--migrated">Migrated</span>`
  return `<span class="rev-phase-badge rev-phase--planning">Planning</span>`
}

function aggregatePhase(mappings: MigrationMapping[]): MigrationPhase {
  const order: Record<MigrationPhase, number> = { planning: 0, migrated: 1, testing: 2, live: 3 }
  let min: MigrationPhase = 'live'
  for (const m of mappings) {
    const p = m.phase ?? 'planning'
    if (order[p] < order[min]) min = p
  }
  return min
}

// ─── SPMT result cross-reference ──────────────────────────────────────────────

function spStatForMapping(m: MigrationMapping, reviewData: ReviewData): { migrated: number; failed: number; skipped: number } | null {
  const sourcePath = m.sourceNode.path
  const items = reviewData.items.filter(i =>
    i.sourcePath === sourcePath || i.sourcePath.startsWith(sourcePath + '/'))
  if (items.length === 0) return null
  return {
    migrated: items.filter(i => i.status === 'Migrated').length,
    failed:   items.filter(i => i.status === 'Failed').length,
    skipped:  items.filter(i => i.status === 'Skipped').length,
  }
}

// ─── Review data loading ──────────────────────────────────────────────────────

async function ensureReviewData(): Promise<ReviewData | null> {
  const state = getState()
  if (state.reviewData) return state.reviewData

  const project = state.currentProject
  const resultUploads = project?.projectData.resultUploads ?? []
  if (resultUploads.length === 0) return null

  const { siteId } = getSpConfig()
  const downloads = await Promise.all(
    resultUploads.map(u => downloadDriveItem(siteId, u.summaryItemId))
  ) as MigrationResultSummary[]

  const combinedItems: MigrationResultItem[] = downloads.flatMap(d => d.items ?? [])
  const tree = buildReviewTree(combinedItems)
  const reviewData: ReviewData = {
    tree,
    items: combinedItems,
    totals: {
      migrated: downloads.reduce((s, d) => s + d.migratedCount, 0),
      failed:   downloads.reduce((s, d) => s + d.failedCount, 0),
      skipped:  downloads.reduce((s, d) => s + d.skippedCount, 0),
      partial:  downloads.reduce((s, d) => s + d.partialCount, 0),
      total:    downloads.reduce((s, d) => s + d.totalCount, 0),
      failedRecycleBin:  combinedItems.filter(i => i.status === 'Failed'  && i.isRecycleBin).length,
      skippedRecycleBin: combinedItems.filter(i => i.status === 'Skipped' && i.isRecycleBin).length,
    },
  }
  setState({ reviewData })
  return reviewData
}

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
          <p class="review-empty-title">No migration results yet</p>
          <p class="review-empty-desc">Upload SPMT result ZIP files on the <strong>Upload</strong> tab. Only destinations with matching results will appear here.</p>
        </div>
      </div>`
    return
  }

  container.innerHTML = `<div class="review-panel"><div class="review-loading"><span class="spinner"></span> Loading migration results…</div></div>`

  let reviewData: ReviewData | null
  try {
    reviewData = await ensureReviewData()
  } catch (err) {
    container.innerHTML = `
      <div class="review-panel">
        <div class="review-empty">
          <div class="review-empty-icon">⚠️</div>
          <p class="review-empty-title">Failed to load results</p>
          <p class="review-empty-desc">${escHtml((err as Error).message)}</p>
        </div>
      </div>`
    return
  }

  if (!reviewData) return

  const allMappings = state.mappings.filter(m => m.targetSite || m.plannedSite)
  const allGroups = buildDestGroups(allMappings)

  // Only show destinations where at least one source folder has matching SPMT results
  const groups = allGroups.filter(g =>
    g.mappings.some(m => spStatForMapping(m, reviewData!) !== null)
  )

  if (groups.length === 0) {
    container.innerHTML = `
      <div class="review-panel">
        <div class="review-empty">
          <div class="review-empty-icon">🔍</div>
          <p class="review-empty-title">No results match current mappings</p>
          <p class="review-empty-desc">Migration results were found but none match the source paths in your current mappings.</p>
        </div>
      </div>`
    return
  }

  const migrationAccount = project.projectData.autoMapSettings?.migrationAccount ?? ''
  renderLayout(container, groups, migrationAccount, reviewData!)
}

// ─── Layout ───────────────────────────────────────────────────────────────────

function renderLayout(container: HTMLElement, groups: DestGroup[], migrationAccount: string, reviewData: ReviewData): void {
  const totalDests = groups.length
  const withAccess = groups.filter(g =>
    g.isOneDrive && g.mappings.some(m => m.accessStatus === 'accessible' || m.accessStatus === 'granted')).length
  const revoked = groups.filter(g =>
    g.isOneDrive && g.mappings.some(m => m.accessStatus === 'revoked')).length

  const phaseCounts: Record<MigrationPhase, number> = { planning: 0, migrated: 0, testing: 0, live: 0 }
  for (const g of groups) {
    for (const m of g.mappings) phaseCounts[m.phase ?? 'planning']++
  }

  container.innerHTML = `
    <div class="review-panel">
      <div class="review-stats-bar">
        <div class="rstat-card">
          <div class="rstat-label">Destinations</div>
          <div class="rstat-value rstat-blue">${totalDests}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Has Access</div>
          <div class="rstat-value rstat-green">${withAccess}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Revoked</div>
          <div class="rstat-value rstat-amber">${revoked}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Planning</div>
          <div class="rstat-value">${phaseCounts.planning}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Migrated</div>
          <div class="rstat-value rstat-blue">${phaseCounts.migrated}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Testing</div>
          <div class="rstat-value rstat-amber">${phaseCounts.testing}</div>
        </div>
        <div class="rstat-card">
          <div class="rstat-label">Live</div>
          <div class="rstat-value rstat-green">${phaseCounts.live}</div>
        </div>
      </div>

      <div class="review-mapping-layout">
        <div class="review-mapping-left">
          <div class="review-col-header">
            <span class="rch-dest">Destination</span>
            <span class="rch-phase">Phase</span>
          </div>
          <ul class="review-dest-list" id="review-dest-list">
            ${groups.map(g => renderDestItemHtml(g, reviewData)).join('')}
          </ul>
        </div>
        <div class="review-mapping-right" id="review-mapping-right">
          <div class="review-right-placeholder">
            <div class="review-placeholder-arrow">←</div>
            <p>Select a destination to see details</p>
          </div>
        </div>
      </div>
    </div>
  `

  wireDestList(container, groups, migrationAccount)
}

function renderDestItemHtml(g: DestGroup, reviewData: ReviewData): string {
  const initials = escHtml(g.displayName.slice(0, 2).toUpperCase())
  const phase = aggregatePhase(g.mappings)
  const accessBadge = g.isOneDrive
    ? `<span class="rev-access-mini">${accessStatusBadge(g.mappings[0]?.accessStatus)}</span>`
    : ''

  // Single mapping — merge into one flat row
  if (g.mappings.length === 1) {
    const m = g.mappings[0]
    const spStat = spStatForMapping(m, reviewData)
    const size = m.sourceNode.sizeBytes > 0 ? formatBytes(m.sourceNode.sizeBytes) : ''
    const spHtml = spStat
      ? `<span class="rev-source-spstat">
          <span class="rss-m" title="Migrated">✓${spStat.migrated}</span>
          ${spStat.failed > 0 ? `<span class="rss-f" title="Failed">✗${spStat.failed}</span>` : ''}
          ${spStat.skipped > 0 ? `<span class="rss-s" title="Skipped">⊘${spStat.skipped}</span>` : ''}
        </span>`
      : ''
    const viewBtn = spStat
      ? `<button class="rev-view-results-btn" data-mapping-id="${escHtml(m.id)}" title="Open full results tree">View →</button>`
      : ''
    const phaseSelect = `
      <select class="rev-phase-select" data-mapping-id="${escHtml(m.id)}" title="Migration phase">
        ${PHASES.map(p =>
          `<option value="${p}" ${(m.phase ?? 'planning') === p ? 'selected' : ''}>${p.charAt(0).toUpperCase() + p.slice(1)}</option>`
        ).join('')}
      </select>`

    return `
      <li class="review-dest-item review-dest-item--flat" data-dest-key="${escHtml(g.key)}">
        <div class="review-dest-row" tabindex="0" role="button">
          <span class="review-dest-avatar">${initials}</span>
          <span class="review-dest-name">${escHtml(g.displayName)}</span>
          ${size ? `<span class="review-source-size">${escHtml(size)}</span>` : ''}
          ${spHtml}
          ${viewBtn}
          ${accessBadge}
          <span class="rev-dest-phase" data-dest-phase="${escHtml(g.key)}">${phaseSelect}</span>
        </div>
      </li>`
  }

  // Multiple mappings — keep expandable list
  const sourceRows = g.mappings.map(m => renderSourceRowHtml(m, reviewData)).join('')

  return `
    <li class="review-dest-item" data-dest-key="${escHtml(g.key)}">
      <div class="review-dest-row" tabindex="0" role="button">
        <span class="review-dest-toggle">▶</span>
        <span class="review-dest-avatar">${initials}</span>
        <span class="review-dest-name">${escHtml(g.displayName)}</span>
        ${accessBadge}
        <span class="rev-dest-phase" data-dest-phase="${escHtml(g.key)}">${phaseBadgeHtml(phase)}</span>
      </div>
      <ul class="review-dest-sources" style="display:none">
        ${sourceRows}
      </ul>
    </li>`
}

function renderSourceRowHtml(m: MigrationMapping, reviewData: ReviewData): string {
  const name = m.sourceNode.name || m.sourceNode.originalPath
  const size = m.sourceNode.sizeBytes > 0 ? formatBytes(m.sourceNode.sizeBytes) : ''
  const spStat = spStatForMapping(m, reviewData)
  const spHtml = spStat
    ? `<span class="rev-source-spstat">
        <span class="rss-m" title="Migrated">✓${spStat.migrated}</span>
        ${spStat.failed > 0 ? `<span class="rss-f" title="Failed">✗${spStat.failed}</span>` : ''}
        ${spStat.skipped > 0 ? `<span class="rss-s" title="Skipped">⊘${spStat.skipped}</span>` : ''}
      </span>`
    : ''

  const phaseSelect = `
    <select class="rev-phase-select" data-mapping-id="${escHtml(m.id)}" title="Migration phase">
      ${PHASES.map(p =>
        `<option value="${p}" ${(m.phase ?? 'planning') === p ? 'selected' : ''}>${p.charAt(0).toUpperCase() + p.slice(1)}</option>`
      ).join('')}
    </select>`

  const viewBtn = spStat
    ? `<button class="rev-view-results-btn" data-mapping-id="${escHtml(m.id)}" title="Open full results tree">View →</button>`
    : ''

  return `
    <li class="review-source-row">
      <span class="review-source-icon">📁</span>
      <span class="review-source-name" title="${escHtml(m.sourceNode.originalPath)}">${escHtml(name)}</span>
      ${size ? `<span class="review-source-size">${escHtml(size)}</span>` : ''}
      ${spHtml}
      ${viewBtn}
      ${phaseSelect}
    </li>`
}

// ─── Wiring ───────────────────────────────────────────────────────────────────

function wireDestList(container: HTMLElement, groups: DestGroup[], migrationAccount: string): void {
  const list = container.querySelector<HTMLElement>('#review-dest-list')!
  const rightPanel = container.querySelector<HTMLElement>('#review-mapping-right')!

  list.querySelectorAll<HTMLElement>('.review-dest-row').forEach(row => {
    const item = row.closest<HTMLElement>('.review-dest-item')!
    const key = item.dataset.destKey!
    const group = groups.find(g => g.key === key)!
    const isFlat = item.classList.contains('review-dest-item--flat')

    row.addEventListener('click', (e) => {
      // Don't trigger row selection when interacting with controls inside the row
      if ((e.target as HTMLElement).closest('select, button')) return

      if (!isFlat) {
        const sources = item.querySelector<HTMLElement>('.review-dest-sources')!
        const toggle = row.querySelector<HTMLElement>('.review-dest-toggle')!
        const isOpen = sources.style.display !== 'none'
        sources.style.display = isOpen ? 'none' : ''
        toggle.textContent = isOpen ? '▶' : '▼'
      }

      list.querySelectorAll('.review-dest-row').forEach(r => r.classList.remove('review-dest-row--selected'))
      row.classList.add('review-dest-row--selected')
      renderRightPanel(rightPanel, group, migrationAccount, (newStatus, mappingId) =>
        handleAccessChanged(newStatus, mappingId, list, groups))
    })
  })

  // "View Results" button click — navigate to full tree view for that mapping
  list.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('.rev-view-results-btn')
    if (!btn) return
    e.stopPropagation()
    const mappingId = btn.dataset.mappingId!
    const mapping = getState().mappings.find(m => m.id === mappingId)
    if (!mapping) return
    const reviewData = getState().reviewData
    if (!reviewData) return
    // Pass the outer container so the back button can re-render the destinations view
    openResultsView(container, mapping, reviewData)
  })

  // Phase select change
  list.addEventListener('change', async (e) => {
    const sel = (e.target as HTMLElement).closest<HTMLSelectElement>('.rev-phase-select')
    if (!sel) return
    const mappingId = sel.dataset.mappingId!
    const newPhase = sel.value as MigrationPhase

    const updated = getState().mappings.map(m => m.id === mappingId ? { ...m, phase: newPhase } : m)
    try {
      await persistProjectMappings(updated)
    } catch (err) {
      console.warn('[Review] Failed to persist phase change:', err)
    }

    // Update aggregate phase badge on the parent destination row (multi-source only)
    const destItem = sel.closest<HTMLElement>('.review-dest-item')
    if (destItem && !destItem.classList.contains('review-dest-item--flat')) {
      const destKey = destItem.dataset.destKey!
      const latestMappings = getState().mappings.filter(m =>
        groups.find(g => g.key === destKey)?.mappings.some(gm => gm.id === m.id))
      const badge = destItem.querySelector<HTMLElement>(`[data-dest-phase="${CSS.escape(destKey)}"]`)
      if (badge) badge.innerHTML = phaseBadgeHtml(aggregatePhase(latestMappings))
    }
  })
}

function renderRightPanel(
  rightPanel: HTMLElement,
  group: DestGroup,
  migrationAccount: string,
  onAccessChanged: (s: OneDriveAccessStatus, id: string) => Promise<void>,
): void {
  if (group.isOneDrive) {
    const rep = group.mappings[0]
    rightPanel.innerHTML = `<div id="rev-person-card-wrap" style="padding:16px;overflow-y:auto;height:100%;box-sizing:border-box;"></div>`
    renderPersonCard({
      mapping: rep,
      migrationAccount,
      container: rightPanel.querySelector<HTMLElement>('#rev-person-card-wrap')!,
      onAccessChanged,
    })
  } else {
    const url = group.webUrl
    const planned = group.mappings[0]?.plannedSite && !group.mappings[0]?.targetSite
    rightPanel.innerHTML = `
      <div style="padding:16px;overflow-y:auto;height:100%;box-sizing:border-box;">
        <div class="person-card">
          <div class="person-card-header">
            <div class="person-card-avatar" style="background:var(--color-primary)">
              ${escHtml(group.displayName.slice(0, 2).toUpperCase())}
            </div>
            <div class="person-card-name">${escHtml(group.displayName)}</div>
          </div>
          <div class="person-card-body">
            ${url ? `<div class="person-card-row">
              <span class="person-card-label">URL</span>
              <a href="${escHtml(url)}" target="_blank" rel="noopener" class="person-card-link">${escHtml(url)}</a>
            </div>` : ''}
            <div class="person-card-row">
              <span class="person-card-label">Mappings</span>
              <span class="person-card-value">${group.mappings.length} source folder${group.mappings.length === 1 ? '' : 's'}</span>
            </div>
            ${planned ? `<div class="person-card-row">
              <span class="person-card-label">Status</span>
              <span class="person-card-value">Planned site</span>
            </div>` : ''}
          </div>
        </div>
      </div>`
  }
}

async function handleAccessChanged(
  newStatus: OneDriveAccessStatus,
  mappingId: string,
  list: HTMLElement,
  groups: DestGroup[],
): Promise<void> {
  try {
    await persistProjectMappings(getState().mappings)
  } catch (err) {
    console.warn('[Review] Failed to persist access change:', err)
  }

  const mapping = getState().mappings.find(m => m.id === mappingId)
  if (!mapping) return
  const destKey = mapping.targetSite?.id ?? ''
  const destItem = list.querySelector<HTMLElement>(`.review-dest-item[data-dest-key="${CSS.escape(destKey)}"]`)
  if (destItem) {
    const mini = destItem.querySelector<HTMLElement>('.rev-access-mini')
    if (mini) mini.innerHTML = accessStatusBadge(newStatus)
  }

  const group = groups.find(g => g.key === destKey)
  if (group) {
    const m = group.mappings.find(m => m.id === mappingId)
    if (m) m.accessStatus = newStatus
  }
}

// ─── Styles ───────────────────────────────────────────────────────────────────

// ─── Results tree view ────────────────────────────────────────────────────────

function openResultsView(container: HTMLElement, mapping: MigrationMapping, reviewData: ReviewData): void {
  const sourcePath = mapping.sourceNode.path
  const sourceName = mapping.sourceNode.name || mapping.sourceNode.originalPath

  // Filter items to just this source path and build a sub-tree
  const filteredItems = reviewData.items.filter(i =>
    i.sourcePath === sourcePath || i.sourcePath.startsWith(sourcePath + '/'))
  const subTree = buildReviewTree(filteredItems)
  const totals = {
    migrated: filteredItems.filter(i => i.status === 'Migrated').length,
    failed:   filteredItems.filter(i => i.status === 'Failed').length,
    skipped:  filteredItems.filter(i => i.status === 'Skipped').length,
    partial:  filteredItems.filter(i => i.status === 'Partial').length,
    total:    filteredItems.length,
    failedRecycleBin:  filteredItems.filter(i => i.status === 'Failed'  && i.isRecycleBin).length,
    skippedRecycleBin: filteredItems.filter(i => i.status === 'Skipped' && i.isRecycleBin).length,
  }

  // Reset tree state
  _statusFilter = 'all'
  _hideRecycleBin = false
  _allItems = filteredItems
  _selectedNode = null
  _expandedPaths = new Set()
  _spFeedEnabled = getState().currentProject?.projectData.sharePointFeedEnabled ?? false

  const pct = (n: number) => totals.total > 0 ? ` (${Math.round(n / totals.total * 100)}%)` : ''

  container.innerHTML = `
    <div class="review-panel">
      <div class="rev-results-back-bar">
        <button class="rev-back-btn" id="rev-back-btn">← Back</button>
        <span class="rev-results-breadcrumb">Results for <strong>${escHtml(sourceName)}</strong></span>
      </div>

      <div class="rev-tabs-bar">
        <button class="rev-tab-btn rev-tab-btn--active" data-tab="tree">Tree View</button>
        <button class="rev-tab-btn" data-tab="validation">File Validation</button>
      </div>

      <div id="rev-tab-tree" class="rev-tab-pane">
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
              ${totals.partial > 0 ? `<div class="rstat-card">
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
                <button class="review-pill" data-filter="Migrated">✓ Migrated</button>
                <button class="review-pill" data-filter="Failed">✗ Failed</button>
                <button class="review-pill" data-filter="Skipped">⊘ Skipped</button>
                ${totals.partial > 0 ? '<button class="review-pill" data-filter="Partial">◐ Partial</button>' : ''}
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
      </div>

      <div id="rev-tab-validation" class="rev-tab-pane" style="display:none">
        <div class="rev-val-wrap" id="rev-val-wrap">
          <div class="rev-val-empty" id="rev-val-empty">
            <div class="rev-val-empty-icon">🔍</div>
            <div class="rev-val-empty-title">File Validation</div>
            <div class="rev-val-empty-desc">
              Finds the highest-level migrated folder, queries the destination for all files
              and folders under that path, then compares them against your SPMT migration results.
            </div>
            <button class="btn btn-primary rev-check-files-btn" id="rev-check-files-btn">Check Files</button>
          </div>
        </div>
      </div>
    </div>`

  const panel = container.querySelector<HTMLElement>('.review-panel')!
  _treeEl = panel.querySelector('#review-tree') as HTMLElement
  _rightPanel = panel.querySelector('#review-right') as HTMLElement

  const rootNodes = subTree.path === '' ? subTree.children : [subTree]
  renderTreeNodes(rootNodes, _treeEl)
  setupResultsFilters(panel, rootNodes)
  setupSpFeedToggle(panel)

  // Tab switching
  panel.querySelectorAll<HTMLButtonElement>('.rev-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      panel.querySelectorAll('.rev-tab-btn').forEach(b => b.classList.remove('rev-tab-btn--active'))
      btn.classList.add('rev-tab-btn--active')
      const tab = btn.dataset.tab!
      panel.querySelectorAll<HTMLElement>('.rev-tab-pane').forEach(p => { p.style.display = 'none' })
      panel.querySelector<HTMLElement>(`#rev-tab-${tab}`)!.style.display = ''
    })
  })

  // Check Files button
  panel.querySelector('#rev-check-files-btn')!.addEventListener('click', () => {
    void runValidation(panel, mapping, filteredItems)
  })

  panel.querySelector('#rev-back-btn')!.addEventListener('click', () => {
    void renderReviewPanel(container)
  })
}

// ─── File Validation ─────────────────────────────────────────────────────────

interface ValidationRow {
  status: 'matched' | 'missing' | 'extra'
  name: string
  relPath: string
  // Source (SPMT)
  sourceUrl?: string   // raw UNC/HTTP source path from SPMT
  sourceSize?: number
  // Destination (Graph / SPMT destination URL)
  destUrl?: string
  title?: string
  createdDateTime?: string
  lastModifiedDateTime?: string
  createdBy?: string
  lastModifiedBy?: string
  versionLabel?: string
  destSize?: number
}

function normalizeDestUrl(url: string): string {
  try { return decodeURIComponent(url).toLowerCase().replace(/\/+$/, '') } catch { return url.toLowerCase().replace(/\/+$/, '') }
}

function truncateUrl(url: string, maxLen = 60): string {
  if (url.length <= maxLen) return url
  // Detect separator (UNC paths use \, URLs use /)
  const sep = url.includes('\\') ? '\\' : '/'
  const parts = url.split(sep).filter(Boolean)
  // Build from the end so the most specific part is always visible
  let result = parts[parts.length - 1]
  let i = parts.length - 2
  while (i >= 0 && result.length + parts[i].length + 1 < maxLen - 1) {
    result = parts[i] + sep + result
    i--
  }
  return '…' + sep + result
}

const VALIDATION_STEPS = ['Analysing results', 'Resolving destination', 'Enumerating files', 'Comparing']

async function runValidation(panel: HTMLElement, mapping: MigrationMapping, filteredItems: MigrationResultItem[]): Promise<void> {
  // Store context so Re-check button can re-invoke
  ;(panel as HTMLElement & { _valCtx?: unknown })._valCtx = { mapping, items: filteredItems }
  const wrap = panel.querySelector<HTMLElement>('#rev-val-wrap')!

  const renderProgress = (stepIdx: number, pct: number, detail: string) => {
    wrap.innerHTML = `
      <div class="rev-val-progress">
        <div class="rev-val-progress-steps">
          ${VALIDATION_STEPS.map((s, i) => `
            <div class="rev-val-pstep ${i < stepIdx ? 'rev-val-pstep--done' : i === stepIdx ? 'rev-val-pstep--active' : ''}">
              <span class="rev-val-pstep-dot">${i < stepIdx ? '✓' : i === stepIdx ? '●' : '○'}</span>
              <span class="rev-val-pstep-label">${escHtml(s)}</span>
            </div>`).join('')}
        </div>
        <div class="rev-val-progress-bar-wrap">
          <div class="rev-val-progress-bar" style="width:${pct}%"></div>
        </div>
        <div class="rev-val-progress-detail">${escHtml(detail)}</div>
      </div>`
  }

  // Lightweight updater for enumeration step — only updates the bar and detail, not the whole DOM
  const updateEnumProgress = (filesFound: number, foldersScanned: number) => {
    const bar = wrap.querySelector<HTMLElement>('.rev-val-progress-bar')
    const detail = wrap.querySelector<HTMLElement>('.rev-val-progress-detail')
    if (bar) {
      // bar grows from 30% → 80% as folders are scanned (unbounded, so we use log scale)
      const grow = Math.min(50, Math.log10(foldersScanned + 1) * 22)
      bar.style.width = `${30 + grow}%`
    }
    if (detail) detail.textContent = `Found ${filesFound.toLocaleString()} file${filesFound !== 1 ? 's' : ''} across ${foldersScanned.toLocaleString()} folder${foldersScanned !== 1 ? 's' : ''} scanned…`
  }

  try {
    // Step 1: find the migrated files (skip recycle bin)
    renderProgress(0, 5, `Analysing SPMT results…`)
    const migratedItems = filteredItems.filter(i => i.status === 'Migrated' && !i.isRecycleBin && i.destination)
    const migratedFiles  = migratedItems.filter(i => i.itemType === 'File')

    if (migratedFiles.length === 0) {
      wrap.innerHTML = `<div class="rev-val-empty"><div class="rev-val-empty-title">No migrated files</div><div class="rev-val-empty-desc">No files with status "Migrated" were found in the SPMT results for this source.</div></div>`
      return
    }

    // Step 2: derive the enumeration root — parent directory of the shallowest destination.
    const shallowest = migratedItems.slice().sort((a, b) => a.destination.length - b.destination.length)[0]
    const rootDestUrl = shallowest.destination.split('/').slice(0, -1).join('/')

    renderProgress(1, 15, `Resolving destination folder…`)
    let ref: { driveId: string; itemId: string } | null = null

    const isPersonalOneDrive = /\/personal\//i.test(rootDestUrl)
    const oneDriveUserId = isPersonalOneDrive ? (mapping.targetSite?.id ?? '') : ''

    if (isPersonalOneDrive && oneDriveUserId) {
      try {
        const urlParsed = new URL(rootDestUrl)
        const parts = urlParsed.pathname.split('/').slice(4).map(decodeURIComponent).filter(Boolean)
        const drivePath = parts.join('/')
        ref = await resolveOneDriveFolderByPath(oneDriveUserId, drivePath)
      } catch (e) {
        throw new Error(`Could not resolve OneDrive folder.\nPath: ${rootDestUrl}\nDetail: ${(e as Error)?.message ?? e}`)
      }
    } else {
      ref = await resolveDriveItemRef(rootDestUrl)
    }

    if (!ref) {
      throw new Error(`Could not look up destination folder:\n${rootDestUrl}`)
    }

    // Step 3: enumerate all items from the destination
    renderProgress(2, 30, `Enumerating destination files — scanning folders…`)
    const destItems = await listDriveItemsRecursive(ref.driveId, ref.itemId, updateEnumProgress)
    const destFiles = destItems.filter(d => !d.isFolder)

    // Build lookup: normalized relative path → DriveItemFlat
    const destMap = new Map<string, DriveItemFlat>()
    for (const d of destFiles) {
      destMap.set(d.relativePath.toLowerCase(), d)
    }

    // Step 4: compare datasets
    renderProgress(3, 82, `Comparing ${migratedFiles.length.toLocaleString()} source files against ${destFiles.length.toLocaleString()} destination files…`)
    const rootNorm = normalizeDestUrl(rootDestUrl)
    const rows: ValidationRow[] = []
    const matchedKeys = new Set<string>()

    for (const src of migratedFiles) {
      const srcNorm = normalizeDestUrl(src.destination)
      let rel = srcNorm.startsWith(rootNorm + '/') ? srcNorm.slice(rootNorm.length + 1) : src.itemName
      // Try exact match first, then case-insensitive
      const destItem = destMap.get(rel) ?? destMap.get(rel.toLowerCase())
      if (destItem) {
        matchedKeys.add(destItem.relativePath.toLowerCase())
        rows.push({
          status: 'matched', name: src.itemName, relPath: rel,
          sourceUrl: src.source, destUrl: src.destination,
          sourceSize: src.fileSizeBytes,
          title: destItem.title, createdDateTime: destItem.createdDateTime,
          lastModifiedDateTime: destItem.lastModifiedDateTime,
          createdBy: destItem.createdBy, lastModifiedBy: destItem.lastModifiedBy,
          versionLabel: destItem.versionLabel, destSize: destItem.size,
        })
      } else {
        rows.push({ status: 'missing', name: src.itemName, relPath: rel, sourceUrl: src.source, destUrl: src.destination, sourceSize: src.fileSizeBytes })
      }
    }

    // Extra files in destination not in source
    for (const d of destFiles) {
      if (!matchedKeys.has(d.relativePath.toLowerCase())) {
        rows.push({
          status: 'extra', name: d.name, relPath: d.relativePath,
          destUrl: `${rootDestUrl}/${d.relativePath}`,
          title: d.title, createdDateTime: d.createdDateTime,
          lastModifiedDateTime: d.lastModifiedDateTime,
          createdBy: d.createdBy, lastModifiedBy: d.lastModifiedBy,
          versionLabel: d.versionLabel, destSize: d.size,
        })
      }
    }

    renderValidationTable(wrap, rows, rootDestUrl)
  } catch (err) {
    wrap.innerHTML = `<div class="rev-val-empty"><div class="rev-val-empty-title">Validation failed</div><div class="rev-val-empty-desc">${escHtml((err as Error)?.message ?? String(err))}</div><button class="btn btn-ghost rev-check-files-btn" id="rev-check-files-btn-retry" style="margin-top:12px">Retry</button></div>`
    wrap.querySelector('#rev-check-files-btn-retry')?.addEventListener('click', () => {
      void runValidation(panel, mapping, filteredItems)
    })
  }
}

function renderValidationTable(wrap: HTMLElement, rows: ValidationRow[], rootUrl: string): void {
  const matched = rows.filter(r => r.status === 'matched').length
  const missing = rows.filter(r => r.status === 'missing').length
  const extra   = rows.filter(r => r.status === 'extra').length

  const fmtDate = (s?: string) => {
    if (!s) return '—'
    try { return new Date(s).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) } catch { return s }
  }

  const statusIcon = (s: ValidationRow['status']) =>
    s === 'matched' ? '<span class="rev-val-s rev-val-s--matched">✓ Matched</span>'
    : s === 'missing' ? '<span class="rev-val-s rev-val-s--missing">✗ Missing</span>'
    : '<span class="rev-val-s rev-val-s--extra">⚠ Extra</span>'

  wrap.innerHTML = `
    <div class="rev-val-summary">
      <span class="rev-val-stat rev-val-stat--ok">✓ ${matched.toLocaleString()} matched</span>
      <span class="rev-val-stat rev-val-stat--err">${missing > 0 ? `✗ ${missing.toLocaleString()} missing` : '✓ None missing'}</span>
      <span class="rev-val-stat rev-val-stat--warn">${extra > 0 ? `⚠ ${extra.toLocaleString()} extra` : '✓ None extra'}</span>
      <span class="rev-val-root" title="${escHtml(rootUrl)}">Root: <code>${escHtml(rootUrl.split('/').slice(-2).join('/'))}</code></span>
      <button class="btn btn-ghost btn-sm rev-val-download-btn" style="margin-left:auto">⬇ Download CSV</button>
      <button class="btn btn-ghost btn-sm rev-val-recheck-btn">Re-check</button>
    </div>
    <div class="rev-val-filter-bar">
      <button class="rev-val-pill rev-val-pill--active" data-vfilter="all">All (${rows.length})</button>
      <button class="rev-val-pill" data-vfilter="matched">✓ Matched (${matched})</button>
      <button class="rev-val-pill" data-vfilter="missing">✗ Missing (${missing})</button>
      <button class="rev-val-pill" data-vfilter="extra">⚠ Extra (${extra})</button>
      <input type="text" class="rev-val-search form-input" placeholder="Filter by name or path…" />
    </div>
    <div class="rev-val-table-wrap">
      <table class="rev-val-table" id="rev-val-table">
        <thead>
          <tr>
            <th style="width:90px">Status</th>
            <th style="width:140px">Name</th>
            <th style="width:200px">Source</th>
            <th style="width:200px">Destination</th>
            <th style="width:110px">Created</th>
            <th style="width:110px">Modified</th>
            <th style="width:130px">Owner</th>
            <th style="width:130px">Modified By</th>
            <th style="width:70px">Version</th>
          </tr>
        </thead>
        <tbody id="rev-val-tbody">
          ${rows.map(r => `
            <tr class="rev-val-row" data-vstatus="${r.status}">
              <td>${statusIcon(r.status)}</td>
              <td class="rev-val-name" title="${escHtml(r.relPath)}">${escHtml(r.name)}</td>
              <td class="rev-val-url rev-val-url--src" data-fullurl="${escHtml(r.sourceUrl ?? '')}" title="Click to expand">
                <span class="rev-val-url-short">${escHtml(r.sourceUrl ? truncateUrl(r.sourceUrl) : '—')}</span>
              </td>
              <td class="rev-val-url rev-val-url--dest" data-fullurl="${escHtml(r.destUrl ?? '')}" title="Click to expand">
                <span class="rev-val-url-short">${escHtml(r.destUrl ? truncateUrl(r.destUrl) : '—')}</span>
              </td>
              <td>${fmtDate(r.createdDateTime)}</td>
              <td>${fmtDate(r.lastModifiedDateTime)}</td>
              <td>${escHtml(r.createdBy ?? '—')}</td>
              <td>${escHtml(r.lastModifiedBy ?? '—')}</td>
              <td>${escHtml(r.versionLabel || '—')}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`

  // Column resize handles
  const table = wrap.querySelector<HTMLElement>('#rev-val-table')!
  table.querySelectorAll<HTMLElement>('th').forEach(th => {
    const handle = document.createElement('div')
    handle.className = 'rev-col-resize-handle'
    th.appendChild(handle)
    handle.addEventListener('mousedown', (e) => {
      const startX = e.clientX
      const startW = th.offsetWidth
      const onMove = (ev: MouseEvent) => { th.style.width = Math.max(50, startW + ev.clientX - startX) + 'px' }
      const onUp   = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp) }
      document.addEventListener('mousemove', onMove)
      document.addEventListener('mouseup', onUp)
      e.preventDefault()
    })
  })

  // URL cell expand on click
  wrap.querySelectorAll<HTMLElement>('.rev-val-url').forEach(cell => {
    cell.addEventListener('click', () => {
      const full = cell.dataset.fullurl ?? ''
      if (!full || full === '—') return
      const isExpanded = cell.classList.contains('rev-val-url--expanded')
      cell.classList.toggle('rev-val-url--expanded', !isExpanded)
      cell.querySelector('.rev-val-url-short')!.textContent = isExpanded ? truncateUrl(full) : full
      cell.title = isExpanded ? 'Click to expand' : 'Click to collapse'
    })
  })

  // CSV download
  wrap.querySelector('.rev-val-download-btn')?.addEventListener('click', () => {
    downloadValidationCsv(rows)
  })

  // Filter pills
  wrap.querySelectorAll<HTMLElement>('.rev-val-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      wrap.querySelectorAll('.rev-val-pill').forEach(p => p.classList.remove('rev-val-pill--active'))
      pill.classList.add('rev-val-pill--active')
      const f = pill.dataset.vfilter!
      const search = (wrap.querySelector<HTMLInputElement>('.rev-val-search')?.value ?? '').toLowerCase()
      applyValidationFilter(wrap, f, search)
    })
  })

  // Search
  wrap.querySelector<HTMLInputElement>('.rev-val-search')?.addEventListener('input', (e) => {
    const search = (e.target as HTMLInputElement).value.toLowerCase()
    const f = wrap.querySelector<HTMLElement>('.rev-val-pill--active')?.dataset.vfilter ?? 'all'
    applyValidationFilter(wrap, f, search)
  })

  // Re-check — re-run via closure
  wrap.querySelector('.rev-val-recheck-btn')?.addEventListener('click', () => {
    const panel = wrap.closest<HTMLElement>('.review-panel')
    if (panel) {
      const ctx = (panel as HTMLElement & { _valCtx?: { mapping: MigrationMapping; items: MigrationResultItem[] } })._valCtx
      if (ctx) void runValidation(panel, ctx.mapping, ctx.items)
    }
  })
}

function downloadValidationCsv(rows: ValidationRow[]): void {
  const csvCell = (v: string | undefined) => {
    const s = v ?? ''
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s
  }
  const headers = ['Status', 'Name', 'Source', 'Destination', 'Created', 'Modified', 'Owner', 'Modified By', 'Version']
  const fmtDate = (s?: string) => {
    if (!s) return ''
    try { return new Date(s).toLocaleDateString(undefined, { year: 'numeric', month: '2-digit', day: '2-digit' }) } catch { return s }
  }
  const lines = [
    headers.join(','),
    ...rows.map(r => [
      r.status, r.name, r.sourceUrl ?? '', r.destUrl ?? '',
      fmtDate(r.createdDateTime), fmtDate(r.lastModifiedDateTime),
      r.createdBy ?? '', r.lastModifiedBy ?? '', r.versionLabel ?? '',
    ].map(csvCell).join(',')),
  ]
  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `file-validation-${new Date().toISOString().slice(0, 10)}.csv`
  a.click()
  URL.revokeObjectURL(url)
}

function applyValidationFilter(wrap: HTMLElement, status: string, search: string): void {
  wrap.querySelectorAll<HTMLElement>('.rev-val-row').forEach(row => {
    const matchStatus = status === 'all' || row.dataset.vstatus === status
    const nameCell = row.querySelector<HTMLElement>('.rev-val-name')
    const matchSearch = !search || (nameCell?.textContent ?? '').toLowerCase().includes(search)
      || (nameCell?.title ?? '').toLowerCase().includes(search)
    row.style.display = matchStatus && matchSearch ? '' : 'none'
  })
}

// ─── Tree rendering ───────────────────────────────────────────────────────────

function renderTreeNodes(nodes: ReviewNode[], container: HTMLElement): void {
  container.innerHTML = ''
  for (const node of nodes) container.appendChild(createReviewNodeEl(node))
}

function createReviewNodeEl(node: ReviewNode): HTMLLIElement {
  const li = document.createElement('li')
  li.className = 'review-node'

  const hasChildren = node.children.length > 0
  const isExpanded = _expandedPaths.has(node.path)

  const row = document.createElement('div')
  row.className = [
    'review-row',
    node.failedCount > 0 ? 'review-row--has-failed' : '',
    _selectedNode?.path === node.path ? 'review-row--selected' : '',
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
  _treeEl?.querySelectorAll('.review-row--selected').forEach(el => el.classList.remove('review-row--selected'))
  rowEl.classList.add('review-row--selected')
  _selectedNode = node
  renderRightPanelContent()
}

function renderRightPanelContent(): void {
  if (!_rightPanel) return
  const itemPanel = _rightPanel.querySelector('#review-item-panel') as HTMLElement
  if (!itemPanel) return

  if (!_selectedNode) {
    itemPanel.innerHTML = `<div class="review-item-placeholder"><span class="review-placeholder-arrow">←</span><p>Select a file or folder</p></div>`
    clearSpContent()
    return
  }

  const items = _allItems.filter(i => i.sourcePath === _selectedNode!.path)
  const item = items[0] ?? null
  const isFolder = _selectedNode.children.length > 0

  itemPanel.innerHTML = isFolder || !item
    ? renderFolderCardHtml(_selectedNode)
    : renderFileCardHtml(_selectedNode, item)

  if (_spFeedEnabled) void loadSpFeed(item)
}

function clearSpContent(): void {
  const el = _rightPanel?.querySelector('#review-sp-content') as HTMLElement | null
  if (el) el.innerHTML = `<div class="review-sp-placeholder">Select an item to load SharePoint details</div>`
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
        <div class="ric ric--migrated"><div class="ric-value">✓ ${node.migratedCount.toLocaleString()}</div><div class="ric-label">Migrated${pct(node.migratedCount)}</div></div>
        <div class="ric ric--failed"><div class="ric-value">✗ ${node.failedCount.toLocaleString()}</div><div class="ric-label">Failed</div></div>
        <div class="ric ric--skipped"><div class="ric-value">⊘ ${node.skippedCount.toLocaleString()}</div><div class="ric-label">Skipped</div></div>
        <div class="ric ric--total"><div class="ric-value">${node.totalCount.toLocaleString()}</div><div class="ric-label">Total</div></div>
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

  spContent.innerHTML = `<div class="review-sp-loading"><span class="spinner"></span><span>Fetching from SharePoint…</span></div>`

  try {
    const details = await resolveSharePointItemByUrl(item.destination)
    spContent.innerHTML = renderSpDetailsHtml(details)
  } catch (err) {
    const msg = (err as Error).message ?? 'Unknown error'
    const reason = msg.includes('404') ? 'Item not found in SharePoint — it may not have been migrated yet.'
      : msg.includes('403') || msg.includes('401') ? 'Access denied — check permissions for this SharePoint site.'
      : msg
    spContent.innerHTML = `<div class="review-sp-error">
      <div class="review-sp-error-heading">⚠ Could not load SharePoint details</div>
      <div class="review-sp-error-detail">${escHtml(reason)}</div>
    </div>`
  }
}

function renderSpDetailsHtml(details: SpDriveItemDetails): string {
  const title = details.listItem?.fields?.Title
  return `<dl class="review-detail-grid review-sp-detail-grid">
    <dt>Name</dt><dd>${escHtml(details.name)}</dd>
    ${title && title !== details.name ? `<dt>Title</dt><dd>${escHtml(title)}</dd>` : ''}
    <dt>Created By</dt><dd>${escHtml(details.createdBy?.user?.displayName ?? '—')}</dd>
    <dt>Created Date</dt><dd>${escHtml(formatDateTime(details.createdDateTime))}</dd>
    <dt>Modified By</dt><dd>${escHtml(details.lastModifiedBy?.user?.displayName ?? '—')}</dd>
    <dt>Modified Date</dt><dd>${escHtml(formatDateTime(details.lastModifiedDateTime))}</dd>
  </dl>`
}

// ─── Results view filters & SP toggle ────────────────────────────────────────

function setupResultsFilters(panel: HTMLElement, rootNodes: ReviewNode[]): void {
  const rebuildTree = (): void => {
    if (!_treeEl) return
    const search = (panel.querySelector('#review-search') as HTMLInputElement)?.value.trim().toLowerCase() ?? ''
    const filtered = filterNodes(rootNodes, _statusFilter, _hideRecycleBin, search)
    renderTreeNodes(filtered, _treeEl)
  }

  panel.querySelector('.review-pill-group')?.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('.review-pill')
    if (!btn) return
    panel.querySelectorAll('.review-pill').forEach(p => p.classList.remove('review-pill--active'))
    btn.classList.add('review-pill--active')
    _statusFilter = btn.dataset.filter ?? 'all'
    rebuildTree()
  })

  panel.querySelector('#review-hide-rb')?.addEventListener('change', (e) => {
    _hideRecycleBin = (e.target as HTMLInputElement).checked
    rebuildTree()
  })

  panel.querySelector('#review-search')?.addEventListener('input', () => rebuildTree())
}

function setupSpFeedToggle(panel: HTMLElement): void {
  panel.querySelector('#review-sp-feed-toggle')?.addEventListener('change', async (e) => {
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
      const items = _allItems.filter(i => i.sourcePath === _selectedNode!.path)
      void loadSpFeed(items[0] ?? null)
    } else if (!enabled) {
      clearSpContent()
    }
  })
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
  if (statusFilter === 'Failed'   && node.failedCount === 0)   return false
  if (statusFilter === 'Skipped'  && node.skippedCount === 0)  return false
  if (statusFilter === 'Partial'  && node.partialCount === 0)  return false
  if (hideRb && node.children.length === 0) {
    const items = _allItems.filter(i => i.sourcePath === node.path)
    if (items.length > 0 && items.every(i => i.isRecycleBin)) return false
  }
  return true
}

// ─── Shared helpers ───────────────────────────────────────────────────────────

function statusBadgeHtml(status: string, isRecycleBin: boolean): string {
  if (isRecycleBin) return `<span class="rbadge rbadge--rb">🗑️ Recycle Bin (${escHtml(status)})</span>`
  if (status === 'Migrated') return `<span class="rbadge rbadge--migrated">✓ Migrated</span>`
  if (status === 'Failed')   return `<span class="rbadge rbadge--failed">✗ Failed</span>`
  if (status === 'Skipped')  return `<span class="rbadge rbadge--skipped">⊘ Skipped</span>`
  if (status === 'Partial')  return `<span class="rbadge rbadge--partial">◐ Partial</span>`
  return `<span class="rbadge">${escHtml(status)}</span>`
}

function formatDateTime(iso: string): string {
  if (!iso) return '—'
  try {
    return new Date(iso).toLocaleString(undefined, {
      year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
    })
  } catch { return iso }
}

function injectReviewStyles(): void {
  if (document.getElementById('review-styles')) return
  const style = document.createElement('style')
  style.id = 'review-styles'
  style.textContent = `
    /* ── Panel shell ── */
    .review-panel { padding: 0; display: flex; flex-direction: column;
      height: calc(100vh - 140px); overflow: hidden; }
    .review-empty { display: flex; flex-direction: column; align-items: center; justify-content: center;
      padding: 80px 24px; text-align: center; }
    .review-empty-icon { font-size: 3rem; margin-bottom: 16px; }
    .review-empty-title { font-size: 1.1rem; font-weight: 600; margin-bottom: 8px; }
    .review-empty-desc { font-size: 0.875rem; color: var(--color-text-muted); max-width: 420px; line-height: 1.5; }

    /* ── Loading ── */
    .review-loading { display: flex; align-items: center; gap: 10px; padding: 40px 24px;
      font-size: 0.9rem; color: var(--color-text-muted); }
    .review-panel .spinner { display: inline-block; width: 14px; height: 14px;
      border: 2px solid currentColor; border-top-color: transparent; border-radius: 50%;
      animation: review-spin 0.7s linear infinite; flex-shrink: 0; }
    @keyframes review-spin { to { transform: rotate(360deg); } }

    /* ── Stats bar ── */
    .review-stats-bar { display: flex; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface); flex-shrink: 0; overflow-x: auto; }
    .rstat-card { flex: 1; min-width: 80px; padding: 10px 14px;
      border-right: 1px solid var(--color-border); }
    .rstat-card:last-child { border-right: none; }
    .rstat-label { font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 3px; }
    .rstat-value { font-size: 1.1rem; font-weight: 700; line-height: 1.1; }
    .rstat-blue { color: var(--color-primary); }
    .rstat-green { color: #107c10; }
    .rstat-amber { color: #7d4200; }

    /* ── Two-panel layout ── */
    .review-mapping-layout { flex: 1; display: grid; grid-template-columns: 3fr 2fr;
      overflow: hidden; min-height: 0; }
    .review-mapping-left { display: flex; flex-direction: column; overflow: hidden;
      border-right: 1px solid var(--color-border); min-height: 0; }
    .review-mapping-right { display: flex; flex-direction: column; overflow: hidden;
      min-width: 260px; min-height: 0; }
    .review-right-placeholder { display: flex; flex-direction: column; align-items: center;
      justify-content: center; height: 100%; gap: 8px; color: var(--color-text-muted);
      font-size: 0.875rem; text-align: center; }
    .review-placeholder-arrow { font-size: 1.5rem; }

    /* ── Column header ── */
    .review-col-header { display: flex; align-items: center; padding: 6px 16px;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border);
      font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.05em; flex-shrink: 0; gap: 8px; }
    .rch-dest { flex: 1; }
    .rch-phase { width: 90px; text-align: right; padding-right: 8px; }

    /* ── Destination list ── */
    .review-dest-list { flex: 1; overflow-y: auto; list-style: none; padding: 0; margin: 0; min-height: 0; }
    .review-dest-item { border-bottom: 1px solid var(--color-border); }
    .review-dest-row { display: flex; align-items: center; gap: 8px; padding: 10px 16px;
      cursor: pointer; transition: background 0.1s; }
    .review-dest-row:hover { background: #f0f6ff; }
    .review-dest-row--selected { background: var(--color-primary-light, #deecf9) !important; }
    .review-dest-toggle { font-size: 0.6rem; color: var(--color-text-muted); width: 10px; flex-shrink: 0; }
    .review-dest-item--flat .review-dest-row { padding-left: 12px; }
    .review-dest-item--flat .review-dest-name { max-width: 180px; }
    .review-dest-item--flat .rev-dest-phase { display: contents; }
    .review-dest-avatar { width: 28px; height: 28px; border-radius: 50%;
      background: var(--color-primary, #0078d4); color: white; font-size: 0.72rem; font-weight: 700;
      display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .review-dest-name { flex: 1; font-size: 0.875rem; font-weight: 500;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .rev-access-mini { flex-shrink: 0; }
    .rev-access-mini .badge { font-size: 0.68rem; padding: 1px 6px; }
    .rev-dest-phase { flex-shrink: 0; }

    /* ── Source rows ── */
    .review-dest-sources { list-style: none; padding: 0; margin: 0;
      background: var(--color-surface-alt); border-top: 1px solid var(--color-border); }
    .review-source-row { display: flex; align-items: center; gap: 8px;
      padding: 7px 16px 7px 44px; border-bottom: 1px solid var(--color-border);
      font-size: 0.82rem; }
    .review-source-row:last-child { border-bottom: none; }
    .review-source-icon { flex-shrink: 0; font-size: 0.9rem; }
    .review-source-name { flex: 1; min-width: 0; white-space: nowrap;
      overflow: hidden; text-overflow: ellipsis; font-family: 'Consolas', monospace;
      font-size: 0.78rem; }
    .review-source-size { flex-shrink: 0; font-size: 0.75rem; color: var(--color-text-muted); }
    .rev-source-spstat { display: flex; gap: 6px; flex-shrink: 0; font-size: 0.72rem; }
    .rss-m { color: #107c10; font-weight: 600; }
    .rss-f { color: var(--color-danger, #a4262c); font-weight: 600; }
    .rss-s { color: var(--color-text-muted); }

    /* ── Phase select ── */
    .rev-phase-select { font-size: 0.75rem; padding: 2px 4px; border-radius: 4px;
      border: 1px solid var(--color-border); background: white; cursor: pointer;
      flex-shrink: 0; font-family: inherit; }

    /* ── Phase badges ── */
    .rev-phase-badge { display: inline-block; padding: 2px 8px; border-radius: 10px;
      font-size: 0.72rem; font-weight: 600; }
    .rev-phase--planning { background: #f3f2f1; color: #605e5c; }
    .rev-phase--migrated { background: rgba(0,120,212,0.12); color: var(--color-primary, #0078d4); }
    .rev-phase--testing  { background: rgba(255,140,0,0.12); color: #7d4200; }
    .rev-phase--live     { background: rgba(16,124,16,0.12); color: #107c10; }

    /* ── Access / status badges ── */
    .badge { display: inline-block; padding: 2px 8px; border-radius: 3px;
      font-size: 0.78rem; font-weight: 600; }
    .badge-neutral { background: var(--color-surface-alt, #f3f2f1); color: var(--color-text-muted, #605e5c); }
    .badge-revoked  { background: #f3f2f1; color: #605e5c; }
    .status-ready   { background: rgba(16,124,16,0.12); color: #107c10; }
    .status-error   { background: rgba(209,52,56,0.12); color: var(--color-danger, #a4262c); }

    /* ── Person card ── */
    .person-card { font-size: 0.875rem; }
    .person-card-header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px;
      padding-bottom: 16px; border-bottom: 1px solid var(--color-border); }
    .person-card-avatar { width: 40px; height: 40px; border-radius: 50%;
      background: var(--color-primary, #0078d4); color: white; font-size: 0.9rem; font-weight: 700;
      display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .person-card-name { font-size: 1rem; font-weight: 600; }
    .person-card-body { display: flex; flex-direction: column; gap: 10px; }
    .person-card-row { display: flex; gap: 8px; }
    .person-card-label { width: 90px; flex-shrink: 0; font-size: 0.75rem; font-weight: 600;
      color: var(--color-text-muted); padding-top: 2px; }
    .person-card-value { flex: 1; min-width: 0; word-break: break-all; font-size: 0.85rem; }
    .person-card-link { color: var(--color-primary, #0078d4); text-decoration: none; font-size: 0.78rem; }
    .person-card-link:hover { text-decoration: underline; }
    .person-card-access-row { align-items: center; }
    .person-card-actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 4px; }
    .person-card-error { font-size: 0.78rem; color: var(--color-danger, #a4262c);
      margin-top: 4px; word-break: break-word; }

    /* ── Buttons ── */
    .btn { border-radius: 4px; font-family: inherit; cursor: pointer; }
    .btn-sm { padding: 4px 10px; font-size: 0.78rem; }
    .btn-warning { background: #fff4ce; border: 1px solid #f3c00a; color: #7d5900; }
    .btn-warning:hover:not(:disabled) { background: #f3c00a; color: #3d2c00; }
    .btn-ghost { background: transparent; border: 1px solid var(--color-border); color: var(--color-text); }
    .btn-ghost:hover:not(:disabled) { background: var(--color-surface-alt); }
    button:disabled { opacity: 0.6; cursor: not-allowed; }

    /* ── View Results button on source rows ── */
    .rev-view-results-btn { font-size: 0.72rem; padding: 2px 7px; border-radius: 4px; flex-shrink: 0;
      border: 1px solid var(--color-border); background: white; cursor: pointer; font-family: inherit;
      color: var(--color-primary, #0078d4); }
    .rev-view-results-btn:hover { background: var(--color-primary-light, #deecf9); }

    /* ── Tabs ── */
    .rev-tabs-bar { display: flex; border-bottom: 2px solid var(--color-border);
      background: var(--color-surface); flex-shrink: 0; padding: 0 16px; gap: 4px; }
    .rev-tab-btn { background: none; border: none; border-bottom: 2px solid transparent;
      padding: 9px 16px; cursor: pointer; font-size: 0.875rem; font-weight: 500;
      color: var(--color-text-muted); margin-bottom: -2px; font-family: inherit; }
    .rev-tab-btn:hover { color: var(--color-text); }
    .rev-tab-btn--active { color: var(--color-primary, #0078d4); border-bottom-color: var(--color-primary, #0078d4); }
    .rev-tab-pane { flex: 1; overflow: hidden; display: flex; flex-direction: column; min-height: 0; }

    /* ── Validation tab ── */
    .rev-val-wrap { flex: 1; display: flex; flex-direction: column; min-height: 0; overflow: hidden; }
    .rev-val-loading { display: flex; align-items: center; gap: 10px; padding: 40px 24px;
      font-size: 0.9rem; color: var(--color-text-muted); }
    .rev-val-progress { padding: 40px 48px; display: flex; flex-direction: column; gap: 20px; max-width: 600px; }
    .rev-val-progress-steps { display: flex; gap: 0; }
    .rev-val-pstep { display: flex; flex-direction: column; align-items: center; flex: 1; gap: 6px; position: relative; }
    .rev-val-pstep + .rev-val-pstep::before { content: ''; position: absolute; left: calc(-50%); top: 10px;
      width: 100%; height: 2px; background: var(--color-border); z-index: 0; }
    .rev-val-pstep--done + .rev-val-pstep::before { background: var(--color-primary, #0078d4); }
    .rev-val-pstep-dot { width: 22px; height: 22px; border-radius: 50%; border: 2px solid var(--color-border);
      background: white; display: flex; align-items: center; justify-content: center;
      font-size: 0.7rem; font-weight: 700; z-index: 1; position: relative; }
    .rev-val-pstep--done .rev-val-pstep-dot { background: var(--color-primary, #0078d4);
      border-color: var(--color-primary, #0078d4); color: white; }
    .rev-val-pstep--active .rev-val-pstep-dot { border-color: var(--color-primary, #0078d4);
      color: var(--color-primary, #0078d4); animation: rev-pulse 1.2s ease-in-out infinite; }
    @keyframes rev-pulse { 0%,100% { box-shadow: 0 0 0 0 rgba(0,120,212,0.35); } 50% { box-shadow: 0 0 0 6px rgba(0,120,212,0); } }
    .rev-val-pstep-label { font-size: 0.72rem; color: var(--color-text-muted); text-align: center; white-space: nowrap; }
    .rev-val-pstep--done .rev-val-pstep-label,
    .rev-val-pstep--active .rev-val-pstep-label { color: var(--color-text); font-weight: 500; }
    .rev-val-progress-bar-wrap { height: 6px; background: var(--color-border); border-radius: 3px; overflow: hidden; }
    .rev-val-progress-bar { height: 100%; background: var(--color-primary, #0078d4); border-radius: 3px;
      transition: width 0.4s ease; }
    .rev-val-progress-detail { font-size: 0.82rem; color: var(--color-text-muted); }
    .rev-val-empty { display: flex; flex-direction: column; align-items: center; justify-content: center;
      padding: 60px 24px; text-align: center; gap: 12px; flex: 1; }
    .rev-val-empty-icon { font-size: 2.5rem; }
    .rev-val-empty-title { font-size: 1rem; font-weight: 600; }
    .rev-val-empty-desc { font-size: 0.85rem; color: var(--color-text-muted); max-width: 480px; line-height: 1.5; }
    .rev-check-files-btn { margin-top: 4px; }
    .rev-val-summary { display: flex; align-items: center; gap: 16px; padding: 10px 16px;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border);
      flex-shrink: 0; flex-wrap: wrap; }
    .rev-val-stat { font-size: 0.82rem; font-weight: 600; }
    .rev-val-stat--ok   { color: #107c10; }
    .rev-val-stat--err  { color: #a4262c; }
    .rev-val-stat--warn { color: #7d4200; }
    .rev-val-root { font-size: 0.78rem; color: var(--color-text-muted); margin-left: auto; }
    .rev-val-root code { font-size: 0.78rem; background: #f3f2f1; padding: 1px 4px; border-radius: 3px; }
    .rev-val-filter-bar { display: flex; align-items: center; gap: 6px; padding: 8px 16px;
      border-bottom: 1px solid var(--color-border); flex-shrink: 0; flex-wrap: wrap; }
    .rev-val-pill { border: 1px solid var(--color-border); background: white; border-radius: 20px;
      padding: 3px 12px; font-size: 0.75rem; cursor: pointer; font-family: inherit; }
    .rev-val-pill--active { background: var(--color-primary, #0078d4); color: white; border-color: var(--color-primary); }
    .rev-val-search { font-size: 0.8rem !important; padding: 3px 8px !important; width: 200px; margin-left: auto; }
    .rev-val-table-wrap { flex: 1; overflow-y: auto; min-height: 0; }
    .rev-val-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; table-layout: fixed; }
    .rev-val-table th { text-align: left; padding: 6px 22px 6px 10px; background: var(--color-surface-alt);
      font-weight: 700; text-transform: uppercase; font-size: 0.65rem; letter-spacing: 0.04em;
      position: sticky; top: 0; border-bottom: 1px solid var(--color-border); white-space: nowrap;
      overflow: hidden; position: relative; user-select: none; }
    .rev-val-table td { padding: 6px 10px; border-bottom: 1px solid var(--color-border);
      vertical-align: middle; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .rev-val-row:nth-child(odd) td  { background: #fff; }
    .rev-val-row:nth-child(even) td { background: #f8f7f6; }
    .rev-val-row:hover td { background: #eef4fb !important; }
    .rev-col-resize-handle { position: absolute; right: 0; top: 0; bottom: 0; width: 5px;
      cursor: col-resize; z-index: 1; }
    .rev-col-resize-handle:hover, .rev-col-resize-handle:active { background: var(--color-primary, #0078d4); opacity: 0.5; }
    .rev-val-name { max-width: 180px; overflow: hidden; text-overflow: ellipsis;
      font-family: 'Consolas', monospace; font-size: 0.76rem; }
    .rev-val-s { font-size: 0.75rem; font-weight: 600; white-space: nowrap; }
    .rev-val-s--matched { color: #107c10; }
    .rev-val-s--missing { color: #a4262c; }
    .rev-val-s--extra   { color: #7d4200; }
    .rev-val-row:hover td { background: #f9f8f7; }
    .rev-val-url { min-width: 180px; max-width: 260px; cursor: pointer; }
    .rev-val-url .rev-val-url-short { display: block; font-family: 'Consolas', monospace;
      font-size: 0.73rem; color: var(--color-primary, #0078d4); white-space: nowrap; }
    .rev-val-url--expanded { max-width: 420px; }
    .rev-val-url--expanded .rev-val-url-short { white-space: normal; word-break: break-all; }
    .rev-val-url:hover .rev-val-url-short { text-decoration: underline; }

    /* ── Back bar (results tree view) ── */
    .rev-results-back-bar { display: flex; align-items: center; gap: 12px; padding: 8px 16px;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .rev-back-btn { font-size: 0.82rem; padding: 4px 10px; border-radius: 4px;
      border: 1px solid var(--color-border); background: white; cursor: pointer; font-family: inherit; }
    .rev-back-btn:hover { background: var(--color-surface-alt); }
    .rev-results-breadcrumb { font-size: 0.82rem; color: var(--color-text-muted); }
    .rev-results-breadcrumb strong { color: var(--color-text); }

    /* ── Results two-panel layout ── */
    .review-layout { flex: 1; display: grid; grid-template-columns: 2fr 1fr; overflow: hidden; min-height: 0; }
    .review-left { display: flex; flex-direction: column; overflow: hidden;
      border-right: 1px solid var(--color-border); min-height: 0; }
    .review-right { display: flex; flex-direction: column; overflow: hidden; min-width: 260px; min-height: 0; }

    /* ── Stats/filter bars ── */
    .rstat-card--danger { border-left: 3px solid var(--color-danger, #a4262c); }
    .rstat-card--skipped { border-left: 3px solid #f3c00a; }
    .rstat-red { color: var(--color-danger, #a4262c); }
    .rstat-sub { font-size: 0.68rem; color: var(--color-text-muted); margin-top: 2px; }
    .review-filter-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 8px;
      padding: 10px 16px; background: var(--color-surface);
      border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .review-search-wrap { flex: 1; min-width: 160px; max-width: 260px; }
    .review-search-input { padding: 5px 10px; font-size: 0.85rem; }
    .review-pill-group { display: flex; gap: 4px; flex-wrap: wrap; }
    .review-pill { padding: 4px 12px; border-radius: 20px; border: 1px solid var(--color-border);
      background: white; font-size: 0.8rem; cursor: pointer; font-family: inherit; }
    .review-pill:hover { background: var(--color-surface-alt); }
    .review-pill.review-pill--active { background: var(--color-primary, #0078d4); color: white; border-color: var(--color-primary, #0078d4); }
    .review-rb-label { display: flex; align-items: center; gap: 6px; font-size: 0.82rem;
      color: var(--color-text-muted); cursor: pointer; white-space: nowrap; }

    /* ── Tree ── */
    .rch-name { flex: 1; padding-left: 64px; }
    .rch-stat { width: 80px; text-align: right; padding-right: 12px; flex-shrink: 0; }
    .rch-migrated { color: #107c10; }
    .rch-failed { color: var(--color-danger, #a4262c); }
    .rch-skipped { color: #605e5c; }
    .review-tree { flex: 1; overflow-y: auto; list-style: none; padding: 0; margin: 0; min-height: 0; }
    .review-node { list-style: none; }
    .review-children { list-style: none; padding: 0; margin: 0 0 0 20px; border-left: 1px solid var(--color-border); }
    .review-row { display: flex; align-items: center; padding: 5px 8px; cursor: pointer;
      transition: background 0.1s; min-height: 32px; }
    .review-row:hover { background: #f0f6ff; }
    .review-row--selected { background: var(--color-primary-light, #deecf9) !important; }
    .review-row--has-failed { border-left: 3px solid var(--color-danger); background: rgba(209,52,56,0.04); }
    .review-row--has-failed:hover { background: rgba(209,52,56,0.09); }
    .review-toggle { width: 16px; font-size: 0.6rem; color: var(--color-text-muted); flex-shrink: 0; }
    .review-icon { width: 22px; text-align: center; flex-shrink: 0; margin-right: 4px; }
    .review-name { flex: 1; min-width: 0; font-size: 0.85rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .rstat { width: 80px; text-align: right; padding-right: 12px; font-size: 0.78rem; flex-shrink: 0; white-space: nowrap; }
    .rstat-migrated { color: #107c10; font-weight: 600; }
    .rstat-failed { color: var(--color-danger, #a4262c); font-weight: 600; }
    .rstat-skipped { color: var(--color-text-muted); }
    .rstat-total { color: var(--color-text-muted); }

    /* ── Right detail panel ── */
    .review-right-header { padding: 12px 16px; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface-alt); flex-shrink: 0; }
    .review-feed-toggle-label { display: flex; align-items: center; gap: 8px; font-size: 0.85rem; cursor: pointer; user-select: none; }
    .review-item-panel { flex: 1; overflow-y: auto; border-bottom: 1px solid var(--color-border); min-height: 0; }
    .review-item-placeholder { display: flex; flex-direction: column; align-items: center;
      justify-content: center; height: 100%; padding: 32px 16px; text-align: center;
      color: var(--color-text-muted); font-size: 0.875rem; gap: 6px; }
    .review-item-card { padding: 16px; }
    .review-item-title-row { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 16px; }
    .review-item-type-icon { font-size: 1.8rem; flex-shrink: 0; margin-top: 2px; }
    .review-item-title-text { min-width: 0; }
    .review-item-name { font-size: 0.9rem; font-weight: 600; font-family: 'Consolas', monospace;
      word-break: break-all; margin-bottom: 3px; }
    .review-item-path { font-size: 0.72rem; color: var(--color-text-muted); font-family: 'Consolas', monospace; word-break: break-all; }
    .review-item-counts { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .ric { padding: 10px 12px; border-radius: 6px; background: var(--color-surface-alt); text-align: center; }
    .ric-value { font-size: 1rem; font-weight: 700; margin-bottom: 2px; }
    .ric-label { font-size: 0.68rem; color: var(--color-text-muted); text-transform: uppercase; letter-spacing: 0.04em; }
    .ric--migrated .ric-value { color: #107c10; }
    .ric--failed .ric-value { color: var(--color-danger, #a4262c); }
    .ric--skipped .ric-value { color: #605e5c; }
    .ric--total .ric-value { color: var(--color-primary, #0078d4); }
    .review-detail-grid { display: grid; grid-template-columns: 110px 1fr; gap: 6px 12px; margin: 0; padding: 0; }
    .review-detail-grid dt { color: var(--color-text-muted); font-size: 0.78rem; font-weight: 600; align-self: start; padding-top: 2px; }
    .review-detail-grid dd { margin: 0; font-size: 0.82rem; word-break: break-all; }
    .review-detail-path { font-family: 'Consolas', monospace; font-size: 0.72rem; color: var(--color-text); }
    .review-detail-path--muted { color: var(--color-text-muted); }
    .review-detail-message { color: var(--color-text); }
    .review-detail-error { color: var(--color-danger, #a4262c); font-family: 'Consolas', monospace; font-size: 0.72rem; }
    .rbadge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.78rem; font-weight: 600; }
    .rbadge--migrated { background: rgba(16,124,16,0.12); color: #107c10; }
    .rbadge--failed { background: rgba(209,52,56,0.12); color: var(--color-danger, #a4262c); }
    .rbadge--skipped { background: var(--color-surface-alt); color: var(--color-text-muted); }
    .rbadge--partial { background: rgba(255,140,0,0.12); color: #7d4200; }
    .rbadge--rb { background: rgba(243,192,10,0.15); color: #7d5900; }

    /* ── SP feed ── */
    .review-sp-section { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-height: 0; }
    .review-sp-header { padding: 8px 16px; font-size: 0.72rem; font-weight: 700;
      text-transform: uppercase; letter-spacing: 0.05em; color: var(--color-text-muted);
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border); flex-shrink: 0; }
    .review-sp-content { flex: 1; overflow-y: auto; min-height: 0; }
    .review-sp-placeholder { padding: 24px 16px; text-align: center; color: var(--color-text-muted); font-size: 0.85rem; }
    .review-sp-loading { display: flex; align-items: center; gap: 10px; padding: 20px 16px; color: var(--color-text-muted); font-size: 0.85rem; }
    .review-sp-error { padding: 14px 16px; }
    .review-sp-error-heading { font-size: 0.82rem; font-weight: 600; color: var(--color-danger, #a4262c); margin-bottom: 6px; }
    .review-sp-error-detail { font-size: 0.78rem; color: var(--color-text-muted); white-space: pre-wrap; word-break: break-all; line-height: 1.5; }
    .review-sp-detail-grid { padding: 16px; }
  `
  document.head.appendChild(style)
}
