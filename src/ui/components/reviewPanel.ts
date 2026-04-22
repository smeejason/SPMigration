import { getState } from '../../state/store'
import { persistProjectMappings } from '../../graph/projectService'
import { renderPersonCard, accessStatusBadge } from './oneDrivePersonCard'
import type { MigrationMapping, MigrationPhase, OneDriveAccessStatus } from '../../types'

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

function spStatForMapping(m: MigrationMapping): { migrated: number; failed: number; skipped: number } | null {
  const reviewData = getState().reviewData
  if (!reviewData) return null
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

// ─── Entry point ──────────────────────────────────────────────────────────────

export async function renderReviewPanel(container: HTMLElement): Promise<void> {
  injectReviewStyles()
  const state = getState()
  const project = state.currentProject
  if (!project) return

  const mappings = state.mappings.filter(m => m.targetSite || m.plannedSite)

  if (mappings.length === 0) {
    container.innerHTML = `
      <div class="review-panel">
        <div class="review-empty">
          <div class="review-empty-icon">🗂️</div>
          <p class="review-empty-title">No mappings yet</p>
          <p class="review-empty-desc">Map source folders to destinations on the <strong>Map</strong> tab first.</p>
        </div>
      </div>`
    return
  }

  const groups = buildDestGroups(mappings)
  const migrationAccount = project.projectData.autoMapSettings?.migrationAccount ?? ''

  renderLayout(container, groups, migrationAccount)
}

// ─── Layout ───────────────────────────────────────────────────────────────────

function renderLayout(container: HTMLElement, groups: DestGroup[], migrationAccount: string): void {
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
            ${groups.map(g => renderDestItemHtml(g)).join('')}
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

function renderDestItemHtml(g: DestGroup): string {
  const initials = escHtml(g.displayName.slice(0, 2).toUpperCase())
  const phase = aggregatePhase(g.mappings)
  const accessBadge = g.isOneDrive
    ? `<span class="rev-access-mini">${accessStatusBadge(g.mappings[0]?.accessStatus)}</span>`
    : ''

  const sourceRows = g.mappings.map(m => renderSourceRowHtml(m)).join('')

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

function renderSourceRowHtml(m: MigrationMapping): string {
  const name = m.sourceNode.name || m.sourceNode.originalPath
  const size = m.sourceNode.sizeBytes > 0 ? formatBytes(m.sourceNode.sizeBytes) : ''
  const spStat = spStatForMapping(m)
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

  return `
    <li class="review-source-row">
      <span class="review-source-icon">📁</span>
      <span class="review-source-name" title="${escHtml(m.sourceNode.originalPath)}">${escHtml(name)}</span>
      ${size ? `<span class="review-source-size">${escHtml(size)}</span>` : ''}
      ${spHtml}
      ${phaseSelect}
    </li>`
}

// ─── Wiring ───────────────────────────────────────────────────────────────────

function wireDestList(container: HTMLElement, groups: DestGroup[], migrationAccount: string): void {
  let selectedKey: string | null = null

  const list = container.querySelector<HTMLElement>('#review-dest-list')!
  const rightPanel = container.querySelector<HTMLElement>('#review-mapping-right')!

  list.querySelectorAll<HTMLElement>('.review-dest-row').forEach(row => {
    const item = row.closest<HTMLElement>('.review-dest-item')!
    const key = item.dataset.destKey!
    const group = groups.find(g => g.key === key)!

    row.addEventListener('click', () => {
      const sources = item.querySelector<HTMLElement>('.review-dest-sources')!
      const toggle = row.querySelector<HTMLElement>('.review-dest-toggle')!
      const isOpen = sources.style.display !== 'none'
      sources.style.display = isOpen ? 'none' : ''
      toggle.textContent = isOpen ? '▶' : '▼'

      if (selectedKey !== key) {
        list.querySelectorAll('.review-dest-row').forEach(r => r.classList.remove('review-dest-row--selected'))
        row.classList.add('review-dest-row--selected')
        selectedKey = key
        renderRightPanel(rightPanel, group, migrationAccount, (newStatus, mappingId) =>
          handleAccessChanged(newStatus, mappingId, list, groups))
      }
    })
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

    // Update aggregate phase badge on the parent destination row
    const destItem = (sel.closest<HTMLElement>('.review-source-row'))?.closest<HTMLElement>('.review-dest-item')
    if (destItem) {
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
  `
  document.head.appendChild(style)
}
