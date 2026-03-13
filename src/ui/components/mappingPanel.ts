import { searchSites, getSiteDrives, saveMappingsFile, searchUsers, getUserDrive, checkUserDriveAccess, grantUserDriveAccess, getUserById } from '../../graph/graphClient'
import { updateProject, getSpConfig } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { TreeNode, MigrationMapping, SharePointSite, SharePointDrive, PlannedSiteTarget, AppUser } from '../../types'

// Live references to mapping tag elements so we can update them without re-rendering
const tagRegistry = new Map<string, HTMLSpanElement>()
// Live references to double-mapped warning icons on each row
const warnRegistry = new Map<string, HTMLSpanElement>()
// Paths (at stat level) that share a target with another path
let _doubleMappedPaths = new Set<string>()
// Callback set by renderMappingPanel to refresh the users-count section of the stats bar
let _statsRefreshCallback: (() => void) | null = null

// ─── Entry point ──────────────────────────────────────────────────────────────

export function renderMappingPanel(container: HTMLElement): void {
  const state = getState()
  const tree = state.treeData

  if (!tree) {
    container.innerHTML = `
      <div class="mapping-empty">
        <p>No TreeSize data loaded. Go to the <strong>Upload</strong> tab first.</p>
      </div>
    `
    return
  }

  // If the top of the tree is a synthetic root (empty path), skip it and render
  // its children directly so the user sees their actual top-level folder(s) first.
  const topNodes = !tree.path ? tree.children : [tree]

  // For stats, prefer nodes at the automap-selected level (e.g. user home-drive folders).
  // Fall back to topNodes when no level has been selected yet.
  const autoMapLevel = state.currentProject?.projectData.autoMapSettings?.selectedLevel ?? -1
  const statNodes = autoMapLevel >= 0 ? collectAtDepth(tree, autoMapLevel) : topNodes
  const stats = buildMappingStats(statNodes)

  const statMappedPaths = new Set(state.mappings.filter(m => m.targetSite || m.plannedSite).map(m => m.sourceNode.path))
  const usersReady = statNodes.filter(n => statMappedPaths.has(n.path)).length
  const usersNotMapped = statNodes.length - usersReady

  // Detect double-mapped: same target used by 2+ stat-level nodes
  {
    const targetToNodePaths = new Map<string, string[]>()
    for (const m of state.mappings) {
      if (statNodes.some(n => n.path === m.sourceNode.path) && (m.targetSite || m.resolvedDisplayName)) {
        const key = m.targetSite?.id ?? m.resolvedDisplayName ?? ''
        if (key) {
          if (!targetToNodePaths.has(key)) targetToNodePaths.set(key, [])
          targetToNodePaths.get(key)!.push(m.sourceNode.path)
        }
      }
    }
    _doubleMappedPaths = new Set([...targetToNodePaths.values()].filter(p => p.length > 1).flat())
  }
  const doubleMappedUserCount = (() => {
    const targetToNodePaths = new Map<string, number>()
    for (const m of state.mappings) {
      if (statNodes.some(n => n.path === m.sourceNode.path) && (m.targetSite || m.resolvedDisplayName)) {
        const key = m.targetSite?.id ?? m.resolvedDisplayName ?? ''
        if (key) targetToNodePaths.set(key, (targetToNodePaths.get(key) ?? 0) + 1)
      }
    }
    return [...targetToNodePaths.values()].filter(c => c > 1).length
  })()

  const statsHtml = statNodes.length > 0 ? `
    <div class="mapping-stats-bar">
      <div class="mstat-card">
        <div class="mstat-label">USERS TO MIGRATE</div>
        <div class="mstat-value mstat-blue" id="mstat-users-ready-val">${usersReady} ready to Migrate</div>
        <div class="mstat-sub mstat-not-mapped" id="mstat-users-unmapped-val">${usersNotMapped} not Mapped</div>
        <div class="mstat-sub mstat-double-mapped-warn" id="mstat-double-mapped-warn" ${doubleMappedUserCount === 0 ? 'style="display:none"' : ''}>⚠ ${doubleMappedUserCount} user${doubleMappedUserCount !== 1 ? 's' : ''} double mapped</div>
      </div>
      <div class="mstat-card">
        <div class="mstat-label">DATA TO MIGRATE</div>
        <div class="mstat-value mstat-green">${formatBytes(stats.migrateBytes) || '0 B'}</div>
        <div class="mstat-sub">${formatBytes(stats.recycleBinBytes) || '0 B'} excluded (recycle bin)</div>
      </div>
      <div class="mstat-card">
        <div class="mstat-label">TOTAL DATA SIZE</div>
        <div class="mstat-value mstat-orange">${formatBytes(stats.totalBytes) || '0 B'}</div>
        <div class="mstat-sub">across all user drives</div>
      </div>
      <div class="mstat-card">
        <div class="mstat-label">FILES TO MIGRATE</div>
        <div class="mstat-value mstat-blue">${stats.migrateFiles.toLocaleString()}</div>
        <div class="mstat-sub">Where ${stats.recycleBinFiles.toLocaleString()} files are in the recycle bin</div>
      </div>
      <div class="mstat-card mstat-card--danger">
        <div class="mstat-label">RECYCLE BIN (EXCLUDED)</div>
        <div class="mstat-value mstat-red">${formatBytes(stats.recycleBinBytes) || '0 B'}</div>
        <div class="mstat-sub">${stats.recycleBinFiles.toLocaleString()} files in ${stats.userCount} user bins</div>
        <div class="mstat-recycle-bar"><div class="mstat-recycle-fill" style="width:${stats.totalBytes > 0 ? Math.round(stats.recycleBinBytes / stats.totalBytes * 100) : 0}%"></div></div>
      </div>
    </div>` : ''

  container.innerHTML = `
    <div class="mapping-panel">
      <div class="mapping-left">
        <div class="mapping-section-header">
          <h3>Source: File System</h3>
          <span class="mapping-hint">Click a folder to map it</span>
        </div>
        ${statsHtml}
        <div class="mapping-search-bar">
          <div class="search-input-wrap">
            <input type="text" id="tree-search" class="form-input mapping-search-input" placeholder="Search by name or path… (press Enter)" autocomplete="off" />
            <button type="button" id="btn-clear-search" class="btn-clear-search" style="display:none" title="Clear search">✕</button>
          </div>
        </div>
        <div class="tree-col-header" id="tree-col-header">
          <span class="tch-name">USER FOLDER</span>
          <span class="tch-col tch-col-mapped">MAPPED TO</span>
          <span class="tch-col">TOTAL SIZE</span>
          <span class="tch-col">RECYCLE BIN</span>
          <span class="tch-col">FILES</span>
          <span class="tch-col">MIGRATE SIZE</span>
        </div>
        <div id="mapping-tree" class="mapping-tree"></div>
        <div id="mapping-search-results" class="mapping-tree" style="display:none"></div>
      </div>
      <div class="mapping-right">
        <div id="mapping-target" class="mapping-target">
          <p class="mapping-placeholder">← Select a folder on the left to configure its target</p>
        </div>
      </div>
    </div>
  `
  injectMappingStyles()

  tagRegistry.clear()
  warnRegistry.clear()
  _statsRefreshCallback = () => refreshUsersStats(container, statNodes)

  const treeEl = container.querySelector('#mapping-tree') as HTMLElement
  const targetEl = container.querySelector('#mapping-target') as HTMLElement

  const ul = document.createElement('ul')
  ul.className = 'tree-list tree-root'

  for (const node of topNodes) {
    ul.appendChild(createMappingNodeEl(node, targetEl, true))
  }

  treeEl.appendChild(ul)

  // Auto-expand if there is only one top-level node
  if (topNodes.length === 1) {
    const rootToggle = ul.querySelector<HTMLButtonElement>('.mapping-toggle-btn:not(.invisible)')
    rootToggle?.click()
  }

  // Auto-expand tree to reveal mapped nodes (without expanding their children)
  const mappedPaths = new Set(
    getState().mappings
      .filter((m) => m.targetSite || m.plannedSite)
      .map((m) => m.sourceNode.path)
  )
  if (mappedPaths.size > 0) {
    autoExpandToMappedNodes(ul, topNodes, mappedPaths)
  }

  // ── Search ────────────────────────────────────────────────────────────────
  const searchInput = container.querySelector('#tree-search') as HTMLInputElement
  const treeDiv = container.querySelector('#mapping-tree') as HTMLElement
  const resultsDiv = container.querySelector('#mapping-search-results') as HTMLElement

  // Pre-collect every node in the tree for fast searching
  const allNodes: TreeNode[] = []
  function collectNodes(node: TreeNode): void {
    allNodes.push(node)
    for (const child of node.children) collectNodes(child)
  }
  for (const n of topNodes) collectNodes(n)

  const clearSearchBtn = container.querySelector('#btn-clear-search') as HTMLButtonElement
  const treeColHeader = container.querySelector('#tree-col-header') as HTMLElement | null

  function clearSearch(): void {
    searchInput.value = ''
    clearSearchBtn.style.display = 'none'
    treeDiv.style.display = ''
    resultsDiv.style.display = 'none'
    resultsDiv.innerHTML = ''
    if (treeColHeader) treeColHeader.style.display = ''
  }

  function runSearch(): void {
    const term = searchInput.value.trim().toLowerCase()
    if (!term) { clearSearch(); return }

    const matches = allNodes.filter(
      (n) => n.name.toLowerCase().includes(term)
    )

    treeDiv.style.display = 'none'
    resultsDiv.style.display = ''
    clearSearchBtn.style.display = ''
    if (treeColHeader) treeColHeader.style.display = 'none'

    if (matches.length === 0) {
      resultsDiv.innerHTML = '<p class="mapping-search-empty">No folders match your search.</p>'
      return
    }

    const ul2 = document.createElement('ul')
    ul2.className = 'tree-list'
    for (const match of matches) {
      const li = createMappingNodeEl(match, targetEl)
      // In search results, inject a full-path subtitle so it's always visible
      const row = li.querySelector<HTMLElement>('.mapping-row')
      if (row && match.originalPath) {
        const pathLabel = document.createElement('span')
        pathLabel.className = 'search-result-path'
        pathLabel.textContent = match.originalPath
        row.insertAdjacentElement('afterend', pathLabel)
      }
      ul2.appendChild(li)
    }
    resultsDiv.innerHTML = ''
    resultsDiv.appendChild(ul2)
  }

  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') runSearch()
    if (e.key === 'Escape') clearSearch()
  })

  clearSearchBtn.addEventListener('click', clearSearch)
}

// ─── Lazy node element factory ────────────────────────────────────────────────

function createMappingNodeEl(node: TreeNode, targetEl: HTMLElement, isRoot = false, isAncestorMapped = false): HTMLLIElement {
  const li = document.createElement('li')
  li.className = `mapping-node${isRoot ? ' mapping-node--root' : ''}`

  const hasChildren = node.children.length > 0
  // All TreeSize rows are directories. Only *-wildcard entries (e.g. "*.*") are loose-file indicators.
  const isFolder = !node.name.includes('*')

  // ── Row ──────────────────────────────────────────────────────────────────
  const row = document.createElement('div')
  row.className = 'mapping-row'
  row.dataset.path = node.path

  // Toggle button (expand/collapse)
  const toggleBtn = document.createElement('button')
  toggleBtn.type = 'button'
  toggleBtn.className = `mapping-toggle-btn${hasChildren ? '' : ' invisible'}`
  const toggleIcon = document.createElement('span')
  toggleIcon.className = 'toggle-icon'
  toggleIcon.textContent = '▶'
  toggleBtn.appendChild(toggleIcon)

  // Icon — folder with optional mapped-badge overlay
  const iconWrap = document.createElement('span')
  iconWrap.className = 'tree-icon-wrap'

  // Name
  const nameEl = document.createElement('span')
  const isLooseFiles = !isFolder
  nameEl.className = `tree-name${isLooseFiles ? ' tree-name--loose' : ''}`
  nameEl.textContent = isLooseFiles ? 'Loose files' : String(node.name || node.path || '(unnamed)')
  if (node.originalPath) nameEl.title = node.originalPath

  // Mapped-to column cell (replaces the old floating tag)
  const tagEl = document.createElement('span')
  tagEl.className = 'tree-col tree-col-mapped'
  const existingMapping = getState().mappings.find((m) => m.sourceNode.path === node.path)
  tagRegistry.set(node.path, tagEl)

  // Double-mapped warning icon
  const warnEl = document.createElement('span')
  warnEl.className = 'row-warn-icon'
  warnEl.title = 'This user is mapped to multiple source folders'
  warnEl.textContent = _doubleMappedPaths.has(node.path) ? '⚠' : ''
  warnRegistry.set(node.path, warnEl)

  // Helper: apply/remove the mapped visual state on this row
  function applyMappedState(isMapped: boolean, siteName?: string, isPlanned = false): void {
    if (isFolder) {
      iconWrap.innerHTML = isMapped
        ? '📁<span class="mapped-folder-badge" aria-hidden="true">✓</span>'
        : '📁'
      iconWrap.className = `tree-icon-wrap${isMapped ? ' tree-icon-wrap--mapped' : ''}`
    } else {
      iconWrap.textContent = '📄'
    }
    if (isMapped) {
      row.classList.add('mapping-row--mapped')
      if (isPlanned) row.classList.add('mapping-row--planned'); else row.classList.remove('mapping-row--planned')
      tagEl.textContent = siteName ? `${siteName}${isPlanned ? ' (planned)' : ''}` : '—'
      tagEl.className = `tree-col tree-col-mapped${isPlanned ? ' tree-col-mapped--planned' : ''}`
    } else {
      row.classList.remove('mapping-row--mapped')
      row.classList.remove('mapping-row--planned')
      tagEl.textContent = '—'
      tagEl.className = 'tree-col tree-col-mapped tree-col-mapped--empty'
    }
  }

  const isMappedInitially = !!(existingMapping?.targetSite || existingMapping?.plannedSite)
  const initialSiteName = existingMapping?.targetSite?.displayName ?? existingMapping?.plannedSite?.displayName
  const isPlannedInitially = !existingMapping?.targetSite && !!existingMapping?.plannedSite
  applyMappedState(isMappedInitially || isAncestorMapped, initialSiteName, isPlannedInitially)

  // Column data cells
  const rbInfo = getRecycleBin(node)
  const migrateBytes = node.sizeBytes - rbInfo.sizeBytes

  const colTotal = document.createElement('span')
  colTotal.className = 'tree-col tree-col-total'
  colTotal.textContent = node.sizeBytes > 0 ? formatBytes(node.sizeBytes) : '—'

  const colRb = document.createElement('span')
  colRb.className = `tree-col tree-col-rb${rbInfo.sizeBytes > 0 ? ' tree-col-rb--has-rb' : ''}`
  colRb.textContent = rbInfo.sizeBytes > 0 ? formatBytes(rbInfo.sizeBytes) : '—'

  const colFiles = document.createElement('span')
  colFiles.className = 'tree-col tree-col-files'
  colFiles.textContent = node.fileCount > 0 ? `${node.fileCount.toLocaleString()} files` : '—'

  const colMigrate = document.createElement('span')
  colMigrate.className = 'tree-col tree-col-migrate'
  colMigrate.textContent = migrateBytes > 0 ? formatBytes(migrateBytes) : '—'

  row.appendChild(toggleBtn)
  row.appendChild(iconWrap)
  row.appendChild(nameEl)
  row.appendChild(warnEl)
  row.appendChild(tagEl)
  row.appendChild(colTotal)
  row.appendChild(colRb)
  row.appendChild(colFiles)
  row.appendChild(colMigrate)
  li.appendChild(row)

  // ── Toggle: lazy-render children on first expand ──────────────────────────
  if (hasChildren) {
    let childrenLoaded = false

    toggleBtn.addEventListener('click', (e) => {
      e.stopPropagation()
      const isOpen = li.classList.contains('mapping-node--open')

      if (isOpen) {
        const childUl = li.querySelector<HTMLElement>(':scope > .tree-children')
        if (childUl) childUl.style.display = 'none'
        li.classList.remove('mapping-node--open')
        toggleIcon.textContent = '▶'
      } else {
        if (!childrenLoaded) {
          const childUl = document.createElement('ul')
          childUl.className = 'tree-list tree-children'
          const isCurrentlyMapped = row.classList.contains('mapping-row--mapped')
          for (const child of node.children) {
            childUl.appendChild(createMappingNodeEl(child, targetEl, false, isCurrentlyMapped))
          }
          li.appendChild(childUl)
          childrenLoaded = true
        } else {
          const childUl = li.querySelector<HTMLElement>(':scope > .tree-children')
          if (childUl) childUl.style.display = ''
        }
        li.classList.add('mapping-node--open')
        toggleIcon.textContent = '▼'
      }
    })
  }

  // ── Row click: open target mapping panel (folders only) ───────────────────
  if (isFolder) {
    row.addEventListener('click', () => {
      document.querySelectorAll('.mapping-row--active').forEach((r) => r.classList.remove('mapping-row--active'))
      row.classList.add('mapping-row--active')

      // Dynamically check if any ancestor is mapped — block mapping and show info panel
      const ancestorMapping = findAncestorMapping(node.path)
      if (ancestorMapping) {
        openBlockedPanel(targetEl, node, ancestorMapping)
        return
      }

      openTargetPanel(targetEl, node, (siteName, isPlanned) => {
        const isSelfMapped = !!siteName
        applyMappedState(isSelfMapped || isAncestorMapped, siteName ?? undefined, isPlanned)
        updateDescendantHighlights(li, isSelfMapped || isAncestorMapped)
        _statsRefreshCallback?.()
      })
    })
  }

  return li
}

// ─── Target panel (right side) ────────────────────────────────────────────────

async function openTargetPanel(
  targetEl: HTMLElement,
  node: TreeNode,
  onMappingChange: (siteName: string | null, isPlanned?: boolean) => void
): Promise<void> {
  if (getState().currentProject?.type === 'OneDrive') {
    await openOneDriveTargetPanel(targetEl, node, onMappingChange)
    return
  }

  const existing = getState().mappings.find((m) => m.sourceNode.path === node.path)
  const initialTab = existing?.plannedSite && !existing?.targetSite ? 'planned' : 'existing'

  const fmtDate = (d?: Date | string) =>
    d ? new Date(d).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '—'
  const lastModStr = fmtDate(node.lastModified)
  const lastAccStr = fmtDate(node.lastAccessed)
  const sizeStr = node.sizeBytes > 0 ? formatBytes(node.sizeBytes) : '—'
  const rb = getRecycleBin(node)
  const rbStr = rb.sizeBytes > 0 ? formatBytes(rb.sizeBytes) : '—'
  const migrateSize = node.sizeBytes - rb.sizeBytes
  const migrateStr = migrateSize > 0 ? formatBytes(migrateSize) : sizeStr
  const fileStr = node.fileCount > 0 ? node.fileCount.toLocaleString() : '—'
  const folderStr = node.folderCount > 0 ? node.folderCount.toLocaleString() : '—'
  const childStr = node.children.length > 0 ? node.children.length.toLocaleString() : '—'

  targetEl.innerHTML = `
    <div class="target-panel">

      <!-- ── Section 1: Folder Summary (collapsible) ── -->
      <div class="target-section" id="summary-section">
        <button type="button" class="target-section-toggle" id="btn-toggle-summary" aria-expanded="true">
          <span class="target-section-title">Folder Summary Information</span>
          <span class="target-section-chevron" aria-hidden="true">▼</span>
        </button>
        <div class="target-section-body" id="summary-body">
          <div class="source-detail-card">
            <div class="source-detail-title">
              <span class="source-detail-icon">📁</span>
              <span class="source-detail-name">${escHtml(String(node.name || node.path))}</span>
            </div>
            <dl class="source-detail-grid">
              <dt>Full Path</dt>
              <dd class="source-detail-path" title="${escHtml(node.originalPath)}">${escHtml(node.originalPath)}</dd>
              <dt>Size</dt><dd>${sizeStr}</dd>
              <dt>Recycle Bin</dt><dd${rb.sizeBytes > 0 ? ' class="detail-recycle"' : ''}>${rbStr}</dd>
              <dt>Migrate Size</dt><dd>${migrateStr}</dd>
              <dt>Files</dt><dd>${fileStr}</dd>
              <dt>Subfolders</dt><dd>${folderStr}</dd>
              <dt>Direct Children</dt><dd>${childStr}</dd>
              <dt>Last Modified</dt><dd>${lastModStr}</dd>
              <dt>Last Accessed</dt><dd>${lastAccStr}</dd>
            </dl>
          </div>
        </div>
      </div>

      <!-- ── Section 2: SharePoint Location (tabbed) ── -->
      <div class="target-section">
        <div class="sp-tabs-bar">
          <button type="button" class="sp-tab${initialTab === 'existing' ? ' sp-tab--active' : ''}" data-tab="existing">Existing SharePoint Location</button>
          <button type="button" class="sp-tab${initialTab === 'planned' ? ' sp-tab--active' : ''}" data-tab="planned">Planned SharePoint Location</button>
        </div>

        <!-- Tab: Existing -->
        <div id="tab-existing" class="sp-tab-panel target-section-body--sp"${initialTab !== 'existing' ? ' style="display:none"' : ''}>
          <div class="form-group">
            <label>SharePoint Site</label>
            <div class="site-search-row">
              <input id="site-search" type="text" class="form-input" placeholder="Search sites…"
                value="${escHtml(existing?.targetSite?.displayName ?? '')}" />
              <button type="button" id="btn-search-sites" class="btn btn-primary btn-sm">Search</button>
            </div>
            <div id="site-results" class="site-results"></div>
            <div id="selected-site" class="selected-badge" style="${existing?.targetSite ? '' : 'display:none'}">
              ✓ ${escHtml(existing?.targetSite?.displayName ?? '')}
              <button type="button" class="btn-clear" id="btn-clear-site">✕</button>
            </div>
          </div>
          <div class="form-group" id="library-group" style="${existing?.targetSite ? '' : 'display:none'}">
            <label>Document Library</label>
            <select id="library-select" class="form-input">
              <option value="">Loading libraries…</option>
            </select>
          </div>
          <div class="form-group">
            <label>Subfolder Path <span class="hint">(optional)</span></label>
            <input id="folder-path" type="text" class="form-input" placeholder="e.g. /Migrations/Phase1"
              value="${escHtml(existing?.targetFolderPath ?? '')}" />
          </div>
          <div class="target-action-row">
            <button type="button" id="btn-save-mapping" class="btn btn-primary">Save Mapping</button>
            ${existing?.targetSite ? `<button type="button" id="btn-remove-mapping" class="btn btn-ghost">Remove</button>` : ''}
          </div>
        </div>

        <!-- Tab: Planned -->
        <div id="tab-planned" class="sp-tab-panel target-section-body--sp"${initialTab !== 'planned' ? ' style="display:none"' : ''}>
          <p class="form-hint" style="margin:0">Define the SharePoint site that will be created for this content.</p>
          <div class="form-group">
            <label>Site Display Name <span class="required">*</span></label>
            <input id="planned-name" type="text" class="form-input" placeholder="e.g. Engineering"
              value="${escHtml(existing?.plannedSite?.displayName ?? '')}" />
          </div>
          <div class="form-group">
            <label>URL Alias <span class="required">*</span></label>
            <div class="alias-row">
              <span class="alias-prefix">.../sites/</span>
              <input id="planned-alias" type="text" class="form-input" placeholder="engineering"
                value="${escHtml(existing?.plannedSite?.alias ?? '')}" />
            </div>
            <small class="form-hint">Letters, numbers, and hyphens only.</small>
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea id="planned-desc" class="form-input" rows="2" placeholder="Optional description">${escHtml(existing?.plannedSite?.description ?? '')}</textarea>
          </div>
          <div class="form-group">
            <label>Template</label>
            <div class="template-row">
              <label class="radio-label">
                <input type="radio" name="planned-template" value="team" checked /> Team site (M365 Group)
              </label>
            </div>
            <small class="form-hint">Team site support only in Phase 1.</small>
          </div>
          <div class="form-group">
            <label>Document Library <span class="hint">(optional)</span></label>
            <input id="planned-library" type="text" class="form-input" placeholder="e.g. Documents"
              value="${escHtml(existing?.plannedSite?.libraryName ?? '')}" />
          </div>
          <div class="form-group">
            <label>Subfolder Path <span class="hint">(optional)</span></label>
            <input id="planned-folder" type="text" class="form-input" placeholder="e.g. /Migrations/Phase1"
              value="${escHtml(existing?.plannedSite?.folderPath ?? '')}" />
          </div>
          <div class="target-action-row">
            <button type="button" id="btn-save-planned" class="btn btn-primary">Save Mapping</button>
            ${existing?.plannedSite ? `<button type="button" id="btn-remove-planned" class="btn btn-ghost">Remove</button>` : ''}
          </div>
        </div>

      </div>
    </div>
  `

  // ── Tab switching ─────────────────────────────────────────────────────────
  targetEl.querySelectorAll<HTMLButtonElement>('.sp-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab!
      targetEl.querySelectorAll('.sp-tab').forEach((b) => b.classList.remove('sp-tab--active'))
      btn.classList.add('sp-tab--active')
      targetEl.querySelectorAll<HTMLElement>('.sp-tab-panel').forEach((p) => {
        p.style.display = p.id === `tab-${tab}` ? '' : 'none'
      })
    })
  })

  // ── Collapsible summary toggle ────────────────────────────────────────────
  const summaryToggleBtn = targetEl.querySelector('#btn-toggle-summary') as HTMLButtonElement
  const summaryBody = targetEl.querySelector('#summary-body') as HTMLElement
  summaryToggleBtn?.addEventListener('click', () => {
    const isOpen = summaryBody.style.display !== 'none'
    summaryBody.style.display = isOpen ? 'none' : ''
    summaryToggleBtn.setAttribute('aria-expanded', String(!isOpen))
    const chevron = summaryToggleBtn.querySelector('.target-section-chevron') as HTMLElement
    if (chevron) chevron.textContent = isOpen ? '▶' : '▼'
  })

  // ── Existing tab logic ────────────────────────────────────────────────────
  let selectedSite: SharePointSite | null = existing?.targetSite ?? null
  let selectedDrive: SharePointDrive | null = existing?.targetDrive ?? null

  if (selectedSite) {
    loadLibraries(targetEl, selectedSite, selectedDrive)
  }

  targetEl.querySelector('#btn-search-sites')?.addEventListener('click', async () => {
    const query = (targetEl.querySelector('#site-search') as HTMLInputElement).value
    const results = targetEl.querySelector('#site-results') as HTMLElement
    results.innerHTML = '<span class="searching">Searching…</span>'
    try {
      const sites = await searchSites(query || '*')
      setState({ sites })
      results.innerHTML = sites.length === 0
        ? '<span class="no-results">No sites found.</span>'
        : sites.map((s) =>
            `<div class="site-result-item" data-id="${escHtml(s.id)}">${escHtml(s.displayName)}<br><small>${escHtml(s.webUrl)}</small></div>`
          ).join('')

      results.querySelectorAll('.site-result-item').forEach((item) => {
        item.addEventListener('click', () => {
          const id = item.getAttribute('data-id')!
          selectedSite = sites.find((s) => s.id === id) ?? null
          if (!selectedSite) return
          results.innerHTML = ''
          const badge = targetEl.querySelector('#selected-site') as HTMLElement
          badge.innerHTML = `✓ ${escHtml(selectedSite.displayName)} <button class="btn-clear" id="btn-clear-site">✕</button>`
          badge.style.display = ''
          attachClearSite()
          loadLibraries(targetEl, selectedSite, null)
        })
      })
    } catch {
      results.innerHTML = '<span class="no-results">Search failed.</span>'
    }
  })

  function attachClearSite(): void {
    targetEl.querySelector('#btn-clear-site')?.addEventListener('click', () => {
      selectedSite = null
      selectedDrive = null
      ;(targetEl.querySelector('#selected-site') as HTMLElement).style.display = 'none'
      ;(targetEl.querySelector('#library-group') as HTMLElement).style.display = 'none'
    })
  }
  attachClearSite()

  targetEl.querySelector('#btn-save-mapping')?.addEventListener('click', async () => {
    const folderPath = (targetEl.querySelector('#folder-path') as HTMLInputElement).value.trim()
    const libSelect = targetEl.querySelector('#library-select') as HTMLSelectElement | null
    if (libSelect && selectedSite) {
      const selId = libSelect.value
      const selName = libSelect.options[libSelect.selectedIndex]?.text ?? ''
      selectedDrive = selId ? { id: selId, name: selName, webUrl: '', driveType: 'documentLibrary' } : null
    }

    const mapping: MigrationMapping = {
      id: node.path,
      sourceNode: node,
      targetSite: selectedSite,
      targetDrive: selectedDrive,
      targetFolderPath: folderPath,
      status: selectedSite ? 'ready' : 'pending',
    }

    const mappings = [
      ...getState().mappings.filter((m) => m.sourceNode.path !== node.path),
      mapping,
    ]
    setState({ mappings })
    onMappingChange(selectedSite?.displayName ?? null, false)

    const saveBtn = targetEl.querySelector('#btn-save-mapping') as HTMLButtonElement
    saveBtn.disabled = true
    saveBtn.textContent = 'Saving…'
    try {
      await persistMappings(mappings)
      saveBtn.textContent = '✓ Saved'
    } catch {
      saveBtn.textContent = '⚠ Save failed — retry'
    } finally {
      saveBtn.disabled = false
      setTimeout(() => { if (saveBtn.textContent !== '⚠ Save failed — retry') saveBtn.textContent = 'Save Mapping' }, 2000)
    }
  })

  targetEl.querySelector('#btn-remove-mapping')?.addEventListener('click', async () => {
    const removeBtn = targetEl.querySelector('#btn-remove-mapping') as HTMLButtonElement
    removeBtn.disabled = true
    removeBtn.textContent = 'Removing…'
    const mappings = getState().mappings.filter((m) => m.sourceNode.path !== node.path)
    setState({ mappings })
    try {
      await persistMappings(mappings)
    } catch {
      removeBtn.disabled = false
      removeBtn.textContent = 'Remove'
      return
    }
    onMappingChange(null)
    removeBtn.remove()
  })

  // ── Planned tab logic ─────────────────────────────────────────────────────
  const plannedNameInput = targetEl.querySelector('#planned-name') as HTMLInputElement
  const plannedAliasInput = targetEl.querySelector('#planned-alias') as HTMLInputElement

  plannedNameInput?.addEventListener('input', () => {
    if (plannedAliasInput.dataset.userEdited) return
    plannedAliasInput.value = plannedNameInput.value
      .toLowerCase().replace(/[^a-z0-9-]/g, '-').replace(/-+/g, '-').slice(0, 60)
  })
  plannedAliasInput?.addEventListener('input', () => { plannedAliasInput.dataset.userEdited = '1' })

  targetEl.querySelector('#btn-save-planned')?.addEventListener('click', async () => {
    const plannedName = plannedNameInput.value.trim()
    const plannedAlias = plannedAliasInput.value.trim()
    const plannedDesc = (targetEl.querySelector('#planned-desc') as HTMLTextAreaElement).value.trim()
    const plannedLibrary = (targetEl.querySelector('#planned-library') as HTMLInputElement).value.trim()
    const plannedFolder = (targetEl.querySelector('#planned-folder') as HTMLInputElement).value.trim()

    if (!plannedName) { plannedNameInput.focus(); return }

    const plannedSite: PlannedSiteTarget = {
      displayName: plannedName,
      alias: plannedAlias,
      description: plannedDesc,
      template: 'team',
      libraryName: plannedLibrary,
      folderPath: plannedFolder,
    }

    const mapping: MigrationMapping = {
      id: node.path,
      sourceNode: node,
      targetSite: null,
      targetDrive: null,
      targetFolderPath: plannedFolder,
      status: 'pending',
      plannedSite,
    }

    const mappings = [
      ...getState().mappings.filter((m) => m.sourceNode.path !== node.path),
      mapping,
    ]
    setState({ mappings })
    onMappingChange(plannedName, true)

    const saveBtn = targetEl.querySelector('#btn-save-planned') as HTMLButtonElement
    saveBtn.disabled = true
    saveBtn.textContent = 'Saving…'
    try {
      await persistMappings(mappings)
      saveBtn.textContent = '✓ Saved'
    } catch {
      saveBtn.textContent = '⚠ Save failed — retry'
    } finally {
      saveBtn.disabled = false
      setTimeout(() => { if (saveBtn.textContent !== '⚠ Save failed — retry') saveBtn.textContent = 'Save Mapping' }, 2000)
    }
  })

  targetEl.querySelector('#btn-remove-planned')?.addEventListener('click', async () => {
    const removeBtn = targetEl.querySelector('#btn-remove-planned') as HTMLButtonElement
    removeBtn.disabled = true
    removeBtn.textContent = 'Removing…'
    const mappings = getState().mappings.filter((m) => m.sourceNode.path !== node.path)
    setState({ mappings })
    try {
      await persistMappings(mappings)
    } catch {
      removeBtn.disabled = false
      removeBtn.textContent = 'Remove'
      return
    }
    onMappingChange(null)
    removeBtn.remove()
  })
}

// ─── OneDrive target panel ─────────────────────────────────────────────────────

async function openOneDriveTargetPanel(
  targetEl: HTMLElement,
  node: TreeNode,
  onMappingChange: (siteName: string | null, isPlanned?: boolean) => void
): Promise<void> {
  const existing = getState().mappings.find((m) => m.sourceNode.path === node.path)
  const existingUser: AppUser | null = existing?.targetSite
    ? { id: existing.targetSite.id, displayName: existing.targetSite.displayName,
        mail: existing.targetSite.webUrl, userPrincipalName: existing.targetSite.webUrl }
    : null
  const projectDefaultSubfolder = getState().currentProject?.projectData.autoMapSettings?.targetFolderPath ?? ''
  // A mapping is using the project default when its targetFolderPath is empty (unset means "use project default")
  const isUsingProjectDefault = projectDefaultSubfolder !== '' && !existing?.targetFolderPath

  const fmtDate = (d?: Date | string) =>
    d ? new Date(d).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '—'

  targetEl.innerHTML = `
    <div class="target-panel">

      <!-- Folder Summary (collapsible) -->
      <div class="target-section" id="summary-section">
        <button type="button" class="target-section-toggle" id="btn-toggle-summary" aria-expanded="true">
          <span class="target-section-title">Folder Summary Information</span>
          <span class="target-section-chevron" aria-hidden="true">▼</span>
        </button>
        <div class="target-section-body" id="summary-body">
          <div class="source-detail-card">
            <div class="source-detail-title">
              <span class="source-detail-icon">📁</span>
              <span class="source-detail-name">${escHtml(String(node.name || node.path))}</span>
            </div>
            <dl class="source-detail-grid">
              <dt>Full Path</dt>
              <dd class="source-detail-path" title="${escHtml(node.originalPath)}">${escHtml(node.originalPath)}</dd>
              <dt>Size</dt><dd>${node.sizeBytes > 0 ? formatBytes(node.sizeBytes) : '—'}</dd>
              <dt>Files</dt><dd>${node.fileCount > 0 ? node.fileCount.toLocaleString() : '—'}</dd>
              <dt>Last Modified</dt><dd>${fmtDate(node.lastModified)}</dd>
            </dl>
          </div>
        </div>
      </div>

      <!-- OneDrive user selection -->
      <div class="target-section">
        <div class="sp-tabs-bar">
          <span class="sp-tab sp-tab--active" style="cursor:default">OneDrive User</span>
        </div>
        <div class="target-section-body--sp">

          <div class="form-group">
            <label>Select User</label>
            <div class="site-search-row">
              <input id="od-user-search" type="text" class="form-input"
                placeholder="Search by name or UPN…"
                value="${escHtml(existingUser?.displayName ?? '')}" />
              <button type="button" id="btn-od-search" class="btn btn-primary btn-sm">Search</button>
            </div>
            <div id="od-user-results" class="site-results"></div>
            <div id="od-selected-user" class="selected-badge" style="${existingUser ? '' : 'display:none'}">
              ✓ ${escHtml(existingUser?.displayName ?? '')}
              <button type="button" class="btn-clear" id="btn-clear-od-user">✕</button>
            </div>
          </div>

          <div id="od-drive-info" style="${existing?.targetSite ? '' : 'display:none'}">
            <div class="od-drive-card">
              <div class="od-drive-row">
                <span class="od-drive-label">Display Name</span>
                <span id="od-user-displayname" class="od-drive-value">${escHtml(existing?.targetSite?.displayName ?? '')}</span>
              </div>
              <div class="od-drive-row">
                <span class="od-drive-label">UPN</span>
                <span id="od-user-upn" class="od-drive-value">${existing?.targetSite ? '⏳ Loading…' : ''}</span>
              </div>
              <div class="od-drive-row">
                <span class="od-drive-label">OneDrive URL</span>
                <span id="od-drive-url" class="od-drive-value">${escHtml(existing?.targetSite?.webUrl ?? '')}</span>
              </div>
              <div class="od-drive-row">
                <span class="od-drive-label">Access</span>
                <span id="od-access-status" class="od-drive-value">⏳ Checking…</span>
              </div>
            </div>
          </div>

          <div class="form-group">
            <label>Subfolder Path <span class="hint">(optional)</span></label>
            ${projectDefaultSubfolder ? `
            <div class="subfolder-mode-row">
              <label class="radio-label">
                <input type="radio" name="od-subfolder-mode" value="project" ${isUsingProjectDefault ? 'checked' : ''} />
                Project default: <code class="subfolder-default-code">${escHtml(projectDefaultSubfolder)}</code>
              </label>
              <label class="radio-label">
                <input type="radio" name="od-subfolder-mode" value="custom" ${!isUsingProjectDefault ? 'checked' : ''} />
                Override for this folder
              </label>
            </div>` : ''}
            <input id="od-folder-path" type="text" class="form-input" placeholder="e.g. Migration/Files"
              value="${escHtml(existing?.targetFolderPath ?? '')}"
              ${projectDefaultSubfolder && isUsingProjectDefault ? 'style="display:none"' : ''} />
          </div>

          <div class="target-action-row">
            <button type="button" id="btn-save-od-mapping" class="btn btn-primary">Save Mapping</button>
            ${existing?.targetSite ? `<button type="button" id="btn-remove-od-mapping" class="btn btn-ghost">Remove</button>` : ''}
          </div>

        </div>
      </div>
    </div>
  `

  // ── Collapsible summary ────────────────────────────────────────────────────
  const summaryToggleBtn = targetEl.querySelector('#btn-toggle-summary') as HTMLButtonElement
  const summaryBody = targetEl.querySelector('#summary-body') as HTMLElement
  summaryToggleBtn?.addEventListener('click', () => {
    const isOpen = summaryBody.style.display !== 'none'
    summaryBody.style.display = isOpen ? 'none' : ''
    summaryToggleBtn.setAttribute('aria-expanded', String(!isOpen))
    ;(summaryToggleBtn.querySelector('.target-section-chevron') as HTMLElement).textContent = isOpen ? '▶' : '▼'
  })

  // ── Subfolder mode radio ────────────────────────────────────────────────────
  if (projectDefaultSubfolder) {
    const folderPathInput = targetEl.querySelector('#od-folder-path') as HTMLInputElement
    targetEl.querySelectorAll<HTMLInputElement>('input[name="od-subfolder-mode"]').forEach(radio => {
      radio.addEventListener('change', () => {
        if (radio.value === 'project') {
          folderPathInput.value = ''
          folderPathInput.style.display = 'none'
        } else {
          folderPathInput.style.display = ''
          folderPathInput.focus()
        }
      })
    })
  }

  // ── State ──────────────────────────────────────────────────────────────────
  let selectedUser: AppUser | null = existingUser
  let selectedDriveId = existing?.targetDrive?.id ?? ''
  let selectedDriveWebUrl = existing?.targetSite?.webUrl ?? ''

  const migrationAccount = getState().currentProject?.projectData.autoMapSettings?.migrationAccount ?? ''

  // Fetch UPN, drive URL, and check access for existing user on load
  if (existing?.targetSite?.id) {
    checkAndShowAccess(targetEl, existing.targetSite.id, migrationAccount)
    getUserById(existing.targetSite.id).then(user => {
      const upnEl = targetEl.querySelector('#od-user-upn') as HTMLElement | null
      if (upnEl) upnEl.textContent = user?.userPrincipalName ?? user?.mail ?? '—'
      const dnEl = targetEl.querySelector('#od-user-displayname') as HTMLElement | null
      if (dnEl && user?.displayName) dnEl.textContent = user.displayName
    })
    // Populate drive URL if not already saved (e.g. auto-mapped before access was granted)
    if (!selectedDriveWebUrl) {
      getUserDrive(existing.targetSite.id).then(drive => {
        if (drive?.webUrl) {
          selectedDriveId = drive.id
          selectedDriveWebUrl = drive.webUrl
          const urlEl = targetEl.querySelector('#od-drive-url') as HTMLElement | null
          if (urlEl) urlEl.textContent = drive.webUrl
        }
      })
    }
  }

  // ── User search ────────────────────────────────────────────────────────────
  targetEl.querySelector('#btn-od-search')?.addEventListener('click', async () => {
    const query = (targetEl.querySelector('#od-user-search') as HTMLInputElement).value.trim()
    const resultsEl = targetEl.querySelector('#od-user-results') as HTMLElement
    resultsEl.innerHTML = '<span class="searching">Searching…</span>'
    try {
      const users = await searchUsers(query)
      if (users.length === 0) {
        resultsEl.innerHTML = '<span class="no-results">No users found.</span>'
        return
      }
      resultsEl.innerHTML = users.map((u) =>
        `<div class="site-result-item" data-uid="${escHtml(u.id)}">
          ${escHtml(u.displayName)}<br>
          <small>${escHtml(u.userPrincipalName ?? u.mail ?? '')}</small>
        </div>`
      ).join('')

      resultsEl.querySelectorAll('.site-result-item').forEach((item) => {
        item.addEventListener('click', async () => {
          const uid = item.getAttribute('data-uid')!
          selectedUser = users.find(u => u.id === uid) ?? null
          if (!selectedUser) return
          resultsEl.innerHTML = ''

          const badge = targetEl.querySelector('#od-selected-user') as HTMLElement
          badge.innerHTML = `✓ ${escHtml(selectedUser.displayName)} <button class="btn-clear" id="btn-clear-od-user">✕</button>`
          badge.style.display = ''
          attachClearUser()

          // Load drive info
          const driveInfo = targetEl.querySelector('#od-drive-info') as HTMLElement
          driveInfo.style.display = ''
          ;(targetEl.querySelector('#od-user-displayname') as HTMLElement).textContent = selectedUser.displayName
          ;(targetEl.querySelector('#od-user-upn') as HTMLElement).textContent = selectedUser.userPrincipalName ?? selectedUser.mail ?? '—'
          ;(targetEl.querySelector('#od-drive-url') as HTMLElement).textContent = '⏳ Loading…'
          ;(targetEl.querySelector('#od-access-status') as HTMLElement).textContent = '⏳ Checking…'

          const drive = await getUserDrive(selectedUser.id)
          selectedDriveId = drive?.id ?? ''
          selectedDriveWebUrl = drive?.webUrl ?? ''
          ;(targetEl.querySelector('#od-drive-url') as HTMLElement).textContent = selectedDriveWebUrl || '—'

          checkAndShowAccess(targetEl, selectedUser.id, migrationAccount)
        })
      })
    } catch {
      resultsEl.innerHTML = '<span class="no-results">Search failed.</span>'
    }
  })

  function attachClearUser(): void {
    targetEl.querySelector('#btn-clear-od-user')?.addEventListener('click', () => {
      selectedUser = null
      selectedDriveId = ''
      selectedDriveWebUrl = ''
      ;(targetEl.querySelector('#od-selected-user') as HTMLElement).style.display = 'none'
      ;(targetEl.querySelector('#od-drive-info') as HTMLElement).style.display = 'none'
    })
  }
  attachClearUser()

  // ── Save ──────────────────────────────────────────────────────────────────
  targetEl.querySelector('#btn-save-od-mapping')?.addEventListener('click', async () => {
    if (!selectedUser) { alert('Select a user first.'); return }
    const folderPath = (targetEl.querySelector('#od-folder-path') as HTMLInputElement).value.trim()

    const mapping: MigrationMapping = {
      id: node.path,
      sourceNode: node,
      targetSite: { id: selectedUser.id, displayName: selectedUser.displayName, webUrl: selectedDriveWebUrl, name: selectedUser.displayName },
      targetDrive: selectedDriveId ? { id: selectedDriveId, name: 'OneDrive', webUrl: selectedDriveWebUrl, driveType: 'personal' } : null,
      targetFolderPath: folderPath,
      status: 'ready',
    }

    const mappings = [...getState().mappings.filter(m => m.sourceNode.path !== node.path), mapping]
    setState({ mappings })
    onMappingChange(selectedUser.displayName, false)

    const saveBtn = targetEl.querySelector('#btn-save-od-mapping') as HTMLButtonElement
    saveBtn.disabled = true
    saveBtn.textContent = 'Saving…'
    try {
      await persistMappings(mappings)
      saveBtn.textContent = '✓ Saved'
    } catch {
      saveBtn.textContent = '⚠ Save failed — retry'
    } finally {
      saveBtn.disabled = false
      setTimeout(() => { if (saveBtn.textContent !== '⚠ Save failed — retry') saveBtn.textContent = 'Save Mapping' }, 2000)
    }
  })

  // ── Remove ────────────────────────────────────────────────────────────────
  targetEl.querySelector('#btn-remove-od-mapping')?.addEventListener('click', async () => {
    const removeBtn = targetEl.querySelector('#btn-remove-od-mapping') as HTMLButtonElement
    removeBtn.disabled = true
    removeBtn.textContent = 'Removing…'
    const mappings = getState().mappings.filter(m => m.sourceNode.path !== node.path)
    setState({ mappings })
    try {
      await persistMappings(mappings)
    } catch {
      removeBtn.disabled = false
      removeBtn.textContent = 'Remove'
      return
    }
    onMappingChange(null)
    removeBtn.remove()
  })
}

async function checkAndShowAccess(targetEl: HTMLElement, userId: string, migrationAccount: string): Promise<void> {
  const statusEl = targetEl.querySelector('#od-access-status') as HTMLElement | null
  if (!statusEl) return
  statusEl.textContent = '⏳ Checking…'
  statusEl.style.color = ''
  try {
    const access = await checkUserDriveAccess(userId)
    if (access === 'accessible') {
      statusEl.textContent = '✓ Accessible'
      statusEl.style.color = 'var(--color-success, #107c10)'
    } else if (access === 'no-access') {
      statusEl.style.color = 'var(--color-danger, #a4262c)'
      if (migrationAccount) {
        statusEl.innerHTML = `✗ No access (or OneDrive not provisioned) &nbsp;<button type="button" id="btn-grant-access" class="btn btn-sm btn-warning" style="font-size:0.75rem;padding:2px 8px;margin-left:4px;">Grant Access</button>`
        statusEl.querySelector('#btn-grant-access')?.addEventListener('click', async () => {
          const btn = statusEl.querySelector('#btn-grant-access') as HTMLButtonElement
          btn.disabled = true
          btn.textContent = 'Granting…'
          try {
            await grantUserDriveAccess(userId, migrationAccount)
            // Re-fetch drive URL now that we have access
            getUserDrive(userId).then(drive => {
              if (drive?.webUrl) {
                const urlEl = targetEl.querySelector('#od-drive-url') as HTMLElement | null
                if (urlEl) urlEl.textContent = drive.webUrl
              }
            })
            await checkAndShowAccess(targetEl, userId, migrationAccount)
          } catch (err) {
            btn.disabled = false
            const msg = (err as Error)?.message ?? String(err)
            btn.textContent = '⚠ Failed — retry'
            btn.title = msg
            // Also show the error inline so it's visible without hovering
            const errEl = document.createElement('div')
            errEl.style.cssText = 'color:var(--color-danger,#a4262c);font-size:0.75rem;margin-top:4px;word-break:break-word;'
            errEl.textContent = msg
            statusEl.appendChild(errEl)
          }
        })
      } else {
        statusEl.textContent = '✗ No access or OneDrive not provisioned — configure Migration Account in settings to grant'
      }
    } else if (access === 'no-drive') {
      statusEl.textContent = '✗ No OneDrive provisioned'
      statusEl.style.color = 'var(--color-danger, #a4262c)'
    } else {
      statusEl.textContent = '⚠ Could not check'
      statusEl.style.color = 'var(--color-danger, #a4262c)'
    }
  } catch {
    statusEl.textContent = '⚠ Could not check'
    statusEl.style.color = 'var(--color-danger, #a4262c)'
  }
}

async function loadLibraries(
  targetEl: HTMLElement,
  site: SharePointSite,
  selected: SharePointDrive | null
): Promise<void> {
  const libGroup = targetEl.querySelector('#library-group') as HTMLElement
  const libSelect = targetEl.querySelector('#library-select') as HTMLSelectElement
  libGroup.style.display = ''
  libSelect.innerHTML = '<option>Loading…</option>'
  try {
    const drives = await getSiteDrives(site.id)
    libSelect.innerHTML = drives
      .map((d) => `<option value="${escHtml(d.id)}" ${selected?.id === d.id ? 'selected' : ''}>${escHtml(d.name)}</option>`)
      .join('')
  } catch {
    libSelect.innerHTML = '<option>Failed to load libraries</option>'
  }
}

async function persistMappings(mappings: MigrationMapping[]): Promise<void> {
  const project = getState().currentProject
  if (!project) return

  const hasUploadFolder = (project.projectData.uploads?.length ?? 0) > 0

  if (hasUploadFolder) {
    // New model: store mappings as a separate file to avoid SP column size limits.
    // sourceNode.children are stripped by saveMappingsFile — they are already in .tree.json.
    const { siteId } = getSpConfig()
    await saveMappingsFile(siteId, project.title, project.id, mappings)

    // Remove any inline mappings from ProjectData but keep a denormalized count
    // so the project list scorecard can display the correct number without loading the file.
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { mappings: _removed, ...restData } = project.projectData
    // Count only entries with actual targets (not Phase 1 not-found/ambiguous entries)
    const mappedCount = mappings.filter(m => m.targetSite || m.plannedSite).length
    const updatedProjectData = { ...restData, mappingCount: mappedCount }
    await updateProject(project.id, { projectData: updatedProjectData })
    setState({ mappings, currentProject: { ...project, projectData: updatedProjectData } })
  } else {
    // Legacy model (no upload folder yet): store inline as before; keep count in sync too.
    const mappedCount = mappings.filter(m => m.targetSite || m.plannedSite).length
    const updatedProjectData = { ...project.projectData, mappings, mappingCount: mappedCount }
    await updateProject(project.id, { projectData: updatedProjectData })
    setState({ mappings, currentProject: { ...project, projectData: updatedProjectData } })
  }
}

// ─── Auto-expand helpers ───────────────────────────────────────────────────────

/**
 * Expands tree nodes so that every mapped node is visible, without expanding
 * the mapped nodes themselves (their children carry the "parent mapped" style).
 */
function autoExpandToMappedNodes(
  rootUl: HTMLUListElement,
  topNodes: TreeNode[],
  mappedPaths: Set<string>
): void {
  for (const mappedPath of mappedPaths) {
    const ancestors = findAncestorPaths(topNodes, mappedPath)
    if (!ancestors) continue
    let container: Element = rootUl
    for (const ancestorPath of ancestors) {
      const li = findDirectChildLi(container, ancestorPath)
      if (!li) break
      if (!li.classList.contains('mapping-node--open')) {
        li.querySelector<HTMLButtonElement>(':scope > .mapping-row > .mapping-toggle-btn:not(.invisible)')?.click()
      }
      const childUl = li.querySelector<HTMLElement>(':scope > .tree-children')
      if (childUl) container = childUl
      else break
    }
  }
}

/** Returns the ordered list of ancestor paths (excluding the target) needed to reach targetPath. */
function findAncestorPaths(nodes: TreeNode[], targetPath: string): string[] | null {
  for (const node of nodes) {
    if (node.path === targetPath) return []
    if (node.children.length > 0) {
      const sub = findAncestorPaths(node.children, targetPath)
      if (sub !== null) return [node.path, ...sub]
    }
  }
  return null
}

/** Finds a direct li.mapping-node child of the given UL whose mapping-row has the given path. */
function findDirectChildLi(ul: Element, path: string): HTMLLIElement | null {
  for (const li of Array.from(ul.children)) {
    const row = (li as HTMLElement).querySelector<HTMLElement>(':scope > .mapping-row')
    if (row?.dataset.path === path) return li as HTMLLIElement
  }
  return null
}

// ─── Ancestor-mapped block helpers ────────────────────────────────────────────

/** Returns the nearest ancestor mapping for a given node path, or null if none. */
function findAncestorMapping(nodePath: string): MigrationMapping | null {
  for (const m of getState().mappings) {
    const sp = m.sourceNode.path
    if (sp && nodePath !== sp && (m.targetSite || m.plannedSite)) {
      if (nodePath.startsWith(sp + '\\') || nodePath.startsWith(sp + '/')) {
        return m
      }
    }
  }
  return null
}

/** Renders the "parent already mapped" info panel in place of the mapping form. */
function openBlockedPanel(targetEl: HTMLElement, node: TreeNode, ancestor: MigrationMapping): void {
  const mappedName = ancestor.targetSite?.displayName ?? ancestor.plannedSite?.displayName ?? '(unknown)'
  const isPlanned = !ancestor.targetSite && !!ancestor.plannedSite

  let destinationHtml = ''
  if (ancestor.targetSite) {
    const rel = computeRelativePath(node.path, ancestor.sourceNode.path)
    const library = ancestor.targetDrive?.name ?? 'Shared Documents'
    const parts: string[] = [ancestor.targetSite.webUrl.replace(/\/$/, ''), library]
    if (ancestor.targetFolderPath) parts.push(ancestor.targetFolderPath.replace(/^[/\\]+/, ''))
    if (rel) parts.push(rel)
    const url = parts.join('/')
    destinationHtml = `<a href="${escHtml(url)}" target="_blank" class="ancestor-url">${escHtml(url)}</a>`
  } else if (ancestor.plannedSite) {
    const ps = ancestor.plannedSite
    const rel = computeRelativePath(node.path, ancestor.sourceNode.path)
    const library = ps.libraryName || 'Documents'
    const parts: string[] = [`[Planned] …/sites/${ps.alias}`, library]
    if (ps.folderPath) parts.push(ps.folderPath.replace(/^[/\\]+/, ''))
    if (rel) parts.push(rel)
    destinationHtml = `<span class="ancestor-url ancestor-url--planned">${escHtml(parts.join('/'))}</span>`
  }

  targetEl.innerHTML = `
    <div class="ancestor-blocked-panel">
      <div class="ancestor-blocked-icon">🔒</div>
      <h4 class="ancestor-blocked-title">Parent folder is already mapped</h4>
      <p class="ancestor-blocked-msg">
        <strong>${escHtml(String(node.name || node.path))}</strong> is a subfolder of a mapped location
        and cannot be mapped separately.
      </p>
      <div class="ancestor-blocked-info">
        <div class="ancestor-info-row">
          <span class="ancestor-info-label">Mapped to</span>
          <span class="ancestor-info-value">${escHtml(mappedName)}${isPlanned ? ' <em>(planned)</em>' : ''}</span>
        </div>
        <div class="ancestor-info-row ancestor-info-row--url">
          <span class="ancestor-info-label">Destination URL</span>
          <div class="ancestor-info-value">${destinationHtml || '—'}</div>
        </div>
      </div>
    </div>
  `
}

function computeRelativePath(nodePath: string, ancestorPath: string): string {
  if (nodePath.startsWith(ancestorPath)) {
    return nodePath.slice(ancestorPath.length).replace(/^[/\\]+/, '').replace(/\\/g, '/')
  }
  return ''
}

// ─── Descendant highlight propagation ─────────────────────────────────────────

function updateDescendantHighlights(parentLi: HTMLLIElement, parentIsMapped: boolean): void {
  const childUl = parentLi.querySelector<HTMLElement>(':scope > .tree-children')
  if (!childUl) return
  childUl.querySelectorAll<HTMLLIElement>(':scope > .mapping-node').forEach((childLi) => {
    const childRow = childLi.querySelector<HTMLElement>(':scope > .mapping-row')
    if (!childRow) return
    const childPath = childRow.dataset.path ?? ''
    const childSelfMapped = !!getState().mappings.find((m) => m.sourceNode.path === childPath && m.targetSite)
    const shouldBeMapped = parentIsMapped || childSelfMapped
    if (shouldBeMapped) {
      childRow.classList.add('mapping-row--mapped')
    } else {
      childRow.classList.remove('mapping-row--mapped')
    }
    updateDescendantHighlights(childLi, shouldBeMapped)
  })
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function collectAtDepth(root: TreeNode, depth: number): TreeNode[] {
  const result: TreeNode[] = []
  function walk(node: TreeNode): void {
    if (node.depth === depth) { result.push(node); return }
    if (node.depth < depth) node.children.forEach(walk)
  }
  walk(root)
  return result
}

function refreshUsersStats(container: HTMLElement, statNodes: TreeNode[]): void {
  const currentMappings = getState().mappings

  // User counts
  const mappedPaths = new Set(currentMappings.filter(m => m.targetSite || m.plannedSite).map(m => m.sourceNode.path))
  const ready = statNodes.filter(n => mappedPaths.has(n.path)).length
  const notMapped = statNodes.length - ready
  const readyEl = container.querySelector('#mstat-users-ready-val')
  const unmappedEl = container.querySelector('#mstat-users-unmapped-val')
  if (readyEl) readyEl.textContent = `${ready} ready to Migrate`
  if (unmappedEl) unmappedEl.textContent = `${notMapped} not Mapped`

  // Double-mapped detection
  const targetToNodePaths = new Map<string, string[]>()
  for (const m of currentMappings) {
    if (statNodes.some(n => n.path === m.sourceNode.path) && (m.targetSite || m.resolvedDisplayName)) {
      const key = m.targetSite?.id ?? m.resolvedDisplayName ?? ''
      if (key) {
        if (!targetToNodePaths.has(key)) targetToNodePaths.set(key, [])
        targetToNodePaths.get(key)!.push(m.sourceNode.path)
      }
    }
  }
  _doubleMappedPaths = new Set([...targetToNodePaths.values()].filter(p => p.length > 1).flat())
  const dmCount = [...targetToNodePaths.values()].filter(p => p.length > 1).length

  // Update warn icons on rendered rows
  warnRegistry.forEach((el, path) => {
    const isDM = _doubleMappedPaths.has(path)
    el.textContent = isDM ? '⚠' : ''
  })

  // Update stats card warning
  const dmWarnEl = container.querySelector<HTMLElement>('#mstat-double-mapped-warn')
  if (dmWarnEl) {
    dmWarnEl.textContent = dmCount > 0 ? `⚠ ${dmCount} user${dmCount !== 1 ? 's' : ''} double mapped` : ''
    dmWarnEl.style.display = dmCount > 0 ? '' : 'none'
  }
}

function getRecycleBin(node: TreeNode): { sizeBytes: number; fileCount: number } {
  let sizeBytes = 0, fileCount = 0
  function walk(n: TreeNode): void {
    if (/^\$recycle\.bin$/i.test(n.name) || /^recycler$/i.test(n.name)) {
      sizeBytes += n.sizeBytes
      fileCount += n.fileCount
      return // don't recurse inside the recycle bin itself
    }
    for (const child of n.children) walk(child)
  }
  for (const child of node.children) walk(child)
  return { sizeBytes, fileCount }
}

interface MappingStats {
  userCount: number; totalBytes: number; migrateBytes: number
  totalFiles: number; migrateFiles: number; recycleBinBytes: number; recycleBinFiles: number
}

function buildMappingStats(nodes: TreeNode[]): MappingStats {
  let totalBytes = 0, totalFiles = 0, recycleBinBytes = 0, recycleBinFiles = 0
  for (const n of nodes) {
    const rb = getRecycleBin(n)
    totalBytes += n.sizeBytes
    totalFiles += n.fileCount
    recycleBinBytes += rb.sizeBytes
    recycleBinFiles += rb.fileCount
  }
  return {
    userCount: nodes.length,
    totalBytes,
    migrateBytes: Math.max(0, totalBytes - recycleBinBytes),
    totalFiles,
    migrateFiles: Math.max(0, totalFiles - recycleBinFiles),
    recycleBinBytes,
    recycleBinFiles,
  }
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return ''
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectMappingStyles(): void {
  if (document.getElementById('mapping-styles')) return
  const style = document.createElement('style')
  style.id = 'mapping-styles'
  style.textContent = `
    .mapping-empty { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .mapping-search-bar { padding: 8px 12px; border-bottom: 1px solid var(--color-border); background: white; }
    .search-input-wrap { position: relative; display: flex; align-items: center; }
    .mapping-search-input { flex: 1; box-sizing: border-box; padding: 6px 32px 6px 10px; font-size: 0.85rem; }
    .btn-clear-search {
      position: absolute; right: 6px; background: none; border: none; cursor: pointer;
      color: var(--color-text-muted); font-size: 0.85rem; line-height: 1; padding: 2px 4px;
      border-radius: 3px;
    }
    .btn-clear-search:hover { background: var(--color-surface-alt); color: var(--color-text); }
    .mapping-search-empty { padding: 16px; color: var(--color-text-muted); font-size: 0.875rem; text-align: center; }
    .search-result-path { display: block; font-size: 0.72rem; color: var(--color-text-muted);
      font-family: 'Consolas', monospace; padding: 0 8px 4px 46px; white-space: nowrap;
      overflow: hidden; text-overflow: ellipsis; }
    .mapping-panel { display: grid; grid-template-columns: 2fr 1fr; height: calc(100vh - 140px); overflow: hidden; }
    .mapping-left, .mapping-right { overflow-y: auto; border-right: 1px solid var(--color-border); }
    .mapping-right { border-right: none; }
    .mapping-section-header { padding: 12px 16px; border-bottom: 1px solid var(--color-border);
      display: flex; align-items: center; justify-content: space-between; background: var(--color-surface-alt);
      position: sticky; top: 0; z-index: 1; }
    .mapping-section-header h3 { font-size: 0.9rem; font-weight: 600; margin: 0; }
    .mapping-hint { font-size: 0.78rem; color: var(--color-text-muted); }

    /* Tree */
    .mapping-tree { padding: 8px; }
    .tree-list { list-style: none; padding: 0; margin: 0; }
    .tree-children { padding-left: 20px; border-left: 1px solid var(--color-border); margin-left: 18px; }
    .mapping-node { margin: 1px 0; }
    .mapping-node--root > .mapping-row { font-weight: 600; }

    /* Row */
    .mapping-row { display: flex; align-items: center; gap: 6px; padding: 5px 8px; border-radius: 4px;
      user-select: none; transition: background 0.1s; cursor: default; }
    .mapping-row[data-path]:not([data-path=""]) { cursor: pointer; }
    .mapping-row:hover { background: var(--color-primary-light); }
    .mapping-row--active { background: var(--color-primary-light); border-left: 3px solid var(--color-primary); }

    .mapping-toggle-btn { background: none; border: none; cursor: pointer; width: 16px;
      font-size: 0.65rem; color: var(--color-text-muted); padding: 0; flex-shrink: 0; }
    .mapping-toggle-btn.invisible { visibility: hidden; pointer-events: none; }
    .toggle-icon { display: block; }

    /* Folder icon with optional mapped-badge */
    .tree-icon-wrap { position: relative; display: inline-flex; flex-shrink: 0; line-height: 1; }
    .mapped-folder-badge {
      position: absolute; bottom: -2px; right: -5px;
      font-size: 0.48rem; font-style: normal; font-weight: 700;
      background: #107c10; color: white; border-radius: 50%;
      width: 9px; height: 9px; display: flex; align-items: center; justify-content: center;
      border: 1px solid white;
    }

    /* Mapped row highlighting */
    .mapping-row--mapped { background: rgba(16, 124, 16, 0.07); }
    .mapping-row--mapped:hover { background: rgba(16, 124, 16, 0.13); }
    .mapping-row--mapped.mapping-row--active { background: rgba(16, 124, 16, 0.13); border-left-color: #107c10; }

    .tree-name { flex: 1; font-size: 0.875rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
    .tree-name--loose { font-style: italic; color: var(--color-text-muted); }
    /* Mapped-to column (replaces floating tag) */
    .tree-col-mapped { width: 140px; text-align: right; font-size: 0.78rem; font-weight: 500;
      color: #107c10; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex-shrink: 0; }
    .tree-col-mapped--planned { color: #7a5900; }
    .tree-col-mapped--empty { color: var(--color-text-muted); font-weight: 400; }
    .tch-col-mapped { width: 140px; }
    /* Double-mapped warning icon on row */
    .row-warn-icon { font-size: 0.7rem; color: #d83b01; flex-shrink: 0; min-width: 12px;
      cursor: help; line-height: 1; }
    /* Double-mapped warning in stats card */
    .mstat-double-mapped-warn { color: #d83b01 !important; font-weight: 700; }

    /* Stats bar */
    .mapping-stats-bar { display: flex; gap: 0; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface); overflow-x: auto; flex-shrink: 0; }
    .mstat-card { flex: 1; min-width: 110px; padding: 8px 12px;
      border-right: 1px solid var(--color-border); }
    .mstat-card:last-child { border-right: none; }
    .mstat-card--danger { border-left: 3px solid var(--color-danger, #a4262c); }
    .mstat-label { font-size: 0.6rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 1px; }
    .mstat-value { font-size: 1.1rem; font-weight: 700; line-height: 1.15; }
    .mstat-sub { font-size: 0.65rem; color: var(--color-text-muted); margin-top: 1px; }
    .mstat-not-mapped { color: #b35c00; font-weight: 600; }
    .mstat-blue { color: var(--color-primary); }
    .mstat-green { color: #107c10; }
    .mstat-orange { color: #d83b01; }
    .mstat-red { color: var(--color-danger, #a4262c); }
    .mstat-recycle-bar { height: 3px; background: var(--color-border); border-radius: 2px;
      margin-top: 5px; overflow: hidden; }
    .mstat-recycle-fill { height: 100%; background: var(--color-danger, #a4262c); border-radius: 2px; }

    /* Column header */
    .tree-col-header { display: flex; align-items: center; padding: 4px 8px 4px 0;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border);
      font-size: 0.62rem; font-weight: 700; color: var(--color-text-muted);
      text-transform: uppercase; letter-spacing: 0.05em;
      position: sticky; top: 0; z-index: 1; flex-shrink: 0; }
    .tch-name { flex: 1; padding-left: 46px; white-space: nowrap; }
    .tch-col { width: 90px; text-align: right; flex-shrink: 0; padding-right: 8px; white-space: nowrap; }

    /* Tree column cells */
    .tree-col { font-size: 0.75rem; color: var(--color-text-muted); white-space: nowrap;
      flex-shrink: 0; width: 90px; text-align: right; }
    .tree-col-rb--has-rb { background: rgba(255, 140, 0, 0.15); color: #b35c00; font-weight: 600;
      padding: 1px 5px; border-radius: 3px; }
    .tree-col-migrate { color: var(--color-text); font-weight: 500; }

    /* Detail grid recycle bin */
    .detail-recycle { color: #b35c00; font-weight: 600; }

    /* Target panel */
    .mapping-placeholder { padding: 32px; text-align: center; color: var(--color-text-muted); font-size: 0.88rem; }
    .target-panel { display: flex; flex-direction: column; }

    /* Two-section layout */
    .target-section { border-bottom: 1px solid var(--color-border); }
    .target-section:last-child { border-bottom: none; }

    .target-section-toggle {
      width: 100%; display: flex; align-items: center; justify-content: space-between;
      padding: 12px 16px; background: var(--color-surface-alt);
      border: none; border-bottom: 1px solid var(--color-border);
      cursor: pointer; text-align: left; font-family: inherit;
    }
    .target-section-toggle:hover { background: var(--color-primary-light); }

    /* Tab bar */
    .sp-tabs-bar { display: flex; border-bottom: 1px solid var(--color-border); background: var(--color-surface-alt); }
    .sp-tab {
      flex: 1; padding: 10px 14px; background: none; border: none;
      border-bottom: 2px solid transparent; cursor: pointer; font-size: 0.8rem; font-weight: 500;
      color: var(--color-text-muted); font-family: inherit; text-align: center;
      transition: color 0.15s, border-color 0.15s;
    }
    .sp-tab:hover { color: var(--color-text); background: var(--color-primary-light); }
    .sp-tab--active { color: var(--color-primary); border-bottom-color: var(--color-primary); font-weight: 600; }

    /* Planned mapping tag */
    .mapping-tag--planned { background: #fff4ce; color: #7a5900; }

    /* Planned form helpers (mirrors siteCreator styles) */
    .alias-row { display: flex; align-items: center; gap: 0; }
    .alias-prefix { background: var(--color-surface-alt); border: 1px solid var(--color-border);
      border-right: none; padding: 8px 10px; border-radius: 4px 0 0 4px; font-size: 0.85rem;
      color: var(--color-text-muted); white-space: nowrap; }
    .alias-row .form-input { border-radius: 0 4px 4px 0; }
    .template-row { margin-bottom: 4px; }
    .radio-label { display: flex; align-items: center; gap: 6px; font-size: 0.88rem; cursor: pointer; }
    .required { color: var(--color-danger); }

    .target-section-title { font-size: 0.9rem; font-weight: 600; color: var(--color-text); }
    .target-section-chevron { font-size: 0.7rem; color: var(--color-text-muted); flex-shrink: 0; }

    .target-section-body { }
    .target-section-body--sp { padding: 16px; display: flex; flex-direction: column; gap: 16px; }
    .target-section-body--sp .form-group { margin-bottom: 0; }
    .target-action-row { display: flex; gap: 8px; padding-top: 4px; }

    /* Source detail card */
    .source-detail-card { background: var(--color-surface-alt); border: 1px solid var(--color-border);
      border-radius: 6px; overflow: hidden; }
    .source-detail-title { display: flex; align-items: center; gap: 8px; padding: 10px 14px;
      border-bottom: 1px solid var(--color-border); background: var(--color-surface); }
    .source-detail-icon { font-size: 1.1rem; flex-shrink: 0; }
    .source-detail-name { font-weight: 600; font-size: 0.9rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
    .source-detail-grid { display: grid; grid-template-columns: auto 1fr; gap: 0; margin: 0; padding: 0; }
    .source-detail-grid dt, .source-detail-grid dd {
      padding: 6px 14px; margin: 0; font-size: 0.82rem;
      border-bottom: 1px solid var(--color-border); }
    .source-detail-grid dt:last-of-type, .source-detail-grid dd:last-of-type { border-bottom: none; }
    .source-detail-grid dt { color: var(--color-text-muted); font-weight: 500; white-space: nowrap;
      background: var(--color-surface); border-right: 1px solid var(--color-border); }
    .source-detail-grid dd { font-family: 'Consolas', monospace; word-break: break-all; }
    .source-detail-path { font-size: 0.78rem; color: var(--color-text-muted); }
    .site-search-row { display: flex; gap: 8px; }
    .site-results { margin-top: 8px; border: 1px solid var(--color-border); border-radius: 4px;
      max-height: 200px; overflow-y: auto; }
    .site-result-item { padding: 8px 12px; cursor: pointer; font-size: 0.85rem;
      border-bottom: 1px solid var(--color-border); }
    .site-result-item:last-child { border-bottom: none; }
    .site-result-item:hover { background: var(--color-primary-light); }
    .site-result-item small { color: var(--color-text-muted); }
    .selected-badge { background: #dff6dd; color: #107c10; padding: 6px 10px; border-radius: 4px;
      font-size: 0.85rem; margin-top: 8px; display: flex; align-items: center; justify-content: space-between; }
    .btn-clear { background: none; border: none; cursor: pointer; color: inherit; font-size: 0.9rem; }
    .searching, .no-results { padding: 8px 12px; font-size: 0.85rem; color: var(--color-text-muted); display: block; }
    .hint { font-size: 0.78rem; color: var(--color-text-muted); font-weight: 400; }

    /* Ancestor-blocked panel */
    .ancestor-blocked-panel {
      display: flex; flex-direction: column; align-items: flex-start; gap: 14px;
      padding: 24px; }
    .ancestor-blocked-icon { font-size: 1.8rem; line-height: 1; }
    .ancestor-blocked-title { font-size: 1rem; font-weight: 600; color: var(--color-text); margin: 0; }
    .ancestor-blocked-msg { font-size: 0.875rem; color: var(--color-text-muted); margin: 0; line-height: 1.5; }
    .ancestor-blocked-info {
      width: 100%; background: var(--color-surface-alt); border: 1px solid var(--color-border);
      border-radius: 6px; overflow: hidden; }
    .ancestor-info-row {
      display: grid; grid-template-columns: 110px 1fr; gap: 8px; align-items: baseline;
      padding: 10px 14px; border-bottom: 1px solid var(--color-border); }
    .ancestor-info-row:last-child { border-bottom: none; }
    .ancestor-info-row--url { align-items: start; }
    .ancestor-info-label { font-size: 0.8rem; font-weight: 600; color: var(--color-text-muted); white-space: nowrap; }
    .ancestor-info-value { font-size: 0.85rem; color: var(--color-text); word-break: break-all; }
    .ancestor-url {
      font-family: 'Consolas', monospace; font-size: 0.8rem; color: var(--color-primary);
      text-decoration: none; word-break: break-all; display: block; }
    .ancestor-url:hover { text-decoration: underline; }
    .ancestor-url--planned { color: var(--color-text-muted); font-style: italic; }

    /* OneDrive drive info card */
    .od-drive-card { background: var(--color-surface-alt); border: 1px solid var(--color-border);
      border-radius: 6px; overflow: hidden; margin-bottom: 0; }
    .od-drive-row { display: grid; grid-template-columns: 100px 1fr; gap: 8px; align-items: baseline;
      padding: 8px 14px; border-bottom: 1px solid var(--color-border); }
    .od-drive-row:last-child { border-bottom: none; }
    .od-drive-label { font-size: 0.8rem; font-weight: 600; color: var(--color-text-muted); white-space: nowrap; }
    .od-drive-value { font-size: 0.82rem; color: var(--color-text); word-break: break-all;
      font-family: 'Consolas', monospace; }

    /* Subfolder mode */
    .subfolder-mode-row { display: flex; flex-direction: column; gap: 4px; margin-bottom: 8px; }
    .subfolder-default-code { font-family: 'Consolas', monospace; font-size: 0.82rem;
      background: var(--color-surface-alt); padding: 1px 5px; border-radius: 3px;
      border: 1px solid var(--color-border); }
  `
  document.head.appendChild(style)
}
