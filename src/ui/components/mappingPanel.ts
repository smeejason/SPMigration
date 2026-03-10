import { searchSites, getSiteDrives } from '../../graph/graphClient'
import { updateProject } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { TreeNode, MigrationMapping, SharePointSite, SharePointDrive } from '../../types'

// Live references to mapping tag elements so we can update them without re-rendering
const tagRegistry = new Map<string, HTMLSpanElement>()

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

  container.innerHTML = `
    <div class="mapping-panel">
      <div class="mapping-left">
        <div class="mapping-section-header">
          <h3>Source: File System</h3>
          <span class="mapping-hint">Click a folder to map it</span>
        </div>
        <div class="mapping-search-bar">
          <div class="search-input-wrap">
            <input type="text" id="tree-search" class="form-input mapping-search-input" placeholder="Search by name or path… (press Enter)" autocomplete="off" />
            <button type="button" id="btn-clear-search" class="btn-clear-search" style="display:none" title="Clear search">✕</button>
          </div>
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

  const treeEl = container.querySelector('#mapping-tree') as HTMLElement
  const targetEl = container.querySelector('#mapping-target') as HTMLElement

  const ul = document.createElement('ul')
  ul.className = 'tree-list tree-root'

  // If the top of the tree is a synthetic root (empty path), skip it and render
  // its children directly so the user sees their actual top-level folder(s) first.
  const topNodes = !tree.path ? tree.children : [tree]
  for (const node of topNodes) {
    ul.appendChild(createMappingNodeEl(node, targetEl, true))
  }

  treeEl.appendChild(ul)

  // Auto-expand if there is only one top-level node
  if (topNodes.length === 1) {
    const rootToggle = ul.querySelector<HTMLButtonElement>('.mapping-toggle-btn:not(.invisible)')
    rootToggle?.click()
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

  function clearSearch(): void {
    searchInput.value = ''
    clearSearchBtn.style.display = 'none'
    treeDiv.style.display = ''
    resultsDiv.style.display = 'none'
    resultsDiv.innerHTML = ''
  }

  function runSearch(): void {
    const term = searchInput.value.trim().toLowerCase()
    if (!term) { clearSearch(); return }

    const matches = allNodes.filter(
      (n) => n.name.toLowerCase().includes(term) || n.originalPath.toLowerCase().includes(term) || n.path.toLowerCase().includes(term)
    )

    treeDiv.style.display = 'none'
    resultsDiv.style.display = ''
    clearSearchBtn.style.display = ''

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

function createMappingNodeEl(node: TreeNode, targetEl: HTMLElement, isRoot = false): HTMLLIElement {
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
  nameEl.className = 'tree-name'
  nameEl.textContent = String(node.name || node.path || '(unnamed)')
  if (node.originalPath) nameEl.title = node.originalPath

  // Size
  const sizeEl = document.createElement('span')
  sizeEl.className = 'tree-size-sm'
  sizeEl.textContent = formatBytes(node.sizeBytes)

  // Mapping tag (shows which site this folder is mapped to)
  const tagEl = document.createElement('span')
  tagEl.className = 'mapping-tag'
  const existingMapping = getState().mappings.find((m) => m.sourceNode.path === node.path)
  tagRegistry.set(node.path, tagEl)

  // Helper: apply/remove the mapped visual state on this row
  function applyMappedState(isMapped: boolean, siteName?: string): void {
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
      tagEl.textContent = siteName ? `→ ${siteName}` : ''
      tagEl.style.display = siteName ? '' : 'none'
    } else {
      row.classList.remove('mapping-row--mapped')
      tagEl.style.display = 'none'
    }
  }

  const isMappedInitially = !!existingMapping?.targetSite
  applyMappedState(isMappedInitially, existingMapping?.targetSite?.displayName)

  row.appendChild(toggleBtn)
  row.appendChild(iconWrap)
  row.appendChild(nameEl)
  row.appendChild(sizeEl)
  row.appendChild(tagEl)
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
          for (const child of node.children) {
            childUl.appendChild(createMappingNodeEl(child, targetEl))
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
      openTargetPanel(targetEl, node, (siteName) => {
        applyMappedState(!!siteName, siteName ?? undefined)
      })
    })
  }

  return li
}

// ─── Target panel (right side) ────────────────────────────────────────────────

async function openTargetPanel(
  targetEl: HTMLElement,
  node: TreeNode,
  onMappingChange: (siteName: string | null) => void
): Promise<void> {
  const existing = getState().mappings.find((m) => m.sourceNode.path === node.path)

  const fmtDate = (d?: Date) =>
    d ? d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '—'
  const lastModStr = fmtDate(node.lastModified)
  const lastAccStr = fmtDate(node.lastAccessed)
  const sizeStr = node.sizeBytes > 0 ? formatBytes(node.sizeBytes) : '—'
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
              <dt>Size</dt>
              <dd>${sizeStr}</dd>
              <dt>Files</dt>
              <dd>${fileStr}</dd>
              <dt>Subfolders</dt>
              <dd>${folderStr}</dd>
              <dt>Direct Children</dt>
              <dd>${childStr}</dd>
              <dt>Last Modified</dt>
              <dd>${lastModStr}</dd>
              <dt>Last Accessed</dt>
              <dd>${lastAccStr}</dd>
            </dl>
          </div>
        </div>
      </div>

      <!-- ── Section 2: Target SharePoint Location ── -->
      <div class="target-section">
        <div class="target-section-header-static">
          <span class="target-section-title">Target SharePoint Location</span>
        </div>
        <div class="target-section-body target-section-body--sp">

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
            ${existing ? `<button type="button" id="btn-remove-mapping" class="btn btn-ghost">Remove</button>` : ''}
          </div>

        </div>
      </div>

    </div>
  `

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

  targetEl.querySelector('#btn-save-mapping')?.addEventListener('click', () => {
    const folderPath = (targetEl.querySelector('#folder-path') as HTMLInputElement).value.trim()

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
    persistMappings(mappings)
    onMappingChange(selectedSite?.displayName ?? null)

    const saveBtn = targetEl.querySelector('#btn-save-mapping') as HTMLButtonElement
    saveBtn.textContent = '✓ Saved'
    setTimeout(() => { saveBtn.textContent = 'Save Mapping' }, 2000)
  })

  targetEl.querySelector('#btn-remove-mapping')?.addEventListener('click', () => {
    const mappings = getState().mappings.filter((m) => m.sourceNode.path !== node.path)
    setState({ mappings })
    persistMappings(mappings)
    onMappingChange(null)
    targetEl.querySelector('#btn-remove-mapping')?.remove()
  })
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
  try {
    await updateProject(project.id, {
      projectData: { ...project.projectData, mappings },
    })
  } catch {
    console.warn('[Mapping] Could not persist mappings to SharePoint')
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

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
    .mapping-panel { display: grid; grid-template-columns: 1fr 1fr; height: calc(100vh - 140px); overflow: hidden; }
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
    .tree-size-sm { font-size: 0.75rem; color: var(--color-text-muted); white-space: nowrap; flex-shrink: 0; }
    .mapping-tag { font-size: 0.72rem; background: #dff6dd; color: #107c10; padding: 2px 6px;
      border-radius: 10px; white-space: nowrap; flex-shrink: 0; }

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

    .target-section-header-static {
      display: flex; align-items: center; padding: 12px 16px;
      background: var(--color-surface-alt); border-bottom: 1px solid var(--color-border);
    }

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
  `
  document.head.appendChild(style)
}
