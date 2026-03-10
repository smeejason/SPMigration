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
          <input type="text" id="tree-search" class="form-input mapping-search-input" placeholder="Search by name or path…" autocomplete="off" />
        </div>
        <div id="mapping-tree" class="mapping-tree"></div>
        <div id="mapping-search-results" class="mapping-tree" style="display:none"></div>
      </div>
      <div class="mapping-right">
        <div class="mapping-section-header">
          <h3>Target: SharePoint</h3>
        </div>
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

  searchInput.addEventListener('input', () => {
    const term = searchInput.value.trim().toLowerCase()
    if (!term) {
      treeDiv.style.display = ''
      resultsDiv.style.display = 'none'
      resultsDiv.innerHTML = ''
      return
    }

    const matches = allNodes.filter(
      (n) => n.name.toLowerCase().includes(term) || n.path.toLowerCase().includes(term)
    )

    treeDiv.style.display = 'none'
    resultsDiv.style.display = ''

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
      if (row && match.path) {
        const pathLabel = document.createElement('span')
        pathLabel.className = 'search-result-path'
        pathLabel.textContent = match.path.replace(/\//g, '\\')
        row.insertAdjacentElement('afterend', pathLabel)
      }
      ul2.appendChild(li)
    }
    resultsDiv.innerHTML = ''
    resultsDiv.appendChild(ul2)
  })
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

  // Icon
  const iconEl = document.createElement('span')
  iconEl.className = 'tree-icon'
  iconEl.textContent = isFolder ? '📁' : '📄'

  // Name
  const nameEl = document.createElement('span')
  nameEl.className = 'tree-name'
  nameEl.textContent = String(node.name || node.path || '(unnamed)')
  if (node.path) nameEl.title = node.path

  // Size
  const sizeEl = document.createElement('span')
  sizeEl.className = 'tree-size-sm'
  sizeEl.textContent = formatBytes(node.sizeBytes)

  // Mapping tag (shows which site this folder is mapped to)
  const tagEl = document.createElement('span')
  tagEl.className = 'mapping-tag'
  const existingMapping = getState().mappings.find((m) => m.sourceNode.path === node.path)
  if (existingMapping?.targetSite) {
    tagEl.textContent = `→ ${existingMapping.targetSite.displayName}`
  } else {
    tagEl.style.display = 'none'
  }
  tagRegistry.set(node.path, tagEl)

  row.appendChild(toggleBtn)
  row.appendChild(iconEl)
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
        // Update the tag in the DOM live
        if (siteName) {
          tagEl.textContent = `→ ${siteName}`
          tagEl.style.display = ''
        } else {
          tagEl.style.display = 'none'
        }
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

      <div class="source-detail-card">
        <div class="source-detail-title">
          <span class="source-detail-icon">📁</span>
          <span class="source-detail-name">${escHtml(String(node.name || node.path))}</span>
        </div>
        <dl class="source-detail-grid">
          <dt>Full Path</dt>
          <dd class="source-detail-path" title="${escHtml(node.path)}">${escHtml(node.path)}</dd>
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

      <div class="form-group">
        <label>SharePoint Site</label>
        <div class="site-search-row">
          <input id="site-search" type="text" class="form-input" placeholder="Search sites…"
            value="${escHtml(existing?.targetSite?.displayName ?? '')}" />
          <button id="btn-search-sites" class="btn btn-primary btn-sm">Search</button>
        </div>
        <div id="site-results" class="site-results"></div>
        <div id="selected-site" class="selected-badge" style="${existing?.targetSite ? '' : 'display:none'}">
          ✓ ${escHtml(existing?.targetSite?.displayName ?? '')}
          <button class="btn-clear" id="btn-clear-site">✕</button>
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

      <button id="btn-save-mapping" class="btn btn-primary" style="margin-top:8px">Save Mapping</button>
      ${existing ? `<button id="btn-remove-mapping" class="btn btn-ghost" style="margin-top:8px;margin-left:8px">Remove</button>` : ''}
    </div>
  `

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
    .mapping-search-input { width: 100%; box-sizing: border-box; padding: 6px 10px; font-size: 0.85rem; }
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
    .tree-icon { flex-shrink: 0; }
    .tree-name { flex: 1; font-size: 0.875rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
    .tree-size-sm { font-size: 0.75rem; color: var(--color-text-muted); white-space: nowrap; flex-shrink: 0; }
    .mapping-tag { font-size: 0.72rem; background: #dff6dd; color: #107c10; padding: 2px 6px;
      border-radius: 10px; white-space: nowrap; flex-shrink: 0; }

    /* Target panel */
    .mapping-placeholder { padding: 32px; text-align: center; color: var(--color-text-muted); font-size: 0.88rem; }
    .target-panel { padding: 16px; display: flex; flex-direction: column; gap: 20px; }

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
