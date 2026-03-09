import { searchSites, getSiteDrives } from '../../graph/graphClient'
import { updateProject } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { TreeNode, MigrationMapping, SharePointSite, SharePointDrive } from '../../types'

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
        <div id="mapping-tree" class="mapping-tree"></div>
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

  renderMappingTree(
    container.querySelector('#mapping-tree') as HTMLElement,
    tree,
    container.querySelector('#mapping-target') as HTMLElement
  )
}

function renderMappingTree(treeEl: HTMLElement, node: TreeNode, targetEl: HTMLElement): void {
  treeEl.innerHTML = `<ul class="tree-list tree-root">${renderMappingNode(node, true)}</ul>`

  treeEl.querySelectorAll('.mapping-row').forEach((row) => {
    row.addEventListener('click', (e) => {
      const target = e.currentTarget as HTMLElement
      const path = target.getAttribute('data-path')!
      const found = findNode(getState().treeData!, path)
      if (found) openTargetPanel(targetEl, found)
      treeEl.querySelectorAll('.mapping-row').forEach((r) => r.classList.remove('mapping-row--active'))
      target.classList.add('mapping-row--active')
    })
  })

  treeEl.querySelectorAll('.mapping-toggle-btn').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation()
      const li = (e.currentTarget as HTMLElement).closest('.mapping-node')!
      const children = li.querySelector('.tree-children') as HTMLElement | null
      const icon = (e.currentTarget as HTMLElement).querySelector('.toggle-icon') as HTMLElement
      if (!children) return
      const isOpen = children.style.display !== 'none'
      children.style.display = isOpen ? 'none' : ''
      icon.textContent = isOpen ? '▶' : '▼'
    })
  })

  // Auto-expand root
  const firstToggle = treeEl.querySelector('.mapping-toggle-btn') as HTMLElement | null
  firstToggle?.click()
}

function renderMappingNode(node: TreeNode, isRoot = false): string {
  const hasChildren = node.children.length > 0
  const mapping = getState().mappings.find((m) => m.sourceNode.path === node.path)
  const mappedLabel = mapping?.targetSite
    ? `<span class="mapping-tag">→ ${mapping.targetSite.displayName}</span>`
    : ''

  return `
    <li class="mapping-node ${isRoot ? 'mapping-node--root' : ''}">
      <div class="mapping-row" data-path="${escHtml(node.path)}">
        <button class="mapping-toggle-btn${hasChildren ? '' : ' invisible'}" type="button">
          <span class="toggle-icon">▶</span>
        </button>
        <span class="tree-icon">${hasChildren ? '📁' : '📄'}</span>
        <span class="tree-name">${escHtml(node.name || node.path)}</span>
        <span class="tree-size-sm">${formatBytes(node.sizeBytes)}</span>
        ${mappedLabel}
      </div>
      ${hasChildren
        ? `<ul class="tree-list tree-children" style="display:none">${node.children.map((c) => renderMappingNode(c)).join('')}</ul>`
        : ''
      }
    </li>
  `
}

async function openTargetPanel(targetEl: HTMLElement, node: TreeNode): Promise<void> {
  const existing = getState().mappings.find((m) => m.sourceNode.path === node.path)

  targetEl.innerHTML = `
    <div class="target-panel">
      <div class="target-source-label">
        <strong>Source:</strong> <code>${escHtml(node.name)}</code>
        <span class="target-source-size">${formatBytes(node.sizeBytes)}</span>
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
        : sites.map((s) => `<div class="site-result-item" data-id="${escHtml(s.id)}">${escHtml(s.displayName)}<br><small>${escHtml(s.webUrl)}</small></div>`).join('')

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
      const badge = targetEl.querySelector('#selected-site') as HTMLElement
      badge.style.display = 'none'
      const libGroup = targetEl.querySelector('#library-group') as HTMLElement
      libGroup.style.display = 'none'
    })
  }
  attachClearSite()

  targetEl.querySelector('#btn-save-mapping')?.addEventListener('click', () => {
    const folderPath = (targetEl.querySelector('#folder-path') as HTMLInputElement).value.trim()
    const libSelect = targetEl.querySelector('#library-select') as HTMLSelectElement
    const drives = getState().sites.length > 0 ? [] : []
    void drives

    const mapping: MigrationMapping = {
      id: node.path,
      sourceNode: node,
      targetSite: selectedSite,
      targetDrive: selectedDrive,
      targetFolderPath: folderPath,
      status: selectedSite ? 'ready' : 'pending',
    }

    void libSelect  // library will be set from selectedDrive

    const mappings = [
      ...getState().mappings.filter((m) => m.sourceNode.path !== node.path),
      mapping,
    ]
    setState({ mappings })
    persistMappings(mappings)

    const saveBtn = targetEl.querySelector('#btn-save-mapping') as HTMLButtonElement
    saveBtn.textContent = '✓ Saved'
    setTimeout(() => { saveBtn.textContent = 'Save Mapping' }, 2000)
  })

  targetEl.querySelector('#btn-remove-mapping')?.addEventListener('click', () => {
    const mappings = getState().mappings.filter((m) => m.sourceNode.path !== node.path)
    setState({ mappings })
    persistMappings(mappings)
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
    libSelect.innerHTML = drives.map((d) =>
      `<option value="${escHtml(d.id)}" ${selected?.id === d.id ? 'selected' : ''}>${escHtml(d.name)}</option>`
    ).join('')
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

function findNode(node: TreeNode, path: string): TreeNode | null {
  if (node.path === path) return node
  for (const child of node.children) {
    const found = findNode(child, path)
    if (found) return found
  }
  return null
}

function formatBytes(bytes: number): string {
  if (!bytes) return ''
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectMappingStyles(): void {
  if (document.getElementById('mapping-styles')) return
  const style = document.createElement('style')
  style.id = 'mapping-styles'
  style.textContent = `
    .mapping-empty { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .mapping-panel { display: grid; grid-template-columns: 1fr 1fr; height: calc(100vh - 140px); overflow: hidden; }
    .mapping-left, .mapping-right { overflow-y: auto; border-right: 1px solid var(--color-border); }
    .mapping-right { border-right: none; }
    .mapping-section-header { padding: 12px 16px; border-bottom: 1px solid var(--color-border);
      display: flex; align-items: center; justify-content: space-between; background: var(--color-surface-alt); position: sticky; top: 0; z-index: 1; }
    .mapping-section-header h3 { font-size: 0.9rem; font-weight: 600; margin: 0; }
    .mapping-hint { font-size: 0.78rem; color: var(--color-text-muted); }
    .mapping-tree { padding: 8px; }
    .mapping-node { margin: 1px 0; }
    .mapping-row { display: flex; align-items: center; gap: 6px; padding: 5px 8px; border-radius: 4px;
      cursor: pointer; transition: background 0.1s; }
    .mapping-row:hover { background: var(--color-primary-light); }
    .mapping-row--active { background: var(--color-primary-light); border-left: 3px solid var(--color-primary); }
    .mapping-toggle-btn { background: none; border: none; cursor: pointer; width: 16px;
      font-size: 0.65rem; color: var(--color-text-muted); padding: 0; flex-shrink: 0; }
    .mapping-toggle-btn.invisible { visibility: hidden; }
    .toggle-icon { display: block; }
    .tree-size-sm { font-size: 0.75rem; color: var(--color-text-muted); margin-left: auto; }
    .mapping-tag { font-size: 0.72rem; background: #dff6dd; color: #107c10; padding: 2px 6px;
      border-radius: 10px; white-space: nowrap; }
    .mapping-placeholder { padding: 32px; text-align: center; color: var(--color-text-muted); font-size: 0.88rem; }
    .target-panel { padding: 16px; }
    .target-source-label { font-size: 0.85rem; margin-bottom: 20px; padding-bottom: 12px;
      border-bottom: 1px solid var(--color-border); }
    .target-source-size { margin-left: 8px; color: var(--color-text-muted); }
    .site-search-row { display: flex; gap: 8px; }
    .site-results { margin-top: 8px; border: 1px solid var(--color-border); border-radius: 4px;
      max-height: 200px; overflow-y: auto; }
    .site-result-item { padding: 8px 12px; cursor: pointer; font-size: 0.85rem; border-bottom: 1px solid var(--color-border); }
    .site-result-item:last-child { border-bottom: none; }
    .site-result-item:hover { background: var(--color-primary-light); }
    .site-result-item small { color: var(--color-text-muted); }
    .selected-badge { background: #dff6dd; color: #107c10; padding: 6px 10px; border-radius: 4px;
      font-size: 0.85rem; margin-top: 8px; display: flex; align-items: center; justify-content: space-between; }
    .btn-clear { background: none; border: none; cursor: pointer; color: inherit; font-size: 0.9rem; }
    .searching, .no-results { padding: 8px 12px; font-size: 0.85rem; color: var(--color-text-muted); display: block; }
    .hint { font-size: 0.78rem; color: var(--color-text-muted); font-weight: 400; }
    .mapping-node--root > .mapping-row { font-weight: 600; }
  `
  document.head.appendChild(style)
}
