import { searchSites, saveIAFile } from '../../graph/graphClient'
import { updateProject, loadProjectIA, getSpConfig } from '../../graph/projectService'
import { getState, setState } from '../../state/store'
import type { IANode, MigrationMapping, SharePointSite } from '../../types'

// ─── Module state ─────────────────────────────────────────────────────────────

let _nodes: IANode[] = []
let _container: HTMLElement | null = null
let _draggedId: string | null = null
let _saving = false
// Site selected inside the right-panel form
let _panelSelectedSite: SharePointSite | null = null

// ─── Entry point ─────────────────────────────────────────────────────────────

export async function renderIADesignerPanel(container: HTMLElement): Promise<void> {
  injectStyles()
  _container = container
  container.innerHTML = `<div class="ia-panel"><div style="padding:1rem;color:var(--color-text-muted)">Loading IA design…</div></div>`
  _nodes = await loadNodes()
  render()
}

// ─── Data helpers ─────────────────────────────────────────────────────────────

async function loadNodes(): Promise<IANode[]> {
  const project = getState().currentProject
  if (!project) return []
  return loadProjectIA(project)
}

function getChildren(parentId: string | null): IANode[] {
  return _nodes.filter(n => n.parentId === parentId).sort((a, b) => a.order - b.order)
}

/** Returns true if potentialDescendantId is in the subtree rooted at ancestorId. */
function isDescendant(ancestorId: string, potentialDescendantId: string): boolean {
  let current = _nodes.find(n => n.id === potentialDescendantId)
  while (current && current.parentId !== null) {
    if (current.parentId === ancestorId) return true
    current = _nodes.find(n => n.id === current!.parentId!)
  }
  return false
}

function getAllDescendants(nodeId: string): IANode[] {
  const result: IANode[] = []
  const collect = (id: string) => {
    const children = _nodes.filter(n => n.parentId === id)
    result.push(...children)
    children.forEach(c => collect(c.id))
  }
  collect(nodeId)
  return result
}

function escHtml(s: string): string {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

// ─── Persistence ──────────────────────────────────────────────────────────────

async function save(): Promise<void> {
  if (_saving) return
  _saving = true
  const setStatus = (msg: string) => {
    const el = document.getElementById('ia-save-status')
    if (el) el.textContent = msg
  }
  setStatus('Saving…')
  try {
    const state = getState()
    const project = state.currentProject
    if (!project) return
    const { siteId } = getSpConfig()
    await saveIAFile(siteId, project.title, project.id, [..._nodes])
    // Strip iaDesign from the column — it now lives in the .ia.json file
    const updatedProjectData = { ...project.projectData, iaDesign: undefined }
    await updateProject(project.id, { projectData: updatedProjectData })
    setState({ currentProject: { ...project, projectData: updatedProjectData } })
    setStatus('Saved')
    setTimeout(() => setStatus(''), 2000)
  } catch (err) {
    setStatus('Save failed')
    console.error('IA Designer save error', err)
  } finally {
    _saving = false
  }
}

// ─── Node operations ──────────────────────────────────────────────────────────

function addNode(title: string, parentId: string | null): IANode {
  const node: IANode = {
    id: crypto.randomUUID(),
    title,
    parentId,
    order: getChildren(parentId).length,
  }
  _nodes.push(node)
  return node
}

function updateNodeData(nodeId: string, updates: Partial<IANode>): void {
  const idx = _nodes.findIndex(n => n.id === nodeId)
  if (idx === -1) return
  _nodes[idx] = { ..._nodes[idx], ...updates }
}

function deleteNode(nodeId: string): void {
  const toDelete = new Set<string>()
  const collect = (id: string) => {
    toDelete.add(id)
    _nodes.filter(n => n.parentId === id).forEach(n => collect(n.id))
  }
  collect(nodeId)
  const parentId = _nodes.find(n => n.id === nodeId)?.parentId ?? null
  _nodes = _nodes.filter(n => !toDelete.has(n.id))
  _nodes
    .filter(n => n.parentId === parentId)
    .sort((a, b) => a.order - b.order)
    .forEach((n, i) => { n.order = i })
}

function moveNode(nodeId: string, newParentId: string | null, insertAt: number): void {
  if (nodeId === newParentId) return
  if (newParentId !== null && isDescendant(nodeId, newParentId)) return

  const node = _nodes.find(n => n.id === nodeId)
  if (!node) return
  const oldParentId = node.parentId

  _nodes
    .filter(n => n.parentId === oldParentId && n.id !== nodeId)
    .sort((a, b) => a.order - b.order)
    .forEach((n, i) => { n.order = i })

  const newSiblings = _nodes
    .filter(n => n.parentId === newParentId && n.id !== nodeId)
    .sort((a, b) => a.order - b.order)

  node.parentId = newParentId
  newSiblings.splice(insertAt, 0, node)
  newSiblings.forEach((n, i) => { n.order = i })
}

// ─── Full render (initial load) ───────────────────────────────────────────────

function render(): void {
  if (!_container) return
  _container.innerHTML = `
    <div class="ia-panel">
      <div class="ia-left">
        <div class="mapping-section-header">
          <h3>IA Designer</h3>
          <div style="display:flex;align-items:center;gap:10px">
            <button class="btn btn-primary btn-sm" id="ia-btn-add-root">+ Add Root Node</button>
            <span class="ia-save-status" id="ia-save-status"></span>
          </div>
        </div>
        <div class="ia-canvas" id="ia-canvas">
          ${buildCanvasHtml()}
        </div>
      </div>
      <div class="ia-right" id="ia-right">
        <div class="mapping-placeholder">
          Click <strong>+ Add Root Node</strong> or a node's <strong>+</strong> / <strong>✎</strong> button to edit.
        </div>
      </div>
    </div>
  `
  attachEvents()
}

// ─── Partial tree render (preserves right panel) ──────────────────────────────

function renderTree(): void {
  const canvas = document.getElementById('ia-canvas')
  if (canvas) canvas.innerHTML = buildCanvasHtml()
}

// ─── Canvas HTML ──────────────────────────────────────────────────────────────

function buildCanvasHtml(): string {
  const roots = getChildren(null)
  if (roots.length === 0) {
    return `<div class="ia-empty">
      No nodes yet.<br>
      Click <strong>+ Add Root Node</strong> to start building your Information Architecture.
    </div>`
  }
  const items = roots.map((n, i) =>
    `${nodeWrapHtml(n)}<div class="ia-gap ia-gap--root" data-parent-id="" data-order="${i + 1}"></div>`
  ).join('')
  return `
    <div class="ia-tree">
      <div class="ia-gap ia-gap--root" data-parent-id="" data-order="0"></div>
      ${items}
    </div>
  `
}

function nodeWrapHtml(node: IANode): string {
  const children = getChildren(node.id)
  const childrenHtml = children.map((c, ci) =>
    `${nodeWrapHtml(c)}<div class="ia-gap" data-parent-id="${escHtml(node.id)}" data-order="${ci + 1}"></div>`
  ).join('')
  return `
    <div class="ia-node-wrap" data-node-id="${escHtml(node.id)}">
      <div class="ia-node-card" draggable="true" data-node-id="${escHtml(node.id)}">
        <span class="ia-drag-handle" title="Drag to reorder">⠿</span>
        <div class="ia-node-body">
          <div class="ia-node-title">${escHtml(node.title)}</div>
          ${siteLabelHtml(node)}
        </div>
        <div class="ia-node-actions">
          <button class="ia-btn ia-btn--icon" data-action="add-child" data-node-id="${escHtml(node.id)}" title="Add child node">+</button>
          <button class="ia-btn ia-btn--icon" data-action="edit" data-node-id="${escHtml(node.id)}" title="Edit node">✎</button>
          <button class="ia-btn ia-btn--icon ia-btn--danger" data-action="delete" data-node-id="${escHtml(node.id)}" title="Delete node">✕</button>
        </div>
      </div>
      <div class="ia-children-wrap">
        <div class="ia-gap" data-parent-id="${escHtml(node.id)}" data-order="0"></div>
        ${childrenHtml}
      </div>
    </div>
  `
}

function siteLabelHtml(node: IANode): string {
  if (node.mappedSiteId && node.mappedSiteName) {
    return `<div class="ia-node-site ia-node-site--existing" title="${escHtml(node.mappedSiteUrl ?? '')}">
      <span class="ia-site-dot ia-site-dot--existing"></span>${escHtml(node.mappedSiteName)}
    </div>`
  }
  if (node.plannedMappingId && node.plannedSiteDisplayName) {
    return `<div class="ia-node-site ia-node-site--planned">
      <span class="ia-site-dot ia-site-dot--planned"></span>${escHtml(node.plannedSiteDisplayName)} <em>(planned)</em>
    </div>`
  }
  return ''
}

// ─── Events ───────────────────────────────────────────────────────────────────

function attachEvents(): void {
  const canvas = document.getElementById('ia-canvas')
  if (!canvas) return

  document.getElementById('ia-btn-add-root')?.addEventListener('click', () => {
    openNodePanel(null, null)
  })

  canvas.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest('[data-action]') as HTMLElement | null
    if (!btn) return
    const action = btn.dataset.action
    const nodeId = btn.dataset.nodeId
    if (!nodeId) return
    if (action === 'add-child') openNodePanel(null, nodeId)
    else if (action === 'edit') openNodePanel(nodeId, null)
    else if (action === 'delete') confirmDelete(nodeId)
  })

  attachDragEvents(canvas)
}

function confirmDelete(nodeId: string): void {
  const node = _nodes.find(n => n.id === nodeId)
  if (!node) return
  const descCount = getAllDescendants(nodeId).length
  const msg = descCount > 0
    ? `Delete "${node.title}" and its ${descCount} child node(s)?`
    : `Delete "${node.title}"?`
  if (!confirm(msg)) return
  deleteNode(nodeId)
  renderTree()
  // Close panel if the node being edited was deleted
  const rightEl = document.getElementById('ia-right')
  const panelNodeId = rightEl?.querySelector<HTMLElement>('[data-editing-node]')?.dataset.editingNode
  if (panelNodeId === nodeId || getAllDescendants(nodeId).some(n => n.id === panelNodeId)) {
    closeNodePanel()
  }
  void save()
}

// ─── Drag and drop ────────────────────────────────────────────────────────────

function attachDragEvents(canvas: HTMLElement): void {
  canvas.addEventListener('dragstart', (e) => {
    const card = (e.target as HTMLElement).closest('.ia-node-card') as HTMLElement | null
    if (!card) return
    _draggedId = card.dataset.nodeId ?? null
    if (_draggedId) {
      e.dataTransfer!.effectAllowed = 'move'
      e.dataTransfer!.setData('text/plain', _draggedId)
      setTimeout(() => card.classList.add('ia-dragging'), 0)
    }
  })

  canvas.addEventListener('dragend', () => {
    _draggedId = null
    canvas.querySelectorAll('.ia-dragging').forEach(el => el.classList.remove('ia-dragging'))
    canvas.querySelectorAll('.ia-drag-over').forEach(el => el.classList.remove('ia-drag-over'))
    canvas.querySelectorAll('.ia-gap--active').forEach(el => el.classList.remove('ia-gap--active'))
  })

  canvas.addEventListener('dragover', (e) => {
    if (!_draggedId) return
    e.preventDefault()
    e.dataTransfer!.dropEffect = 'move'

    const target = e.target as HTMLElement
    const gap = target.closest('.ia-gap') as HTMLElement | null
    const card = target.closest('.ia-node-card') as HTMLElement | null

    canvas.querySelectorAll('.ia-drag-over').forEach(el => el.classList.remove('ia-drag-over'))
    canvas.querySelectorAll('.ia-gap--active').forEach(el => el.classList.remove('ia-gap--active'))

    if (gap) {
      const rawParentId = gap.dataset.parentId ?? ''
      const targetParentId = rawParentId === '' ? null : rawParentId
      const valid = targetParentId !== _draggedId &&
        (targetParentId === null || !isDescendant(_draggedId, targetParentId))
      if (valid) gap.classList.add('ia-gap--active')
    } else if (card) {
      const targetNodeId = card.dataset.nodeId
      if (targetNodeId && targetNodeId !== _draggedId && !isDescendant(_draggedId, targetNodeId)) {
        card.classList.add('ia-drag-over')
      }
    }
  })

  canvas.addEventListener('dragleave', (e) => {
    const related = e.relatedTarget as HTMLElement | null
    if (!related || !canvas.contains(related)) {
      canvas.querySelectorAll('.ia-drag-over').forEach(el => el.classList.remove('ia-drag-over'))
      canvas.querySelectorAll('.ia-gap--active').forEach(el => el.classList.remove('ia-gap--active'))
    }
  })

  canvas.addEventListener('drop', (e) => {
    e.preventDefault()
    if (!_draggedId) return

    const target = e.target as HTMLElement
    const gap = target.closest('.ia-gap') as HTMLElement | null
    const card = target.closest('.ia-node-card') as HTMLElement | null

    if (gap) {
      const rawParentId = gap.dataset.parentId ?? ''
      const targetParentId = rawParentId === '' ? null : rawParentId
      const insertAt = parseInt(gap.dataset.order ?? '0', 10)
      if (targetParentId !== _draggedId &&
          (targetParentId === null || !isDescendant(_draggedId, targetParentId))) {
        moveNode(_draggedId, targetParentId, insertAt)
      }
    } else if (card) {
      const targetNodeId = card.dataset.nodeId
      if (targetNodeId && targetNodeId !== _draggedId && !isDescendant(_draggedId, targetNodeId)) {
        moveNode(_draggedId, targetNodeId, getChildren(targetNodeId).length)
      }
    }

    _draggedId = null
    renderTree()
    void save()
  })
}

// ─── Right panel ─────────────────────────────────────────────────────────────

function closeNodePanel(): void {
  const rightEl = document.getElementById('ia-right')
  if (rightEl) {
    rightEl.innerHTML = `<div class="mapping-placeholder">
      Click <strong>+ Add Root Node</strong> or a node's <strong>+</strong> / <strong>✎</strong> button to edit.
    </div>`
  }
}

function openNodePanel(nodeId: string | null, parentId: string | null): void {
  const rightEl = document.getElementById('ia-right')
  if (!rightEl) return

  const node = nodeId ? _nodes.find(n => n.id === nodeId) ?? null : null
  _panelSelectedSite = node?.mappedSiteId
    ? { id: node.mappedSiteId, name: node.mappedSiteName ?? '', displayName: node.mappedSiteName ?? '', webUrl: node.mappedSiteUrl ?? '' }
    : null

  const state = getState()
  const allMappings: MigrationMapping[] = (
    state.mappings.length > 0 ? state.mappings : (state.currentProject?.projectData?.mappings ?? [])
  )
  const plannedSites = allMappings
    .filter(m => m.plannedSite)
    .map(m => ({ id: m.id, displayName: m.plannedSite!.displayName }))

  let activeTab: 'none' | 'existing' | 'planned' = 'none'
  if (node?.mappedSiteId) activeTab = 'existing'
  else if (node?.plannedMappingId) activeTab = 'planned'

  const isNew = !node
  const heading = isNew ? (parentId ? 'Add Child Node' : 'Add Root Node') : 'Edit Node'

  const tabBtn = (tab: string, label: string) =>
    `<button class="sp-tab${activeTab === tab ? ' sp-tab--active' : ''}" data-tab="${tab}">${label}</button>`

  rightEl.innerHTML = `
    <div class="target-panel" data-editing-node="${escHtml(nodeId ?? '')}">

      <div class="mapping-section-header" style="position:sticky;top:0;z-index:2">
        <h3 class="ia-panel-heading">${heading}</h3>
        <button class="btn-icon" id="ia-panel-close" title="Close">✕</button>
      </div>

      <!-- Node title -->
      <div class="target-section">
        <div class="target-section-body--sp">
          <div class="form-group" style="margin-bottom:0">
            <label for="ia-panel-title">Title <span class="required">*</span></label>
            <input id="ia-panel-title" type="text" class="form-input"
              value="${escHtml(node?.title ?? '')}"
              placeholder="e.g. Intranet, HR Department…"
              autocomplete="off" />
          </div>
        </div>
      </div>

      <!-- Site mapping (tabbed) -->
      <div class="target-section" style="flex:1">
        <div class="sp-tabs-bar">
          ${tabBtn('none', 'No Mapping')}
          ${tabBtn('existing', 'Existing Site')}
          ${tabBtn('planned', 'Planned Site')}
        </div>

        <!-- Tab: No mapping -->
        <div id="ia-tab-none" class="sp-tab-panel target-section-body--sp"
          ${activeTab !== 'none' ? 'style="display:none"' : ''}>
          <p style="margin:0;font-size:0.85rem;color:var(--color-text-muted)">
            This node has no site mapping. Switch to another tab to link it to a SharePoint site.
          </p>
        </div>

        <!-- Tab: Existing site -->
        <div id="ia-tab-existing" class="sp-tab-panel target-section-body--sp"
          ${activeTab !== 'existing' ? 'style="display:none"' : ''}>
          <div class="form-group" style="margin-bottom:0">
            <label>Search SharePoint sites</label>
            <div class="site-search-row">
              <input id="ia-site-search" type="text" class="form-input"
                value="${escHtml(node?.mappedSiteName ?? '')}"
                placeholder="Site name…" autocomplete="off" />
              <button id="ia-btn-search" class="btn btn-primary btn-sm">Search</button>
            </div>
            <div id="ia-site-results" class="site-results"></div>
            <div id="ia-selected-badge" class="selected-badge"
              style="${_panelSelectedSite ? '' : 'display:none'}">
              ✓ ${escHtml(_panelSelectedSite?.displayName ?? '')}
              <button class="btn-clear" id="ia-clear-site">✕</button>
            </div>
          </div>
        </div>

        <!-- Tab: Planned site -->
        <div id="ia-tab-planned" class="sp-tab-panel target-section-body--sp"
          ${activeTab !== 'planned' ? 'style="display:none"' : ''}>
          <div class="form-group" style="margin-bottom:0">
            <label for="ia-planned-select">Planned site</label>
            ${plannedSites.length === 0
              ? `<p class="ia-hint">No planned sites yet. On the Map tab, create mappings that point to new (not-yet-created) sites — they will appear here.</p>`
              : `<select id="ia-planned-select" class="form-input">
                  <option value="">— Select —</option>
                  ${plannedSites.map(ps =>
                    `<option value="${escHtml(ps.id)}" ${node?.plannedMappingId === ps.id ? 'selected' : ''}>${escHtml(ps.displayName)}</option>`
                  ).join('')}
                 </select>`}
          </div>
        </div>
      </div>

      <!-- Footer -->
      <div class="ia-panel-footer">
        <button class="btn btn-primary btn-sm" id="ia-panel-save">Save</button>
        <button class="btn btn-ghost btn-sm" id="ia-panel-cancel">Cancel</button>
      </div>

    </div>
  `

  // Focus title
  ;(document.getElementById('ia-panel-title') as HTMLInputElement | null)?.focus()

  // Close / Cancel
  document.getElementById('ia-panel-close')?.addEventListener('click', closeNodePanel)
  document.getElementById('ia-panel-cancel')?.addEventListener('click', closeNodePanel)

  // Tab switching
  rightEl.querySelectorAll<HTMLElement>('.sp-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const t = tab.dataset.tab ?? 'none'
      rightEl.querySelectorAll('.sp-tab').forEach(b => b.classList.remove('sp-tab--active'))
      tab.classList.add('sp-tab--active')
      rightEl.querySelectorAll<HTMLElement>('.sp-tab-panel').forEach(p => {
        p.style.display = p.id === `ia-tab-${t}` ? '' : 'none'
      })
    })
  })

  // Search button
  document.getElementById('ia-btn-search')?.addEventListener('click', () => {
    const query = (document.getElementById('ia-site-search') as HTMLInputElement)?.value ?? ''
    void runSiteSearch(query)
  })

  // Search on Enter in search input
  ;(document.getElementById('ia-site-search') as HTMLInputElement | null)
    ?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault()
        const query = (e.target as HTMLInputElement).value
        void runSiteSearch(query)
      }
    })

  // Clear site
  attachClearSite()

  // Save
  document.getElementById('ia-panel-save')?.addEventListener('click', () => {
    const titleInput = document.getElementById('ia-panel-title') as HTMLInputElement | null
    const titleVal = titleInput?.value.trim() ?? ''
    if (!titleVal) { titleInput?.focus(); return }

    const activeTabEl = rightEl.querySelector<HTMLElement>('.sp-tab--active')
    const tabVal = activeTabEl?.dataset.tab ?? 'none'
    const mappingUpdates = resolveMappingUpdates(tabVal, plannedSites, rightEl)

    if (isNew) {
      const newNode = addNode(titleVal, parentId ?? null)
      Object.assign(newNode, mappingUpdates)
    } else if (node) {
      updateNodeData(node.id, { title: titleVal, ...mappingUpdates })
    }

    closeNodePanel()
    renderTree()
    void save()
  })
}

function attachClearSite(): void {
  document.getElementById('ia-clear-site')?.addEventListener('click', () => {
    _panelSelectedSite = null
    const badge = document.getElementById('ia-selected-badge')
    if (badge) badge.style.display = 'none'
    const search = document.getElementById('ia-site-search') as HTMLInputElement | null
    if (search) search.value = ''
    const results = document.getElementById('ia-site-results')
    if (results) results.innerHTML = ''
  })
}

function resolveMappingUpdates(
  tab: string,
  plannedSites: Array<{ id: string; displayName: string }>,
  container: HTMLElement
): Partial<IANode> {
  const updates: Partial<IANode> = {
    mappedSiteId: undefined,
    mappedSiteName: undefined,
    mappedSiteUrl: undefined,
    plannedMappingId: undefined,
    plannedSiteDisplayName: undefined,
  }
  if (tab === 'existing' && _panelSelectedSite) {
    updates.mappedSiteId = _panelSelectedSite.id
    updates.mappedSiteName = _panelSelectedSite.displayName
    updates.mappedSiteUrl = _panelSelectedSite.webUrl
  } else if (tab === 'planned') {
    const sel = container.querySelector<HTMLSelectElement>('#ia-planned-select')
    const selId = sel?.value
    if (selId) {
      const ps = plannedSites.find(p => p.id === selId)
      if (ps) {
        updates.plannedMappingId = ps.id
        updates.plannedSiteDisplayName = ps.displayName
      }
    }
  }
  return updates
}

async function runSiteSearch(query: string): Promise<void> {
  const resultsEl = document.getElementById('ia-site-results')
  if (!resultsEl) return

  resultsEl.innerHTML = '<span class="searching">Searching…</span>'

  try {
    const sites = await searchSites(query || '*')
    if (!document.getElementById('ia-site-results')) return // panel closed

    if (sites.length === 0) {
      resultsEl.innerHTML = '<span class="no-results">No sites found.</span>'
      return
    }

    resultsEl.innerHTML = sites.slice(0, 15).map(s =>
      `<div class="site-result-item"
        data-id="${escHtml(s.id)}"
        data-name="${escHtml(s.displayName)}"
        data-url="${escHtml(s.webUrl)}">
        ${escHtml(s.displayName)}<br><small>${escHtml(s.webUrl)}</small>
      </div>`
    ).join('')

    resultsEl.querySelectorAll<HTMLElement>('.site-result-item').forEach(item => {
      item.addEventListener('click', () => {
        _panelSelectedSite = {
          id: item.dataset.id!,
          name: item.dataset.name!,
          displayName: item.dataset.name!,
          webUrl: item.dataset.url!,
        }
        resultsEl.innerHTML = ''

        const search = document.getElementById('ia-site-search') as HTMLInputElement | null
        if (search) search.value = _panelSelectedSite.displayName

        const badge = document.getElementById('ia-selected-badge')
        if (badge) {
          badge.innerHTML = `✓ ${escHtml(_panelSelectedSite.displayName)} <button class="btn-clear" id="ia-clear-site">✕</button>`
          badge.style.display = ''
          attachClearSite()
        }
      })
    })
  } catch (err) {
    const el = document.getElementById('ia-site-results')
    if (el) el.innerHTML = '<span class="no-results">Search failed.</span>'
    console.error('IA Designer site search error', err)
  }
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectStyles(): void {
  if (document.getElementById('ia-designer-styles')) return
  const style = document.createElement('style')
  style.id = 'ia-designer-styles'
  style.textContent = `
    /* Two-column layout (same structure as .mapping-panel but 3fr/1fr) */
    .ia-panel { display: grid; grid-template-columns: 3fr 1fr; height: calc(100vh - 140px); overflow: hidden; }
    .ia-left { display: flex; flex-direction: column; overflow: hidden;
      border-right: 1px solid var(--color-border); }
    .ia-right { overflow-y: auto; }

    /* Canvas */
    .ia-canvas { flex: 1; overflow: auto; padding: 24px 32px; }
    .ia-save-status { font-size: 0.8rem; color: var(--color-text-muted); }
    .ia-empty { text-align: center; color: var(--color-text-muted); font-size: 0.9rem;
      padding: 72px 20px; line-height: 2; }

    /* Tree */
    .ia-tree { display: flex; flex-direction: column; min-width: max-content; }
    .ia-node-wrap { display: flex; flex-direction: column; }
    .ia-children-wrap { margin-left: 28px; padding-left: 20px;
      border-left: 2px solid var(--color-border); }

    /* Gap drop zones */
    .ia-gap { height: 8px; border-radius: 3px;
      transition: height 0.12s, background 0.12s, box-shadow 0.12s; }
    .ia-gap--active { height: 4px !important; background: var(--color-primary) !important;
      box-shadow: 0 0 0 2px rgba(0,120,212,0.25); }

    /* Node cards */
    .ia-node-card { display: flex; align-items: center; gap: 8px; padding: 9px 12px;
      border: 1.5px solid var(--color-border); border-radius: 8px; background: white;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
      transition: box-shadow 0.15s, border-color 0.15s, background 0.15s;
      min-width: 200px; width: fit-content; max-width: 360px;
      user-select: none; cursor: default; }
    .ia-node-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.12); }
    .ia-node-card.ia-drag-over { border-color: var(--color-primary);
      box-shadow: 0 0 0 3px rgba(0,120,212,0.2); background: #f0f6ff; }
    .ia-node-card.ia-dragging { opacity: 0.35; }
    .ia-drag-handle { cursor: grab; color: var(--color-text-muted); font-size: 1rem;
      padding: 0 2px; line-height: 1; flex-shrink: 0; }
    .ia-drag-handle:active { cursor: grabbing; }
    .ia-node-body { flex: 1; min-width: 0; }
    .ia-node-title { font-weight: 600; font-size: 0.9rem; overflow: hidden;
      text-overflow: ellipsis; white-space: nowrap; }
    .ia-node-site { font-size: 0.76rem; margin-top: 2px; display: flex; align-items: center;
      gap: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .ia-node-site--existing { color: #005a9e; }
    .ia-node-site--planned { color: #7a5800; }
    .ia-site-dot { display: inline-block; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
    .ia-site-dot--existing { background: #0078d4; }
    .ia-site-dot--planned { background: #f0a30a; }
    .ia-node-actions { display: flex; gap: 2px; flex-shrink: 0; opacity: 0; transition: opacity 0.12s; }
    .ia-node-card:hover .ia-node-actions,
    .ia-node-card:focus-within .ia-node-actions { opacity: 1; }

    /* Node action buttons */
    .ia-btn { border: none; border-radius: 4px; cursor: pointer;
      transition: background 0.1s, color 0.1s; }
    .ia-btn--icon { background: none; border: none; color: var(--color-text-muted);
      font-size: 0.95rem; padding: 3px 6px; }
    .ia-btn--icon:hover { background: var(--color-surface-alt); color: var(--color-text); }
    .ia-btn--danger:hover { background: #fde7e9; color: #c50f1f; }

    /* Right panel */
    .ia-panel-heading { font-size: 0.9rem; font-weight: 600; margin: 0; }
    .ia-panel-footer { padding: 14px 16px; border-top: 1px solid var(--color-border);
      display: flex; gap: 8px; }
    .ia-hint { font-size: 0.8rem; color: var(--color-text-muted); margin: 6px 0 0; line-height: 1.4; }

    /* Shared classes (mirrors mapping panel — idempotent if mapping panel also loaded) */
    .mapping-section-header { padding: 10px 16px; border-bottom: 1px solid var(--color-border);
      display: flex; align-items: center; justify-content: space-between;
      background: var(--color-surface-alt); }
    .mapping-section-header h3 { font-size: 0.9rem; font-weight: 600; margin: 0; }
    .mapping-placeholder { padding: 32px 16px; text-align: center;
      color: var(--color-text-muted); font-size: 0.85rem; line-height: 1.6; }
    .target-panel { display: flex; flex-direction: column; }
    .target-section { border-bottom: 1px solid var(--color-border); }
    .target-section:last-child { border-bottom: none; }
    .target-section-body--sp { padding: 16px; display: flex; flex-direction: column; gap: 16px; }
    .target-section-body--sp .form-group { margin-bottom: 0; }
    .sp-tabs-bar { display: flex; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface-alt); }
    .sp-tab { flex: 1; padding: 9px 10px; background: none; border: none;
      border-bottom: 2px solid transparent; cursor: pointer; font-size: 0.78rem; font-weight: 500;
      color: var(--color-text-muted); font-family: inherit; text-align: center;
      transition: color 0.15s, border-color 0.15s; }
    .sp-tab:hover { color: var(--color-text); background: var(--color-primary-light); }
    .sp-tab--active { color: var(--color-primary); border-bottom-color: var(--color-primary); font-weight: 600; }
    .site-search-row { display: flex; gap: 8px; }
    .site-results { margin-top: 8px; border: 1px solid var(--color-border); border-radius: 4px;
      overflow: hidden; empty-cells: hide; }
    .site-results:empty { display: none; }
    .site-result-item { padding: 8px 12px; cursor: pointer; font-size: 0.85rem;
      border-bottom: 1px solid var(--color-border); }
    .site-result-item:last-child { border-bottom: none; }
    .site-result-item:hover { background: var(--color-primary-light); }
    .site-result-item small { color: var(--color-text-muted); }
    .selected-badge { background: #dff6dd; color: #107c10; padding: 6px 10px;
      border-radius: 4px; font-size: 0.85rem; margin-top: 8px;
      display: flex; align-items: center; gap: 6px; }
    .btn-clear { background: none; border: none; cursor: pointer;
      color: inherit; font-size: 0.9rem; }
    .searching, .no-results { padding: 8px 12px; font-size: 0.85rem;
      color: var(--color-text-muted); display: block; }
  `
  document.head.appendChild(style)
}
