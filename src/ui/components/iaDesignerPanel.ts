import { searchSites } from '../../graph/graphClient'
import { updateProject } from '../../graph/projectService'
import { getState, setState } from '../../state/store'
import type { IANode, MigrationMapping, SharePointSite } from '../../types'

// ─── Module state ─────────────────────────────────────────────────────────────

let _nodes: IANode[] = []
let _container: HTMLElement | null = null
let _draggedId: string | null = null
let _saving = false
// Working site selection inside the edit modal
let _modalSelectedSite: SharePointSite | null = null

// ─── Entry point ─────────────────────────────────────────────────────────────

export function renderIADesignerPanel(container: HTMLElement): void {
  injectStyles()
  _container = container
  _nodes = loadNodes()
  render()
}

// ─── Data helpers ─────────────────────────────────────────────────────────────

function loadNodes(): IANode[] {
  return [...(getState().currentProject?.projectData?.iaDesign ?? [])]
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
    const updated = {
      ...project,
      projectData: { ...project.projectData, iaDesign: [..._nodes] },
    }
    await updateProject(updated)
    setState({ currentProject: updated })
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
  // Renumber remaining siblings
  _nodes
    .filter(n => n.parentId === parentId)
    .sort((a, b) => a.order - b.order)
    .forEach((n, i) => { n.order = i })
}

function moveNode(nodeId: string, newParentId: string | null, insertAt: number): void {
  if (nodeId === newParentId) return
  // Prevent dropping a node into one of its own descendants
  if (newParentId !== null && isDescendant(nodeId, newParentId)) return

  const node = _nodes.find(n => n.id === nodeId)
  if (!node) return
  const oldParentId = node.parentId

  // Reorder old siblings (without the moved node)
  _nodes
    .filter(n => n.parentId === oldParentId && n.id !== nodeId)
    .sort((a, b) => a.order - b.order)
    .forEach((n, i) => { n.order = i })

  // Collect new siblings (excluding the moving node, in case same parent)
  const newSiblings = _nodes
    .filter(n => n.parentId === newParentId && n.id !== nodeId)
    .sort((a, b) => a.order - b.order)

  // Update node's parent
  node.parentId = newParentId

  // Insert at the desired position and renumber
  newSiblings.splice(insertAt, 0, node)
  newSiblings.forEach((n, i) => { n.order = i })
}

// ─── Main render ──────────────────────────────────────────────────────────────

function render(): void {
  if (!_container) return
  _container.innerHTML = `
    <div class="ia-panel">
      <div class="ia-toolbar">
        <button class="btn btn-primary btn-sm" id="ia-btn-add-root">+ Add Root Node</button>
        <span class="ia-save-status" id="ia-save-status"></span>
      </div>
      <div class="ia-canvas" id="ia-canvas">
        ${buildCanvasHtml()}
      </div>
    </div>
  `
  attachEvents()
}

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
    openNodeModal(null, null)
  })

  // Delegated click handler for node action buttons
  canvas.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest('[data-action]') as HTMLElement | null
    if (!btn) return
    const action = btn.dataset.action
    const nodeId = btn.dataset.nodeId
    if (!nodeId) return
    if (action === 'add-child') openNodeModal(null, nodeId)
    else if (action === 'edit') openNodeModal(nodeId, null)
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
  render()
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
      // Allow if we're not dropping into the dragged node itself
      const valid = targetParentId !== _draggedId &&
        (targetParentId === null || !isDescendant(_draggedId, targetParentId))
      if (valid) gap.classList.add('ia-gap--active')
    } else if (card) {
      const targetNodeId = card.dataset.nodeId
      if (
        targetNodeId &&
        targetNodeId !== _draggedId &&
        !isDescendant(_draggedId, targetNodeId)
      ) {
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
    render()
    void save()
  })
}

// ─── Node modal ───────────────────────────────────────────────────────────────

function openNodeModal(nodeId: string | null, parentId: string | null): void {
  const modalRoot = document.querySelector('#modal-root') as HTMLElement | null
  if (!modalRoot) return

  const node = nodeId ? _nodes.find(n => n.id === nodeId) ?? null : null
  _modalSelectedSite = node?.mappedSiteId
    ? { id: node.mappedSiteId, name: node.mappedSiteName ?? '', displayName: node.mappedSiteName ?? '', webUrl: node.mappedSiteUrl ?? '' }
    : null

  // Collect planned sites from project data (mappings with plannedSite)
  const state = getState()
  const allMappings: MigrationMapping[] = (
    state.mappings.length > 0 ? state.mappings : (state.currentProject?.projectData?.mappings ?? [])
  )
  const plannedSites = allMappings
    .filter(m => m.plannedSite)
    .map(m => ({ id: m.id, displayName: m.plannedSite!.displayName }))

  let mappingType: 'none' | 'existing' | 'planned' = 'none'
  if (node?.mappedSiteId) mappingType = 'existing'
  else if (node?.plannedMappingId) mappingType = 'planned'

  const isNew = !node
  const title = isNew ? (parentId ? 'Add Child Node' : 'Add Root Node') : 'Edit Node'

  modalRoot.innerHTML = `
    <div class="form-overlay">
      <div class="form-dialog ia-dialog">
        <div class="form-dialog-header">
          <h2>${title}</h2>
          <button class="btn-icon" id="ia-modal-close" title="Close">✕</button>
        </div>
        <div class="ia-modal-body">

          <div class="form-group">
            <label for="ia-modal-title">Title <span class="required">*</span></label>
            <input id="ia-modal-title" type="text" class="form-input"
              value="${escHtml(node?.title ?? '')}"
              placeholder="e.g. Intranet, HR Department, Finance…"
              autocomplete="off" />
          </div>

          <div class="form-group">
            <label>Site Mapping <span class="hint">(optional)</span></label>
            <div class="ia-radio-group">
              <label class="radio-label">
                <input type="radio" name="ia-map-type" value="none"
                  ${mappingType === 'none' ? 'checked' : ''} />
                No site mapping
              </label>
              <label class="radio-label">
                <input type="radio" name="ia-map-type" value="existing"
                  ${mappingType === 'existing' ? 'checked' : ''} />
                Link to existing SharePoint site
              </label>
              <label class="radio-label">
                <input type="radio" name="ia-map-type" value="planned"
                  ${plannedSites.length === 0 ? 'disabled' : ''}
                  ${mappingType === 'planned' ? 'checked' : ''} />
                Link to planned site (from Map tab)
                ${plannedSites.length === 0 ? '<span class="ia-hint-inline"> — no planned sites yet</span>' : ''}
              </label>
            </div>
          </div>

          <div id="ia-section-existing" ${mappingType !== 'existing' ? 'class="ia-hidden"' : ''}>
            <div class="form-group">
              <label for="ia-site-search">Search for site</label>
              <div class="ia-search-wrap">
                <input id="ia-site-search" type="text" class="form-input"
                  placeholder="Start typing to search…"
                  value="${escHtml(node?.mappedSiteName ?? '')}"
                  autocomplete="off" />
                <ul id="ia-site-results" class="ia-dropdown ia-hidden"></ul>
              </div>
              <div id="ia-selected-site" class="${_modalSelectedSite ? '' : 'ia-hidden'} ia-selected-site">
                ${_modalSelectedSite ? selectedSiteHtml(_modalSelectedSite) : ''}
              </div>
            </div>
          </div>

          <div id="ia-section-planned" ${mappingType !== 'planned' ? 'class="ia-hidden"' : ''}>
            <div class="form-group">
              <label for="ia-planned-select">Planned site</label>
              ${plannedSites.length === 0
                ? `<p class="ia-hint">No planned sites found. In the Map tab, add mappings that point to new (not-yet-created) sites.</p>`
                : `<select id="ia-planned-select" class="form-input">
                    <option value="">— Select —</option>
                    ${plannedSites.map(ps =>
                      `<option value="${escHtml(ps.id)}"
                        ${node?.plannedMappingId === ps.id ? 'selected' : ''}>
                        ${escHtml(ps.displayName)}
                       </option>`
                    ).join('')}
                   </select>`}
            </div>
          </div>

        </div>
        <div class="ia-modal-footer">
          <button class="btn btn-primary" id="ia-modal-save">Save</button>
          <button class="btn btn-ghost" id="ia-modal-cancel">Cancel</button>
        </div>
      </div>
    </div>
  `

  // Focus title input
  ;(document.getElementById('ia-modal-title') as HTMLInputElement | null)?.focus()

  const closeModal = () => { modalRoot.innerHTML = '' }
  document.getElementById('ia-modal-close')?.addEventListener('click', closeModal)
  document.getElementById('ia-modal-cancel')?.addEventListener('click', closeModal)

  // Radio: toggle sections
  modalRoot.querySelectorAll<HTMLInputElement>('input[name="ia-map-type"]').forEach(radio => {
    radio.addEventListener('change', () => {
      const v = radio.value
      document.getElementById('ia-section-existing')?.classList.toggle('ia-hidden', v !== 'existing')
      document.getElementById('ia-section-planned')?.classList.toggle('ia-hidden', v !== 'planned')
    })
  })

  // Site search with debounce
  let searchTimer: ReturnType<typeof setTimeout>
  const searchInput = document.getElementById('ia-site-search') as HTMLInputElement | null
  searchInput?.addEventListener('input', () => {
    clearTimeout(searchTimer)
    searchTimer = setTimeout(() => void runSiteSearch(searchInput.value), 300)
  })

  // Save button
  document.getElementById('ia-modal-save')?.addEventListener('click', () => {
    const titleInput = document.getElementById('ia-modal-title') as HTMLInputElement | null
    const titleVal = titleInput?.value.trim() ?? ''
    if (!titleVal) {
      titleInput?.focus()
      return
    }

    const activeType = (
      modalRoot.querySelector<HTMLInputElement>('input[name="ia-map-type"]:checked')
    )?.value ?? 'none'

    const mappingUpdates = resolveMappingUpdates(activeType, plannedSites, modalRoot)

    if (isNew) {
      const newNode = addNode(titleVal, parentId ?? null)
      Object.assign(newNode, mappingUpdates)
    } else if (node) {
      updateNodeData(node.id, { title: titleVal, ...mappingUpdates })
    }

    closeModal()
    render()
    void save()
  })
}

function selectedSiteHtml(site: SharePointSite): string {
  return `
    <span class="ia-selected-label">Selected:</span>
    <a href="${escHtml(site.webUrl)}" target="_blank" rel="noopener">${escHtml(site.displayName)}</a>
    <button class="ia-btn ia-btn--sm ia-btn--ghost" id="ia-clear-site">Clear</button>
  `
}

function resolveMappingUpdates(
  mappingType: string,
  plannedSites: Array<{ id: string; displayName: string }>,
  modalRoot: HTMLElement
): Partial<IANode> {
  // Always clear all mapping fields first
  const updates: Partial<IANode> = {
    mappedSiteId: undefined,
    mappedSiteName: undefined,
    mappedSiteUrl: undefined,
    plannedMappingId: undefined,
    plannedSiteDisplayName: undefined,
  }
  if (mappingType === 'existing' && _modalSelectedSite) {
    updates.mappedSiteId = _modalSelectedSite.id
    updates.mappedSiteName = _modalSelectedSite.displayName
    updates.mappedSiteUrl = _modalSelectedSite.webUrl
  } else if (mappingType === 'planned') {
    const sel = modalRoot.querySelector<HTMLSelectElement>('#ia-planned-select')
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

  const q = query.trim()
  if (q.length < 2) {
    resultsEl.innerHTML = ''
    resultsEl.classList.add('ia-hidden')
    return
  }

  resultsEl.innerHTML = '<li class="ia-dropdown-item ia-dropdown-loading">Searching…</li>'
  resultsEl.classList.remove('ia-hidden')

  try {
    const sites = await searchSites(q)
    if (!document.getElementById('ia-site-results')) return // modal was closed

    if (sites.length === 0) {
      resultsEl.innerHTML = '<li class="ia-dropdown-item ia-dropdown-empty">No sites found</li>'
      return
    }

    resultsEl.innerHTML = sites.slice(0, 10).map(s =>
      `<li class="ia-dropdown-item" role="option"
        data-site-id="${escHtml(s.id)}"
        data-site-name="${escHtml(s.displayName)}"
        data-site-url="${escHtml(s.webUrl)}">
        <span class="ia-site-result-name">${escHtml(s.displayName)}</span>
        <span class="ia-site-result-url">${escHtml(s.webUrl)}</span>
      </li>`
    ).join('')

    resultsEl.querySelectorAll<HTMLElement>('.ia-dropdown-item[data-site-id]').forEach(item => {
      item.addEventListener('click', () => {
        _modalSelectedSite = {
          id: item.dataset.siteId!,
          name: item.dataset.siteName!,
          displayName: item.dataset.siteName!,
          webUrl: item.dataset.siteUrl!,
        }
        const searchInput = document.getElementById('ia-site-search') as HTMLInputElement | null
        if (searchInput) searchInput.value = _modalSelectedSite.displayName

        const selectedDiv = document.getElementById('ia-selected-site')
        if (selectedDiv) {
          selectedDiv.innerHTML = selectedSiteHtml(_modalSelectedSite)
          selectedDiv.classList.remove('ia-hidden')
          document.getElementById('ia-clear-site')?.addEventListener('click', () => {
            _modalSelectedSite = null
            if (searchInput) searchInput.value = ''
            selectedDiv.classList.add('ia-hidden')
            resultsEl.classList.add('ia-hidden')
          })
        }
        resultsEl.classList.add('ia-hidden')
      })
    })
  } catch (err) {
    const el = document.getElementById('ia-site-results')
    if (el) el.innerHTML = '<li class="ia-dropdown-item ia-dropdown-empty">Search failed</li>'
    console.error('IA Designer site search error', err)
  }
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectStyles(): void {
  if (document.getElementById('ia-designer-styles')) return
  const style = document.createElement('style')
  style.id = 'ia-designer-styles'
  style.textContent = `
    /* Panel layout */
    .ia-panel { display: flex; flex-direction: column; height: calc(100vh - 140px); overflow: hidden; }
    .ia-toolbar { display: flex; align-items: center; gap: 12px; padding: 10px 16px;
      border-bottom: 1px solid var(--color-border); background: var(--color-surface-alt); flex-shrink: 0; }
    .ia-save-status { font-size: 0.8rem; color: var(--color-text-muted); }
    .ia-canvas { flex: 1; overflow: auto; padding: 24px 32px; }
    .ia-empty { text-align: center; color: var(--color-text-muted); font-size: 0.9rem;
      padding: 72px 20px; line-height: 2; }

    /* Tree */
    .ia-tree { display: flex; flex-direction: column; min-width: max-content; }
    .ia-node-wrap { display: flex; flex-direction: column; }
    .ia-children-wrap { margin-left: 28px; padding-left: 20px;
      border-left: 2px solid var(--color-border); }

    /* Gap drop zones */
    .ia-gap { height: 8px; border-radius: 3px; transition: height 0.12s, background 0.12s, box-shadow 0.12s; }
    .ia-gap--active { height: 4px !important; background: var(--color-primary) !important;
      box-shadow: 0 0 0 2px rgba(0,120,212,0.25); }

    /* Node cards */
    .ia-node-card { display: flex; align-items: center; gap: 8px; padding: 9px 12px;
      border: 1.5px solid var(--color-border); border-radius: 8px; background: white;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06); transition: box-shadow 0.15s, border-color 0.15s, background 0.15s;
      min-width: 200px; width: fit-content; max-width: 420px; user-select: none; cursor: default; }
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

    /* Buttons */
    .ia-btn { border: none; border-radius: 4px; cursor: pointer; transition: background 0.1s, color 0.1s; }
    .ia-btn--icon { background: none; border: none; color: var(--color-text-muted);
      font-size: 0.95rem; padding: 3px 6px; }
    .ia-btn--icon:hover { background: var(--color-surface-alt); color: var(--color-text); }
    .ia-btn--danger:hover { background: #fde7e9; color: #c50f1f; }
    .ia-btn--sm { padding: 2px 8px; font-size: 0.78rem; }
    .ia-btn--ghost { background: none; border: 1px solid var(--color-border);
      color: var(--color-text); padding: 2px 8px; }
    .ia-btn--ghost:hover { background: var(--color-surface-alt); }

    /* Modal */
    .ia-dialog { min-width: 460px; max-width: 560px; }
    .ia-modal-body { padding: 0 24px 4px; display: flex; flex-direction: column; gap: 0; }
    .ia-modal-body .form-group { margin-bottom: 16px; }
    .ia-modal-footer { padding: 16px 24px; border-top: 1px solid var(--color-border);
      display: flex; gap: 8px; }
    .ia-hidden { display: none !important; }
    .ia-radio-group { display: flex; flex-direction: column; gap: 6px; }
    .ia-hint { font-size: 0.8rem; color: var(--color-text-muted); margin: 6px 0 0; line-height: 1.4; }
    .ia-hint-inline { font-size: 0.78rem; color: var(--color-text-muted); }

    /* Site search */
    .ia-search-wrap { position: relative; }
    .ia-dropdown { position: absolute; top: calc(100% + 2px); left: 0; right: 0; background: white;
      border: 1px solid var(--color-border); border-radius: 6px; box-shadow: var(--shadow);
      z-index: 40; list-style: none; padding: 4px 0; margin: 0; max-height: 200px; overflow-y: auto; }
    .ia-dropdown-item { padding: 8px 12px; cursor: pointer; display: flex; flex-direction: column; gap: 1px; }
    .ia-dropdown-item:hover { background: var(--color-surface-alt); }
    .ia-dropdown-loading, .ia-dropdown-empty { color: var(--color-text-muted); cursor: default; font-size: 0.85rem; }
    .ia-site-result-name { font-size: 0.875rem; font-weight: 500; }
    .ia-site-result-url { font-size: 0.75rem; color: var(--color-text-muted); }
    .ia-selected-site { display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
      font-size: 0.85rem; margin-top: 8px; }
    .ia-selected-label { color: var(--color-text-muted); }
    .ia-selected-site a { color: var(--color-primary); text-decoration: none; }
    .ia-selected-site a:hover { text-decoration: underline; }
  `
  document.head.appendChild(style)
}
