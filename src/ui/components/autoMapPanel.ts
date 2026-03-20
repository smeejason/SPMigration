import {
  findUserForOneDrive,
  getUserDrive,
  saveMappingsFile,
  searchUsers,
} from '../../graph/graphClient'
import { updateProject, getSpConfig } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { TreeNode, MigrationMapping, OneDriveMatchStatus, AppUser } from '../../types'

// Light-red folder with X — used for "Can't Find" flagged folders
const CANT_FIND_FOLDER_ICON = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 17" width="18" height="15" style="display:inline-block;vertical-align:-2px;flex-shrink:0"><path d="M1 3A1 1 0 0 1 2 2H7.5L9.5 4H18A1 1 0 0 1 19 5V15A1 1 0 0 1 18 16H2A1 1 0 0 1 1 15V3Z" fill="#ffd6d6" stroke="#e08080" stroke-width="0.5"/><line x1="8" y1="8" x2="13" y2="13" stroke="#c50000" stroke-width="1.8" stroke-linecap="round"/><line x1="13" y1="8" x2="8" y2="13" stroke="#c50000" stroke-width="1.8" stroke-linecap="round"/></svg>`

// ─── Entry point ──────────────────────────────────────────────────────────────

export function renderAutoMapPanel(container: HTMLElement): void {
  injectAutoMapStyles()
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

  const settings = state.currentProject?.projectData.autoMapSettings
  let selectedLevel = settings?.selectedLevel ?? -1
  const topNodes = !tree.path ? tree.children : [tree]
  // Phase 1 results are stored in state.mappings with matchStatus set
  const existingMappings = state.mappings.filter(m => m.matchStatus !== undefined)
  const allMappings = state.mappings
  const phase1Done = existingMappings.length > 0
  // hasMatchedUsers was used by the removed Phase 2 section
  const totalAtLevel = selectedLevel >= 0 ? countNodesAtDepth(tree, selectedLevel) : 0

  container.innerHTML = `
    <div class="automap-panel">
      <div class="automap-left">
        <div class="automap-section-header">
          <h3>Source: File System</h3>
          <span class="automap-hint">Click any folder to select that level</span>
        </div>
        <div id="automap-tree-wrap" class="automap-tree" data-selected-level="${selectedLevel}"></div>
      </div>
      <div class="automap-right">
        <div class="automap-right-inner">

          <div class="automap-settings">
            <div class="form-group">
              <label>Migration Account</label>
              <div class="people-picker" id="migration-people-picker">
                <div class="people-picker-input-wrap">
                  <input id="migration-account-search" type="text" class="form-input people-picker-input"
                    placeholder="Search for user…"
                    autocomplete="off"
                    value="${escHtml(settings?.migrationAccount ?? '')}" />
                  <button type="button" class="people-picker-clear" id="migration-account-clear" style="display:${settings?.migrationAccount ? '' : 'none'}">✕</button>
                </div>
                <div class="people-picker-dropdown" id="migration-account-dropdown" style="display:none"></div>
                <input type="hidden" id="migration-account" value="${escHtml(settings?.migrationAccount ?? '')}" />
              </div>
            </div>
          </div>

          <div class="level-banner" id="level-banner">
            <span id="level-label" class="level-label-text">${selectedLevel >= 0 ? levelBannerText(tree, selectedLevel) : 'No level selected — click a folder in the tree'}</span>
            <button id="btn-confirm-level" class="btn btn-primary btn-sm" ${selectedLevel >= 0 ? '' : 'disabled'}>Confirm Level</button>
          </div>

          <div class="automap-phase-section">
            <div class="phase-header"><span class="phase-num">Phase 1</span> Find &amp; Map Users</div>
            <button id="btn-phase1" class="btn btn-primary" ${phase1Done || selectedLevel >= 0 ? '' : 'disabled'}>
              ${phase1Done ? 'Re-run Phase 1' : 'Find &amp; Map Users'}
            </button>
            <div id="phase1-progress" ${phase1Done ? '' : 'style="display:none"'}>
              <div class="progress-bar-wrap">
                <div id="phase1-bar" class="progress-bar" style="width:${totalAtLevel > 0 ? Math.round(existingMappings.length / totalAtLevel * 100) : 0}%"></div>
              </div>
              <div class="progress-stats">
                <span id="phase1-count">${existingMappings.length} / ${totalAtLevel}</span>
                <span class="stat-matched" id="phase1-matched">✓ Matched: ${existingMappings.filter(m => m.matchStatus === 'matched').length}</span>
                <span class="stat-notfound" id="phase1-notfound">✗ Not found: ${existingMappings.filter(m => m.matchStatus === 'not-found').length}</span>
                <span class="stat-ambiguous" id="phase1-ambiguous">? Ambiguous: ${existingMappings.filter(m => m.matchStatus === 'ambiguous').length}</span>
                <span class="stat-error" id="phase1-error">${existingMappings.filter(m => m.matchStatus === 'error').length > 0 ? `⚠ Errors: ${existingMappings.filter(m => m.matchStatus === 'error').length}` : ''}</span>
              </div>
            </div>
          </div>



        </div>
      </div>
    </div>
  `

  // ── Tree ──────────────────────────────────────────────────────────────────
  const treeWrap = container.querySelector('#automap-tree-wrap') as HTMLElement
  const ul = document.createElement('ul')
  ul.className = 'tree-list tree-root'

  const onLevelSelect = (depth: number): void => {
    selectedLevel = depth
    treeWrap.dataset.selectedLevel = String(depth)
    ;(container.querySelector('#level-label') as HTMLElement).textContent = levelBannerText(tree, depth)
    ;(container.querySelector('#btn-confirm-level') as HTMLButtonElement).disabled = false
    if (!(container.querySelector('#btn-phase1') as HTMLButtonElement).textContent?.includes('Re-run')) {
      ;(container.querySelector('#btn-phase1') as HTMLButtonElement).disabled = false
    }
  }

  for (const node of topNodes) {
    ul.appendChild(createAutoMapNodeEl(node, onLevelSelect, allMappings, true))
  }
  treeWrap.appendChild(ul)

  // Auto-expand: if a level is selected, open all ancestor depths so that level
  // is visible; otherwise just open the root when there's a single top node.
  if (selectedLevel >= 0) {
    expandToLevel(treeWrap, selectedLevel)
  } else if (topNodes.length === 1) {
    ul.querySelector<HTMLButtonElement>('.automap-toggle-btn:not(.invisible)')?.click()
  }

  // ── People Picker: Migration Account ──────────────────────────────────────
  wirePeoplePicker(
    container,
    'migration-account-search',
    'migration-account-dropdown',
    'migration-account-clear',
    'migration-account',
    async (upn) => {
      const project = getState().currentProject
      if (!project) return
      const existing = project.projectData.autoMapSettings
      const updatedData = {
        ...project.projectData,
        autoMapSettings: {
          selectedLevel: existing?.selectedLevel ?? selectedLevel,
          migrationAccount: upn,
          targetFolderPath: existing?.targetFolderPath ?? '',
        },
      }
      try {
        await updateProject(project.id, { projectData: updatedData })
        setState({ currentProject: { ...project, projectData: updatedData } })
      } catch { /* non-fatal */ }
    }
  )

  // ── Confirm Level ──────────────────────────────────────────────────────────
  const confirmBtn = container.querySelector('#btn-confirm-level') as HTMLButtonElement
  confirmBtn.addEventListener('click', async () => {
    const migrationAccount = (container.querySelector('#migration-account') as HTMLInputElement).value.trim()
    const project = getState().currentProject!
    const updatedData = { ...project.projectData, autoMapSettings: { selectedLevel, migrationAccount, targetFolderPath: '' } }
    try {
      await updateProject(project.id, { projectData: updatedData })
      setState({ currentProject: { ...project, projectData: updatedData } })
    } catch { /* non-fatal */ }
    confirmBtn.textContent = '✓ Confirmed'
    confirmBtn.disabled = true
    ;(container.querySelector('#btn-phase1') as HTMLButtonElement).disabled = false
  })

  // ── Phase 1 ───────────────────────────────────────────────────────────────
  const phase1Btn = container.querySelector('#btn-phase1') as HTMLButtonElement
  phase1Btn.addEventListener('click', async () => {
    if (selectedLevel < 0) return
    const nodesToProcess = collectNodesAtDepth(tree, selectedLevel)
    if (nodesToProcess.length === 0) return

    phase1Btn.disabled = true
    phase1Btn.textContent = 'Running…'
    ;(container.querySelector('#phase1-progress') as HTMLElement).style.display = ''

    await runPhase1(container, nodesToProcess, '')

    phase1Btn.textContent = 'Re-run Phase 1'
    phase1Btn.disabled = false
  })
}

// ─── Tree node factory ────────────────────────────────────────────────────────

function createAutoMapNodeEl(
  node: TreeNode,
  onLevelSelect: (depth: number) => void,
  existingMappings: MigrationMapping[],
  isRoot = false
): HTMLLIElement {
  const li = document.createElement('li')
  li.className = `automap-node${isRoot ? ' automap-node--root' : ''}`

  const hasChildren = node.children.length > 0
  const isFolder = !node.name.includes('*')

  const autoMapping    = existingMappings.find(m => m.id === node.path && m.matchStatus === 'matched')
  const manualMapping  = existingMappings.find(m => m.id === node.path && m.matchStatus === undefined && !!(m.targetSite || m.targetDrive || m.plannedSite))
  const cantFindMapping = existingMappings.find(m => m.id === node.path && m.matchStatus === 'cant-find')
  const isCantFind = !!cantFindMapping

  const row = document.createElement('div')
  row.className = `automap-row${(autoMapping || manualMapping) ? ' automap-row--mapped' : ''}${isCantFind ? ' automap-row--cant-find' : ''}`
  row.dataset.path = node.path
  row.dataset.depth = String(node.depth)

  // Toggle button
  const toggleBtn = document.createElement('button')
  toggleBtn.type = 'button'
  toggleBtn.className = `automap-toggle-btn${hasChildren ? '' : ' invisible'}`
  const toggleIcon = document.createElement('span')
  toggleIcon.className = 'toggle-icon'
  toggleIcon.textContent = '▶'
  toggleBtn.appendChild(toggleIcon)

  // Folder icon — light-red folder with X when flagged as can't find
  const iconWrap = document.createElement('span')
  iconWrap.className = 'automap-icon'
  iconWrap.innerHTML = isCantFind ? CANT_FIND_FOLDER_ICON : (isFolder ? '📁' : '📄')

  // Name
  const nameEl = document.createElement('span')
  nameEl.className = 'automap-name'
  nameEl.textContent = isFolder ? String(node.name || node.path || '(unnamed)') : 'Loose files'
  if (node.originalPath) nameEl.title = node.originalPath

  // Level badge (1-indexed)
  const levelBadge = document.createElement('span')
  levelBadge.className = 'automap-level-badge'
  levelBadge.textContent = `L${node.depth + 1}`

  // Map-status badge: Auto | Manual | Can't Find | (empty)
  const mapBadge = document.createElement('span')
  mapBadge.dataset.mapBadgeFor = node.path
  if (autoMapping) {
    mapBadge.className = 'automap-map-badge map-badge--auto'
    mapBadge.textContent = 'Auto'
  } else if (manualMapping) {
    mapBadge.className = 'automap-map-badge map-badge--manual'
    mapBadge.textContent = 'Manual'
  } else if (isCantFind) {
    mapBadge.className = 'automap-map-badge map-badge--cant-find'
    mapBadge.textContent = "Can't Find"
  } else {
    mapBadge.className = 'automap-map-badge'
  }

  // Can't Find toggle button — quick flag for this folder
  const cantFindBtn = document.createElement('button')
  cantFindBtn.type = 'button'
  cantFindBtn.className = `automap-cant-find-btn${isCantFind ? ' is-cant-find' : ''}`
  cantFindBtn.dataset.cantFindFor = node.path
  cantFindBtn.textContent = isCantFind ? '↩ Clear' : "Can't Find"
  cantFindBtn.title = isCantFind ? 'Remove Can\'t Find flag' : 'Flag this folder as Can\'t Find'
  cantFindBtn.addEventListener('click', async (e) => {
    e.stopPropagation()
    const isCF = getState().mappings.find(m => m.id === node.path)?.matchStatus === 'cant-find'
    if (isCF) {
      setState({ mappings: getState().mappings.filter(m => m.id !== node.path) })
      iconWrap.innerHTML = isFolder ? '📁' : '📄'
      mapBadge.className = 'automap-map-badge'
      mapBadge.textContent = ''
      cantFindBtn.classList.remove('is-cant-find')
      cantFindBtn.textContent = "Can't Find"
      cantFindBtn.title = 'Flag this folder as Can\'t Find'
      row.classList.remove('automap-row--cant-find')
    } else {
      const mapping: MigrationMapping = {
        id: node.path,
        sourceNode: node,
        targetSite: null,
        targetDrive: null,
        targetFolderPath: '',
        status: 'error',
        matchStatus: 'cant-find',
        accessStatus: 'unknown',
        resolvedDisplayName: folderNameToDisplayName(node.name),
      }
      setState({ mappings: [...getState().mappings.filter(m => m.id !== node.path), mapping] })
      iconWrap.innerHTML = CANT_FIND_FOLDER_ICON
      mapBadge.className = 'automap-map-badge map-badge--cant-find'
      mapBadge.textContent = "Can't Find"
      cantFindBtn.classList.add('is-cant-find')
      cantFindBtn.textContent = '↩ Clear'
      cantFindBtn.title = "Remove Can't Find flag"
      row.classList.add('automap-row--cant-find')
      row.classList.remove('automap-row--mapped')
    }
    await persistCantFindMappings()
  })

  row.appendChild(toggleBtn)
  row.appendChild(iconWrap)
  row.appendChild(nameEl)
  row.appendChild(levelBadge)
  row.appendChild(mapBadge)
  row.appendChild(cantFindBtn)
  li.appendChild(row)

  // Level selection on row click
  if (isFolder) {
    row.addEventListener('click', () => { onLevelSelect(node.depth) })
  }

  // Lazy-render children
  if (hasChildren) {
    let childrenLoaded = false
    toggleBtn.addEventListener('click', (e) => {
      e.stopPropagation()
      const isOpen = li.classList.contains('automap-node--open')
      if (isOpen) {
        li.querySelector<HTMLElement>(':scope > .tree-children')!.style.display = 'none'
        li.classList.remove('automap-node--open')
        toggleIcon.textContent = '▶'
      } else {
        if (!childrenLoaded) {
          const childUl = document.createElement('ul')
          childUl.className = 'tree-list tree-children'
          for (const child of node.children) {
            childUl.appendChild(createAutoMapNodeEl(child, onLevelSelect, getState().mappings))
          }
          li.appendChild(childUl)
          childrenLoaded = true
        } else {
          li.querySelector<HTMLElement>(':scope > .tree-children')!.style.display = ''
        }
        li.classList.add('automap-node--open')
        toggleIcon.textContent = '▼'
      }
    })
  }

  return li
}

// ─── Phase 1: find & map users ────────────────────────────────────────────────

async function runPhase1(
  container: HTMLElement,
  nodes: TreeNode[],
  targetFolderPath: string
): Promise<void> {
  const BATCH_SIZE = 5
  const accumulated: MigrationMapping[] = []

  const barEl = container.querySelector('#phase1-bar') as HTMLElement
  const countEl = container.querySelector('#phase1-count') as HTMLElement
  const matchedEl = container.querySelector('#phase1-matched') as HTMLElement
  const notFoundEl = container.querySelector('#phase1-notfound') as HTMLElement
  const ambiguousEl = container.querySelector('#phase1-ambiguous') as HTMLElement
  const errorEl = container.querySelector('#phase1-error') as HTMLElement

  const updateUI = (): void => {
    const total = nodes.length
    const done = accumulated.length
    barEl.style.width = `${total > 0 ? Math.round(done / total * 100) : 100}%`
    countEl.textContent = `${done} / ${total}`
    matchedEl.textContent = `✓ Matched: ${accumulated.filter(m => m.matchStatus === 'matched').length}`
    notFoundEl.textContent = `✗ Not found: ${accumulated.filter(m => m.matchStatus === 'not-found').length}`
    ambiguousEl.textContent = `? Ambiguous: ${accumulated.filter(m => m.matchStatus === 'ambiguous').length}`
    const errCount = accumulated.filter(m => m.matchStatus === 'error').length
    if (errorEl) errorEl.textContent = errCount > 0 ? `⚠ Errors: ${errCount}` : ''
  }

  for (let i = 0; i < nodes.length; i += BATCH_SIZE) {
    const batch = nodes.slice(i, i + BATCH_SIZE)

    await Promise.all(batch.map(async (node) => {
      const resolvedDisplayName = folderNameToDisplayName(node.name)
      let matchStatus: OneDriveMatchStatus = 'error'
      let matchedUser: AppUser | null = null
      let driveId = ''
      let driveWebUrl = ''
      let errorMsg: string | undefined

      try {
        const result = await findUserForOneDrive(resolvedDisplayName)
        matchedUser = result.user
        matchStatus = result.status
        if (matchedUser) {
          const drive = await getUserDrive(matchedUser.id)
          driveId = drive?.id ?? ''
          driveWebUrl = drive?.webUrl ?? ''
        }
      } catch (err) {
        matchStatus = 'error'
        errorMsg = (err as Error).message
      }

      // Store ALL Phase 1 results (including not-found/ambiguous/error) in state.mappings
      // so AutoMap and Map page share a single source of truth.
      const mapping: MigrationMapping = {
        id: node.path,
        sourceNode: node,
        targetSite: matchedUser
          ? { id: matchedUser.id, displayName: matchedUser.displayName, webUrl: driveWebUrl, name: matchedUser.displayName }
          : null,
        targetDrive: driveId
          ? { id: driveId, name: 'OneDrive', webUrl: driveWebUrl, driveType: 'personal' }
          : null,
        targetFolderPath,
        status: matchedUser ? 'ready' : 'error',
        matchStatus,
        accessStatus: 'unknown',
        resolvedDisplayName,
        ...(errorMsg ? { notes: errorMsg } : {}),
      }
      accumulated.push(mapping)

      // Update map-status badge to Auto
      const mapBadgeEl = container.querySelector(`[data-map-badge-for="${CSS.escape(node.path)}"]`) as HTMLElement | null
      if (mapBadgeEl) {
        mapBadgeEl.className = 'automap-map-badge map-badge--auto'
        mapBadgeEl.textContent = 'Auto'
      }
      const rowEl = container.querySelector(`.automap-row[data-path="${CSS.escape(node.path)}"]`) as HTMLElement | null
      if (rowEl) rowEl.classList.add('automap-row--mapped')
    }))

    // Merge accumulated into state.mappings: keep manual entries (no matchStatus), replace Phase 1 entries
    const manualMappings = getState().mappings.filter(m => m.matchStatus === undefined)
    setState({ mappings: [...manualMappings, ...accumulated] })
    updateUI()
    await new Promise(r => setTimeout(r, 0))
  }

  // Persist after Phase 1 complete
  try {
    const project = getState().currentProject!
    const { siteId } = getSpConfig()

    // Keep manual mappings that aren't overridden by Phase 1 results
    const accIds = new Set(accumulated.map(m => m.id))
    const manualMappings = getState().mappings.filter(m => m.matchStatus === undefined && !accIds.has(m.id))
    const merged: MigrationMapping[] = [...manualMappings, ...accumulated]

    await saveMappingsFile(siteId, project.title, project.id, merged)

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { mappings: _removed, ...restData } = project.projectData
    const updatedData = {
      ...restData,
      oneDriveMappingCount: accumulated.length,
      mappingCount: merged.filter(m => m.targetSite || m.plannedSite).length,
    }
    await updateProject(project.id, { projectData: updatedData })
    setState({ mappings: merged, currentProject: { ...project, projectData: updatedData } })
  } catch (err) {
    console.warn('[AutoMap] Failed to persist mappings:', err)
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function persistCantFindMappings(): Promise<void> {
  try {
    const state = getState()
    const project = state.currentProject!
    const { siteId } = getSpConfig()
    await saveMappingsFile(siteId, project.title, project.id, state.mappings)
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { mappings: _r, ...restData } = project.projectData
    const updatedData = {
      ...restData,
      mappingCount: state.mappings.filter(m => m.targetSite || m.plannedSite).length,
    }
    await updateProject(project.id, { projectData: updatedData })
    setState({ currentProject: { ...project, projectData: updatedData } })
  } catch (err) {
    console.warn("[AutoMap] Failed to persist Can't Find flag:", err)
  }
}

function folderNameToDisplayName(name: string): string {
  return name
    .replace(/[._-]+/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(Boolean)
    .map(w => w[0].toUpperCase() + w.slice(1).toLowerCase())
    .join(' ')
}

/**
 * Expand the tree so that nodes at `targetDepth` are visible.
 * Works depth-by-depth because children are lazy-rendered on toggle click.
 */
function expandToLevel(treeWrap: HTMLElement, targetDepth: number): void {
  for (let d = 0; d < targetDepth; d++) {
    treeWrap.querySelectorAll<HTMLElement>(`.automap-row[data-depth="${d}"]`).forEach(row => {
      const li = row.closest<HTMLElement>('.automap-node')
      if (li && !li.classList.contains('automap-node--open')) {
        li.querySelector<HTMLButtonElement>('.automap-toggle-btn:not(.invisible)')?.click()
      }
    })
  }
}

function collectNodesAtDepth(root: TreeNode, depth: number): TreeNode[] {
  const result: TreeNode[] = []
  function walk(node: TreeNode): void {
    if (node.depth === depth) { result.push(node); return }
    if (node.depth < depth) node.children.forEach(walk)
  }
  walk(root)
  return result
}

function countNodesAtDepth(root: TreeNode, depth: number): number {
  return collectNodesAtDepth(root, depth).length
}

function levelBannerText(tree: TreeNode, depth: number): string {
  const count = countNodesAtDepth(tree, depth)
  return `Level ${depth + 1} selected — ${count} folder${count !== 1 ? 's' : ''} will be mapped`
}


function wirePeoplePicker(
  container: HTMLElement,
  searchId: string,
  dropdownId: string,
  clearId: string,
  hiddenId: string,
  onChange?: (upn: string) => void
): void {
  const searchInput = container.querySelector(`#${searchId}`) as HTMLInputElement
  const dropdown = container.querySelector(`#${dropdownId}`) as HTMLElement
  const clearBtn = container.querySelector(`#${clearId}`) as HTMLElement
  const hidden = container.querySelector(`#${hiddenId}`) as HTMLInputElement

  let debounce: ReturnType<typeof setTimeout> | null = null

  const closeDropdown = (): void => {
    dropdown.style.display = 'none'
    dropdown.innerHTML = ''
  }

  const selectUser = (displayName: string, upn: string): void => {
    searchInput.value = `${displayName} (${upn})`
    hidden.value = upn
    clearBtn.style.display = ''
    closeDropdown()
    onChange?.(upn)
  }

  searchInput.addEventListener('input', () => {
    const q = searchInput.value.trim()
    if (!q) { hidden.value = ''; clearBtn.style.display = 'none'; closeDropdown(); return }
    if (debounce) clearTimeout(debounce)
    debounce = setTimeout(async () => {
      try {
        const users = await searchUsers(q)
        if (users.length === 0) { closeDropdown(); return }
        dropdown.innerHTML = ''
        users.forEach(u => {
          const item = document.createElement('div')
          item.className = 'people-picker-item'
          item.innerHTML = `<span class="pp-name">${escHtml(u.displayName)}</span><span class="pp-upn">${escHtml(u.userPrincipalName ?? '')}</span>`
          item.addEventListener('mousedown', (e) => {
            e.preventDefault()
            selectUser(u.displayName, u.userPrincipalName ?? '')
          })
          dropdown.appendChild(item)
        })
        dropdown.style.display = ''
      } catch { closeDropdown() }
    }, 300)
  })

  searchInput.addEventListener('blur', () => { setTimeout(closeDropdown, 150) })

  clearBtn.addEventListener('click', () => {
    searchInput.value = ''
    hidden.value = ''
    clearBtn.style.display = 'none'
    searchInput.focus()
    closeDropdown()
    onChange?.('')
  })
}


function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectAutoMapStyles(): void {
  if (document.getElementById('automap-styles')) return
  const style = document.createElement('style')
  style.id = 'automap-styles'
  style.textContent = `
    .automap-panel { display: grid; grid-template-columns: 2fr 1fr; height: calc(100vh - 140px); overflow: hidden; }
    .automap-left, .automap-right { overflow-y: auto; }
    .automap-left { border-right: 1px solid var(--color-border); }
    .automap-section-header { padding: 12px 16px; border-bottom: 1px solid var(--color-border);
      display: flex; align-items: center; justify-content: space-between;
      background: var(--color-surface-alt); position: sticky; top: 0; z-index: 1; }
    .automap-section-header h3 { font-size: 0.9rem; font-weight: 600; margin: 0; }
    .automap-hint { font-size: 0.78rem; color: var(--color-text-muted); }
    .automap-hint-inline { font-size: 0.78rem; color: var(--color-text-muted); font-weight: 400; }

    /* Tree */
    .tree-list { list-style: none; padding: 0; margin: 0; }
    .tree-children { padding-left: 20px; border-left: 1px solid var(--color-border); margin-left: 18px; }
    .automap-tree { padding: 8px; }
    .automap-node { margin: 1px 0; }
    .automap-node--root > .automap-row { font-weight: 600; }
    .automap-row { display: flex; align-items: center; gap: 6px; padding: 5px 8px; border-radius: 4px;
      cursor: pointer; user-select: none; transition: background 0.1s; }
    .automap-row--mapped { background: rgba(16, 124, 16, 0.06); }
    .automap-row:hover { background: var(--color-primary-light); }
    .automap-toggle-btn { background: none; border: none; cursor: pointer; width: 16px;
      font-size: 0.65rem; color: var(--color-text-muted); padding: 0; flex-shrink: 0; }
    .automap-toggle-btn.invisible { visibility: hidden; pointer-events: none; }
    .automap-icon { flex-shrink: 0; }
    .automap-name { flex: 1; font-size: 0.875rem; font-family: 'Consolas', monospace;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
    .automap-level-badge { font-size: 0.68rem; font-weight: 700; background: var(--color-surface-alt);
      color: var(--color-text-muted); border: 1px solid var(--color-border);
      padding: 1px 5px; border-radius: 10px; flex-shrink: 0; white-space: nowrap; }
    .automap-map-badge { font-size: 0.7rem; font-weight: 700; flex-shrink: 0;
      padding: 1px 8px; border-radius: 10px; min-width: 52px; text-align: center; }
    .map-badge--auto      { background: #dff6dd; color: #107c10; }
    .map-badge--manual    { background: #e8f4fd; color: #0078d4; }
    .map-badge--cant-find { background: #fde7e9; color: #a4262c; }
    .automap-row--cant-find { background: rgba(164, 38, 44, 0.05) !important; }
    .automap-cant-find-btn {
      font-size: 0.68rem; font-weight: 600; padding: 1px 7px; border-radius: 10px;
      border: 1px solid #ddb0b0; background: transparent; color: #a4262c;
      cursor: pointer; flex-shrink: 0; white-space: nowrap; opacity: 0.5; transition: opacity 0.12s;
    }
    .automap-cant-find-btn:hover { opacity: 1; background: #fde7e9; }
    .automap-cant-find-btn.is-cant-find { background: #fde7e9; opacity: 1; border-color: #a4262c; }

    /* Level highlighting via data attribute on tree container */
    ${Array.from({ length: 16 }, (_, i) =>
      `.automap-tree[data-selected-level="${i}"] .automap-row[data-depth="${i}"] {
        background: rgba(0, 120, 212, 0.1); border-left: 3px solid var(--color-primary); }`
    ).join('\n    ')}
    ${Array.from({ length: 16 }, (_, i) =>
      `.automap-tree[data-selected-level="${i}"] .automap-row[data-depth="${i}"] .automap-level-badge {
        background: var(--color-primary); color: white; border-color: var(--color-primary); }`
    ).join('\n    ')}

    /* Right panel */
    .automap-right-inner { padding: 20px; display: flex; flex-direction: column; gap: 20px; }
    .automap-settings { display: flex; flex-direction: column; gap: 12px; }
    .automap-settings .form-group { margin-bottom: 0; }

    /* Level banner */
    .level-banner { display: flex; align-items: center; justify-content: space-between; gap: 12px;
      padding: 12px 16px; background: var(--color-surface-alt);
      border: 1px solid var(--color-border); border-radius: 6px; }
    .level-label-text { font-size: 0.875rem; color: var(--color-text); flex: 1; }

    /* Phase sections */
    .automap-phase-section { border: 1px solid var(--color-border); border-radius: 6px;
      overflow: hidden; }
    .phase-header { padding: 10px 14px; background: var(--color-surface-alt);
      border-bottom: 1px solid var(--color-border);
      font-size: 0.875rem; font-weight: 600; color: var(--color-text); }
    .phase-num { display: inline-flex; align-items: center; justify-content: center;
      background: var(--color-primary); color: white; border-radius: 50%;
      width: 20px; height: 20px; font-size: 0.72rem; font-weight: 700; margin-right: 6px; }
    .automap-phase-section .btn { margin: 14px; }
    .progress-bar-wrap { margin: 0 14px 8px; height: 8px; background: var(--color-border);
      border-radius: 4px; overflow: hidden; }
    .progress-bar { height: 100%; background: var(--color-primary); border-radius: 4px;
      transition: width 0.25s ease; }
    .progress-stats { padding: 0 14px 14px; display: flex; flex-wrap: wrap; gap: 8px 16px;
      font-size: 0.8rem; }
    .stat-matched { color: #107c10; }
    .stat-notfound { color: #a4262c; }
    .stat-ambiguous { color: #7a5900; }
    .stat-error { color: #a4262c; }

    /* People picker */
    .people-picker { position: relative; }
    .people-picker-input-wrap { display: flex; align-items: center; position: relative; }
    .people-picker-input { flex: 1; padding-right: 28px !important; }
    .people-picker-clear { position: absolute; right: 6px; background: none; border: none;
      cursor: pointer; color: var(--color-text-muted); font-size: 0.75rem; padding: 2px 4px;
      line-height: 1; border-radius: 3px; }
    .people-picker-clear:hover { background: var(--color-border); color: var(--color-text); }
    .people-picker-dropdown { position: absolute; top: calc(100% + 2px); left: 0; right: 0;
      background: var(--color-surface); border: 1px solid var(--color-border);
      border-radius: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.12);
      z-index: 100; max-height: 200px; overflow-y: auto; }
    .people-picker-item { display: flex; flex-direction: column; padding: 8px 12px;
      cursor: pointer; border-bottom: 1px solid var(--color-border); gap: 1px; }
    .people-picker-item:last-child { border-bottom: none; }
    .people-picker-item:hover { background: var(--color-primary-light); }
    .pp-name { font-size: 0.875rem; font-weight: 500; color: var(--color-text); }
    .pp-upn { font-size: 0.75rem; color: var(--color-text-muted); }
  `
  document.head.appendChild(style)
}
