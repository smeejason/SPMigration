import { setState, getState } from '../../state/store'
import { updateProject, getSpConfig } from '../../graph/projectService'
import { getOrCreateProjectFolder, uploadFileToDrive, downloadDriveItem, saveMappingsFile } from '../../graph/graphClient'
import type { TreeNode, MigrationMapping, ExcelUpload } from '../../types'

// ─── Entry point ──────────────────────────────────────────────────────────────

export function renderUploadPanel(container: HTMLElement): void {
  const state = getState()
  const project = state.currentProject
  const existingTree = state.treeData
  const uploads = project?.projectData.uploads ?? []
  const activeId = project?.projectData.activeUploadId
    ?? (uploads.length > 0 ? uploads[uploads.length - 1].id : undefined)

  container.innerHTML = `
    <div class="upload-panel">

      ${uploads.length > 0 ? `
      <div class="panel-section">
        <h3>Upload History</h3>
        <p class="panel-desc">All TreeSize reports uploaded for this project, stored in SharePoint. The <strong>Active</strong> report drives the mapping view.</p>
        <div class="upload-history-list" id="history-list">
          ${renderHistoryItems(uploads, activeId, existingTree)}
        </div>
      </div>
      ` : ''}

      <div class="conflict-warning" id="conflict-warning" style="display:none">
        <div class="conflict-warning-header">
          <span>⚠ <strong>Mapping conflicts detected</strong></span>
          <button type="button" id="btn-dismiss-conflicts" class="btn-dismiss-conflicts">✕</button>
        </div>
        <p id="conflict-msg" class="conflict-msg"></p>
        <ul id="conflict-list" class="conflict-list"></ul>
      </div>

      <div class="panel-section">
        <h3>${uploads.length > 0 ? 'Upload New Report' : 'Upload TreeSize Report'}</h3>
        <p class="panel-desc">Export your file server structure from TreeSize Pro or Free as <strong>.csv</strong> or <strong>.xlsx</strong>, then upload it here. Each report is saved to SharePoint automatically.</p>

        <div id="drop-zone" class="drop-zone ${existingTree ? 'drop-zone--has-file' : ''}">
          <input type="file" id="file-input" accept=".csv,.xlsx,.xls" style="display:none" />
          <div class="drop-zone-content">
            <div class="drop-icon">📂</div>
            <p class="drop-label">Drag & drop your TreeSize export here, or</p>
            <button type="button" id="btn-browse" class="btn btn-primary btn-sm">Browse files</button>
            <p class="drop-hint">Accepts .csv and .xlsx · Saved to SharePoint automatically</p>
          </div>
        </div>

        <div id="upload-status" class="upload-status" style="display:none"></div>
      </div>

    </div>
  `
  injectUploadStyles()
  setupDropZone(container)
  setupHistoryButtons(container)
}

// ─── History ──────────────────────────────────────────────────────────────────

function renderHistoryItems(uploads: ExcelUpload[], activeId?: string, activeTree?: TreeNode | null): string {
  return [...uploads].reverse().map((u) => {
    const isActive = u.id === activeId
    const date = formatDate(new Date(u.uploadedAt))

    // Prefer stored stats; fall back to live tree for the active item (covers legacy uploads)
    let rowCount = u.rowCount
    let topFolderName = u.topFolderName
    let fileCount = u.fileCount
    let folderCount = u.folderCount
    let sizeBytes = u.sizeBytes
    if (isActive && activeTree && rowCount === undefined) {
      const topNode = findTopDataNode(activeTree)
      rowCount = countAllNodes(activeTree)
      topFolderName = topNode.name || topNode.path || 'Root'
      fileCount = activeTree.fileCount
      folderCount = activeTree.folderCount
      sizeBytes = activeTree.sizeBytes
    }
    const hasStats = rowCount !== undefined

    return `
      <div class="history-item${isActive ? ' history-item--active' : ''}">
        <div class="history-item-header">
          <span class="history-item-icon">📊</span>
          <div class="history-item-info">
            <span class="history-item-name" title="${escHtml(u.fileName)}">${escHtml(u.fileName)}</span>
            <span class="history-item-meta">${date}${hasStats && rowCount !== undefined ? ` · ${rowCount.toLocaleString()} rows` : ''}</span>
          </div>
          ${isActive
            ? '<span class="history-active-badge">● Active</span>'
            : `<button type="button" class="btn btn-ghost btn-sm history-switch-btn" data-upload-id="${escHtml(u.id)}">Use This</button>`
          }
        </div>
        ${hasStats ? `
        <div class="history-item-detail">
          <span class="history-detail-folder">📁 ${escHtml(topFolderName ?? '')}</span>
          <span class="history-detail-divider">·</span>
          <span class="history-detail-stat">${(folderCount ?? 0).toLocaleString()} folders</span>
          <span class="history-detail-divider">·</span>
          <span class="history-detail-stat">${(fileCount ?? 0).toLocaleString()} files</span>
          <span class="history-detail-divider">·</span>
          <span class="history-detail-stat">${formatBytes(sizeBytes ?? 0)}</span>
        </div>
        ` : ''}
      </div>
    `
  }).join('')
}

function setupHistoryButtons(container: HTMLElement): void {
  container.querySelector('#history-list')?.addEventListener('click', async (e) => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('.history-switch-btn')
    if (!btn || btn.disabled) return

    const uploadId = btn.dataset.uploadId!
    const project = getState().currentProject
    const upload = project?.projectData.uploads?.find((u) => u.id === uploadId)
    if (!project || !upload) return

    btn.disabled = true
    btn.textContent = 'Loading…'

    try {
      const { siteId } = getSpConfig()
      const tree = (await downloadDriveItem(siteId, upload.treeItemId)) as TreeNode
      const conflicts = detectMappingConflicts(tree, getState().mappings)

      const newProjectData = { ...project.projectData, activeUploadId: uploadId }
      await updateProject(project.id, { projectData: newProjectData })
      setState({ treeData: tree, currentProject: { ...project, projectData: newProjectData } })

      renderUploadPanel(container)
      if (conflicts.length > 0) showConflictWarning(container, conflicts)
    } catch (err) {
      btn.disabled = false
      btn.textContent = 'Use This'
      const status = container.querySelector('#upload-status') as HTMLElement
      status.className = 'upload-status upload-status--error'
      status.textContent = `Failed to switch: ${(err as Error).message}`
      status.style.display = 'block'
    }
  })
}

// ─── Drop zone / upload flow ──────────────────────────────────────────────────

function setupDropZone(container: HTMLElement): void {
  const dropZone = container.querySelector('#drop-zone') as HTMLElement
  const fileInput = container.querySelector('#file-input') as HTMLInputElement
  const browseBtn = container.querySelector('#btn-browse') as HTMLButtonElement

  browseBtn.addEventListener('click', () => fileInput.click())
  fileInput.addEventListener('change', () => {
    if (fileInput.files?.[0]) void handleFile(container, fileInput.files[0])
  })
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault()
    dropZone.classList.add('drop-zone--active')
  })
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drop-zone--active'))
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault()
    dropZone.classList.remove('drop-zone--active')
    const file = e.dataTransfer?.files[0]
    if (file) void handleFile(container, file)
  })
}

async function handleFile(container: HTMLElement, file: File): Promise<void> {
  const status = container.querySelector('#upload-status') as HTMLElement

  function setStatus(type: 'info' | 'success' | 'error', msg: string, spin = false): void {
    status.className = `upload-status upload-status--${type}`
    status.innerHTML = spin ? `<span class="spinner"></span>${msg}` : msg
    status.style.display = 'block'
  }

  setStatus('info', `Parsing <strong>${escHtml(file.name)}</strong> — this may take a moment for large files…`, true)

  let tree: TreeNode
  try {
    tree = await parseFileWithWorker(file)
  } catch (err) {
    setStatus('error', `Parse error: ${(err as Error).message ?? 'Unknown error'}`)
    return
  }

  const project = getState().currentProject
  if (!project) {
    // No active project — just update state (fallback, shouldn't normally happen)
    setState({ treeData: tree })
    setStatus('success', `✓ Parsed — ${formatSummary(tree)}`)
    return
  }

  setStatus('info', 'Uploading to SharePoint…', true)

  try {
    const { siteId } = getSpConfig()
    const folderId = await getOrCreateProjectFolder(siteId, project.title, project.id)

    const ts = Date.now().toString()
    const safeName = file.name.replace(/["*:<>?/\\|#%]/g, '_')

    setStatus('info', `Uploading <strong>${escHtml(file.name)}</strong>…`, true)
    const excelItemId = await uploadFileToDrive(siteId, folderId, `${ts}_${safeName}`, await file.arrayBuffer())

    setStatus('info', 'Saving report data…', true)
    const treeItemId = await uploadFileToDrive(
      siteId, folderId, `${ts}_${safeName}.tree.json`, JSON.stringify(tree)
    )

    const topNode = findTopDataNode(tree)
    const newUpload: ExcelUpload = {
      id: ts,
      fileName: file.name,
      uploadedAt: new Date().toISOString(),
      excelItemId,
      treeItemId,
      rowCount: countAllNodes(tree),
      topFolderName: topNode.name || topNode.path || 'Root',
      fileCount: tree.fileCount,
      folderCount: tree.folderCount,
      sizeBytes: tree.sizeBytes,
    }

    // Detect mapping conflicts against currently mapped folders
    const conflicts = detectMappingConflicts(tree, getState().mappings)
    let updatedMappings = getState().mappings
    if (conflicts.length > 0) {
      const errorIds = new Set(conflicts.map((c) => c.id))
      updatedMappings = updatedMappings.map((m) =>
        errorIds.has(m.id) ? { ...m, status: 'error' as const } : m
      )
    }

    // Save mappings as a separate file (keeps the list item field small)
    if (updatedMappings.length > 0) {
      setStatus('info', 'Saving mappings…', true)
      await saveMappingsFile(siteId, project.title, project.id, updatedMappings)
    }

    // ProjectData holds only lightweight metadata — no inline treeData or mappings
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { treeData: _t, mappings: _m, ...restData } = project.projectData
    const newProjectData = {
      ...restData,
      uploads: [...(project.projectData.uploads ?? []), newUpload],
      activeUploadId: ts,
    }

    await updateProject(project.id, { projectData: newProjectData })
    setState({
      treeData: tree,
      mappings: updatedMappings,
      currentProject: { ...project, projectData: newProjectData },
    })

    // Re-render to show updated history and stats, then restore status
    renderUploadPanel(container)
    const newStatus = container.querySelector('#upload-status') as HTMLElement
    newStatus.className = 'upload-status upload-status--success'
    newStatus.textContent = `✓ Saved to SharePoint — ${formatSummary(tree)}`
    newStatus.style.display = 'block'

    if (conflicts.length > 0) showConflictWarning(container, conflicts)
  } catch (err) {
    setStatus('error', `Upload failed: ${(err as Error).message}`)
  }
}

// ─── Worker ───────────────────────────────────────────────────────────────────

function parseFileWithWorker(file: File): Promise<TreeNode> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(
      new URL('../../parsers/parseWorker.ts', import.meta.url), { type: 'module' }
    )
    worker.postMessage(file)
    worker.onmessage = (e: MessageEvent<{ ok: boolean; tree?: TreeNode; error?: string }>) => {
      worker.terminate()
      if (e.data.ok && e.data.tree) resolve(e.data.tree)
      else reject(new Error(e.data.error ?? 'Parse failed'))
    }
    worker.onerror = (e) => {
      worker.terminate()
      reject(new Error(e.message ?? 'Worker error'))
    }
  })
}

// ─── Conflict detection ───────────────────────────────────────────────────────

function detectMappingConflicts(newTree: TreeNode, mappings: MigrationMapping[]): MigrationMapping[] {
  const allPaths = new Set<string>()
  function collect(n: TreeNode): void {
    allPaths.add(n.path)
    for (const child of n.children) collect(child)
  }
  collect(newTree)
  // Only flag mappings that have a site target — pending/unconfigured ones don't matter
  return mappings.filter((m) => m.targetSite && !allPaths.has(m.sourceNode.path))
}

function showConflictWarning(container: HTMLElement, conflicts: MigrationMapping[]): void {
  const warning = container.querySelector('#conflict-warning') as HTMLElement
  const msg = container.querySelector('#conflict-msg') as HTMLElement
  const list = container.querySelector('#conflict-list') as HTMLElement

  msg.textContent = `${conflicts.length} mapped folder${conflicts.length !== 1 ? 's' : ''} could not be found in the new report. They remain in your mappings with an error status — review them in the Map tab.`
  list.innerHTML = conflicts
    .map((m) => `<li><code>${escHtml(m.sourceNode.originalPath || m.sourceNode.path)}</code></li>`)
    .join('')
  warning.style.display = ''

  container.querySelector('#btn-dismiss-conflicts')?.addEventListener('click', () => {
    warning.style.display = 'none'
  }, { once: true })
}

// ─── Tree helpers ──────────────────────────────────────────────────────────────

function countAllNodes(node: TreeNode): number {
  return 1 + node.children.reduce((s, c) => s + countAllNodes(c), 0)
}

/**
 * Finds the first node in the tree that has real data (size or file count > 0).
 * Walks down single-child implicit ancestor nodes (e.g. a UNC server node like
 * \\BHFP03 that was created synthetically because the report starts at
 * \\BHFP03\NationalDataDrive\ — the server segment has no row of its own).
 * Stops as soon as it finds a node with data, or when the path branches.
 */
function findTopDataNode(node: TreeNode): TreeNode {
  if (node.sizeBytes > 0 || node.fileCount > 0) return node
  if (node.children.length === 1) return findTopDataNode(node.children[0])
  return node
}

function formatSummary(tree: TreeNode): string {
  return `${formatBytes(tree.sizeBytes)} · ${tree.fileCount.toLocaleString()} files · ${tree.folderCount.toLocaleString()} folders`
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function formatDate(d: Date): string {
  return d.toLocaleDateString(undefined, {
    year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit',
  })
}

function escHtml(s: string): string {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectUploadStyles(): void {
  if (document.getElementById('upload-styles')) return
  const style = document.createElement('style')
  style.id = 'upload-styles'
  style.textContent = `
    .upload-panel { padding: 24px; max-width: 900px; }
    .panel-section { margin-bottom: 32px; }
    .panel-section h3 { font-size: 1.05rem; font-weight: 600; margin-bottom: 8px; }
    .panel-desc { font-size: 0.88rem; color: var(--color-text-muted); margin-bottom: 16px; }

    /* Upload history list */
    .upload-history-list { border: 1px solid var(--color-border); border-radius: 6px; overflow: hidden; }
    .history-item { border-bottom: 1px solid var(--color-border); }
    .history-item:last-child { border-bottom: none; }
    .history-item--active { background: rgba(16, 124, 16, 0.06); }
    .history-item-header { display: flex; align-items: center; gap: 12px; padding: 10px 14px; }
    .history-item-icon { font-size: 1.2rem; flex-shrink: 0; }
    .history-item-info { flex: 1; min-width: 0; }
    .history-item-name { display: block; font-size: 0.875rem; font-weight: 500;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-family: 'Consolas', monospace; }
    .history-item-meta { font-size: 0.78rem; color: var(--color-text-muted); }
    .history-active-badge { font-size: 0.78rem; color: #107c10; font-weight: 600;
      white-space: nowrap; flex-shrink: 0; }
    .history-item-detail {
      display: flex; align-items: center; flex-wrap: wrap; gap: 4px;
      padding: 0 14px 10px 42px; font-size: 0.8rem; color: var(--color-text-muted);
    }
    .history-detail-folder { font-family: 'Consolas', monospace; color: var(--color-text);
      font-weight: 500; max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .history-detail-divider { color: var(--color-border); }
    .history-detail-stat { white-space: nowrap; }

    /* Conflict warning */
    .conflict-warning { margin-bottom: 24px; background: #fff4ce; border: 1px solid #f3e06b;
      border-left: 4px solid #f3c00a; border-radius: 6px; padding: 12px 16px; }
    .conflict-warning-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
    .btn-dismiss-conflicts { background: none; border: none; cursor: pointer; color: #7d5900;
      font-size: 0.9rem; padding: 2px 4px; border-radius: 3px; line-height: 1; }
    .btn-dismiss-conflicts:hover { background: rgba(0,0,0,0.08); }
    .conflict-msg { font-size: 0.85rem; color: #7d5900; margin: 0 0 6px; }
    .conflict-list { font-size: 0.82rem; color: #7d5900; padding-left: 20px; margin: 4px 0 0;
      max-height: 120px; overflow-y: auto; }
    .conflict-list li { margin: 2px 0; }
    .conflict-list code { font-family: 'Consolas', monospace; font-size: 0.78rem; }

    /* Drop zone */
    .drop-zone { border: 2px dashed var(--color-border); border-radius: 8px; padding: 40px 24px;
      text-align: center; transition: border-color 0.15s, background 0.15s; cursor: default; }
    .drop-zone--active, .drop-zone:hover { border-color: var(--color-primary); background: var(--color-primary-light); }
    .drop-zone--has-file { border-color: var(--color-success); }
    .drop-icon { font-size: 2.5rem; margin-bottom: 12px; }
    .drop-label { font-size: 0.9rem; color: var(--color-text-muted); margin-bottom: 12px; }
    .drop-hint { font-size: 0.8rem; color: var(--color-text-muted); margin-top: 8px; }

    /* Status */
    .upload-status { padding: 10px 14px; border-radius: 4px; font-size: 0.875rem; margin-top: 12px;
      display: flex; align-items: center; gap: 8px; }
    .upload-status--info { background: #deecf9; color: #005a9e; }
    .upload-status--success { background: #dff6dd; color: #107c10; }
    .upload-status--error { background: #fde7e9; color: #a4262c; }

    /* Spinner */
    .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid currentColor;
      border-top-color: transparent; border-radius: 50%; animation: spin 0.7s linear infinite; flex-shrink: 0; }
    @keyframes spin { to { transform: rotate(360deg); } }

  `
  document.head.appendChild(style)
}
