import { setState, getState } from '../../state/store'
import { updateProject } from '../../graph/projectService'
import type { TreeNode } from '../../types'

export function renderUploadPanel(container: HTMLElement): void {
  const existingTree = getState().treeData

  container.innerHTML = `
    <div class="upload-panel">
      <div class="panel-section">
        <h3>Upload TreeSize Report</h3>
        <p class="panel-desc">Export your file server structure from TreeSize Pro or Free as <strong>.csv</strong> or <strong>.xlsx</strong>, then upload it here.</p>

        <div id="drop-zone" class="drop-zone ${existingTree ? 'drop-zone--has-file' : ''}">
          <input type="file" id="file-input" accept=".csv,.xlsx,.xls" style="display:none" />
          <div class="drop-zone-content">
            <div class="drop-icon">📂</div>
            <p class="drop-label">${existingTree ? 'Replace file — drag & drop or' : 'Drag & drop your TreeSize export here, or'}</p>
            <button type="button" id="btn-browse" class="btn btn-primary btn-sm">Browse files</button>
            <p class="drop-hint">Accepts .csv and .xlsx</p>
          </div>
        </div>

        <div id="upload-status" class="upload-status" style="display:none"></div>
      </div>

      <div id="stats-section" class="panel-section" style="${existingTree ? '' : 'display:none'}">
        <h3>File System Summary</h3>
        <div id="stats-cards" class="stats-grid">
          ${existingTree ? renderStatsCards(existingTree) : ''}
        </div>
      </div>
    </div>
  `
  injectUploadStyles()
  setupDropZone(container)
}

// ─── Stats ────────────────────────────────────────────────────────────────────

function renderStatsCards(tree: TreeNode): string {
  const { totalFiles, totalFolders, totalBytes } = computeStats(tree)
  return `
    <div class="stat-card">
      <div class="stat-icon">📄</div>
      <div class="stat-value">${totalFiles.toLocaleString()}</div>
      <div class="stat-label">Total Files</div>
    </div>
    <div class="stat-card">
      <div class="stat-icon">📁</div>
      <div class="stat-value">${totalFolders.toLocaleString()}</div>
      <div class="stat-label">Total Folders</div>
    </div>
    <div class="stat-card">
      <div class="stat-icon">💾</div>
      <div class="stat-value">${formatBytes(totalBytes)}</div>
      <div class="stat-label">Space Used</div>
    </div>
  `
}

function computeStats(root: TreeNode): { totalFiles: number; totalFolders: number; totalBytes: number } {
  // Use the highest node's values directly — TreeSize stores cumulative totals on each row
  return {
    totalFiles: root.fileCount,
    totalFolders: root.folderCount,
    totalBytes: root.sizeBytes,
  }
}

// ─── Drop zone / file handling ────────────────────────────────────────────────

function setupDropZone(container: HTMLElement): void {
  const dropZone = container.querySelector('#drop-zone') as HTMLElement
  const fileInput = container.querySelector('#file-input') as HTMLInputElement
  const browseBtn = container.querySelector('#btn-browse') as HTMLButtonElement

  browseBtn.addEventListener('click', () => fileInput.click())
  fileInput.addEventListener('change', () => {
    if (fileInput.files?.[0]) handleFile(container, fileInput.files[0])
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
    if (file) handleFile(container, file)
  })
}

function handleFile(container: HTMLElement, file: File): void {
  const status = container.querySelector('#upload-status') as HTMLElement
  const statsSection = container.querySelector('#stats-section') as HTMLElement
  const statsCards = container.querySelector('#stats-cards') as HTMLElement

  const fileName = String(file?.name ?? 'file')

  status.className = 'upload-status upload-status--info'
  status.innerHTML = `
    <span class="spinner"></span>
    Parsing <strong>${escHtml(fileName)}</strong> — this may take a moment for large files…
  `
  status.style.display = 'block'
  statsSection.style.display = 'none'

  // Run parsing in a Web Worker so the UI stays responsive
  const worker = new Worker(new URL('../../parsers/parseWorker.ts', import.meta.url), { type: 'module' })
  worker.postMessage(file)

  worker.onmessage = (e: MessageEvent<{ ok: boolean; tree?: TreeNode; error?: string }>) => {
    worker.terminate()
    if (!e.data.ok || !e.data.tree) {
      status.className = 'upload-status upload-status--error'
      status.textContent = `Error: ${e.data.error ?? 'Unknown error'}`
      return
    }

    const tree = e.data.tree
    setState({ treeData: tree })

    const stats = computeStats(tree)
    status.className = 'upload-status upload-status--success'
    status.textContent = `✓ Parsed successfully — ${formatBytes(stats.totalBytes)} · ${stats.totalFiles.toLocaleString()} files · ${stats.totalFolders.toLocaleString()} folders`

    statsCards.innerHTML = renderStatsCards(tree)
    statsSection.style.display = ''

    // Persist to SharePoint (non-critical)
    const project = getState().currentProject
    if (project) {
      updateProject(project.id, {
        projectData: { ...project.projectData, treeData: tree },
      })
        .then(() => {
          setState({
            currentProject: { ...project, projectData: { ...project.projectData, treeData: tree } },
          })
        })
        .catch(() => console.warn('[Upload] Could not persist tree to SharePoint'))
    }
  }

  worker.onerror = (e) => {
    worker.terminate()
    status.className = 'upload-status upload-status--error'
    status.textContent = `Error: ${e.message ?? 'Parse worker failed'}`
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
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

    /* Drop zone */
    .drop-zone { border: 2px dashed var(--color-border); border-radius: 8px; padding: 40px 24px;
      text-align: center; transition: border-color 0.15s, background 0.15s; cursor: default; }
    .drop-zone--active, .drop-zone:hover { border-color: var(--color-primary); background: var(--color-primary-light); }
    .drop-zone--has-file { border-color: var(--color-success); }
    .drop-icon { font-size: 2.5rem; margin-bottom: 12px; }
    .drop-label { font-size: 0.9rem; color: var(--color-text-muted); margin-bottom: 12px; }
    .drop-hint { font-size: 0.8rem; color: var(--color-text-muted); margin-top: 8px; }
    .btn-sm { padding: 6px 14px; font-size: 0.85rem; }

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

    /* Stats cards */
    .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 12px; }
    .stat-card { background: var(--color-surface, #faf9f8); border: 1px solid var(--color-border);
      border-radius: 8px; padding: 20px 24px; display: flex; flex-direction: column; align-items: center;
      gap: 6px; text-align: center; }
    .stat-icon { font-size: 2rem; }
    .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--color-primary, #0078d4);
      font-variant-numeric: tabular-nums; }
    .stat-label { font-size: 0.8rem; color: var(--color-text-muted); text-transform: uppercase;
      letter-spacing: 0.05em; font-weight: 500; }
  `
  document.head.appendChild(style)
}
