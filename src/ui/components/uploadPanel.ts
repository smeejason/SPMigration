import { parseTreeSizeFile } from '../../parsers/treeSizeParser'
import { setState, getState } from '../../state/store'
import { updateProject } from '../../graph/projectService'
import { renderTreeView } from './treeView'
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

      <div id="tree-preview" class="panel-section" style="${existingTree ? '' : 'display:none'}">
        <div class="tree-header-row">
          <h3>File System Tree</h3>
          ${existingTree ? `<div class="tree-summary">${treeSummary(existingTree)}</div>` : ''}
        </div>
        <div id="tree-container" class="tree-container"></div>
      </div>
    </div>
  `
  injectUploadStyles()

  // Show existing tree if available
  if (existingTree) {
    const treeContainer = container.querySelector('#tree-container') as HTMLElement
    renderTreeView(treeContainer, existingTree)
  }

  setupDropZone(container)
}

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

async function handleFile(container: HTMLElement, file: File): Promise<void> {
  const status = container.querySelector('#upload-status') as HTMLElement
  const treePreview = container.querySelector('#tree-preview') as HTMLElement
  const treeContainer = container.querySelector('#tree-container') as HTMLElement

  status.className = 'upload-status upload-status--info'
  status.textContent = `Parsing "${file.name}"…`
  status.style.display = 'block'

  try {
    const tree = await parseTreeSizeFile(file)
    setState({ treeData: tree })

    status.className = 'upload-status upload-status--success'
    status.textContent = `✓ Parsed successfully — ${treeSummary(tree)}`

    // Show tree
    treePreview.style.display = ''
    treePreview.querySelector('.tree-header-row')!.innerHTML = `
      <h3>File System Tree</h3>
      <div class="tree-summary">${treeSummary(tree)}</div>
    `
    renderTreeView(treeContainer, tree)

    // Persist to SharePoint
    const project = getState().currentProject
    if (project) {
      try {
        await updateProject(project.id, {
          projectData: { ...project.projectData, treeData: tree },
        })
        setState({
          currentProject: { ...project, projectData: { ...project.projectData, treeData: tree } },
        })
      } catch {
        // Non-critical — tree is in state, save failed silently
        console.warn('[Upload] Could not persist tree to SharePoint')
      }
    }
  } catch (err) {
    status.className = 'upload-status upload-status--error'
    status.textContent = `Error: ${(err as Error).message}`
  }
}

function treeSummary(tree: TreeNode): string {
  const totalFiles = countFiles(tree)
  const size = formatBytes(tree.sizeBytes)
  return `${size} · ${totalFiles.toLocaleString()} files`
}

function countFiles(node: TreeNode): number {
  if (node.children.length === 0) return node.fileCount
  return node.fileCount + node.children.reduce((s, c) => s + countFiles(c), 0)
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function injectUploadStyles(): void {
  if (document.getElementById('upload-styles')) return
  const style = document.createElement('style')
  style.id = 'upload-styles'
  style.textContent = `
    .upload-panel { padding: 24px; max-width: 900px; }
    .panel-section { margin-bottom: 32px; }
    .panel-section h3 { font-size: 1.05rem; font-weight: 600; margin-bottom: 8px; }
    .panel-desc { font-size: 0.88rem; color: var(--color-text-muted); margin-bottom: 16px; }
    .drop-zone { border: 2px dashed var(--color-border); border-radius: 8px; padding: 40px 24px;
      text-align: center; transition: border-color 0.15s, background 0.15s; cursor: default; }
    .drop-zone--active, .drop-zone:hover { border-color: var(--color-primary); background: var(--color-primary-light); }
    .drop-zone--has-file { border-color: var(--color-success); }
    .drop-icon { font-size: 2.5rem; margin-bottom: 12px; }
    .drop-label { font-size: 0.9rem; color: var(--color-text-muted); margin-bottom: 12px; }
    .drop-hint { font-size: 0.8rem; color: var(--color-text-muted); margin-top: 8px; }
    .btn-sm { padding: 6px 14px; font-size: 0.85rem; }
    .upload-status { padding: 10px 14px; border-radius: 4px; font-size: 0.875rem; margin-top: 12px; }
    .upload-status--info { background: #deecf9; color: #005a9e; }
    .upload-status--success { background: #dff6dd; color: #107c10; }
    .upload-status--error { background: #fde7e9; color: #a4262c; }
    .tree-header-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
    .tree-summary { font-size: 0.85rem; color: var(--color-text-muted); }
    .tree-container { border: 1px solid var(--color-border); border-radius: 6px; overflow: auto;
      max-height: 500px; background: white; }
  `
  document.head.appendChild(style)
}
