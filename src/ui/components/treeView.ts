import type { TreeNode } from '../../types'

export function renderTreeView(container: HTMLElement, root: TreeNode): void {
  container.innerHTML = `<ul class="tree-list tree-root">${renderNode(root, true)}</ul>`
  injectTreeStyles()
  attachToggleHandlers(container)
}

function renderNode(node: TreeNode, isRoot = false): string {
  const hasChildren = node.children.length > 0
  const icon = hasChildren ? '📁' : '📄'
  const sizeLabel = formatBytes(node.sizeBytes)
  const filesLabel = node.fileCount > 0 ? `${node.fileCount.toLocaleString()} files` : ''

  return `
    <li class="tree-node ${isRoot ? 'tree-node--root' : ''}">
      <div class="tree-row" data-has-children="${hasChildren}">
        <span class="tree-toggle">${hasChildren ? '▶' : ' '}</span>
        <span class="tree-icon">${icon}</span>
        <span class="tree-name">${escHtml(node.name || node.path)}</span>
        <span class="tree-meta">
          ${sizeLabel ? `<span class="tree-size">${sizeLabel}</span>` : ''}
          ${filesLabel ? `<span class="tree-files">${filesLabel}</span>` : ''}
        </span>
      </div>
      ${hasChildren
        ? `<ul class="tree-list tree-children" style="display:none">${node.children.map((c) => renderNode(c)).join('')}</ul>`
        : ''
      }
    </li>
  `
}

function attachToggleHandlers(container: HTMLElement): void {
  container.querySelectorAll('.tree-row[data-has-children="true"]').forEach((row) => {
    row.addEventListener('click', () => {
      const li = row.closest('.tree-node')!
      const children = li.querySelector('.tree-children') as HTMLElement | null
      const toggle = row.querySelector('.tree-toggle') as HTMLElement
      if (!children) return
      const isOpen = children.style.display !== 'none'
      children.style.display = isOpen ? 'none' : ''
      toggle.textContent = isOpen ? '▶' : '▼'
      row.classList.toggle('tree-row--open', !isOpen)
    })
  })

  // Auto-expand root
  const rootRow = container.querySelector('.tree-node--root > .tree-row') as HTMLElement | null
  rootRow?.click()
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes === 0) return ''
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function injectTreeStyles(): void {
  if (document.getElementById('tree-styles')) return
  const style = document.createElement('style')
  style.id = 'tree-styles'
  style.textContent = `
    .tree-list { list-style: none; padding: 0; margin: 0; }
    .tree-root { padding: 8px; }
    .tree-children { padding-left: 20px; border-left: 1px solid var(--color-border); margin-left: 12px; }
    .tree-node { margin: 1px 0; }
    .tree-row { display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-radius: 4px;
      cursor: pointer; user-select: none; transition: background 0.1s; }
    .tree-row:hover { background: var(--color-primary-light); }
    .tree-row--open > .tree-toggle { color: var(--color-primary); }
    .tree-toggle { width: 14px; font-size: 0.65rem; color: var(--color-text-muted); flex-shrink: 0; }
    .tree-icon { flex-shrink: 0; }
    .tree-name { flex: 1; font-size: 0.875rem; font-family: 'Consolas', monospace; white-space: nowrap;
      overflow: hidden; text-overflow: ellipsis; }
    .tree-meta { display: flex; gap: 12px; flex-shrink: 0; }
    .tree-size { font-size: 0.78rem; color: var(--color-primary); font-weight: 500; }
    .tree-files { font-size: 0.78rem; color: var(--color-text-muted); }
    .tree-node--root > .tree-row { font-weight: 600; }
  `
  document.head.appendChild(style)
}
