import type { TreeNode } from '../../types'

// ─── Public API ───────────────────────────────────────────────────────────────

export function renderTreeView(container: HTMLElement, root: TreeNode): void {
  injectTreeStyles()
  container.innerHTML = ''

  // Column header bar
  const header = document.createElement('div')
  header.className = 'tree-col-header'
  header.innerHTML = `
    <span class="tree-col-name">Name</span>
    <span class="tree-col-date">Last Modified</span>
    <span class="tree-col-size">Size</span>
    <span class="tree-col-count">Items</span>
  `
  container.appendChild(header)

  const ul = document.createElement('ul')
  ul.className = 'tree-list tree-root'
  ul.appendChild(createNodeEl(root, true))
  container.appendChild(ul)

  // Auto-expand root
  const rootRow = ul.querySelector<HTMLElement>('.tree-row[data-has-children="true"]')
  rootRow?.click()
}

// ─── Node element factory (lazy children) ────────────────────────────────────

function createNodeEl(node: TreeNode, isRoot = false): HTMLLIElement {
  const li = document.createElement('li')
  li.className = `tree-node${isRoot ? ' tree-node--root' : ''}`

  const hasChildren = node.children.length > 0
  // All TreeSize rows are directories. Only *-wildcard entries (e.g. "*.*") are loose-file indicators.
  const isFolder = !node.name.includes('*')

  const row = document.createElement('div')
  row.className = 'tree-row'
  row.dataset.hasChildren = String(hasChildren)
  row.dataset.loaded = 'false'
  if (node.path) row.title = node.path

  const childCount = node.fileCount + node.folderCount
  const sizeLabel = formatBytes(node.sizeBytes)
  const dateLabel = node.lastModified ? formatDate(node.lastModified) : '—'
  const countLabel = childCount > 0 ? childCount.toLocaleString() : '—'

  row.innerHTML = `
    <span class="tree-col-name tree-name-cell">
      <span class="tree-toggle">${hasChildren ? '▶' : '\u00a0'}</span>
      <span class="tree-icon">${isFolder ? '📁' : '📄'}</span>
      <span class="tree-name">${escHtml(String(node.name || node.path || '(unnamed)'))}</span>
    </span>
    <span class="tree-col-date tree-date">${escHtml(dateLabel)}</span>
    <span class="tree-col-size tree-size">${escHtml(sizeLabel)}</span>
    <span class="tree-col-count tree-count">${escHtml(countLabel)}</span>
  `

  if (hasChildren) {
    row.addEventListener('click', () => toggleNode(li, node, row))
  }

  li.appendChild(row)
  return li
}

function toggleNode(li: HTMLLIElement, node: TreeNode, row: HTMLElement): void {
  const isOpen = li.classList.contains('tree-node--open')
  const toggle = row.querySelector<HTMLElement>('.tree-toggle')!

  if (isOpen) {
    // Collapse — just hide, keep DOM intact
    const childUl = li.querySelector<HTMLElement>(':scope > .tree-children')
    if (childUl) childUl.style.display = 'none'
    li.classList.remove('tree-node--open')
    row.classList.remove('tree-row--open')
    toggle.textContent = '▶'
  } else {
    // Expand — lazy-render children on first open
    if (row.dataset.loaded === 'false') {
      const childUl = document.createElement('ul')
      childUl.className = 'tree-list tree-children'
      for (const child of node.children) {
        childUl.appendChild(createNodeEl(child))
      }
      li.appendChild(childUl)
      row.dataset.loaded = 'true'
    } else {
      const childUl = li.querySelector<HTMLElement>(':scope > .tree-children')
      if (childUl) childUl.style.display = ''
    }
    li.classList.add('tree-node--open')
    row.classList.add('tree-row--open')
    toggle.textContent = '▼'
  }
}

// ─── Formatters ───────────────────────────────────────────────────────────────

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '—'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1)
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function formatDate(d: Date | string): string {
  const date = d instanceof Date ? d : new Date(d)
  if (isNaN(date.getTime())) return '—'
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' })
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectTreeStyles(): void {
  if (document.getElementById('tree-styles')) return
  const style = document.createElement('style')
  style.id = 'tree-styles'
  style.textContent = `
    /* Column layout — shared between header and rows */
    .tree-col-header,
    .tree-row {
      display: grid;
      grid-template-columns: 1fr 140px 90px 70px;
      align-items: center;
      gap: 0;
    }

    /* Header bar */
    .tree-col-header {
      padding: 6px 12px;
      border-bottom: 1px solid var(--color-border);
      background: var(--color-surface, #f5f5f5);
      font-size: 0.75rem;
      font-weight: 600;
      color: var(--color-text-muted);
      text-transform: uppercase;
      letter-spacing: 0.04em;
      position: sticky;
      top: 0;
      z-index: 1;
    }
    .tree-col-date, .tree-col-size, .tree-col-count { text-align: right; padding-right: 12px; }
    .tree-col-name { text-align: left; }

    /* Tree structure */
    .tree-list { list-style: none; padding: 0; margin: 0; }
    .tree-root { }
    .tree-children { padding-left: 20px; border-left: 1px solid var(--color-border); margin-left: 18px; }
    .tree-node { margin: 0; }

    /* Row */
    .tree-row {
      padding: 5px 12px;
      border-radius: 4px;
      cursor: default;
      user-select: none;
      transition: background 0.1s;
      min-height: 30px;
    }
    .tree-row[data-has-children="true"] { cursor: pointer; }
    .tree-row:hover { background: var(--color-primary-light, #f0f4ff); }
    .tree-row--open { background: var(--color-primary-light, #f0f4ff); }

    /* Name cell */
    .tree-name-cell {
      display: flex;
      align-items: center;
      gap: 5px;
      min-width: 0;
    }
    .tree-toggle {
      width: 14px;
      font-size: 0.6rem;
      color: var(--color-text-muted);
      flex-shrink: 0;
    }
    .tree-row--open > .tree-name-cell > .tree-toggle { color: var(--color-primary, #0078d4); }
    .tree-icon { flex-shrink: 0; font-size: 0.95rem; }
    .tree-name {
      font-size: 0.875rem;
      font-family: 'Consolas', 'Cascadia Code', monospace;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .tree-node--root > .tree-row > .tree-name-cell > .tree-name { font-weight: 600; }

    /* Meta columns */
    .tree-date {
      font-size: 0.78rem;
      color: var(--color-text-muted);
      text-align: right;
      padding-right: 12px;
    }
    .tree-size {
      font-size: 0.78rem;
      color: var(--color-primary, #0078d4);
      font-weight: 500;
      text-align: right;
      padding-right: 12px;
    }
    .tree-count {
      font-size: 0.78rem;
      color: var(--color-text-muted);
      text-align: right;
      padding-right: 12px;
    }
  `
  document.head.appendChild(style)
}
