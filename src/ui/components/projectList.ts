import { getProjects, deleteProject, loadProjectTree, loadProjectMappings } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import { getCurrentUser } from '../../auth/authService'
import type { MigrationProject } from '../../types'

export async function renderProjectList(container: HTMLElement): Promise<void> {
  injectProjectStyles()
  container.innerHTML = `<div class="loading-spinner">Loading projects…</div>`
  await loadProjects(container)
}

async function loadProjects(container: HTMLElement): Promise<void> {
  try {
    const allProjects = await getProjects()
    setState({ projects: allProjects })
    const currentUser = getCurrentUser()
    const userEmail = currentUser?.mail?.toLowerCase() ?? ''
    const myProjects = userEmail
      ? allProjects.filter((p) =>
          p.owners.some((o) => o.email.toLowerCase() === userEmail)
        )
      : allProjects
    renderProjectCards(container, myProjects)
  } catch (err) {
    const isConfigMissing = !import.meta.env.VITE_SP_SITE_ID
    container.innerHTML = `
      <div class="projects-empty">
        ${isConfigMissing
          ? `<p class="error-text">SharePoint list not configured. Set <code>VITE_SP_SITE_ID</code> and <code>VITE_SP_LIST_ID</code> in your <code>.env.local</code> file.</p>`
          : `<p class="error-text">Could not load projects: ${(err as Error).message}</p>`
        }
      </div>
    `
  }
}

function renderProjectCards(container: HTMLElement, projects: MigrationProject[]): void {
  if (projects.length === 0) {
    container.innerHTML = `
      <div class="projects-empty">
        <p>No projects yet. Click <strong>+ New Project</strong> to get started.</p>
      </div>
    `
    return
  }

  container.innerHTML = `
    <div class="project-grid">
      ${projects.map((p) => projectCardHtml(p)).join('')}
    </div>
  `

  container.querySelectorAll('[data-project-id]').forEach((card) => {
    const id = card.getAttribute('data-project-id')!

    card.querySelector('.project-open')?.addEventListener('click', async () => {
      const project = getState().projects.find((p) => p.id === id)
      if (!project) return

      const btn = card.querySelector('.project-open') as HTMLButtonElement
      btn.disabled = true
      btn.textContent = 'Opening…'

      try {
        const [treeData, mappings] = await Promise.all([
          loadProjectTree(project),
          loadProjectMappings(project),
        ])
        setState({
          currentProject: project,
          treeData,
          mappings,
          oneDriveMappings: [],
          ui: { activeView: 'project-upload', loading: false, error: null },
        })
      } catch {
        btn.disabled = false
        btn.textContent = 'Open'
        alert('Could not load project data from SharePoint. Please try again.')
      }
    })

    card.querySelector('.project-delete')?.addEventListener('click', async (e) => {
      e.stopPropagation()
      if (!confirm(`Delete project "${getState().projects.find((p) => p.id === id)?.title}"?`)) return
      try {
        await deleteProject(id)
        const updated = getState().projects.filter((p) => p.id !== id)
        setState({ projects: updated })
        renderProjectCards(container, updated)
      } catch (err) {
        alert(`Delete failed: ${(err as Error).message}`)
      }
    })
  })
}


const SHAREPOINT_LOGO = `<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
  <circle cx="16" cy="21" r="13" fill="#036c70"/>
  <circle cx="27" cy="26" r="10" fill="#1a9ba1"/>
  <circle cx="35" cy="21" r="8" fill="#37c6d0"/>
  <rect x="8" y="30" width="32" height="10" rx="3" fill="#1a9ba1"/>
</svg>`

const ONEDRIVE_LOGO = `<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
  <path d="M29.5 18C27.6 12.2 22.2 8 15.7 8 8.7 8 3 13.7 3 20.7c0 .4 0 .8.1 1.2A11.5 11.5 0 0 0 4 44h32.5C41.7 44 46 39.7 46 34.5a9.5 9.5 0 0 0-7.2-9.2A12 12 0 0 0 29.5 18z" fill="#0078d4"/>
</svg>`

function projectCardHtml(p: MigrationProject): string {
  const stats = p.projectData
  const uploadCount = stats.uploads?.length ?? 0
  const sizeLabel = uploadCount > 0
    ? `${uploadCount} upload${uploadCount !== 1 ? 's' : ''}`
    : stats.treeData ? formatBytes(stats.treeData.sizeBytes) : '—'
  const mappingCount = stats.mappingCount ?? (stats.mappings ?? []).length
  const modified = p.lastModified ? formatDate(p.lastModified) : '—'
  const isOneDrive = (p.type ?? 'SharePoint') === 'OneDrive'
  const ownerNames = p.owners.map((o) => escHtml(o.displayName || o.email)).join(', ')

  return `
    <div class="project-card" data-project-id="${p.id}">
      <div class="project-type-logo" title="${isOneDrive ? 'OneDrive' : 'SharePoint'}">${isOneDrive ? ONEDRIVE_LOGO : SHAREPOINT_LOGO}</div>
      <div class="project-card-header">
        <div class="project-card-title-wrap">
          <h3 class="project-name">${escHtml(p.title)}</h3>
          ${p.description ? `<p class="project-desc">${escHtml(p.description)}</p>` : ''}
        </div>
      </div>
      <div class="project-stats">
        <span>📦 ${sizeLabel}</span>
        <span>🗺 ${mappingCount} mapping${mappingCount !== 1 ? 's' : ''}</span>
        <span>📅 ${modified}</span>
      </div>
      ${ownerNames ? `<div class="project-owners">👤 ${ownerNames}</div>` : ''}
      <div class="project-actions">
        <button class="btn btn-primary btn-sm project-open">Open</button>
        <button class="btn btn-ghost btn-sm project-delete" title="Delete project">🗑</button>
      </div>
    </div>
  `
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 B'
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

function formatDate(d: Date): string {
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectProjectStyles(): void {
  if (document.getElementById('project-styles')) return
  const style = document.createElement('style')
  style.id = 'project-styles'
  style.textContent = `
    .project-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
      gap: 16px;
      padding: 24px 32px;
      align-items: start;
    }
    .project-card {
      position: relative;
      background: white; border: 1px solid var(--color-border); border-radius: 8px;
      padding: 20px; transition: box-shadow 0.15s; display: flex; flex-direction: column; gap: 10px;
    }
    .project-card:hover { box-shadow: var(--shadow); }
    .project-type-logo {
      position: absolute; top: 14px; right: 16px;
      width: 52px; height: 52px; opacity: 0.92; pointer-events: none;
    }
    .project-type-logo svg { width: 100%; height: 100%; }
    .project-card-header {
      display: flex; justify-content: space-between; align-items: flex-start; gap: 12px;
      padding-right: 64px;
    }
    .project-card-title-wrap { flex: 1; min-width: 0; }
    .project-card-badges { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; }
    .project-name { font-size: 1.05rem; font-weight: 600; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .project-desc { font-size: 0.85rem; color: var(--color-text-muted); }
    .project-stats { display: flex; gap: 12px; font-size: 0.82rem; color: var(--color-text-muted); flex-wrap: wrap; }
    .project-owners { font-size: 0.82rem; color: var(--color-text-muted); }
    .project-actions { display: flex; gap: 8px; margin-top: auto; }
    .status-badge { padding: 3px 10px; border-radius: 12px; font-size: 0.78rem; font-weight: 600; white-space: nowrap; }
    .status-planning { background: #deecf9; color: #005a9e; }
    .status-in-progress { background: #fff4ce; color: #7d5900; }
    .status-completed { background: #dff6dd; color: #107c10; }
    .status-on-hold { background: #f3f2f1; color: #605e5c; }
    .projects-empty { padding: 48px 32px; color: var(--color-text-muted); }
    .error-text { color: var(--color-danger); }
    .loading-spinner { padding: 48px 32px; color: var(--color-text-muted); }
  `
  document.head.appendChild(style)
}
