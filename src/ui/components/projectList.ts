import { getProjects, deleteProject, loadProjectTree, loadProjectMappings } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { AppUser, MigrationProject } from '../../types'

export async function renderProjectList(container: HTMLElement): Promise<void> {
  injectProjectStyles()
  container.innerHTML = `<div class="loading-spinner">Loading projects…</div>`
  await loadProjects(container)
}

async function loadProjects(container: HTMLElement): Promise<void> {
  try {
    const allProjects = await getProjects()
    const currentUser = getState().auth.user
    const projects = filterByOwnership(allProjects, currentUser)
    setState({ projects })
    renderProjectCards(container, projects)
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

function filterByOwnership(projects: MigrationProject[], user: AppUser | null): MigrationProject[] {
  if (!user) return []
  return projects.filter((p) => {
    if (p.owners.length === 0) return true
    return p.owners.some((o) => o.email === user.mail || o.id === user.id)
  })
}

function projectCardHtml(p: MigrationProject): string {
  const stats = p.projectData
  const uploadCount = stats.uploads?.length ?? 0
  const sizeLabel = uploadCount > 0
    ? `${uploadCount} upload${uploadCount !== 1 ? 's' : ''}`
    : stats.treeData ? formatBytes(stats.treeData.sizeBytes) : '—'
  const mappingCount = stats.mappingCount ?? (stats.mappings ?? []).length
  const modified = p.lastModified ? formatDate(p.lastModified) : '—'
  const statusClass = p.status.toLowerCase().replace(' ', '-')
  const typeClass = (p.type ?? 'SharePoint') === 'OneDrive' ? 'type-onedrive' : 'type-sharepoint'
  const ownerNames = p.owners.map((o) => escHtml(o.displayName || o.email)).join(', ')

  return `
    <div class="project-card" data-project-id="${p.id}">
      <div class="project-card-header">
        <div class="project-card-title-wrap">
          <h3 class="project-name">${escHtml(p.title)}</h3>
          ${p.description ? `<p class="project-desc">${escHtml(p.description)}</p>` : ''}
        </div>
        <div class="project-card-badges">
          <span class="type-badge ${typeClass}">${escHtml(p.type ?? 'SharePoint')}</span>
          <span class="status-badge status-${statusClass}">${escHtml(p.status)}</span>
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
      background: white; border: 1px solid var(--color-border); border-radius: 8px;
      padding: 20px; transition: box-shadow 0.15s; display: flex; flex-direction: column; gap: 10px;
    }
    .project-card:hover { box-shadow: var(--shadow); }
    .project-card-header {
      display: flex; justify-content: space-between; align-items: flex-start; gap: 12px;
    }
    .project-card-title-wrap { flex: 1; min-width: 0; }
    .project-card-badges { display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; }
    .type-badge { padding: 2px 8px; border-radius: 12px; font-size: 0.72rem; font-weight: 600; }
    .type-sharepoint { background: #e8f0fe; color: #1a56db; }
    .type-onedrive { background: #e8f4fd; color: #0078d4; }
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
