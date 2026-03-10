import { getProjects, deleteProject, loadProjectTree, loadProjectMappings } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { AppUser, MigrationProject } from '../../types'

export async function renderProjectList(container: HTMLElement): Promise<void> {
  container.innerHTML = `
    <div class="projects-page">
      <div class="projects-header">
        <h2>Your Migration Projects</h2>
        <button id="btn-new-project" class="btn btn-primary">+ New Project</button>
      </div>
      <div id="projects-body">
        <div class="loading-spinner">Loading projects…</div>
      </div>
    </div>
  `
  injectProjectStyles()

  container.querySelector('#btn-new-project')!.addEventListener('click', () => {
    setState({ currentProject: null, ui: { activeView: 'projects', loading: false, error: null } })
    // Signal to app shell to open new project form
    container.dispatchEvent(new CustomEvent('new-project', { bubbles: true }))
  })

  await loadProjects(container)
}

async function loadProjects(container: HTMLElement): Promise<void> {
  const body = container.querySelector('#projects-body') as HTMLElement
  try {
    const allProjects = await getProjects()
    // Only show projects where the current user is listed as an owner.
    // Projects with no owners (legacy data) are shown to everyone.
    const currentUser = getState().auth.user
    const projects = filterByOwnership(allProjects, currentUser)
    setState({ projects })
    renderProjectCards(body, projects, container)
  } catch (err) {
    const isConfigMissing = !import.meta.env.VITE_SP_SITE_ID
    body.innerHTML = `
      <div class="projects-empty">
        ${isConfigMissing
          ? `<p class="error-text">SharePoint list not configured. Set <code>VITE_SP_SITE_ID</code> and <code>VITE_SP_LIST_ID</code> in your <code>.env.local</code> file.</p>`
          : `<p class="error-text">Could not load projects: ${(err as Error).message}</p>`
        }
      </div>
    `
  }
}

function renderProjectCards(
  body: HTMLElement,
  projects: MigrationProject[],
  container: HTMLElement
): void {
  if (projects.length === 0) {
    body.innerHTML = `
      <div class="projects-empty">
        <p>No projects yet. Create your first migration project to get started.</p>
      </div>
    `
    return
  }

  body.innerHTML = `
    <div class="project-grid">
      ${projects.map((p) => projectCardHtml(p)).join('')}
    </div>
  `

  body.querySelectorAll('[data-project-id]').forEach((card) => {
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
        const projects = getState().projects.filter((p) => p.id !== id)
        setState({ projects })
        renderProjectCards(body, projects, container)
      } catch (err) {
        alert(`Delete failed: ${(err as Error).message}`)
      }
    })
  })
}

function filterByOwnership(projects: MigrationProject[], user: AppUser | null): MigrationProject[] {
  if (!user) return []
  return projects.filter((p) => {
    // Legacy projects with no owners are visible to everyone
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
  const mappingCount = (stats.mappings ?? []).length
  const modified = p.lastModified ? formatDate(p.lastModified) : '—'
  const statusClass = p.status.toLowerCase().replace(' ', '-')
  const ownerNames = p.owners.map((o) => escHtml(o.displayName || o.email)).join(', ')

  return `
    <div class="project-card" data-project-id="${p.id}">
      <div class="project-card-header">
        <div>
          <h3 class="project-name">${escHtml(p.title)}</h3>
          ${p.description ? `<p class="project-desc">${escHtml(p.description)}</p>` : ''}
        </div>
        <span class="status-badge status-${statusClass}">${escHtml(p.status)}</span>
      </div>
      <div class="project-stats">
        <span>📦 ${sizeLabel}</span>
        <span>🗺 ${mappingCount} mapping${mappingCount !== 1 ? 's' : ''}</span>
        <span>📅 ${modified}</span>
      </div>
      ${ownerNames ? `<div class="project-owners">👤 ${ownerNames}</div>` : ''}
      <div class="project-actions">
        <button class="btn btn-primary project-open">Open</button>
        <button class="btn btn-ghost project-delete" title="Delete project">🗑</button>
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
    .projects-page { padding: 32px; max-width: 900px; margin: 0 auto; }
    .projects-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px; }
    .projects-header h2 { font-size: 1.4rem; font-weight: 600; }
    .project-grid { display: grid; gap: 16px; }
    .project-card { background: white; border: 1px solid var(--color-border); border-radius: 8px;
      padding: 20px; transition: box-shadow 0.15s; }
    .project-card:hover { box-shadow: var(--shadow); }
    .project-card-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; gap: 12px; }
    .project-name { font-size: 1.05rem; font-weight: 600; margin-bottom: 4px; }
    .project-desc { font-size: 0.85rem; color: var(--color-text-muted); }
    .project-stats { display: flex; gap: 16px; font-size: 0.85rem; color: var(--color-text-muted); margin-bottom: 8px; flex-wrap: wrap; }
    .project-owners { font-size: 0.82rem; color: var(--color-text-muted); margin-bottom: 16px; }
    .project-actions { display: flex; gap: 8px; }
    .status-badge { padding: 3px 10px; border-radius: 12px; font-size: 0.78rem; font-weight: 600; white-space: nowrap; }
    .status-planning { background: #deecf9; color: #005a9e; }
    .status-in-progress { background: #fff4ce; color: #7d5900; }
    .status-completed { background: #dff6dd; color: #107c10; }
    .status-on-hold { background: #f3f2f1; color: #605e5c; }
    .projects-empty { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .error-text { color: var(--color-danger); }
    .loading-spinner { padding: 48px; text-align: center; color: var(--color-text-muted); }
    .btn-ghost { background: transparent; border: 1px solid var(--color-border); color: var(--color-text-muted); }
    .btn-ghost:hover { background: var(--color-surface-alt); }
  `
  document.head.appendChild(style)
}
