import { createProject, updateProject } from '../../graph/projectService'
import { setState, getState } from '../../state/store'
import type { MigrationProject, ProjectStatus, SharePointUser } from '../../types'

export function renderProjectForm(
  container: HTMLElement,
  project: MigrationProject | null,
  onSave: (project: MigrationProject) => void,
  onCancel: () => void
): void {
  const isEdit = project !== null
  const statuses: ProjectStatus[] = ['Planning', 'In Progress', 'Completed', 'On Hold']

  container.innerHTML = `
    <div class="form-overlay">
      <div class="form-dialog">
        <div class="form-dialog-header">
          <h2>${isEdit ? 'Edit Project' : 'New Migration Project'}</h2>
          <button id="btn-cancel-form" class="btn-icon" title="Close">✕</button>
        </div>
        <form id="project-form" class="project-form">
          <div class="form-group">
            <label for="f-title">Project Name <span class="required">*</span></label>
            <input id="f-title" type="text" class="form-input" required
              value="${escHtml(project?.title ?? '')}" placeholder="e.g. Contoso File Share Migration" />
          </div>
          <div class="form-group">
            <label for="f-desc">Description</label>
            <textarea id="f-desc" class="form-input" rows="3"
              placeholder="Optional notes about this migration project">${escHtml(project?.description ?? '')}</textarea>
          </div>
          <div class="form-group">
            <label for="f-status">Status</label>
            <select id="f-status" class="form-input">
              ${statuses.map((s) => `<option value="${s}" ${(project?.status ?? 'Planning') === s ? 'selected' : ''}>${s}</option>`).join('')}
            </select>
          </div>
          <div id="form-error" class="form-error" style="display:none"></div>
          <div class="form-actions">
            <button type="button" id="btn-cancel" class="btn btn-ghost">Cancel</button>
            <button type="submit" id="btn-save" class="btn btn-primary">
              ${isEdit ? 'Save Changes' : 'Create Project'}
            </button>
          </div>
        </form>
      </div>
    </div>
  `
  injectFormStyles()

  const closeForm = (): void => {
    container.innerHTML = ''
    onCancel()
  }

  container.querySelector('#btn-cancel-form')!.addEventListener('click', closeForm)
  container.querySelector('#btn-cancel')!.addEventListener('click', closeForm)

  container.querySelector('#project-form')!.addEventListener('submit', async (e) => {
    e.preventDefault()
    const title = (container.querySelector('#f-title') as HTMLInputElement).value.trim()
    const description = (container.querySelector('#f-desc') as HTMLTextAreaElement).value.trim()
    const status = (container.querySelector('#f-status') as HTMLSelectElement).value as ProjectStatus
    const errorEl = container.querySelector('#form-error') as HTMLElement
    const saveBtn = container.querySelector('#btn-save') as HTMLButtonElement

    if (!title) {
      errorEl.textContent = 'Project name is required.'
      errorEl.style.display = 'block'
      return
    }

    saveBtn.disabled = true
    saveBtn.textContent = isEdit ? 'Saving…' : 'Creating…'
    errorEl.style.display = 'none'

    try {
      let saved: MigrationProject
      if (isEdit && project) {
        await updateProject(project.id, { title, description, status })
        saved = { ...project, title, description, status }
        // Update in store
        setState({
          projects: getState().projects.map((p) => (p.id === saved.id ? saved : p)),
        })
      } else {
        const authUser = getState().auth.user
        const owner: SharePointUser | undefined = authUser
          ? { id: authUser.id, displayName: authUser.displayName, email: authUser.mail }
          : undefined
        saved = await createProject({ title, description, status, owner })
        setState({ projects: [...getState().projects, saved] })
      }
      container.innerHTML = ''
      onSave(saved)
    } catch (err) {
      errorEl.textContent = `Failed to save: ${(err as Error).message}`
      errorEl.style.display = 'block'
      saveBtn.disabled = false
      saveBtn.textContent = isEdit ? 'Save Changes' : 'Create Project'
    }
  })
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectFormStyles(): void {
  if (document.getElementById('form-styles')) return
  const style = document.createElement('style')
  style.id = 'form-styles'
  style.textContent = `
    .form-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: flex;
      align-items: center; justify-content: center; z-index: 100; padding: 24px; }
    .form-dialog { background: white; border-radius: 8px; width: 100%; max-width: 480px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.25); }
    .form-dialog-header { display: flex; justify-content: space-between; align-items: center;
      padding: 20px 24px 0; margin-bottom: 20px; }
    .form-dialog-header h2 { font-size: 1.15rem; font-weight: 600; }
    .btn-icon { background: none; border: none; cursor: pointer; font-size: 1.1rem; color: var(--color-text-muted);
      padding: 4px; border-radius: 4px; }
    .btn-icon:hover { background: var(--color-surface-alt); }
    .project-form { padding: 0 24px 24px; }
    .form-group { margin-bottom: 16px; }
    .form-group label { display: block; font-size: 0.85rem; font-weight: 600; margin-bottom: 6px; }
    .required { color: var(--color-danger); }
    .form-input { width: 100%; padding: 8px 12px; border: 1px solid var(--color-border);
      border-radius: 4px; font-family: inherit; font-size: 0.9rem; outline: none; }
    .form-input:focus { border-color: var(--color-primary); box-shadow: 0 0 0 2px var(--color-primary-light); }
    textarea.form-input { resize: vertical; }
    .form-error { padding: 10px 12px; background: #fde7e9; color: #a4262c;
      border-radius: 4px; font-size: 0.85rem; margin-bottom: 12px; }
    .form-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 24px; }
  `
  document.head.appendChild(style)
}
