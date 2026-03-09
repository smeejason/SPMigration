import { createProject, updateProject } from '../../graph/projectService'
import { searchUsers } from '../../graph/graphClient'
import { setState, getState } from '../../state/store'
import type { MigrationProject, ProjectStatus, SharePointUser, AppUser } from '../../types'

export function renderProjectForm(
  container: HTMLElement,
  project: MigrationProject | null,
  onSave: (project: MigrationProject) => void,
  onCancel: () => void
): void {
  const isEdit = project !== null
  const statuses: ProjectStatus[] = ['Planning', 'In Progress', 'Completed', 'On Hold']

  // Seed owners: on create pre-fill with current user; on edit use existing owners
  const currentAuthUser = getState().auth.user
  const initialOwners: SharePointUser[] = isEdit
    ? project.owners
    : currentAuthUser
      ? [authUserToOwner(currentAuthUser)]
      : []

  // Mutable working list of selected owners
  let selectedOwners: SharePointUser[] = [...initialOwners]

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
          <div class="form-group">
            <label>Owners <span class="required">*</span></label>
            <div id="owners-chips" class="owners-chips"></div>
            <div class="owners-search-wrap">
              <input id="f-owners-search" type="text" class="form-input" placeholder="Search people…" autocomplete="off" />
              <ul id="owners-dropdown" class="owners-dropdown" style="display:none"></ul>
            </div>
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

  // ── Render owner chips ────────────────────────────────────────────────────

  function renderChips(): void {
    const chipsEl = container.querySelector('#owners-chips') as HTMLElement
    chipsEl.innerHTML = selectedOwners.map((o) => `
      <span class="owner-chip" data-id="${escHtml(o.id)}">
        ${escHtml(o.displayName || o.email)}
        <button type="button" class="chip-remove" data-id="${escHtml(o.id)}" title="Remove">✕</button>
      </span>
    `).join('')

    chipsEl.querySelectorAll('.chip-remove').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = (btn as HTMLElement).dataset.id!
        selectedOwners = selectedOwners.filter((o) => o.id !== id)
        renderChips()
      })
    })
  }

  renderChips()

  // ── People search / dropdown ──────────────────────────────────────────────

  const searchInput = container.querySelector('#f-owners-search') as HTMLInputElement
  const dropdown = container.querySelector('#owners-dropdown') as HTMLUListElement
  let searchTimer: ReturnType<typeof setTimeout> | null = null

  searchInput.addEventListener('input', () => {
    if (searchTimer) clearTimeout(searchTimer)
    const q = searchInput.value.trim()
    if (q.length < 2) { dropdown.style.display = 'none'; return }

    searchTimer = setTimeout(async () => {
      try {
        const users = await searchUsers(q)
        const available = users.filter((u) => !selectedOwners.some((o) => o.id === u.id))
        if (available.length === 0) { dropdown.style.display = 'none'; return }

        dropdown.innerHTML = available.map((u) => `
          <li class="owners-option" data-id="${escHtml(u.id)}"
              data-name="${escHtml(u.displayName)}" data-email="${escHtml(u.mail)}">
            <span class="opt-name">${escHtml(u.displayName)}</span>
            <span class="opt-email">${escHtml(u.mail)}</span>
          </li>
        `).join('')
        dropdown.style.display = 'block'

        dropdown.querySelectorAll('.owners-option').forEach((li) => {
          li.addEventListener('click', () => {
            const el = li as HTMLElement
            selectedOwners.push({ id: el.dataset.id!, displayName: el.dataset.name!, email: el.dataset.email! })
            renderChips()
            searchInput.value = ''
            dropdown.style.display = 'none'
          })
        })
      } catch {
        dropdown.style.display = 'none'
      }
    }, 300)
  })

  // Close dropdown when clicking outside
  document.addEventListener('click', (e) => {
    if (!container.contains(e.target as Node)) dropdown.style.display = 'none'
  }, { once: false, capture: true })

  // ── Form cancel / submit ──────────────────────────────────────────────────

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
    if (selectedOwners.length === 0) {
      errorEl.textContent = 'At least one owner is required.'
      errorEl.style.display = 'block'
      return
    }

    saveBtn.disabled = true
    saveBtn.textContent = isEdit ? 'Saving…' : 'Creating…'
    errorEl.style.display = 'none'

    try {
      let saved: MigrationProject
      if (isEdit && project) {
        const updatedProjectData = { ...project.projectData, owners: selectedOwners }
        await updateProject(project.id, { title, description, status, owners: selectedOwners, projectData: updatedProjectData })
        saved = { ...project, title, description, status, owners: selectedOwners, projectData: updatedProjectData }
        setState({
          projects: getState().projects.map((p) => (p.id === saved.id ? saved : p)),
        })
      } else {
        saved = await createProject({ title, description, status, owners: selectedOwners })
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

function authUserToOwner(u: AppUser): SharePointUser {
  return { id: u.id, displayName: u.displayName, email: u.mail }
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
      border-radius: 4px; font-family: inherit; font-size: 0.9rem; outline: none; box-sizing: border-box; }
    .form-input:focus { border-color: var(--color-primary); box-shadow: 0 0 0 2px var(--color-primary-light); }
    textarea.form-input { resize: vertical; }
    .owners-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; min-height: 0; }
    .owner-chip { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px 3px 10px;
      background: var(--color-primary-light, #deecf9); color: var(--color-primary, #005a9e);
      border-radius: 12px; font-size: 0.82rem; font-weight: 500; }
    .chip-remove { background: none; border: none; cursor: pointer; font-size: 0.75rem;
      color: inherit; opacity: 0.7; padding: 0 1px; line-height: 1; }
    .chip-remove:hover { opacity: 1; }
    .owners-search-wrap { position: relative; }
    .owners-dropdown { position: absolute; top: 100%; left: 0; right: 0; background: white;
      border: 1px solid var(--color-border); border-radius: 4px; box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      list-style: none; margin: 2px 0 0; padding: 4px 0; z-index: 200; max-height: 200px; overflow-y: auto; }
    .owners-option { padding: 8px 12px; cursor: pointer; display: flex; flex-direction: column; gap: 1px; }
    .owners-option:hover { background: var(--color-surface-alt, #f3f2f1); }
    .opt-name { font-size: 0.88rem; font-weight: 500; }
    .opt-email { font-size: 0.78rem; color: var(--color-text-muted); }
    .form-error { padding: 10px 12px; background: #fde7e9; color: #a4262c;
      border-radius: 4px; font-size: 0.85rem; margin-bottom: 12px; }
    .form-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 24px; }
  `
  document.head.appendChild(style)
}
