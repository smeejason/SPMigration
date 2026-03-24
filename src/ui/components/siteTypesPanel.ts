import { getSiteDesigns, searchUsers } from '../../graph/graphClient'
import { updateProject } from '../../graph/projectService'
import { getState, setState } from '../../state/store'
import type { SiteType, UserRef } from '../../types'

// ─── Entry point ──────────────────────────────────────────────────────────────

export function renderSiteTypesPanel(container: HTMLElement): void {
  injectSiteTypesStyles()
  render(container)
}

// ─── Main render ──────────────────────────────────────────────────────────────

function render(container: HTMLElement, editingId: string | null = null): void {
  const siteTypes = getSiteTypes()
  const editing = editingId ? siteTypes.find(s => s.id === editingId) ?? null : null

  container.innerHTML = `
    <div class="st-panel">
      <!-- Left: list -->
      <div class="st-list-col">
        <div class="st-list-header">
          <h3 class="st-list-title">Site Types</h3>
          <p class="st-list-sub">Define reusable site presets. When mapping a folder to a new SharePoint site, pick a type to pre-fill its settings.</p>
        </div>
        <div class="st-list" id="st-list">
          ${siteTypes.length === 0
            ? `<div class="st-empty">No site types defined yet — add one using the form.</div>`
            : siteTypes.map(st => siteTypeCardHtml(st, editingId === st.id)).join('')}
        </div>
      </div>

      <!-- Right: form -->
      <div class="st-form-col">
        <div class="st-form-wrap">
          <h3 class="st-form-title">${editing ? 'Edit Site Type' : 'Add Site Type'}</h3>
          <form id="st-form" autocomplete="off">
            <input type="hidden" id="st-id" value="${escHtml(editing?.id ?? '')}" />

            <div class="form-group">
              <label for="st-name">Name <span class="required">*</span></label>
              <input id="st-name" type="text" class="form-input" placeholder="e.g. Department Site"
                value="${escHtml(editing?.name ?? '')}" required />
            </div>

            <div class="form-group">
              <label class="checkbox-label">
                <input type="checkbox" id="st-create-team"
                  ${editing?.createTeam ? 'checked' : ''} />
                Also create a Microsoft Teams team
              </label>
              <small class="form-hint">Provisions a Teams team linked to this M365 Group site.</small>
            </div>

            <div class="form-group">
              <label for="st-desc">Description</label>
              <textarea id="st-desc" class="form-input" rows="2"
                placeholder="When should this type be used?">${escHtml(editing?.description ?? '')}</textarea>
            </div>

            <div class="form-group">
              <label for="st-library">Default Library</label>
              <input id="st-library" type="text" class="form-input" placeholder="e.g. Shared Documents"
                value="${escHtml(editing?.defaultLibrary ?? '')}" />
            </div>

            <div class="form-group">
              <label for="st-subfolder">Default Subfolder</label>
              <input id="st-subfolder" type="text" class="form-input" placeholder="e.g. /Migration/Phase1"
                value="${escHtml(editing?.defaultSubfolder ?? '')}" />
            </div>

            <div class="form-group">
              <label>Org Site Design <span class="hint">(optional)</span></label>
              <div class="st-design-row">
                <select id="st-design-select" class="form-input st-design-select">
                  <option value="">— None —</option>
                  ${editing?.siteDesignId
                    ? `<option value="${escHtml(editing.siteDesignId)}" selected>${escHtml(editing.siteDesignName ?? editing.siteDesignId)}</option>`
                    : ''}
                </select>
                <button type="button" id="btn-load-designs" class="btn btn-secondary btn-sm">Load</button>
              </div>
              <small class="form-hint" id="st-design-hint">Click Load to fetch your organisation's published site designs.</small>
            </div>

            <div class="form-group">
              <label>Default Owners</label>
              <div class="st-people-chips" id="st-owners-chips"></div>
              <div class="st-people-search-wrap">
                <input id="st-owners-search" type="text" class="form-input" placeholder="Search people…" autocomplete="off" />
                <ul id="st-owners-dropdown" class="st-people-dropdown" style="display:none"></ul>
              </div>
            </div>

            <div class="form-group">
              <label>Default Members</label>
              <div class="st-people-chips" id="st-members-chips"></div>
              <div class="st-people-search-wrap">
                <input id="st-members-search" type="text" class="form-input" placeholder="Search people…" autocomplete="off" />
                <ul id="st-members-dropdown" class="st-people-dropdown" style="display:none"></ul>
              </div>
            </div>

            <div id="st-error" class="form-error" style="display:none"></div>

            <div class="st-form-actions">
              <button type="submit" id="btn-st-save" class="btn btn-primary">
                ${editing ? 'Save Changes' : 'Add Site Type'}
              </button>
              ${editing ? `<button type="button" id="btn-st-cancel" class="btn btn-ghost">Cancel</button>` : ''}
            </div>
          </form>
        </div>
      </div>
    </div>
  `

  // ── People state ─────────────────────────────────────────────────────────
  let owners: UserRef[] = editing ? [...editing.owners] : []
  let members: UserRef[] = editing ? [...editing.members] : []

  renderChips(container, '#st-owners-chips', owners, () => renderChips(container, '#st-owners-chips', owners))
  renderChips(container, '#st-members-chips', members, () => renderChips(container, '#st-members-chips', members))

  attachPeopleSearch(container, '#st-owners-search', '#st-owners-dropdown', owners, () =>
    renderChips(container, '#st-owners-chips', owners))
  attachPeopleSearch(container, '#st-members-search', '#st-members-dropdown', members, () =>
    renderChips(container, '#st-members-chips', members))

  // ── Load org site designs ─────────────────────────────────────────────────
  container.querySelector('#btn-load-designs')?.addEventListener('click', async () => {
    const btn = container.querySelector('#btn-load-designs') as HTMLButtonElement
    const hintEl = container.querySelector('#st-design-hint') as HTMLElement
    const select = container.querySelector<HTMLSelectElement>('#st-design-select')!

    btn.disabled = true
    btn.textContent = 'Loading…'
    hintEl.textContent = 'Fetching site designs from your tenant…'

    try {
      const designs = await getSiteDesigns()
      // WebTemplate 64 = Team site
      const filtered = designs.filter(d => !d.webTemplate || d.webTemplate === '64')

      const currentVal = select.value
      select.innerHTML = '<option value="">— None —</option>'
      filtered.forEach(d => {
        const opt = document.createElement('option')
        opt.value = d.id
        opt.textContent = d.title
        if (d.description) opt.title = d.description
        if (d.id === currentVal) opt.selected = true
        select.appendChild(opt)
      })

      hintEl.textContent = filtered.length === 0
        ? 'No custom site designs found for this template type.'
        : `${filtered.length} design${filtered.length !== 1 ? 's' : ''} loaded.`
    } catch (err) {
      hintEl.textContent = `Failed to load designs: ${(err as Error).message}`
    } finally {
      btn.disabled = false
      btn.textContent = 'Load'
    }
  })

  // ── Edit / Delete card buttons ────────────────────────────────────────────
  container.querySelectorAll<HTMLButtonElement>('.btn-st-edit').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id!
      render(container, id)
    })
  })

  container.querySelectorAll<HTMLButtonElement>('.btn-st-delete').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.id!
      const name = btn.dataset.name!
      if (!confirm(`Delete site type "${name}"?`)) return
      await saveSiteTypes(getSiteTypes().filter(s => s.id !== id))
      render(container, null)
    })
  })

  // ── Cancel edit ───────────────────────────────────────────────────────────
  container.querySelector('#btn-st-cancel')?.addEventListener('click', () => render(container, null))

  // ── Form submit ───────────────────────────────────────────────────────────
  container.querySelector('#st-form')?.addEventListener('submit', async (e) => {
    e.preventDefault()
    const errorEl = container.querySelector('#st-error') as HTMLElement
    const saveBtn = container.querySelector('#btn-st-save') as HTMLButtonElement

    const name = (container.querySelector('#st-name') as HTMLInputElement).value.trim()
    const template: 'team' = 'team'
    const desc = (container.querySelector('#st-desc') as HTMLTextAreaElement).value.trim()
    const library = (container.querySelector('#st-library') as HTMLInputElement).value.trim()
    const subfolder = (container.querySelector('#st-subfolder') as HTMLInputElement).value.trim()
    const createTeam = (container.querySelector('#st-create-team') as HTMLInputElement).checked
    const designSelect = container.querySelector<HTMLSelectElement>('#st-design-select')!
    const siteDesignId = designSelect.value || undefined
    const siteDesignName = siteDesignId
      ? designSelect.options[designSelect.selectedIndex]?.text ?? undefined
      : undefined
    const idVal = (container.querySelector('#st-id') as HTMLInputElement).value

    if (!name) { errorEl.textContent = 'Name is required.'; errorEl.style.display = ''; return }

    const siteType: SiteType = {
      id: idVal || crypto.randomUUID(),
      name,
      template,
      description: desc || undefined,
      defaultLibrary: library || undefined,
      defaultSubfolder: subfolder || undefined,
      siteDesignId,
      siteDesignName,
      createTeam: createTeam || undefined,
      owners,
      members,
    }

    errorEl.style.display = 'none'
    saveBtn.disabled = true
    saveBtn.textContent = 'Saving…'

    try {
      const existing = getSiteTypes()
      const updated = idVal
        ? existing.map(s => s.id === idVal ? siteType : s)
        : [...existing, siteType]
      await saveSiteTypes(updated)
      render(container, null)
    } catch (err) {
      errorEl.textContent = `Save failed: ${(err as Error).message}`
      errorEl.style.display = ''
      saveBtn.disabled = false
      saveBtn.textContent = editing ? 'Save Changes' : 'Add Site Type'
    }
  })
}

// ─── People picker helpers ────────────────────────────────────────────────────

function renderChips(container: HTMLElement, chipsSelector: string, people: UserRef[], onChange?: () => void): void {
  const el = container.querySelector(chipsSelector) as HTMLElement | null
  if (!el) return
  el.innerHTML = people.map(p => `
    <span class="st-chip" data-id="${escHtml(p.id)}">
      ${escHtml(p.displayName)}
      <button type="button" class="st-chip-remove" data-id="${escHtml(p.id)}" title="Remove">✕</button>
    </span>`).join('')
  el.querySelectorAll<HTMLButtonElement>('.st-chip-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id!
      const idx = people.findIndex(p => p.id === id)
      if (idx !== -1) people.splice(idx, 1)
      renderChips(container, chipsSelector, people, onChange)
      onChange?.()
    })
  })
}

function attachPeopleSearch(
  container: HTMLElement,
  inputSelector: string,
  dropdownSelector: string,
  people: UserRef[],
  onChange: () => void
): void {
  const input = container.querySelector<HTMLInputElement>(inputSelector)
  const dropdown = container.querySelector<HTMLUListElement>(dropdownSelector)
  if (!input || !dropdown) return

  let debounceTimer: ReturnType<typeof setTimeout>

  input.addEventListener('input', () => {
    clearTimeout(debounceTimer)
    const q = input.value.trim()
    if (!q) { dropdown.style.display = 'none'; return }
    debounceTimer = setTimeout(async () => {
      try {
        const users = await searchUsers(q)
        const available = users.filter(u => !people.some(p => p.id === u.id))
        if (available.length === 0) { dropdown.style.display = 'none'; return }
        dropdown.innerHTML = available.map(u => `
          <li class="st-people-option" data-id="${escHtml(u.id)}"
              data-name="${escHtml(u.displayName)}"
              data-email="${escHtml(u.mail ?? u.userPrincipalName ?? '')}">
            <span class="st-person-name">${escHtml(u.displayName)}</span>
            <span class="st-person-email">${escHtml(u.mail ?? u.userPrincipalName ?? '')}</span>
          </li>`).join('')
        dropdown.style.display = ''
        dropdown.querySelectorAll<HTMLLIElement>('.st-people-option').forEach(li => {
          li.addEventListener('click', () => {
            people.push({ id: li.dataset.id!, displayName: li.dataset.name!, email: li.dataset.email! })
            input.value = ''
            dropdown.style.display = 'none'
            onChange()
          })
        })
      } catch { dropdown.style.display = 'none' }
    }, 250)
  })

  document.addEventListener('click', (e) => {
    if (!input.contains(e.target as Node) && !dropdown.contains(e.target as Node)) {
      dropdown.style.display = 'none'
    }
  }, { capture: true })
}

// ─── Card HTML ────────────────────────────────────────────────────────────────

function siteTypeCardHtml(st: SiteType, isActive: boolean): string {
  const templateLabel = st.template === 'team' ? 'Team' : 'Communication'
  const templateClass = st.template === 'team' ? 'st-badge--team' : 'st-badge--comms'
  const ownerCount = st.owners.length
  const memberCount = st.members.length
  const hasPeople = ownerCount > 0 || memberCount > 0

  return `
    <div class="st-card${isActive ? ' st-card--active' : ''}">
      <div class="st-card-header">
        <span class="st-card-name">${escHtml(st.name)}</span>
        <span class="st-badge ${templateClass}">${templateLabel}</span>
        ${st.createTeam ? '<span class="st-badge st-badge--teams">+ Teams</span>' : ''}
      </div>
      ${st.description ? `<div class="st-card-desc">${escHtml(st.description)}</div>` : ''}
      <div class="st-card-meta">
        ${st.defaultLibrary ? `<span title="Library">📚 ${escHtml(st.defaultLibrary)}</span>` : ''}
        ${st.defaultSubfolder ? `<span title="Subfolder">📂 ${escHtml(st.defaultSubfolder)}</span>` : ''}
        ${st.siteDesignName ? `<span title="Site design">🎨 ${escHtml(st.siteDesignName)}</span>` : ''}
        ${hasPeople ? `<span title="Owners / Members">👥 ${ownerCount} owner${ownerCount !== 1 ? 's' : ''}, ${memberCount} member${memberCount !== 1 ? 's' : ''}</span>` : ''}
      </div>
      <div class="st-card-actions">
        <button class="btn btn-ghost btn-sm btn-st-edit" data-id="${escHtml(st.id)}">Edit</button>
        <button class="btn btn-ghost btn-sm btn-st-delete" data-id="${escHtml(st.id)}" data-name="${escHtml(st.name)}">Delete</button>
      </div>
    </div>
  `
}

// ─── Persistence ──────────────────────────────────────────────────────────────

function getSiteTypes(): SiteType[] {
  return getState().currentProject?.projectData.siteTypes ?? []
}

async function saveSiteTypes(siteTypes: SiteType[]): Promise<void> {
  const project = getState().currentProject
  if (!project) return
  const updatedData = { ...project.projectData, siteTypes }
  await updateProject(project.id, { projectData: updatedData })
  setState({ currentProject: { ...project, projectData: updatedData } })
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

// ─── Styles ───────────────────────────────────────────────────────────────────

function injectSiteTypesStyles(): void {
  if (document.getElementById('site-types-styles')) return
  const style = document.createElement('style')
  style.id = 'site-types-styles'
  style.textContent = `
    .st-panel { display: grid; grid-template-columns: 1fr 1fr; height: calc(100vh - 140px);
      overflow: hidden; }
    .st-list-col { overflow-y: auto; border-right: 1px solid var(--color-border);
      display: flex; flex-direction: column; }
    .st-form-col { overflow-y: auto; }
    .st-list-header { padding: 20px 20px 12px; border-bottom: 1px solid var(--color-border);
      background: var(--color-surface-alt); position: sticky; top: 0; z-index: 1; }
    .st-list-title { font-size: 0.95rem; font-weight: 700; margin: 0 0 4px; }
    .st-list-sub { font-size: 0.78rem; color: var(--color-text-muted); margin: 0; line-height: 1.4; }
    .st-list { padding: 12px; display: flex; flex-direction: column; gap: 10px; }
    .st-empty { padding: 32px 16px; text-align: center; color: var(--color-text-muted); font-size: 0.875rem; }

    .st-card { border: 1px solid var(--color-border); border-radius: 8px; padding: 14px 16px;
      background: white; display: flex; flex-direction: column; gap: 6px; transition: box-shadow 0.15s; }
    .st-card:hover { box-shadow: var(--shadow); }
    .st-card--active { border-color: var(--color-primary); box-shadow: 0 0 0 2px rgba(0,120,212,0.15); }
    .st-card-header { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .st-card-name { font-weight: 600; font-size: 0.9rem; flex: 1; }
    .st-card-desc { font-size: 0.8rem; color: var(--color-text-muted); }
    .st-card-meta { display: flex; flex-wrap: wrap; gap: 10px; font-size: 0.78rem; color: var(--color-text-muted); }
    .st-card-actions { display: flex; gap: 6px; margin-top: 4px; }

    .st-badge { padding: 2px 8px; border-radius: 10px; font-size: 0.72rem; font-weight: 600; white-space: nowrap; }
    .st-badge--team { background: #deecf9; color: #005a9e; }
    .st-badge--comms { background: #f3f2f1; color: #323130; }
    .st-badge--teams { background: #e8d5fb; color: #5c2d91; }

    .st-form-wrap { padding: 20px 24px; max-width: 480px; }
    .st-form-title { font-size: 0.95rem; font-weight: 700; margin: 0 0 20px; }
    .st-radio-group { display: flex; flex-direction: column; gap: 6px; }
    .radio-label, .checkbox-label { display: flex; align-items: center; gap: 6px; font-size: 0.875rem; cursor: pointer; }
    .st-design-row { display: flex; gap: 8px; align-items: center; }
    .st-design-select { flex: 1; }
    .st-form-actions { display: flex; gap: 8px; padding-top: 8px; }

    .st-people-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; min-height: 0; }
    .st-chip { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px 3px 10px;
      background: #deecf9; color: #005a9e; border-radius: 12px; font-size: 0.8rem; font-weight: 500; }
    .st-chip-remove { background: none; border: none; cursor: pointer; color: inherit;
      font-size: 0.75rem; padding: 0 1px; line-height: 1; opacity: 0.7; }
    .st-chip-remove:hover { opacity: 1; }

    .st-people-search-wrap { position: relative; }
    .st-people-dropdown { position: absolute; top: 100%; left: 0; right: 0; background: white;
      border: 1px solid var(--color-border); border-radius: 4px; box-shadow: var(--shadow);
      z-index: 20; list-style: none; padding: 0; margin: 2px 0 0; max-height: 180px; overflow-y: auto; }
    .st-people-option { padding: 8px 12px; cursor: pointer; display: flex; flex-direction: column; gap: 1px; }
    .st-people-option:hover { background: var(--color-surface-alt); }
    .st-person-name { font-size: 0.875rem; font-weight: 500; }
    .st-person-email { font-size: 0.75rem; color: var(--color-text-muted); }

    .form-hint { font-size: 0.78rem; color: var(--color-text-muted); display: block; margin-top: 3px; }
    .required { color: var(--color-danger); }
    .hint { font-size: 0.78rem; color: var(--color-text-muted); font-weight: 400; }
  `
  document.head.appendChild(style)
}
