import { createTeamSite, waitForGroupSite } from '../../graph/graphClient'
import { setState, getState } from '../../state/store'
import type { SiteRequest, SiteTemplate } from '../../types'

export function renderSiteCreator(container: HTMLElement): void {
  container.innerHTML = `
    <div class="creator-panel">
      <div class="panel-section">
        <h3>Create New SharePoint Site</h3>
        <p class="panel-desc">New sites are created as Microsoft 365 group-connected Team sites. The site will be available in the Mapping panel once provisioned.</p>
        <form id="site-creator-form" class="site-form">
          <div class="form-group">
            <label for="site-name">Site Display Name <span class="required">*</span></label>
            <input id="site-name" type="text" class="form-input" required
              placeholder="e.g. Engineering" />
          </div>
          <div class="form-group">
            <label for="site-alias">URL Alias <span class="required">*</span></label>
            <div class="alias-row">
              <span class="alias-prefix">.../sites/</span>
              <input id="site-alias" type="text" class="form-input" required
                placeholder="engineering" pattern="[a-zA-Z0-9\-]+" title="Letters, numbers, and hyphens only" />
            </div>
            <small class="form-hint">Letters, numbers, and hyphens only. Cannot be changed after creation.</small>
          </div>
          <div class="form-group">
            <label for="site-desc">Description</label>
            <textarea id="site-desc" class="form-input" rows="2" placeholder="Optional description"></textarea>
          </div>
          <div class="form-group">
            <label>Template</label>
            <div class="template-row">
              <label class="radio-label">
                <input type="radio" name="site-template" value="team" checked /> Team site (M365 Group)
              </label>
            </div>
            <small class="form-hint">Communication sites require SharePoint admin permissions. Team site support only in Phase 1.</small>
          </div>
          <div id="creator-error" class="form-error" style="display:none"></div>
          <button type="submit" id="btn-create-site" class="btn btn-primary">Create Site</button>
        </form>
      </div>

      <div class="panel-section" id="pending-sites-section" style="${getState().pendingSiteCreations.length === 0 ? 'display:none' : ''}">
        <h3>Pending / Created Sites</h3>
        <div id="pending-sites-list"></div>
      </div>
    </div>
  `
  injectCreatorStyles()
  renderPendingList(container)

  // Auto-fill alias from name
  const nameInput = container.querySelector('#site-name') as HTMLInputElement
  const aliasInput = container.querySelector('#site-alias') as HTMLInputElement
  nameInput.addEventListener('input', () => {
    if (aliasInput.dataset.userEdited) return
    aliasInput.value = nameInput.value.toLowerCase().replace(/[^a-z0-9-]/g, '-').replace(/-+/g, '-').slice(0, 60)
  })
  aliasInput.addEventListener('input', () => { aliasInput.dataset.userEdited = '1' })

  container.querySelector('#site-creator-form')!.addEventListener('submit', async (e) => {
    e.preventDefault()
    const name = nameInput.value.trim()
    const alias = aliasInput.value.trim()
    const desc = (container.querySelector('#site-desc') as HTMLTextAreaElement).value.trim()
    const errorEl = container.querySelector('#creator-error') as HTMLElement
    const createBtn = container.querySelector('#btn-create-site') as HTMLButtonElement

    if (!name || !alias) return

    const request: SiteRequest = {
      id: crypto.randomUUID(),
      displayName: name,
      alias,
      description: desc,
      template: 'team' as SiteTemplate,
      status: 'creating',
    }

    setState({ pendingSiteCreations: [...getState().pendingSiteCreations, request] })
    renderPendingList(container)

    createBtn.disabled = true
    errorEl.style.display = 'none'

    try {
      const groupId = await createTeamSite(request)
      updateRequest(request.id, { status: 'creating' })

      const site = await waitForGroupSite(groupId)
      updateRequest(request.id, { status: 'created', createdSite: site })

      // Add to sites cache
      setState({ sites: [...getState().sites, site] })
      renderPendingList(container)

      // Reset form
      ;(container.querySelector('#site-creator-form') as HTMLFormElement).reset()
      delete aliasInput.dataset.userEdited
    } catch (err) {
      updateRequest(request.id, { status: 'failed', error: (err as Error).message })
      renderPendingList(container)
      errorEl.textContent = `Failed: ${(err as Error).message}`
      errorEl.style.display = 'block'
    } finally {
      createBtn.disabled = false
    }
  })
}

function updateRequest(id: string, patch: Partial<SiteRequest>): void {
  setState({
    pendingSiteCreations: getState().pendingSiteCreations.map((r) =>
      r.id === id ? { ...r, ...patch } : r
    ),
  })
}

function renderPendingList(container: HTMLElement): void {
  const section = container.querySelector('#pending-sites-section') as HTMLElement
  const list = container.querySelector('#pending-sites-list') as HTMLElement
  const pending = getState().pendingSiteCreations

  if (pending.length === 0) { section.style.display = 'none'; return }
  section.style.display = ''

  list.innerHTML = pending.map((r) => `
    <div class="pending-site pending-site--${r.status}">
      <div class="pending-site-name">${escHtml(r.displayName)}</div>
      <div class="pending-site-meta">.../sites/${escHtml(r.alias)}</div>
      <div class="pending-site-status">
        ${r.status === 'creating' ? '⏳ Provisioning…'
          : r.status === 'created' ? '✅ Created'
          : r.status === 'failed' ? `❌ Failed: ${escHtml(r.error ?? '')}`
          : '⏳ Pending'}
      </div>
    </div>
  `).join('')
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function injectCreatorStyles(): void {
  if (document.getElementById('creator-styles')) return
  const style = document.createElement('style')
  style.id = 'creator-styles'
  style.textContent = `
    .creator-panel { padding: 24px; max-width: 600px; }
    .site-form { max-width: 480px; }
    .alias-row { display: flex; align-items: center; gap: 0; }
    .alias-prefix { background: var(--color-surface-alt); border: 1px solid var(--color-border);
      border-right: none; padding: 8px 10px; border-radius: 4px 0 0 4px; font-size: 0.85rem;
      color: var(--color-text-muted); white-space: nowrap; }
    .alias-row .form-input { border-radius: 0 4px 4px 0; }
    .template-row { margin-bottom: 4px; }
    .radio-label { display: flex; align-items: center; gap: 6px; font-size: 0.88rem; cursor: pointer; }
    .form-hint { font-size: 0.78rem; color: var(--color-text-muted); }
    .pending-site { padding: 10px 14px; border: 1px solid var(--color-border); border-radius: 6px;
      margin-bottom: 8px; }
    .pending-site--created { border-color: var(--color-success); background: #f0fff0; }
    .pending-site--failed { border-color: var(--color-danger); background: #fff0f0; }
    .pending-site-name { font-weight: 600; font-size: 0.9rem; }
    .pending-site-meta { font-size: 0.8rem; color: var(--color-text-muted); margin: 2px 0; }
    .pending-site-status { font-size: 0.82rem; margin-top: 4px; }
  `
  document.head.appendChild(style)
}
