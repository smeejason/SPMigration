import { getState, setState, subscribe } from '../state/store'
import { signOut, getCurrentUser } from '../auth/authService'
import { renderAuthPanel } from './components/authPanel'
import { renderProjectList } from './components/projectList'
import { renderProjectForm } from './components/projectForm'
import { renderUploadPanel } from './components/uploadPanel'
import { renderMappingPanel } from './components/mappingPanel'
import { renderSiteCreator } from './components/siteCreator'
import { renderSummaryPanel } from './components/summaryPanel'
import { renderAutoMapPanel } from './components/autoMapPanel'
import { renderReviewPanel } from './components/reviewPanel'
import type { AppState, MigrationProject } from '../types'

// ─── Waffle SVG ───────────────────────────────────────────────────────────────

const WAFFLE_SVG = `
  <svg width="18" height="18" viewBox="0 0 18 18" fill="currentColor" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
    <rect x="0"   y="0"   width="5" height="5" rx="1"/>
    <rect x="6.5" y="0"   width="5" height="5" rx="1"/>
    <rect x="13"  y="0"   width="5" height="5" rx="1"/>
    <rect x="0"   y="6.5" width="5" height="5" rx="1"/>
    <rect x="6.5" y="6.5" width="5" height="5" rx="1"/>
    <rect x="13"  y="6.5" width="5" height="5" rx="1"/>
    <rect x="0"   y="13"  width="5" height="5" rx="1"/>
    <rect x="6.5" y="13"  width="5" height="5" rx="1"/>
    <rect x="13"  y="13"  width="5" height="5" rx="1"/>
  </svg>`

// ─── App shell HTML ───────────────────────────────────────────────────────────

function shellHtml(user: string, projectName?: string): string {
  return `
    <header class="app-header">
      <div class="header-left">
        <div class="waffle-wrap">
          <button id="btn-waffle" class="waffle-btn" title="Apps" aria-label="Apps">
            ${WAFFLE_SVG}
          </button>
          <div id="waffle-menu" class="waffle-menu" hidden>
            <div class="waffle-menu-item" id="waffle-projects">Projects</div>
          </div>
        </div>
        <span class="app-logo">SP Migration Planner</span>
        ${projectName ? `<span class="header-separator">›</span><span class="header-project">${escHtml(projectName)}</span>` : ''}
      </div>
      <div class="header-right">
        <span class="header-user">${escHtml(user)}</span>
        <button id="btn-signout" class="btn btn-ghost btn-sm">Sign out</button>
      </div>
    </header>
    <div id="modal-root"></div>
    <main id="app-main"></main>
  `
}

function projectsContextualNavHtml(): string {
  return `
    <div class="contextual-nav">
      <div class="contextual-nav-left">
        <span class="contextual-nav-title">Projects</span>
      </div>
      <div class="contextual-nav-right">
        <button id="btn-new-project" class="btn btn-primary btn-sm">+ New Project</button>
      </div>
    </div>
    <div id="workspace-panel" class="workspace-panel"></div>
  `
}

function projectWorkspaceHtml(projectTitle: string, projectType: string): string {
  const tabs = projectType === 'OneDrive'
    ? `
      <button class="tab-btn" data-view="project-upload">Upload</button>
      <button class="tab-btn" data-view="project-automap">Auto Map</button>
      <button class="tab-btn" data-view="project-map">Map</button>
      <button class="tab-btn" data-view="project-summary">Summary</button>
      <button class="tab-btn" data-view="project-review">Review</button>`
    : `
      <button class="tab-btn" data-view="project-upload">Upload</button>
      <button class="tab-btn" data-view="project-map">Map</button>
      <button class="tab-btn" data-view="project-sites">Create Sites</button>
      <button class="tab-btn" data-view="project-summary">Summary</button>
      <button class="tab-btn" data-view="project-review">Review</button>`
  return `
    <nav class="workspace-tabs">
      ${tabs}
      <div class="tab-spacer"></div>
      <span class="workspace-project-name">${escHtml(projectTitle)}</span>
      <button id="btn-back-projects" class="btn btn-ghost btn-sm">← Projects</button>
    </nav>
    <div id="workspace-panel" class="workspace-panel"></div>
  `
}

// ─── Main render function ─────────────────────────────────────────────────────

export function mountApp(root: HTMLElement): void {
  injectShellStyles()

  let prevView: AppState['ui']['activeView'] | null = null

  const render = (state: AppState): void => {
    const { activeView } = state.ui
    if (activeView === prevView) return
    prevView = activeView

    if (activeView === 'login') {
      root.innerHTML = ''
      renderAuthPanel(root)
      return
    }

    const user = state.auth.user?.displayName ?? state.auth.user?.userPrincipalName ?? ''
    const project = state.currentProject

    if (activeView === 'projects') {
      root.innerHTML = shellHtml(user)
      attachSignOut(root)
      attachWaffle(root)

      const main = root.querySelector('#app-main') as HTMLElement
      main.innerHTML = projectsContextualNavHtml()

      const panel = main.querySelector('#workspace-panel') as HTMLElement
      void renderProjectList(panel)

      main.querySelector('#btn-new-project')?.addEventListener('click', () => {
        const modal = root.querySelector('#modal-root') as HTMLElement
        renderProjectForm(
          modal,
          null,
          (saved: MigrationProject) => {
            setState({
              currentProject: saved,
              treeData: null,
              mappings: [],
              oneDriveMappings: [],
              sites: [],
              pendingSiteCreations: [],
              reviewData: null,
              ui: { activeView: 'project-upload', loading: false, error: null },
            })
          },
          () => { /* cancelled */ }
        )
      })
      return
    }

    if (!project) {
      setState({ ui: { activeView: 'projects', loading: false, error: null } })
      return
    }

    // Project workspace views
    root.innerHTML = shellHtml(user)
    attachSignOut(root)
    attachWaffle(root)

    const main = root.querySelector('#app-main') as HTMLElement
    main.innerHTML = projectWorkspaceHtml(project.title, project.type)

    const panel = main.querySelector('#workspace-panel') as HTMLElement
    const tabs = main.querySelectorAll('.tab-btn[data-view]')

    const setActiveTab = (view: string): void => {
      tabs.forEach((t) => t.classList.toggle('tab-btn--active', t.getAttribute('data-view') === view))
    }

    const renderPanel = (view: AppState['ui']['activeView']): void => {
      panel.innerHTML = ''
      if (view === 'project-upload') renderUploadPanel(panel)
      else if (view === 'project-automap') renderAutoMapPanel(panel)
      else if (view === 'project-map') renderMappingPanel(panel)
      else if (view === 'project-sites') renderSiteCreator(panel)
      else if (view === 'project-summary') renderSummaryPanel(panel)
      else if (view === 'project-review') void renderReviewPanel(panel)
      setActiveTab(view)
    }

    renderPanel(activeView)

    tabs.forEach((tab) => {
      tab.addEventListener('click', () => {
        const view = tab.getAttribute('data-view') as AppState['ui']['activeView']
        prevView = null
        renderPanel(view)
        prevView = view
        setActiveTab(view)
        setState({ ui: { activeView: view, loading: false, error: null } })
        prevView = view
      })
    })

    main.querySelector('#btn-back-projects')?.addEventListener('click', () => {
      setState({ currentProject: null, treeData: null, mappings: [], oneDriveMappings: [], sites: [], pendingSiteCreations: [], reviewData: null, ui: { activeView: 'projects', loading: false, error: null } })
    })
  }

  // Initial render
  render(getState())
  subscribe(render)
}

function attachSignOut(root: HTMLElement): void {
  root.querySelector('#btn-signout')?.addEventListener('click', async () => {
    try {
      await signOut()
    } catch {
      // Ignore sign-out errors
    }
    setState({
      auth: { user: null, isAuthenticated: false },
      currentProject: null,
      projects: [],
      ui: { activeView: 'login', loading: false, error: null },
    })
  })
}

function attachWaffle(root: HTMLElement): void {
  const btn = root.querySelector('#btn-waffle') as HTMLElement | null
  const menu = root.querySelector('#waffle-menu') as HTMLElement | null
  if (!btn || !menu) return

  btn.addEventListener('click', (e) => {
    e.stopPropagation()
    menu.hidden = !menu.hidden
  })

  const closeOnOutsideClick = (): void => {
    if (!document.contains(btn)) {
      document.removeEventListener('click', closeOnOutsideClick)
      return
    }
    menu.hidden = true
  }
  document.addEventListener('click', closeOnOutsideClick)

  root.querySelector('#waffle-projects')?.addEventListener('click', () => {
    menu.hidden = true
    setState({ currentProject: null, treeData: null, mappings: [], oneDriveMappings: [], sites: [], pendingSiteCreations: [], reviewData: null, ui: { activeView: 'projects', loading: false, error: null } })
  })
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function injectShellStyles(): void {
  if (document.getElementById('shell-styles')) return
  const style = document.createElement('style')
  style.id = 'shell-styles'
  style.textContent = `
    .app-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 16px 0 4px; height: 48px; background: var(--color-primary);
      color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.15); flex-shrink: 0;
    }
    .header-left { display: flex; align-items: center; gap: 8px; }
    .app-logo { font-weight: 700; font-size: 1rem; letter-spacing: 0.01em; }
    .header-separator { opacity: 0.6; font-size: 1.2rem; }
    .header-project { font-size: 0.9rem; opacity: 0.9; max-width: 300px;
      overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .header-right { display: flex; align-items: center; gap: 12px; }
    .header-user { font-size: 0.85rem; opacity: 0.85; }
    .app-header .btn-ghost {
      color: rgba(255,255,255,0.9); border-color: rgba(255,255,255,0.4);
      background: transparent;
    }
    .app-header .btn-ghost:hover { background: rgba(255,255,255,0.15); }

    /* Waffle */
    .waffle-wrap { position: relative; display: flex; align-items: center; }
    .waffle-btn {
      display: flex; align-items: center; justify-content: center;
      width: 40px; height: 40px; background: transparent; border: none;
      border-radius: 4px; cursor: pointer; color: rgba(255,255,255,0.9);
      flex-shrink: 0;
    }
    .waffle-btn:hover { background: rgba(255,255,255,0.15); }
    .waffle-menu {
      position: absolute; top: calc(100% + 6px); left: 0;
      background: white; border: 1px solid var(--color-border);
      border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,0.18);
      min-width: 180px; z-index: 300; padding: 6px 0;
    }
    .waffle-menu-item {
      padding: 10px 16px; font-size: 0.875rem; cursor: pointer;
      color: var(--color-text); font-weight: 500;
    }
    .waffle-menu-item:hover { background: var(--color-surface-alt); }

    /* Contextual nav (projects page) */
    .contextual-nav {
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 32px; height: 48px; background: white;
      border-bottom: 1px solid var(--color-border); flex-shrink: 0;
    }
    .contextual-nav-left { display: flex; align-items: center; gap: 12px; }
    .contextual-nav-right { display: flex; align-items: center; gap: 8px; }
    .contextual-nav-title { font-size: 1rem; font-weight: 600; color: var(--color-text); }

    /* Workspace tabs (project open) */
    #app-main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
    .workspace-tabs {
      display: flex; align-items: center; gap: 2px; padding: 0 16px;
      background: white; border-bottom: 1px solid var(--color-border);
      height: 44px; flex-shrink: 0;
    }
    .workspace-project-name {
      font-size: 0.8rem; color: var(--color-text-muted);
      padding: 0 8px; white-space: nowrap; overflow: hidden;
      text-overflow: ellipsis; max-width: 220px;
    }
    .tab-btn {
      padding: 8px 16px; background: none; border: none; border-bottom: 3px solid transparent;
      font-family: inherit; font-size: 0.875rem; cursor: pointer; color: var(--color-text-muted);
      transition: color 0.15s; margin-bottom: -1px;
    }
    .tab-btn:hover { color: var(--color-text); }
    .tab-btn--active { color: var(--color-primary); border-bottom-color: var(--color-primary); font-weight: 600; }
    .tab-spacer { flex: 1; }
    .workspace-panel { flex: 1; overflow-y: auto; background: var(--color-bg); }
    #modal-root { position: relative; z-index: 99; }
  `
  document.head.appendChild(style)

  if (document.getElementById('shared-styles')) return
  const shared = document.createElement('style')
  shared.id = 'shared-styles'
  shared.textContent = `
    .btn { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
      border: 1px solid transparent; border-radius: 4px; font-family: inherit;
      font-size: 0.875rem; cursor: pointer; transition: background 0.15s, border-color 0.15s; }
    .btn-primary { background: var(--color-primary); color: white; border-color: var(--color-primary); }
    .btn-primary:hover:not(:disabled) { background: var(--color-primary-dark); border-color: var(--color-primary-dark); }
    .btn-primary:disabled { opacity: 0.55; cursor: not-allowed; }
    .btn-ghost { background: transparent; color: var(--color-text); border-color: var(--color-border); }
    .btn-ghost:hover { background: var(--color-surface-alt); }
    .btn-sm { padding: 5px 12px; font-size: 0.82rem; }
    .form-group { margin-bottom: 16px; }
    .form-group label { display: block; font-size: 0.85rem; font-weight: 600; margin-bottom: 6px; }
    .form-input { width: 100%; padding: 8px 12px; border: 1px solid var(--color-border);
      border-radius: 4px; font-family: inherit; font-size: 0.875rem; outline: none; background: white; }
    .form-input:focus { border-color: var(--color-primary); box-shadow: 0 0 0 2px var(--color-primary-light); }
    textarea.form-input { resize: vertical; }
    .form-error { padding: 10px 12px; background: #fde7e9; color: #a4262c;
      border-radius: 4px; font-size: 0.85rem; }
    .required { color: var(--color-danger); }
    .panel-section { margin-bottom: 32px; }
    .panel-section h3 { font-size: 1rem; font-weight: 600; margin-bottom: 6px; }
    .panel-desc { font-size: 0.875rem; color: var(--color-text-muted); margin-bottom: 16px; line-height: 1.5; }
  `
  document.head.appendChild(shared)
}

// Re-export getCurrentUser for use by other modules
export { getCurrentUser }
