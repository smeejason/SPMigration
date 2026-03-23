import { signIn, getCurrentUser } from '../../auth/authService'
import { getRootSite } from '../../graph/graphClient'
import { setState, getState } from '../../state/store'

export function renderAuthPanel(container: HTMLElement): void {
  container.innerHTML = `
    <div class="auth-page">
      <div class="auth-card">
        <div class="auth-logo">
          <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
            <rect width="48" height="48" rx="8" fill="#0078d4"/>
            <path d="M12 24L24 12L36 24L24 36L12 24Z" fill="white" opacity="0.9"/>
            <path d="M24 12L36 24L24 36" fill="white" opacity="0.4"/>
          </svg>
        </div>
        <h1 class="auth-title">SP Migration Planner</h1>
        <p class="auth-subtitle">Sign in with your Microsoft 365 account to plan and orchestrate SharePoint migrations.</p>
        <button id="btn-signin" class="btn btn-primary btn-large">
          <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
            <path d="M10 0C4.48 0 0 4.48 0 10s4.48 10 10 10 10-4.48 10-10S15.52 0 10 0zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13v6l5 3-.75 1.23L8 12V5h1z"/>
          </svg>
          Sign in with Microsoft
        </button>
        <div id="auth-status" class="auth-status" style="display:none"></div>
      </div>
    </div>
  `
  injectAuthStyles()

  container.querySelector('#btn-signin')!.addEventListener('click', async () => {
    const btn = container.querySelector('#btn-signin') as HTMLButtonElement
    const status = container.querySelector('#auth-status') as HTMLElement
    btn.disabled = true
    btn.textContent = 'Signing in…'
    status.style.display = 'none'

    try {
      const user = await signIn()
      setState({ auth: { user, isAuthenticated: true } })

      // Verify Graph connectivity
      status.className = 'auth-status auth-status--info'
      status.textContent = 'Verifying Graph API connection…'
      status.style.display = 'block'
      await getRootSite()

      // Navigate to projects
      setState({ ui: { activeView: 'projects', loading: false, error: null } })
    } catch (err) {
      console.error('[Auth] Sign-in failed', err)
      status.className = 'auth-status auth-status--error'
      status.textContent = `Sign-in failed: ${(err as Error).message}`
      status.style.display = 'block'
      btn.disabled = false
      btn.innerHTML = `<svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
        <path d="M10 0C4.48 0 0 4.48 0 10s4.48 10 10 10 10-4.48 10-10S15.52 0 10 0zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13v6l5 3-.75 1.23L8 12V5h1z"/>
      </svg> Sign in with Microsoft`
    }
  })

  // If already authenticated, skip login
  if (getState().auth.isAuthenticated) {
    const user = getCurrentUser()
    if (user) {
      setState({ auth: { user, isAuthenticated: true }, ui: { activeView: 'projects', loading: false, error: null } })
    }
  }
}

function injectAuthStyles(): void {
  if (document.getElementById('auth-styles')) return
  const style = document.createElement('style')
  style.id = 'auth-styles'
  style.textContent = `
    .auth-page {
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
      padding: 24px;
    }
    .auth-card {
      background: white;
      border-radius: 12px;
      padding: 48px 40px;
      max-width: 420px;
      width: 100%;
      text-align: center;
      box-shadow: 0 20px 60px rgba(0,0,0,0.2);
    }
    .auth-logo { margin-bottom: 24px; }
    .auth-title { font-size: 1.5rem; font-weight: 600; color: #323130; margin-bottom: 12px; }
    .auth-subtitle { font-size: 0.9rem; color: #605e5c; margin-bottom: 32px; line-height: 1.5; }
    .btn { display: inline-flex; align-items: center; gap: 8px; padding: 10px 20px;
      border: none; border-radius: 4px; font-family: inherit; font-size: 0.9rem;
      cursor: pointer; transition: background 0.15s; }
    .btn-primary { background: #0078d4; color: white; }
    .btn-primary:hover:not(:disabled) { background: #005a9e; }
    .btn-primary:disabled { opacity: 0.6; cursor: not-allowed; }
    .btn-large { padding: 14px 28px; font-size: 1rem; width: 100%; justify-content: center; }
    .auth-status { margin-top: 16px; padding: 10px 14px; border-radius: 4px; font-size: 0.85rem; }
    .auth-status--info { background: #deecf9; color: #005a9e; }
    .auth-status--error { background: #fde7e9; color: #a4262c; }
  `
  document.head.appendChild(style)
}
