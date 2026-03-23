import { initAuth, isAuthenticated, getCurrentUser } from './auth/authService'
import { setState } from './state/store'
import { mountApp } from './ui/app'

async function bootstrap(): Promise<void> {
  const root = document.querySelector<HTMLDivElement>('#app')!

  // Show a loading state while MSAL initialises
  root.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:Segoe UI,sans-serif;color:#605e5c;">
      Loading…
    </div>
  `

  try {
    await initAuth()

    if (isAuthenticated()) {
      const user = getCurrentUser()
      if (user) {
        setState({
          auth: { user, isAuthenticated: true },
          ui: { activeView: 'projects', loading: false, error: null },
        })
      }
    }
  } catch (err) {
    console.error('[Bootstrap] Auth init failed', err)
    // Continue anyway — the auth panel will handle sign-in
  }

  mountApp(root)
}

bootstrap().catch(console.error)
