import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from '@azure/msal-browser'

const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const tenantId = import.meta.env.VITE_MICROSOFT_TENANT_ID

export const microsoftAuthConfigured = Boolean(clientId && tenantId)
const apiScope = import.meta.env.VITE_MICROSOFT_API_SCOPE?.trim() || ''
export const ticketsApiUrl = import.meta.env.VITE_TICKETS_API_URL?.trim() ||
  (import.meta.env.DEV ? 'http://localhost:3001' : '')

let client: PublicClientApplication | undefined
let initialization: Promise<AccountInfo | null> | undefined

function getClient(): PublicClientApplication {
  if (!microsoftAuthConfigured) {
    throw new Error(
      'Set VITE_MICROSOFT_CLIENT_ID and VITE_MICROSOFT_TENANT_ID.',
    )
  }
  client ??= new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: window.location.origin,
      postLogoutRedirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: 'sessionStorage',
    },
  })
  return client
}

export async function initializeMicrosoftAuth(): Promise<AccountInfo | null> {
  if (!initialization) {
    initialization = (async () => {
      const msal = getClient()
      await msal.initialize()
      const redirectResult = await msal.handleRedirectPromise()
      const account = redirectResult?.account ?? msal.getAllAccounts()[0] ?? null
      if (account) msal.setActiveAccount(account)
      return account
    })()
  }
  try {
    return await initialization
  } catch (error) {
    initialization = undefined
    throw error
  }
}

export async function signInWithMicrosoft(): Promise<AccountInfo> {
  const msal = getClient()
  await initializeMicrosoftAuth()
  const result: AuthenticationResult = await msal.loginPopup({
    scopes: ['openid', 'profile', 'email'],
    prompt: 'select_account',
  })
  if (!result.account) {
    throw new Error('Microsoft did not return an account.')
  }
  msal.setActiveAccount(result.account)
  return result.account
}

export async function signOutMicrosoft(account: AccountInfo): Promise<void> {
  const msal = getClient()
  await msal.logoutPopup({ account, postLogoutRedirectUri: window.location.origin })
  msal.setActiveAccount(null)
  initialization = Promise.resolve(null)
}

export async function getMicrosoftApiToken(): Promise<string> {
  if (!apiScope) {
    throw new Error(
      'Microsoft sign-in succeeded, but ticket access needs VITE_MICROSOFT_API_SCOPE. Create and grant the API delegated scope in Microsoft Entra, copy its exact full URI into .env.local, and restart the app.',
    )
  }
  const msal = getClient()
  await initializeMicrosoftAuth()
  const account = msal.getActiveAccount() ?? msal.getAllAccounts()[0]
  if (!account) throw new Error('Sign in with your organization Microsoft account first.')

  try {
    const result = await msal.acquireTokenSilent({ account, scopes: [apiScope] })
    return result.accessToken
  } catch (error) {
    if (isMissingApiScope(error)) {
      throw missingApiScopeError()
    }
    if (!(error instanceof InteractionRequiredAuthError)) throw error
    try {
      const result = await msal.acquireTokenPopup({ account, scopes: [apiScope] })
      return result.accessToken
    } catch (tokenError) {
      if (isMissingApiScope(tokenError)) throw missingApiScopeError()
      throw tokenError
    }
  }
}

function missingApiScopeError(): Error {
  return new Error(
    `Microsoft sign-in succeeded, but the ticket API permission is not configured in Microsoft Entra ID. This app requests "${apiScope}". In App registrations, open the API application that owns this scope, select API permissions, add its delegated permission to this SPA, and grant admin consent if required. Set VITE_MICROSOFT_API_SCOPE to the exact full scope URI shown by Entra and restart the app.`,
  )
}

function isMissingApiScope(error: unknown): boolean {
  const message =
    error instanceof Error
      ? error.message
      : typeof error === 'object' && error !== null && 'errorMessage' in error
        ? String(error.errorMessage)
        : ''
  return /AADSTS65005|AADSTS650053|scope .+ does not exist|scope .+ doesn't exist/i.test(
    message,
  )
}
