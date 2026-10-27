import {
  getMicrosoftApiToken,
  getMicrosoftGraphToken,
  ticketsApiUrl,
} from './microsoftAuth'
import type { IncidentTicket, TicketPriority, TicketStatus } from './tickets'

type TicketApiRecord = Omit<
  IncidentTicket,
  'createdAt' | 'updatedAt' | 'resolvedAt'
> & {
  createdAt: string
  updatedAt: string
  resolvedAt?: string
}

type TicketListResponse = {
  tickets: TicketApiRecord[]
  isAdmin: boolean
  isGlobalAdmin: boolean
}

export type AppUserRole = 'admin' | 'user'

export type AppUserAccess = {
  userId: string
  email: string
  displayName: string
  role: AppUserRole
  isGlobalAdmin: boolean
  isEntraAdmin: boolean
  lastSeenAt: string
}

type DirectoryUser = {
  id: string
  displayName?: string | null
  mail?: string | null
  userPrincipalName?: string | null
}

type GraphUsersResponse = {
  value: DirectoryUser[]
  '@odata.nextLink'?: string
}

function isLocalApiUrl(url: string): boolean {
  try {
    const hostname = new URL(url).hostname
    return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]'
  } catch {
    return false
  }
}

async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  if (!ticketsApiUrl) {
    throw new Error(
      'Set VITE_TICKETS_API_URL to the public HTTPS URL of the deployed Onedesk API, then redeploy the frontend.',
    )
  }
  const token = await getMicrosoftApiToken()
  let response: Response
  try {
    response = await fetch(`${ticketsApiUrl}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(init.body ? { 'Content-Type': 'application/json' } : {}),
        ...init.headers,
      },
    })
  } catch (error) {
    if (!(error instanceof TypeError)) throw error
    if (import.meta.env.PROD && isLocalApiUrl(ticketsApiUrl)) {
      throw new Error(
        'The deployed frontend is configured with a local API address. Deploy the Onedesk API, set VITE_TICKETS_API_URL in Vercel to its public HTTPS URL, and redeploy the frontend.',
        { cause: error },
      )
    }
    throw new Error(
      `Could not connect to the Onedesk API at ${ticketsApiUrl}. Check the API terminal for Supabase setup errors, then start it with "npm run start:api".`,
      { cause: error },
    )
  }
  if (response.status === 404) {
    throw new Error(
      `No Onedesk API route was found at ${ticketsApiUrl}. Check that this URL points to the Onedesk API server.`,
    )
  }
  if (!response.ok) {
    const result: unknown = await response.json().catch(() => null)
    const message =
      typeof result === 'object' &&
      result !== null &&
      'error' in result &&
      typeof result.error === 'string'
        ? result.error
        : `Incident API request failed (${response.status}).`
    throw new Error(message)
  }
  if (response.status === 204) return undefined as T
  return (await response.json()) as T
}

export async function listTickets(): Promise<{
  tickets: IncidentTicket[]
  isAdmin: boolean
  isGlobalAdmin: boolean
}> {
  const result = await request<TicketListResponse>('/tickets')
  return {
    tickets: result.tickets.map((ticket) => {
      const { createdAt, updatedAt, resolvedAt, ...fields } = ticket
      return {
        ...fields,
        createdAt: new Date(createdAt),
        updatedAt: new Date(updatedAt),
        ...(resolvedAt ? { resolvedAt: new Date(resolvedAt) } : {}),
      }
    }),
    isAdmin: result.isAdmin,
    isGlobalAdmin: result.isGlobalAdmin,
  }
}

export async function listAppUsers(): Promise<AppUserAccess[]> {
  const graphToken = await getMicrosoftGraphToken()
  let nextUrl: string | undefined =
    'https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=100'

  while (nextUrl) {
    const graphUrl = new URL(nextUrl)
    if (
      graphUrl.protocol !== 'https:' ||
      graphUrl.hostname !== 'graph.microsoft.com'
    ) {
      throw new Error('Microsoft Graph returned an unexpected paging URL.')
    }

    const graphResponse = await fetch(graphUrl, {
      headers: { Authorization: `Bearer ${graphToken}` },
    })
    if (!graphResponse.ok) {
      const errorResponse: unknown = await graphResponse.json().catch(() => null)
      const graphMessage =
        typeof errorResponse === 'object' &&
        errorResponse !== null &&
        'error' in errorResponse &&
        typeof errorResponse.error === 'object' &&
        errorResponse.error !== null &&
        'message' in errorResponse.error &&
        typeof errorResponse.error.message === 'string'
          ? errorResponse.error.message
          : `Microsoft Graph request failed (${graphResponse.status}).`
      if (graphResponse.status === 403) {
        throw new Error(
          `Microsoft Graph denied the directory user request. Add the delegated User.ReadBasic.All permission to the OneDesk app registration, grant tenant admin consent, and retry. Details: ${graphMessage}`,
        )
      }
      throw new Error(graphMessage)
    }

    const result = (await graphResponse.json()) as GraphUsersResponse
    const users = result.value.map((directoryUser) => ({
      userId: directoryUser.id,
      email:
        directoryUser.mail ||
        directoryUser.userPrincipalName ||
        directoryUser.id,
      displayName: directoryUser.displayName || directoryUser.userPrincipalName || directoryUser.id,
    }))
    await request('/admin/users/sync', {
      method: 'POST',
      body: JSON.stringify({ users }),
    })
    nextUrl = result['@odata.nextLink']
  }

  const result = await request<{ users: AppUserAccess[] }>('/admin/users')
  return result.users
}

export async function updateUserRole(
  userId: string,
  role: AppUserRole,
): Promise<void> {
  await request(`/admin/users/${encodeURIComponent(userId)}/role`, {
    method: 'PATCH',
    body: JSON.stringify({ role }),
  })
}

export async function createTicket(input: {
  title: string
  description: string
  category: string
  location: string
  priority: TicketPriority
}): Promise<{ notificationWarning?: string }> {
  return request('/tickets', {
    method: 'POST',
    body: JSON.stringify(input),
  })
}

export async function updateTicket(
  id: string,
  input: { status: TicketStatus; assignedTo: string },
): Promise<{ notificationWarning?: string }> {
  return request(`/tickets/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    body: JSON.stringify(input),
  })
}
