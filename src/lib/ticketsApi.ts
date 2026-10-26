import { getMicrosoftApiToken, ticketsApiUrl } from './microsoftAuth'
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
      'Set VITE_TICKETS_API_URL to the public HTTPS URL of the deployed Incident Desk API, then redeploy the frontend.',
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
        'The deployed frontend is configured with a local API address. Deploy the Incident Desk API, set VITE_TICKETS_API_URL in Vercel to its public HTTPS URL, and redeploy the frontend.',
        { cause: error },
      )
    }
    throw new Error(
      `Could not connect to the Incident Desk API at ${ticketsApiUrl}. Check the API terminal for Supabase setup errors, then start it with "npm run start:api".`,
      { cause: error },
    )
  }
  if (response.status === 404) {
    throw new Error(
      `No Incident Desk API route was found at ${ticketsApiUrl}. Check that this URL points to the Incident Desk API server.`,
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
  }
}

export async function createTicket(input: {
  title: string
  description: string
  category: string
  location: string
  priority: TicketPriority
}): Promise<void> {
  await request('/tickets', {
    method: 'POST',
    body: JSON.stringify(input),
  })
}

export async function updateTicket(
  id: string,
  input: { status: TicketStatus; assignedTo: string },
): Promise<void> {
  await request(`/tickets/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    body: JSON.stringify(input),
  })
}
