import { randomUUID } from 'node:crypto'
import express from 'express'
import { createClient } from '@supabase/supabase-js'
import { createRemoteJWKSet, errors, jwtVerify } from 'jose'

function requiredEnvironment(name) {
  const value = process.env[name]?.trim()
  if (!value) throw new Error(`${name} must be set in .env.local.`)
  return value
}

const tenantId = requiredEnvironment('MICROSOFT_TENANT_ID')
const spaClientId = requiredEnvironment('MICROSOFT_CLIENT_ID')
const apiAudience =
  process.env.MICROSOFT_API_AUDIENCE?.trim() || spaClientId
const requiredScope =
  process.env.MICROSOFT_API_SCOPE_NAME?.trim() || 'Notifications.Send'
const adminRole =
  process.env.MICROSOFT_ADMIN_ROLE?.trim() || 'IncidentDesk.Admin'
const adminEmails = new Set(
  (process.env.MICROSOFT_ADMIN_EMAILS || '')
    .split(',')
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean),
)
const supabaseUrl =
  process.env.SUPABASE_URL?.trim() || process.env.VITE_SUPABASE_URL?.trim()
const supabaseSecretKey = process.env.SUPABASE_SECRET_KEY?.trim()
if (!supabaseUrl) {
  throw new Error('SUPABASE_URL or VITE_SUPABASE_URL must be set in .env.local.')
}
if (!supabaseSecretKey) {
  throw new Error('SUPABASE_SECRET_KEY must be set in .env.local; a publishable key cannot write ticket data.')
}
if (supabaseSecretKey.startsWith('sb_publishable_')) {
  throw new Error('SUPABASE_SECRET_KEY must be a server-only secret key, not a publishable key.')
}
if (
  supabaseSecretKey === process.env.VITE_SUPABASE_PUBLISHABLE_KEY?.trim()
) {
  throw new Error('SUPABASE_SECRET_KEY cannot be the configured publishable key.')
}
if (!supabaseSecretKey.startsWith('sb_secret_')) {
  const payload = supabaseSecretKey.split('.')[1]
  let role
  try {
    role = payload ? JSON.parse(Buffer.from(payload, 'base64url').toString()).role : null
  } catch {
    role = null
  }
  if (role !== 'service_role') {
    throw new Error('SUPABASE_SECRET_KEY must be a Supabase secret key or legacy service_role key.')
  }
}
const port = Number.parseInt(process.env.API_PORT || '3001', 10)
const host = process.env.API_HOST?.trim() || '127.0.0.1'
const allowedOrigins = new Set(
  (process.env.APP_ORIGINS ||
    'http://localhost:5173,http://127.0.0.1:5173')
    .split(',')
    .map((origin) => origin.trim())
    .filter(Boolean),
)

if (allowedOrigins.has('*')) {
  throw new Error('APP_ORIGINS must list explicit origins; wildcard CORS is not allowed.')
}
if (!Number.isInteger(port) || port < 1 || port > 65535) {
  throw new Error('API_PORT must be a valid TCP port.')
}

const jwks = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`),
)
const app = express()
const supabase = createClient(supabaseUrl, supabaseSecretKey, {
  auth: {
    autoRefreshToken: false,
    detectSessionInUrl: false,
    persistSession: false,
  },
})

function ticketFromRow(row) {
  return {
    id: row.id,
    title: row.title,
    description: row.description,
    category: row.category,
    location: row.location,
    priority: row.priority,
    status: row.status,
    reporterName: row.reporter_name,
    reporterEmail: row.reporter_email,
    ownerUid: row.owner_uid,
    assignedTo: row.assigned_to,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    ...(row.resolved_at ? { resolvedAt: row.resolved_at } : {}),
  }
}

function ticketToRow(ticket) {
  return {
    id: ticket.id,
    title: ticket.title,
    description: ticket.description,
    category: ticket.category,
    location: ticket.location,
    priority: ticket.priority,
    status: ticket.status,
    reporter_name: ticket.reporterName,
    reporter_email: ticket.reporterEmail,
    owner_uid: ticket.ownerUid,
    assigned_to: ticket.assignedTo,
    created_at: ticket.createdAt,
    updated_at: ticket.updatedAt,
    ...(ticket.resolvedAt ? { resolved_at: ticket.resolvedAt } : {}),
  }
}

function throwIfSupabaseError(error) {
  if (error) {
    throw new Error(`Supabase request failed: ${error.message}`, {
      cause: error,
    })
  }
}

async function verifyTicketsTable() {
  const { error } = await supabase.from('tickets').select('id').limit(0)
  throwIfSupabaseError(error)
}

app.disable('x-powered-by')
app.use((request, response, next) => {
  response.setHeader('Cache-Control', 'no-store')
  response.setHeader('Vary', 'Origin')
  const origin = request.header('origin')

  if (origin && !allowedOrigins.has(origin)) {
    response.status(403).json({ error: 'This application origin is not allowed.' })
    return
  }
  if (origin) response.setHeader('Access-Control-Allow-Origin', origin)
  response.setHeader('Access-Control-Allow-Methods', 'GET, POST, PATCH, OPTIONS')
  response.setHeader(
    'Access-Control-Allow-Headers',
    'Authorization, Content-Type',
  )
  if (request.method === 'OPTIONS') {
    response.status(204).end()
    return
  }
  next()
})
app.use(express.json({ limit: '32kb' }))

app.get('/healthz', async (_request, response) => {
  await verifyTicketsTable()
  response.json({ status: 'ok' })
})

app.use(async (request, response, next) => {
  const match = request.header('authorization')?.match(/^Bearer\s+(\S+)$/i)
  if (!match) {
    response.status(401).json({ error: 'A Microsoft access token is required.' })
    return
  }

  try {
    const { payload } = await jwtVerify(match[1], jwks, {
      issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
      audience: apiAudience,
    })
    if (
      payload.tid !== tenantId ||
      typeof payload.oid !== 'string' ||
      payload.oid.length === 0 ||
      payload.azp !== spaClientId ||
      typeof payload.scp !== 'string' ||
      !payload.scp.split(/\s+/).includes(requiredScope)
    ) {
      response.status(401).json({
        error: 'A delegated Microsoft API token from the allowed organization is required.',
      })
      return
    }
    request.identity = payload
    next()
  } catch (error) {
    if (error instanceof errors.JOSEError) {
      response.status(401).json({ error: 'The Microsoft access token is invalid or expired.' })
      return
    }
    next(error)
  }
})

const hasAdminRole = (identity) => {
  if (Array.isArray(identity.roles) && identity.roles.includes(adminRole)) {
    return true
  }
  const emails = [identity.preferred_username, identity.email]
    .filter((email) => typeof email === 'string')
    .map((email) => email.trim().toLowerCase())
  return emails.some((email) => adminEmails.has(email))
}
const ticketStatuses = new Set(['Open', 'In Progress', 'Resolved'])
const ticketPriorities = new Set(['Low', 'Medium', 'High', 'Urgent'])
const ticketCategories = new Set([
  'Hardware',
  'Software',
  'Network',
  'Access',
  'Email & collaboration',
  'Other',
])
const stringField = (value) => (typeof value === 'string' ? value.trim() : '')
const asRecord = (value) =>
  typeof value === 'object' && value !== null && !Array.isArray(value)
    ? value
    : null

app.get('/tickets', async (request, response) => {
  const isAdmin = hasAdminRole(request.identity)
  let query = supabase
    .from('tickets')
    .select('*')
    .order('created_at', { ascending: false })
  if (!isAdmin) query = query.eq('owner_uid', request.identity.oid)
  const { data, error } = await query
  throwIfSupabaseError(error)
  response.json({
    tickets: (data ?? []).map(ticketFromRow),
    isAdmin,
  })
})

app.post('/tickets', async (request, response) => {
  const identity = request.identity
  const input = asRecord(request.body)
  if (!input) {
    response.status(400).json({ error: 'The incident details are invalid.' })
    return
  }
  const title = stringField(input.title)
  const description = stringField(input.description)
  const category = stringField(input.category)
  const location = stringField(input.location)
  const priority = stringField(input.priority)

  if (
    !title ||
    title.length < 4 ||
    title.length > 120 ||
    !description ||
    description.length < 10 ||
    description.length > 4000 ||
    !location ||
    location.length > 100 ||
    !ticketCategories.has(category) ||
    !ticketPriorities.has(priority)
  ) {
    response.status(400).json({ error: 'The incident details are invalid.' })
    return
  }

  const now = new Date().toISOString()
  const ticket = {
    id: randomUUID().replaceAll('-', ''),
    title,
    description,
    category,
    location,
    priority,
    status: 'Open',
    reporterName:
      identity.name ||
      identity.preferred_username ||
      identity.email ||
      identity.oid,
    reporterEmail:
      identity.preferred_username || identity.email || identity.oid,
    ownerUid: identity.oid,
    assignedTo: '',
    createdAt: now,
    updatedAt: now,
  }
  const { data, error } = await supabase
    .from('tickets')
    .insert(ticketToRow(ticket))
    .select('id')
    .single()
  throwIfSupabaseError(error)
  response.status(201).json({ id: data.id })
})

app.patch('/tickets/:ticketId', async (request, response) => {
  if (!hasAdminRole(request.identity)) {
    response.status(403).json({ error: 'An administrator role is required.' })
    return
  }

  const ticketId = request.params.ticketId
  const input = asRecord(request.body)
  if (
    typeof ticketId !== 'string' ||
    !/^(?:[a-f\d]{24}|[a-f\d]{32}|[a-f\d-]{36})$/i.test(ticketId) ||
    !input
  ) {
    response.status(400).json({ error: 'The ticket update is invalid.' })
    return
  }
  const status = stringField(input.status)
  const assignedTo = stringField(input.assignedTo)
  if (
    !ticketStatuses.has(status) ||
    assignedTo.length > 120 ||
    Object.keys(input).some((key) => !['status', 'assignedTo'].includes(key))
  ) {
    response.status(400).json({ error: 'The ticket update is invalid.' })
    return
  }

  const now = new Date().toISOString()
  const { data, error } = await supabase
    .from('tickets')
    .update({
      status,
      assigned_to: assignedTo,
      updated_at: now,
      resolved_at: status === 'Resolved' ? now : null,
    })
    .eq('id', ticketId)
    .select('id')
  throwIfSupabaseError(error)
  if (!data?.length) {
    response.status(404).json({ error: 'The incident was not found.' })
    return
  }
  response.status(204).end()
})

app.use((request, response) => {
  response.status(404).json({ error: `The API route ${request.path} was not found.` })
})

app.use((error, _request, response, _next) => {
  if (error?.type === 'entity.parse.failed') {
    response.status(400).json({ error: 'The request body is not valid JSON.' })
    return
  }
  if (error?.cause?.code === 'PGRST205') {
    response.status(503).json({
      error:
        'Supabase table public.tickets is unavailable. Run supabase/schema.sql in the SQL Editor for the project configured by SUPABASE_URL.',
    })
    return
  }
  console.error('Incident API request failed.', error)
  response.status(500).json({ error: 'The incident request could not be completed.' })
})

try {
  await verifyTicketsTable()
  app.listen(port, host, () => {
    console.log(`Incident API listening on http://localhost:${port}`)
    console.log('Connected to Supabase.')
  })
} catch (error) {
  console.error('Could not connect to Supabase or start the incident API.', error)
  process.exitCode = 1
}
