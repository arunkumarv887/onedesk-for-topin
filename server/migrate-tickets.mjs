import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { createClient } from '@supabase/supabase-js'

const supabaseUrl =
  process.env.SUPABASE_URL?.trim() || process.env.VITE_SUPABASE_URL?.trim()
const supabaseSecretKey = process.env.SUPABASE_SECRET_KEY?.trim()
if (!supabaseUrl || !supabaseSecretKey) {
  throw new Error(
    'Set SUPABASE_URL and the server-only SUPABASE_SECRET_KEY in .env.local.',
  )
}
if (supabaseSecretKey.startsWith('sb_publishable_')) {
  throw new Error('A Supabase publishable key cannot import server-side ticket data.')
}
if (
  supabaseSecretKey === process.env.VITE_SUPABASE_PUBLISHABLE_KEY?.trim()
) {
  throw new Error('The Supabase publishable key cannot import server-side ticket data.')
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
    throw new Error('Use a Supabase secret key or legacy service_role key for the import.')
  }
}

const ticketsFile = resolve(
  process.cwd(),
  process.env.TICKETS_FILE?.trim() || 'server-data/tickets.json',
)
const tickets = JSON.parse(await readFile(ticketsFile, 'utf8'))
if (!Array.isArray(tickets)) {
  throw new Error(`Ticket data file "${ticketsFile}" must contain a JSON array.`)
}

const rows = tickets.map((ticket) => ({
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
  assigned_to: ticket.assignedTo ?? '',
  created_at: ticket.createdAt,
  updated_at: ticket.updatedAt,
  ...(ticket.resolvedAt ? { resolved_at: ticket.resolvedAt } : {}),
}))

if (rows.some((ticket) => !ticket.id || !ticket.owner_uid)) {
  throw new Error('Every ticket must have an ID and owner before it can be imported.')
}

const supabase = createClient(supabaseUrl, supabaseSecretKey, {
  auth: {
    autoRefreshToken: false,
    detectSessionInUrl: false,
    persistSession: false,
  },
})
if (rows.length > 0) {
  const { error } = await supabase
    .from('tickets')
    .upsert(rows, { onConflict: 'id' })
  if (error) throw new Error(`Supabase ticket import failed: ${error.message}`)
}
console.log(`Imported ${rows.length} incident(s) into Supabase.`)
