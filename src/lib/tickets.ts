export type TicketStatus = 'Open' | 'In Progress' | 'Resolved'
export type TicketPriority = 'Low' | 'Medium' | 'High' | 'Urgent'

export interface IncidentTicket {
  id: string
  ticketNumber: string
  title: string
  description: string
  category: string
  location: string
  priority: TicketPriority
  status: TicketStatus
  reporterName: string
  reporterEmail: string
  ownerUid: string
  assignedTo: string
  createdAt: Date
  updatedAt: Date
  resolvedAt?: Date
}

export const ticketStatuses: TicketStatus[] = [
  'Open',
  'In Progress',
  'Resolved',
]

export const ticketCategories = [
  'Hardware',
  'Software',
  'Network',
  'Access',
  'Email & collaboration',
  'Other',
]
