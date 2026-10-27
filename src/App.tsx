import {
  Activity,
  AlertCircle,
  ArrowDownToLine,
  Check,
  CheckCircle2,
  ClipboardList,
  Clock3,
  CircleDot,
  LayoutDashboard,
  LogOut,
  Moon,
  Plus,
  RefreshCw,
  Search,
  Settings,
  ShieldCheck,
  Sun,
  TicketCheck,
  TriangleAlert,
  X,
} from 'lucide-react'
import {
  type AccountInfo,
} from '@azure/msal-browser'
import companyLogo from './assets/companyLogo.png'
import {
  useCallback,
  useEffect,
  useMemo,
  useState,
  type FormEvent,
  type ReactNode,
} from 'react'
import './App.css'
import {
  initializeMicrosoftAuth,
  microsoftAuthConfigured,
  signInWithMicrosoft,
  signOutMicrosoft,
} from './lib/microsoftAuth'
import {
  createTicket,
  listAppUsers,
  listTickets,
  updateTicket,
  updateUserRole,
  type AppUserAccess,
  type AppUserRole,
} from './lib/ticketsApi'
import {
  ticketCategories,
  ticketStatuses,
  type IncidentTicket,
  type TicketPriority,
  type TicketStatus,
} from './lib/tickets'

type Page = 'dashboard' | 'all' | 'open' | 'progress' | 'resolved' | 'settings'
type TicketForm = {
  title: string
  description: string
  category: string
  location: string
  priority: TicketPriority
}

const emptyForm: TicketForm = {
  title: '',
  description: '',
  category: 'Hardware',
  location: '',
  priority: 'Medium',
}

const monthFormatter = new Intl.DateTimeFormat('en', { month: 'short' })
const dateFormatter = new Intl.DateTimeFormat('en', {
  month: 'short',
  day: 'numeric',
  year: 'numeric',
})

function initials(name: string): string {
  return name
    .split(/[\s.@]+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0].toUpperCase())
    .join('')
}

function dateInMonth(date: Date, month: string): boolean {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}` === month
}

function microsoftSignInError(error: unknown): string {
  const message = error instanceof Error ? error.message : ''
  if (/AADSTS65005|AADSTS650053|scope .+ does not exist|scope .+ doesn't exist/i.test(message)) {
    return 'Microsoft sign-in reached Entra ID, but its requested ticket API scope does not exist in the app registration. First sign in with Microsoft, then configure the API delegated scope under Expose an API and grant it under API permissions. The ticket API scope is requested separately from Microsoft sign-in.'
  }
  return message ? `Microsoft sign-in failed: ${message}` : 'Microsoft sign-in failed.'
}

function exportCsv(tickets: IncidentTicket[], month: string): void {
  const columns = [
    'Ticket ID',
    'Title',
    'Category',
    'Priority',
    'Status',
    'Reporter',
    'Reporter email',
    'Location',
    'Created',
  ]
  const rows = tickets.map((ticket) => [
    ticket.ticketNumber,
    ticket.title,
    ticket.category,
    ticket.priority,
    ticket.status,
    ticket.reporterName,
    ticket.reporterEmail,
    ticket.location,
    ticket.createdAt.toISOString(),
  ])
  const csv = [columns, ...rows]
    .map((row) =>
      row
        .map((cell) => `"${String(cell).replaceAll('"', '""')}"`)
        .join(','),
    )
    .join('\r\n')
  const file = new Blob([csv], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(file)
  const link = document.createElement('a')
  link.href = url
  link.download = `incidents-${month}.csv`
  link.click()
  URL.revokeObjectURL(url)
}

function StatusBadge({ status }: { status: TicketStatus }) {
  const className = status.toLowerCase().replace(' ', '-')
  return <span className={`badge ${className}`}>{status}</span>
}

function MicrosoftMark() {
  return (
    <svg className="login-microsoft-icon" viewBox="0 0 16 16" aria-hidden="true">
      <path fill="#f35325" d="M1 1h6.2v6.2H1z" />
      <path fill="#81bc06" d="M8.8 1H15v6.2H8.8z" />
      <path fill="#05a6f0" d="M1 8.8h6.2V15H1z" />
      <path fill="#ffba08" d="M8.8 8.8H15V15H8.8z" />
    </svg>
  )
}

function App() {
  const [user, setUser] = useState<AccountInfo | null>(null)
  const [isAdmin, setIsAdmin] = useState(false)
  const [isGlobalAdmin, setIsGlobalAdmin] = useState(false)
  const [isNightMode, setIsNightMode] = useState(
    () => localStorage.getItem('onedesk-theme') === 'dark',
  )
  const [authReady, setAuthReady] = useState(!microsoftAuthConfigured)
  const [tickets, setTickets] = useState<IncidentTicket[]>([])
  const [loadedAccountId, setLoadedAccountId] = useState<string | null>(null)
  const [page, setPage] = useState<Page>('dashboard')
  const [search, setSearch] = useState('')
  const [selectedTicket, setSelectedTicket] = useState<IncidentTicket | null>(null)
  const [showCreate, setShowCreate] = useState(false)
  const [form, setForm] = useState<TicketForm>(emptyForm)
  const [assignedDraft, setAssignedDraft] = useState('')
  const [selectedMonth, setSelectedMonth] = useState(() => {
    const now = new Date()
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
  })
  const [today] = useState(() => new Date())
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [dataError, setDataError] = useState('')
  const [statusDraft, setStatusDraft] = useState<TicketStatus>('Open')
  const [appUsers, setAppUsers] = useState<AppUserAccess[]>([])
  const [loadingAppUsers, setLoadingAppUsers] = useState(false)
  const [savingRoleFor, setSavingRoleFor] = useState<string | null>(null)
  const [settingsError, setSettingsError] = useState('')

  useEffect(() => {
    localStorage.setItem('onedesk-theme', isNightMode ? 'dark' : 'light')
  }, [isNightMode])

  useEffect(() => {
    if (!microsoftAuthConfigured) return
    let active = true
    void initializeMicrosoftAuth()
      .then((account) => {
        if (active) {
          setLoadedAccountId(null)
          setUser(account)
        }
      })
      .catch((authError: unknown) => {
        if (active) {
          setError(
            authError instanceof Error
              ? `Could not initialize Microsoft sign-in: ${authError.message}`
              : 'Could not initialize Microsoft sign-in.',
          )
        }
      })
      .finally(() => {
        if (active) setAuthReady(true)
      })
    return () => {
      active = false
    }
  }, [])

  useEffect(() => {
    if (user && !import.meta.env.VITE_MICROSOFT_API_SCOPE?.trim()) {
      setDataError(
        'Microsoft sign-in is active, but ticket access is not configured. Create and grant the API delegated scope in Microsoft Entra, set VITE_MICROSOFT_API_SCOPE to its exact full URI, then restart the app.',
      )
    }
  }, [user])

  const loadTickets = useCallback(async () => {
    if (!user) return
    try {
      const result = await listTickets()
      setTickets(result.tickets)
      setIsAdmin(result.isAdmin)
      setIsGlobalAdmin(result.isGlobalAdmin)
      setDataError('')
    } catch (loadError) {
      setDataError(
        loadError instanceof Error
          ? `Could not load incidents: ${loadError.message}`
          : 'Could not load incidents.',
      )
    } finally {
      setLoadedAccountId(user.homeAccountId)
    }
  }, [user])

  async function handleLoadAppUsers(): Promise<void> {
    setLoadingAppUsers(true)
    setSettingsError('')
    try {
      setAppUsers(await listAppUsers())
    } catch (loadError) {
      setSettingsError(
        loadError instanceof Error
          ? `Could not load user access: ${loadError.message}`
          : 'Could not load user access.',
      )
    } finally {
      setLoadingAppUsers(false)
    }
  }

  useEffect(() => {
    if (!user) return
    let active = true
    void Promise.resolve().then(() => {
      if (active) return loadTickets()
    })
    return () => {
      active = false
    }
  }, [user, loadTickets])

  const openCount = tickets.filter((ticket) => ticket.status === 'Open').length
  const progressCount = tickets.filter(
    (ticket) => ticket.status === 'In Progress',
  ).length
  const resolvedCount = tickets.filter(
    (ticket) => ticket.status === 'Resolved',
  ).length
  const resolvedToday = tickets.filter((ticket) => {
    const resolvedDate = ticket.resolvedAt
    return (
      ticket.status === 'Resolved' &&
      resolvedDate?.toDateString() === today.toDateString()
    )
  }).length

  const sixMonthTrend = useMemo(() => {
    return Array.from({ length: 6 }, (_, index) => {
      const date = new Date(
        today.getFullYear(),
        today.getMonth() - (5 - index),
        1,
      )
      const monthTickets = tickets.filter(
        (ticket) =>
          ticket.createdAt.getFullYear() === date.getFullYear() &&
          ticket.createdAt.getMonth() === date.getMonth(),
      ).length
      return { label: monthFormatter.format(date), count: monthTickets }
    })
  }, [tickets, today])

  const currentMax = Math.max(1, ...sixMonthTrend.map((month) => month.count))

  const visibleTickets = useMemo(() => {
    let result = [...tickets]
    if (page === 'open') result = result.filter((ticket) => ticket.status === 'Open')
    if (page === 'progress') {
      result = result.filter((ticket) => ticket.status === 'In Progress')
    }
    if (page === 'resolved') {
      result = result.filter((ticket) => ticket.status === 'Resolved')
    }
    const term = search.trim().toLowerCase()
    if (term) {
      result = result.filter((ticket) =>
        [
          ticket.title,
          ticket.description,
          ticket.category,
          ticket.reporterName,
          ticket.reporterEmail,
          ticket.ticketNumber,
        ].some((field) => field.toLowerCase().includes(term)),
      )
    }
    return result
  }, [tickets, page, search])

  const monthlyTickets = tickets.filter((ticket) =>
    dateInMonth(ticket.createdAt, selectedMonth),
  )
  const loadingTickets = Boolean(user) && loadedAccountId !== user?.homeAccountId

  async function handleSignIn(): Promise<void> {
    if (!microsoftAuthConfigured) {
      setError('Microsoft sign-in is not configured. Add the required Entra ID values to the environment.')
      return
    }
    setBusy(true)
    setError('')
    try {
      const account = await signInWithMicrosoft()
      setTickets([])
      setIsAdmin(false)
      setIsGlobalAdmin(false)
      setLoadedAccountId(null)
      setUser(account)
    } catch (signInError) {
      setError(microsoftSignInError(signInError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSignOut(): Promise<void> {
    if (!user) return
    setError('')
    try {
      await signOutMicrosoft(user)
      setUser(null)
      setTickets([])
      setIsAdmin(false)
      setIsGlobalAdmin(false)
      setLoadedAccountId(null)
      setPage('dashboard')
    } catch (signOutError) {
      setError(
        signOutError instanceof Error
          ? `Could not sign out: ${signOutError.message}`
          : 'Could not sign out.',
      )
    }
  }

  async function handleCreateTicket(event: FormEvent<HTMLFormElement>): Promise<void> {
    event.preventDefault()
    if (!user) {
      setError('You must be signed in before creating an incident.')
      return
    }
    setBusy(true)
    setError('')
    try {
      await createTicket(form)
      await loadTickets()
      setForm(emptyForm)
      setShowCreate(false)
      setPage('all')
    } catch (createError) {
      setError(
        createError instanceof Error
          ? `Could not create incident: ${createError.message}`
          : 'Could not create incident.',
      )
    } finally {
      setBusy(false)
    }
  }

  async function handleStatusUpdate(): Promise<void> {
    if (!selectedTicket || !isAdmin) return
    setBusy(true)
    setError('')
    try {
      await updateTicket(selectedTicket.id, {
        status: statusDraft,
        assignedTo: assignedDraft.trim(),
      })
      await loadTickets()
      setSelectedTicket(null)
    } catch (updateError) {
      setError(
        updateError instanceof Error
          ? `Could not update incident: ${updateError.message}`
          : 'Could not update incident.',
      )
    } finally {
      setBusy(false)
    }
  }

  async function handleAccessChange(
    appUser: AppUserAccess,
    role: AppUserRole,
  ): Promise<void> {
    setSavingRoleFor(appUser.userId)
    setSettingsError('')
    try {
      await updateUserRole(appUser.userId, role)
      setAppUsers((currentUsers) =>
        currentUsers.map((currentUser) =>
          currentUser.userId === appUser.userId
            ? { ...currentUser, role }
            : currentUser,
        ),
      )
    } catch (updateError) {
      setSettingsError(
        updateError instanceof Error
          ? `Could not update access for ${appUser.displayName}: ${updateError.message}`
          : `Could not update access for ${appUser.displayName}.`,
      )
    } finally {
      setSavingRoleFor(null)
    }
  }

  function openTicket(ticket: IncidentTicket): void {
    setSelectedTicket(ticket)
    setStatusDraft(ticket.status)
    setAssignedDraft(ticket.assignedTo)
  }

  if (!authReady) {
    return (
      <div className="login-screen">
        <div className="login-brand-panel">
          <img className="login-brand-logo" src={companyLogo} alt="" />
          <h1>Topin Technologies</h1>
        </div>
        <div className="login-form-panel">
          <div className="loading-state" role="status">Checking sign-in…</div>
        </div>
      </div>
    )
  }

  if (!user) {
    return (
      <main className="login-screen">
        <section className="login-brand-panel" aria-label="Topin Technologies">
          <img className="login-brand-logo" src={companyLogo} alt="" />
          <h1>Topin Technologies</h1>
          <p>IT support, made simple.</p>
        </section>
        <section className="login-form-panel">
          <div className="login-card" aria-labelledby="login-heading">
            <div className="login-eyebrow">ONEDESK WORKSPACE</div>
            <h1 id="login-heading">Sign in to OneDesk</h1>
            <p>Sign in with your organization’s Microsoft account to report and track IT incidents.</p>
            {error && <div className="alert" role="alert"><AlertCircle size={15} />{error}</div>}
            <button
              className="button"
              type="button"
              onClick={() => void handleSignIn()}
              disabled={!microsoftAuthConfigured || busy}
            >
              <MicrosoftMark />
              {busy ? 'Connecting…' : 'Continue with Microsoft'}
            </button>
            {!microsoftAuthConfigured && (
              <div className="config-notice" role="status">
                <ShieldCheck size={15} />
                <span>
                  Configure <code>VITE_MICROSOFT_CLIENT_ID</code>, <code>VITE_MICROSOFT_TENANT_ID</code>, and the Firestore project ID to enable organization-only sign-in.
                </span>
              </div>
            )}
            <div className="login-footer">Secure access for your organization</div>
          </div>
        </section>
      </main>
    )
  }

  const pageTitles: Record<Page, string> = {
    dashboard: 'Incident Dashboard',
    all: 'All Incidents',
    open: 'Open Incidents',
    progress: 'Work in Progress',
    resolved: 'Resolved Incidents',
    settings: 'User access',
  }
  const pageDescriptions: Record<Page, string> = {
    dashboard: 'A clear overview of your team’s reported incidents.',
    all: isAdmin ? 'Review and manage every incident in your organization.' : 'Review incidents you have reported.',
    open: 'Incidents waiting to be picked up.',
    progress: 'Incidents currently being worked on.',
    resolved: 'Completed incidents and their history.',
    settings: 'Manage who can administer OneDesk.',
  }

  return (
    <div className="app-shell" data-theme={isNightMode ? 'dark' : 'light'}>
      <aside className="sidebar">
        <div className="sidebar-brand">
          <Brand onClick={() => {
            setPage('dashboard')
            void loadTickets()
          }} />
        </div>
        <p className="nav-label">Incident management</p>
        <nav className="nav-list" aria-label="Incident navigation">
          <NavButton page="dashboard" activePage={page} onSelect={setPage} icon={<LayoutDashboard size={17} />} label="Dashboard" />
          <NavButton page="all" activePage={page} onSelect={setPage} icon={<ClipboardList size={17} />} label="All incidents" count={tickets.length} />
          <NavButton page="open" activePage={page} onSelect={setPage} icon={<CircleDot size={17} />} label="Open" count={openCount} />
          <NavButton page="progress" activePage={page} onSelect={setPage} icon={<Activity size={17} />} label="Work in progress" count={progressCount} />
          <NavButton page="resolved" activePage={page} onSelect={setPage} icon={<TicketCheck size={17} />} label="Resolved" count={resolvedCount} />
        </nav>
        <div className="sidebar-footer">
          {isGlobalAdmin && (
            <>
              <div className="sidebar-footer-divider" aria-hidden="true" />
              <NavButton
                page="settings"
                activePage={page}
                onSelect={(selectedPage) => {
                  setPage(selectedPage)
                  void handleLoadAppUsers()
                }}
                icon={<Settings size={17} />}
                label="Settings"
              />
            </>
          )}
          <div className="sidebar-bottom">
            <strong>Topin Technologies</strong>
            <p>{isAdmin ? 'Administrator access' : 'Your reported incidents'}<br />Securely powered by Microsoft</p>
          </div>
        </div>
      </aside>

      <div className="main-area">
        <header className="topbar">
          <div className="topbar-context">Workspace <span aria-hidden="true">/</span> <strong>{pageTitles[page]}</strong></div>
          <div className="user-area">
            <label className="theme-toggle topbar-theme-toggle">
              <span>Dark mode</span>
              <input
                type="checkbox"
                role="switch"
                aria-label="Dark mode"
                checked={isNightMode}
                onChange={(event) => setIsNightMode(event.target.checked)}
              />
              {isNightMode ? <Moon size={16} aria-hidden="true" /> : <Sun size={16} aria-hidden="true" />}
            </label>
            <div className="user-details">
              <strong>{user.name || user.username}</strong>
              <span>{isAdmin ? 'IT administrator' : 'Employee'}</span>
            </div>
            <div className="avatar" aria-label={`Signed in as ${user.name || user.username}`}>
              {initials(user.name || user.username || 'U')}
            </div>
            <button className="icon-button" type="button" aria-label="Sign out" title="Sign out" onClick={() => void handleSignOut()}>
              <LogOut size={16} />
            </button>
          </div>
        </header>

        <main className="page-content">
          {error && <div className="alert" role="alert"><AlertCircle size={15} />{error}</div>}
          {dataError && <div className="alert" role="alert"><AlertCircle size={15} />{dataError}</div>}
          <div className="page-heading">
            <div>
              <p className="eyebrow">INCIDENT MANAGEMENT</p>
              <h1>{pageTitles[page]}</h1>
              <p>{pageDescriptions[page]}</p>
            </div>
            {page !== 'settings' && (
              <div className="heading-actions">
                <input
                  className="month-input"
                  aria-label="Report month"
                  type="month"
                  value={selectedMonth}
                  onChange={(event) => setSelectedMonth(event.target.value)}
                />
                <button
                  className="button secondary"
                  type="button"
                  onClick={() => exportCsv(monthlyTickets, selectedMonth)}
                  title={`Download ${monthlyTickets.length} incidents for ${selectedMonth}`}
                >
                  <ArrowDownToLine size={15} />
                  Monthly report
                </button>
                <button className="button" type="button" onClick={() => setShowCreate(true)}>
                  <Plus size={16} />
                  Report incident
                </button>
              </div>
            )}
          </div>

          {page === 'dashboard' && (
            <>
              <section className="metric-grid" aria-label="Incident statistics">
                <MetricCard page="open" onSelect={setPage} title="Open incidents" value={openCount} note="Awaiting action" icon={<TriangleAlert size={15} />} color="amber" />
                <MetricCard page="progress" onSelect={setPage} title="In progress" value={progressCount} note="Being worked on" icon={<Clock3 size={15} />} color="blue" />
                <MetricCard page="resolved" onSelect={setPage} title="Resolved today" value={resolvedToday} note="Completed today" icon={<CheckCircle2 size={15} />} color="green" />
                <MetricCard page="all" onSelect={setPage} title="Total incidents" value={tickets.length} note={isAdmin ? 'Across your organization' : 'Reported by you'} icon={<ClipboardList size={15} />} color="violet" />
              </section>

              <section className="overview-grid">
                <div className="panel">
                  <div className="panel-heading">
                    <div>
                      <h2>Incident volume</h2>
                      <p>New incidents reported over the last six months.</p>
                    </div>
                    <span className="chart-legend"><i /> Incidents</span>
                  </div>
                  <div className="chart" role="img" aria-label={`Monthly incident counts: ${sixMonthTrend.map(({ label, count }) => `${label} ${count}`).join(', ')}`}>
                    {sixMonthTrend.map((month) => (
                      <div className="chart-column" key={month.label}>
                        <strong>{month.count}</strong>
                        <div className="chart-bar-wrap">
                          <div className="chart-bar" style={{ height: `${Math.max(4, (month.count / currentMax) * 100)}%` }} />
                        </div>
                        <span>{month.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="panel">
                  <div className="panel-heading">
                    <div>
                      <h2>Incidents by status</h2>
                      <p>Current status breakdown.</p>
                    </div>
                  </div>
                  <div className="status-summary">
                    <StatusSummary label="Open" count={openCount} total={tickets.length} color="open" />
                    <StatusSummary label="In progress" count={progressCount} total={tickets.length} color="progress" />
                    <StatusSummary label="Resolved" count={resolvedCount} total={tickets.length} color="resolved" />
                  </div>
                </div>
              </section>
              <TicketTable
                title="Recent incidents"
                tickets={tickets.slice(0, 5)}
                loading={loadingTickets}
                onSelect={openTicket}
                search={search}
                setSearch={setSearch}
                onViewAll={() => setPage('all')}
              />
            </>
          )}

          {page === 'settings' && (
            <>
              <section className="panel access-panel" aria-labelledby="access-heading">
                <div className="panel-heading access-panel-heading">
                  <div>
                    <h2 id="access-heading">OneDesk user access</h2>
                    <p>Choose who can view and manage incidents across the organization. This list is synced from your Microsoft Entra directory.</p>
                  </div>
                  <button
                    className="button secondary refresh-users-button"
                    type="button"
                    onClick={() => void handleLoadAppUsers()}
                    disabled={loadingAppUsers}
                  >
                    <RefreshCw size={14} className={loadingAppUsers ? 'spinning' : undefined} />
                    Refresh users
                  </button>
                </div>
                {settingsError && <div className="alert" role="alert"><AlertCircle size={15} />{settingsError}</div>}
                {settingsError ? null : loadingAppUsers ? (
                  <div className="empty-state" role="status">Loading users…</div>
                ) : appUsers.length === 0 ? (
                  <div className="empty-state">No users found yet. A user must successfully access OneDesk or have reported an incident to appear here.</div>
                ) : (
                  <div className="access-list">
                    {appUsers.map((appUser) => {
                      const locked = appUser.isGlobalAdmin || appUser.isEntraAdmin
                      return (
                        <div className="access-row" key={appUser.userId}>
                          <div className="access-identity">
                            <strong>{appUser.displayName}</strong>
                            <span>{appUser.email}</span>
                            <span className="access-user-id">ID: {appUser.userId}</span>
                          </div>
                          {locked ? (
                            <div className="access-protected">
                              <span className="access-role-badge">Admin</span>
                              <span>{appUser.isGlobalAdmin ? 'Global admin' : 'Managed by Entra ID'}</span>
                            </div>
                          ) : (
                            <fieldset
                              className="role-switch"
                              aria-label={`Access for ${appUser.displayName}`}
                              disabled={savingRoleFor === appUser.userId}
                            >
                              <label>
                                <input
                                  type="radio"
                                  name={`role-${appUser.userId}`}
                                  value="user"
                                  checked={appUser.role === 'user'}
                                  onChange={() => void handleAccessChange(appUser, 'user')}
                                />
                                User
                              </label>
                              <label>
                                <input
                                  type="radio"
                                  name={`role-${appUser.userId}`}
                                  value="admin"
                                  checked={appUser.role === 'admin'}
                                  onChange={() => void handleAccessChange(appUser, 'admin')}
                                />
                                Admin
                              </label>
                            </fieldset>
                          )}
                        </div>
                      )
                    })}
                  </div>
                )}
              </section>
            </>
          )}

          {page !== 'dashboard' && page !== 'settings' && (
            <TicketTable
              title={pageTitles[page]}
              tickets={visibleTickets}
              loading={loadingTickets}
              onSelect={openTicket}
              search={search}
              setSearch={setSearch}
              month={selectedMonth}
              onMonthChange={setSelectedMonth}
              onExport={() => exportCsv(monthlyTickets, selectedMonth)}
            />
          )}
        </main>
      </div>

      {showCreate && (
        <Modal title="Report a new incident" subtitle="Tell us what happened and we’ll get it to the right team." onClose={() => setShowCreate(false)}>
          <form onSubmit={(event) => void handleCreateTicket(event)}>
            <div className="modal-body">
              {error && <div className="alert" role="alert"><AlertCircle size={15} />{error}</div>}
              <div className="form-grid">
                <Field label="Incident title" className="full" id="incident-title">
                  <input
                    id="incident-title"
                    required
                    minLength={4}
                    maxLength={120}
                    placeholder="A short summary of the issue"
                    value={form.title}
                    onChange={(event) => setForm({ ...form, title: event.target.value })}
                  />
                </Field>
                <Field label="Category" id="incident-category">
                  <select id="incident-category" value={form.category} onChange={(event) => setForm({ ...form, category: event.target.value })}>
                    {ticketCategories.map((category) => <option key={category}>{category}</option>)}
                  </select>
                </Field>
                <Field label="Priority" id="incident-priority">
                  <select id="incident-priority" value={form.priority} onChange={(event) => setForm({ ...form, priority: event.target.value as TicketPriority })}>
                    {(['Low', 'Medium', 'High', 'Urgent'] as const).map((priority) => <option key={priority}>{priority}</option>)}
                  </select>
                </Field>
                <Field label="Location" className="full" id="incident-location">
                  <input
                    id="incident-location"
                    required
                    maxLength={100}
                    placeholder="Office, floor, or remote"
                    value={form.location}
                    onChange={(event) => setForm({ ...form, location: event.target.value })}
                  />
                </Field>
                <Field label="Description" className="full" id="incident-description">
                  <textarea
                    id="incident-description"
                    required
                    minLength={10}
                    maxLength={4000}
                    placeholder="Describe the issue, what you were doing, and any steps already tried."
                    value={form.description}
                    onChange={(event) => setForm({ ...form, description: event.target.value })}
                  />
                  <span className="field-hint">Include enough detail to help the IT team investigate.</span>
                </Field>
              </div>
            </div>
            <div className="modal-footer">
              <button className="button secondary" type="button" onClick={() => setShowCreate(false)}>Cancel</button>
              <button className="button" type="submit" disabled={busy}>
                <Plus size={14} /> {busy ? 'Submitting…' : 'Submit incident'}
              </button>
            </div>
          </form>
        </Modal>
      )}

      {selectedTicket && (
        <Modal
          title={selectedTicket.title}
          subtitle={selectedTicket.ticketNumber}
          onClose={() => setSelectedTicket(null)}
        >
          <div className="modal-body">
            <div className="detail-meta">
              <StatusBadge status={selectedTicket.status} />
              <span className={`priority ${selectedTicket.priority.toLowerCase()}`}>{selectedTicket.priority} priority</span>
              <span>Reported {dateFormatter.format(selectedTicket.createdAt)}</span>
            </div>
            <div className="detail-two-col">
              <DetailValue label="Category" value={selectedTicket.category} />
              <DetailValue label="Location" value={selectedTicket.location} />
              <DetailValue label="Reported by" value={selectedTicket.reporterName} />
              <DetailValue label="Assigned to" value={selectedTicket.assignedTo || 'Not assigned'} />
            </div>
            <div className="detail-section">
              <h3>Description</h3>
              <p>{selectedTicket.description}</p>
            </div>
            {isAdmin && (
              <>
                <div className="field">
                  <label htmlFor="ticket-status">Update status</label>
                  <select
                    id="ticket-status"
                    className="inline-status"
                    value={statusDraft}
                    onChange={(event) => setStatusDraft(event.target.value as TicketStatus)}
                  >
                    {ticketStatuses.map((status) => <option key={status}>{status}</option>)}
                  </select>
                </div>
                <div className="field">
                  <label htmlFor="ticket-assignee">Assigned to</label>
                  <input
                    id="ticket-assignee"
                    value={assignedDraft}
                    maxLength={120}
                    placeholder="Name or email (optional)"
                    onChange={(event) => setAssignedDraft(event.target.value)}
                  />
                </div>
              </>
            )}
            {error && <div className="alert" role="alert"><AlertCircle size={15} />{error}</div>}
          </div>
          <div className="modal-footer">
            <button className="button secondary" type="button" onClick={() => setSelectedTicket(null)}>Close</button>
            {isAdmin && (
              <button className="button" type="button" onClick={() => void handleStatusUpdate()} disabled={busy || (statusDraft === selectedTicket.status && assignedDraft === selectedTicket.assignedTo)}>
                <Check size={14} /> {busy ? 'Saving…' : 'Save status'}
              </button>
            )}
          </div>
        </Modal>
      )}
    </div>
  )
}

function Brand({ onClick }: { onClick: () => void }) {
  return (
    <button className="brand" type="button" onClick={onClick} aria-label="Go to dashboard">
      <img className="brand-logo" src={companyLogo} alt="" />
      <span>Topin Technologies</span>
    </button>
  )
}

function NavButton({
  page,
  activePage,
  onSelect,
  icon,
  label,
  count,
}: {
  page: Page
  activePage: Page
  onSelect: (page: Page) => void
  icon: ReactNode
  label: string
  count?: number
}) {
  return (
    <button className={`nav-item ${page === activePage ? 'active' : ''}`} type="button" onClick={() => onSelect(page)}>
      {icon}
      <span>{label}</span>
      {count !== undefined && <span className="nav-count">{count}</span>}
    </button>
  )
}

function MetricCard({
  page,
  onSelect,
  title,
  value,
  note,
  icon,
  color,
}: {
  page: Page
  onSelect: (page: Page) => void
  title: string
  value: number
  note: string
  icon: ReactNode
  color: string
}) {
  return (
    <button
      className="metric-card"
      type="button"
      onClick={() => onSelect(page)}
      aria-label={`${title}: ${value}. Open ${page === 'all' ? 'all incidents' : `${page} incidents`}`}
    >
      <div className="metric-top">
        <span>{title}</span>
        <span className={`metric-icon ${color}`}>{icon}</span>
      </div>
      <div className="metric-value">
        <strong>{value}</strong>
        <span>{note}</span>
      </div>
    </button>
  )
}

function StatusSummary({
  label,
  count,
  total,
  color,
}: {
  label: string
  count: number
  total: number
  color: string
}) {
  const width = total === 0 ? 0 : (count / total) * 100
  return (
    <div className="status-row">
      <span>{label}</span>
      <div className="progress-track" aria-label={`${label}: ${count}`}>
        <div className={`progress-fill ${color}`} style={{ width: `${width}%` }} />
      </div>
      <strong>{count}</strong>
    </div>
  )
}

function TicketTable({
  title,
  tickets,
  loading,
  onSelect,
  search,
  setSearch,
  onViewAll,
}: {
  title: string
  tickets: IncidentTicket[]
  loading: boolean
  onSelect: (ticket: IncidentTicket) => void
  search: string
  setSearch: (search: string) => void
  month?: string
  onMonthChange?: (month: string) => void
  onExport?: () => void
  onViewAll?: () => void
}) {
  return (
    <section className="panel table-panel">
      <div className="table-toolbar">
        <div>
          <h2>{title}</h2>
          <p>{tickets.length} {tickets.length === 1 ? 'incident' : 'incidents'}</p>
        </div>
        <div className="table-controls">
          <label className="search-box">
            <Search size={14} aria-hidden="true" />
            <input
              aria-label="Search incidents"
              placeholder="Search incidents"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
            />
          </label>
          {onViewAll && (
            <button className="button secondary" type="button" onClick={onViewAll}>
              View all
            </button>
          )}
        </div>
      </div>
      {loading ? (
        <div className="loading-state" role="status"><p>Loading incidents…</p></div>
      ) : tickets.length === 0 ? (
        <div className="empty-state">
          <ClipboardList size={23} />
          <strong>No incidents to show</strong>
          <p>New incidents will appear here once reported.</p>
        </div>
      ) : (
        <div className="table-wrap">
          <table className="ticket-table">
            <thead>
              <tr>
                <th>Ticket</th>
                <th>Status</th>
                <th>Priority</th>
                <th>Category</th>
                <th>Reported by</th>
                <th>Created</th>
              </tr>
            </thead>
            <tbody>
              {tickets.map((ticket) => (
                <tr key={ticket.id} onClick={() => onSelect(ticket)} tabIndex={0} onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault()
                    onSelect(ticket)
                  }
                }}>
                  <td>
                    <div className="ticket-title">
                      <strong>{ticket.title}</strong>
                      <span className="ticket-id">{ticket.ticketNumber}</span>
                    </div>
                  </td>
                  <td><StatusBadge status={ticket.status} /></td>
                  <td><span className={`priority ${ticket.priority.toLowerCase()}`}>{ticket.priority}</span></td>
                  <td>{ticket.category}</td>
                  <td>{ticket.reporterName}</td>
                  <td>{dateFormatter.format(ticket.createdAt)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  )
}

function Field({
  label,
  className = '',
  id,
  children,
}: {
  label: string
  className?: string
  id?: string
  children: ReactNode
}) {
  return (
    <div className={`field ${className}`}>
      <label htmlFor={id}>{label}</label>
      {children}
    </div>
  )
}

function DetailValue({ label, value }: { label: string; value: string }) {
  return (
    <div className="detail-section">
      <h3>{label}</h3>
      <p>{value}</p>
    </div>
  )
}

function Modal({
  title,
  subtitle,
  onClose,
  children,
}: {
  title: string
  subtitle: string
  onClose: () => void
  children: ReactNode
}) {
  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', onKeyDown)
    return () => document.removeEventListener('keydown', onKeyDown)
  }, [onClose])

  return (
    <div className="modal-backdrop" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose()
    }}>
      <section className="modal" role="dialog" aria-modal="true" aria-label={title}>
        <header className="modal-header">
          <div>
            <h2>{title}</h2>
            <p>{subtitle}</p>
          </div>
          <button className="icon-button" type="button" aria-label="Close dialog" onClick={onClose}><X size={17} /></button>
        </header>
        {children}
      </section>
    </div>
  )
}

export default App
