create table if not exists public.tickets (
  id text primary key,
  title text not null check (char_length(title) between 4 and 120),
  description text not null check (char_length(description) between 10 and 4000),
  category text not null check (
    category in ('Hardware', 'Software', 'Network', 'Access', 'Email & collaboration', 'Other')
  ),
  location text not null check (char_length(location) between 1 and 100),
  priority text not null check (priority in ('Low', 'Medium', 'High', 'Urgent')),
  status text not null check (status in ('Open', 'In Progress', 'Resolved')),
  reporter_name text not null,
  reporter_email text not null,
  owner_uid text not null,
  assigned_to text not null default '',
  created_at timestamptz not null,
  updated_at timestamptz not null,
  resolved_at timestamptz
);

create index if not exists tickets_owner_created_at_idx
  on public.tickets (owner_uid, created_at desc);

create index if not exists tickets_created_at_idx
  on public.tickets (created_at desc);

alter table public.tickets enable row level security;
revoke all on public.tickets from public, anon, authenticated;
grant all on public.tickets to service_role;
