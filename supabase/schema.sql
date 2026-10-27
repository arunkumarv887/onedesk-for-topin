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

alter table public.tickets
  add column if not exists ticket_number text;

with current_max as (
  select
    (created_at at time zone 'UTC')::date as incident_date,
    max((regexp_match(ticket_number, '^INC-\d{4}-\d{2}-\d{2}-(\d+)$'))[1]::bigint) as last_number
  from public.tickets
  where ticket_number ~ '^INC-\d{4}-\d{2}-\d{2}-\d+$'
  group by (created_at at time zone 'UTC')::date
), unnumbered_tickets as (
  select
    id,
    incident_date,
    coalesce(current_max.last_number, 0) + row_number() over (
      partition by unnumbered.incident_date
      order by created_at, id
    ) as incident_number
  from (
    select
      id,
      created_at,
      (created_at at time zone 'UTC')::date as incident_date
    from public.tickets
    where ticket_number is null
  ) as unnumbered
  left join current_max using (incident_date)
)
update public.tickets as ticket
set ticket_number = format(
  'INC-%s-%s',
  to_char(unnumbered_tickets.incident_date, 'YYYY-MM-DD'),
  unnumbered_tickets.incident_number
)
from unnumbered_tickets
where ticket.id = unnumbered_tickets.id;

create table if not exists public.incident_counters (
  incident_date date primary key,
  last_number bigint not null check (last_number > 0)
);

insert into public.incident_counters (incident_date, last_number)
select
  (created_at at time zone 'UTC')::date,
  max((regexp_match(ticket_number, '^INC-\d{4}-\d{2}-\d{2}-(\d+)$'))[1]::bigint)
from public.tickets
where ticket_number ~ '^INC-\d{4}-\d{2}-\d{2}-\d+$'
group by (created_at at time zone 'UTC')::date
on conflict (incident_date) do update
set last_number = greatest(
  public.incident_counters.last_number,
  excluded.last_number
);

create index if not exists tickets_owner_created_at_idx
  on public.tickets (owner_uid, created_at desc);

create index if not exists tickets_created_at_idx
  on public.tickets (created_at desc);

create unique index if not exists tickets_ticket_number_key
  on public.tickets (ticket_number);

alter table public.tickets
  alter column ticket_number set not null;

create or replace function public.next_incident_number(p_incident_date date)
returns bigint
language sql
volatile
security definer
set search_path = public
as $$
  insert into public.incident_counters (incident_date, last_number)
  values (p_incident_date, 1)
  on conflict (incident_date) do update
    set last_number = public.incident_counters.last_number + 1
  returning last_number;
$$;

create or replace function public.assign_ticket_number()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  incident_date date;
  incident_number bigint;
begin
  if new.ticket_number is null then
    incident_date := (new.created_at at time zone 'UTC')::date;
    incident_number := public.next_incident_number(incident_date);
    new.ticket_number := format(
      'INC-%s-%s',
      to_char(incident_date, 'YYYY-MM-DD'),
      incident_number
    );
  end if;
  return new;
end;
$$;

drop trigger if exists tickets_assign_ticket_number on public.tickets;
create trigger tickets_assign_ticket_number
before insert on public.tickets
for each row execute function public.assign_ticket_number();

create table if not exists public.app_users (
  object_id text primary key,
  email text not null,
  display_name text not null,
  is_entra_admin boolean not null default false,
  last_seen_at timestamptz not null default now()
);

create table if not exists public.user_roles (
  object_id text primary key references public.app_users (object_id) on delete cascade,
  role text not null check (role in ('admin', 'user')),
  updated_by text not null,
  updated_at timestamptz not null default now()
);

create index if not exists app_users_email_idx
  on public.app_users (email);

alter table public.tickets enable row level security;
alter table public.app_users enable row level security;
alter table public.user_roles enable row level security;
alter table public.incident_counters enable row level security;
revoke all on public.tickets from public, anon, authenticated;
revoke all on public.app_users from public, anon, authenticated;
revoke all on public.user_roles from public, anon, authenticated;
revoke all on public.incident_counters from public, anon, authenticated;
revoke all on function public.next_incident_number(date) from public, anon, authenticated;
revoke all on function public.assign_ticket_number() from public, anon, authenticated;
grant all on public.tickets to service_role;
grant all on public.app_users to service_role;
grant all on public.user_roles to service_role;
grant all on public.incident_counters to service_role;
grant execute on function public.next_incident_number(date) to service_role;
