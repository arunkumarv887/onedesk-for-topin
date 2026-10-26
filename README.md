# Incident Desk

Incident Desk uses Microsoft Entra ID for sign-in and Supabase Postgres for incident storage. The Node.js API validates Microsoft tokens and accesses Supabase with a server-only secret key; Supabase credentials are never sent to the browser.

## Configuration

Copy `.env.example` to `.env.local` and configure:

- `VITE_MICROSOFT_CLIENT_ID`, `VITE_MICROSOFT_TENANT_ID`, and `VITE_MICROSOFT_API_SCOPE` for the single-tenant Entra SPA.
- `MICROSOFT_CLIENT_ID`, `MICROSOFT_TENANT_ID`, `MICROSOFT_API_AUDIENCE`, and `MICROSOFT_API_SCOPE_NAME` for API token validation. The client ID must match the SPA, the audience must match the access token's `aud`, and the scope name must match the delegated scope in Entra.
- `MICROSOFT_ADMIN_EMAILS` is a comma-separated server-side allowlist of Entra email/UPN identities with administrator access. `MICROSOFT_ADMIN_ROLE` can also grant access through an Entra app role. Only use trusted identities in the allowlist.
- `SUPABASE_URL` and `SUPABASE_SECRET_KEY` for server-side database access. Use the project URL and a Supabase secret key (or legacy `service_role` key); do not use the publishable key for writes and never prefix the secret key with `VITE_`.
- `TICKETS_FILE` is only used by the one-time migration script to import the previous local ticket data.
- `VITE_TICKETS_API_URL`, `API_HOST`, and `API_PORT` for the browser API endpoint and server listener. The local default binds only to `127.0.0.1`; set `API_HOST=0.0.0.0` when deploying behind a secured web host. `APP_ORIGINS` must list the exact allowed frontend origins, comma-separated.

Only `VITE_MICROSOFT_*` and `VITE_TICKETS_API_URL` are used by the frontend. Never put database credentials in a `VITE_` variable. The publishable Supabase key is not needed by this server-mediated design.

## Run locally

1. In the Supabase SQL Editor, run [`supabase/schema.sql`](./supabase/schema.sql) to create the `tickets` table, indexes, and deny direct client table access with RLS.
2. In the Supabase dashboard, create a server-only secret key and add it to `.env.local` as `SUPABASE_SECRET_KEY`. Set `SUPABASE_URL` to the project URL. Do not paste the secret into chat or expose it in a `VITE_` variable.
3. Install dependencies with `npm install`.
4. Import the prior local incident data once with `node --env-file=.env.local server/migrate-tickets.mjs`. The import is idempotent by ticket ID; keep the local JSON file as a backup until you verify the records in Supabase.
5. Start the API on `http://127.0.0.1:3002` with `npm run start:api`; it verifies the Supabase connection before listening. Start the frontend with `npm run dev -- --host localhost`, then open `http://localhost:5173`.

If the API reports that `public.tickets` is unavailable, run `supabase/schema.sql` in the SQL Editor for the exact Supabase project configured by `SUPABASE_URL`, then restart the API.

The API validates Microsoft access-token signature, issuer, audience, tenant, SPA client, and delegated scope before accessing incident data. Reporters can create and view their own incidents. Users listed in `MICROSOFT_ADMIN_EMAILS` or assigned the `IncidentDesk.Admin` app role can view all incidents and update status or assignment. The Supabase secret bypasses RLS, so it is used only by the API after Microsoft authorization; browser clients cannot query the table directly.

## Incident features

- Incident dashboard with status totals, six-month volume trend, and recent tickets.
- Searchable incident list, status filters, incident details, and admin status/assignment updates.
- Incident submission with category, priority, location, and description.
- Monthly CSV report for incidents visible to the signed-in user.

## Commands

- `npm run dev` — start the Vite frontend.
- `npm run start:api` — start the Microsoft-authenticated Supabase API.
- `npm run build` — type-check and build the frontend.
- `npm run lint` — lint the frontend source.
