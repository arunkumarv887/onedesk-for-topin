# Onedesk

Onedesk uses Microsoft Entra ID for sign-in and Supabase Postgres for incident storage. The Node.js API validates Microsoft tokens and accesses Supabase with a server-only secret key; Supabase credentials are never sent to the browser.

## Configuration

Copy `.env.example` to `.env.local` and configure:

- `VITE_MICROSOFT_CLIENT_ID`, `VITE_MICROSOFT_TENANT_ID`, and `VITE_MICROSOFT_API_SCOPE` for the single-tenant Entra SPA.
- Add the Microsoft Graph delegated `User.ReadBasic.All` permission to the SPA app registration and grant tenant admin consent so global admins can list the Entra directory from Settings.
- `MICROSOFT_CLIENT_ID`, `MICROSOFT_TENANT_ID`, `MICROSOFT_API_AUDIENCE`, and `MICROSOFT_API_SCOPE_NAME` for API token validation. The client ID must match the SPA, the audience must match the access token's `aud`, and the scope name must match the delegated scope in Entra.
- `MICROSOFT_ADMIN_EMAILS` is a comma-separated server-side allowlist of global administrators. These identities can grant or revoke OneDesk admin access for signed-in users. `MICROSOFT_ADMIN_ROLE` can also grant protected admin access through an Entra app role. Only use trusted identities in the allowlist.
- `SUPABASE_URL` and `SUPABASE_SECRET_KEY` for server-side database access. Use the project URL and a Supabase secret key (or legacy `service_role` key); do not use the publishable key for writes and never prefix the secret key with `VITE_`.
- `TICKETS_FILE` is only used by the one-time migration script to import the previous local ticket data.
- `VITE_TICKETS_API_URL`, `API_HOST`, and `API_PORT` for the browser API endpoint and server listener. The local default binds only to `127.0.0.1`; set `API_HOST=0.0.0.0` when deploying behind a secured web host. `APP_ORIGINS` must list the exact allowed frontend origins, comma-separated.

Only `VITE_MICROSOFT_*` and `VITE_TICKETS_API_URL` are used by the frontend. Never put database credentials in a `VITE_` variable. The publishable Supabase key is not needed by this server-mediated design.

## Run locally

1. In the Supabase SQL Editor, run [`supabase/schema.sql`](./supabase/schema.sql) to create the incident and user-access tables, indexes, and deny direct client table access with RLS.
2. In the Supabase dashboard, create a server-only secret key and add it to `.env.local` as `SUPABASE_SECRET_KEY`. Set `SUPABASE_URL` to the project URL. Do not paste the secret into chat or expose it in a `VITE_` variable.
3. Install dependencies with `npm install`.
4. Import the prior local incident data once with `node --env-file=.env.local server/migrate-tickets.mjs`. The import is idempotent by ticket ID; keep the local JSON file as a backup until you verify the records in Supabase.
5. Start the API on `http://127.0.0.1:3002` with `npm run start:api`; it verifies the Supabase connection before listening. Start the frontend with `npm run dev -- --host localhost`, then open `http://localhost:5173`.

If you already have a Supabase project, run `supabase/schema.sql` again to add the user-access tables; the script is safe to re-run and preserves existing incidents. If the API reports that database tables are unavailable, run `supabase/schema.sql` in the SQL Editor for the exact Supabase project configured by `SUPABASE_URL`, then restart the API.

## Deploying to Vercel

The repository includes a Vercel serverless API route at `api/[...path].mjs` alongside the Vite frontend. The frontend uses the same-origin `/api` route in production, even if `VITE_TICKETS_API_URL` is unset or incorrectly points to localhost.

Set these variables in Vercel Project Settings → Environment Variables:

- `VITE_MICROSOFT_CLIENT_ID`, `VITE_MICROSOFT_TENANT_ID`, and `VITE_MICROSOFT_API_SCOPE` for the frontend build.
- Add Microsoft Graph delegated `User.ReadBasic.All` to the SPA app registration and grant tenant admin consent to enable directory user listing in Settings.
- `MICROSOFT_CLIENT_ID`, `MICROSOFT_TENANT_ID`, `MICROSOFT_API_AUDIENCE`, and `MICROSOFT_API_SCOPE_NAME` for API token validation.
- `SUPABASE_URL` and `SUPABASE_SECRET_KEY` for server-side database access. Never prefix the secret with `VITE_`.
- `MICROSOFT_ADMIN_EMAILS` and `MICROSOFT_ADMIN_ROLE` for admin access.

Do not set `VITE_TICKETS_API_URL` to localhost in Vercel. Leave it unset to use `/api`, or set it to `/api`. Redeploy after changing the environment variables. Add the Vercel deployment URLs to the Entra SPA redirect URI configuration. The API allows same-origin calls and continues to require valid Microsoft access tokens for ticket operations.

The `/api/healthz` endpoint checks the Supabase tickets table. The API uses Vercel's request/response lifecycle and does not run `npm run start:api` in production.

The API validates Microsoft access-token signature, issuer, audience, tenant, SPA client, and delegated scope before accessing incident data. Reporters can create and view their own incidents. Global administrators listed in `MICROSOFT_ADMIN_EMAILS` can manage user roles from Settings. Settings loads tenant users through the signed-in global admin's Microsoft Graph delegated `User.ReadBasic.All` permission, then syncs their directory details to OneDesk. App users assigned the `IncidentDesk.Admin` Entra role keep protected admin access. Admins can view all incidents and update status or assignment. The Supabase secret bypasses RLS, so it is used only by the API after Microsoft authorization; browser clients cannot query the tables directly.

## Incident features

- Incident dashboard with status totals, six-month volume trend, and recent tickets.
- Incident identifiers use `INC-YYYY-MM-DD-N`, with a daily sequence starting at 1 (UTC).
- Searchable incident list, status filters, incident details, and admin status/assignment updates.
- Global-admin settings page with radio controls to grant or revoke OneDesk admin access.
- Incident submission with category, priority, location, and description.
- Monthly CSV report for incidents visible to the signed-in user.

## Commands

- `npm run dev` — start the Vite frontend.
- `npm run start:api` — start the Microsoft-authenticated Supabase API.
- `npm run build` — type-check and build the frontend.
- `npm run lint` — lint the frontend source.
