# Relay

Next.js website and client portal for managed email operations.

## Run locally

Node 22.13+; `npm ci`, then `npm run dev`. Open http://localhost:3000. Copy `.env.example` to `.env.local` to configure Outlook. No secrets are required to preview the site.

## Current functionality

- Responsive website and four-step client brief with persistent, authenticated Supabase storage on Vercel and a localhost-only preview fallback.
- Portal overview, connection setup screens and help.
- Microsoft OAuth authorization-code flow with PKCE/state, delegated `User.Read` and `Mail.Read`, server-side account allowlist, authenticated-encryption session cookies and no refresh token.
- Read latest 25 messages from the connected user's own Inbox (sender, subject, date, preview and safe Outlook link). Only GET requests to Microsoft Graph; reading does not mark messages read.
- Sessions expire after at most one hour. Reconnect to continue. Disconnect clears local browser session; revoke the application's consent in Microsoft to remove permission entirely.

**Not implemented:** automatic sorting/labeling, background polling, full-body reader, attachments, shared mailboxes, Google OAuth/Sheets writes, Telegram pairing, billing and shared organization membership. Account setup details are preferences, not connected accounts.

## Deploy to Vercel

Import `eimis1990/relay`. Framework: **Next.js**. Root directory: repository root. Build: `npm run build`. Keep the default Next.js output setting. Select Node 22 or newer supported by Vercel.

The website can deploy without secrets. Microsoft connections stay disabled until all variables below are configured. Hosted workspace saves require the database configuration below. The same verified Microsoft session opens the private workspace. Missing database configuration and failed authentication stop saves; Vercel never uses a local JSON file.

### Microsoft setup (one-time)

1. In Microsoft Entra admin center, open **App registrations → New registration**.
2. Name the app **Relay**. To support both business and personal Outlook, choose **Accounts in any organizational directory and personal Microsoft accounts**. This code uses the `common` endpoint. Choose a Web application, not SPA.
3. Add exact Web redirect URI: `https://YOUR-PROJECT.vercel.app/api/outlook/callback`. Optionally add `http://localhost:3000/api/outlook/callback` for development. Use a stable Vercel URL, not changing preview URLs.
4. Under **API permissions**, add Microsoft Graph **delegated** `User.Read` and `Mail.Read`. Do not add application-wide, sending or write permissions. Your organization's administrator may need to approve consent.
5. Copy the **Application (client) ID**. Under **Certificates & secrets**, create a client secret and privately copy its **Value**, not its ID. Track its expiry date.
6. Add the following Vercel environment variables, then redeploy:

| Variable | Value |
|---|---|
| `APP_URL` | `https://YOUR-PROJECT.vercel.app` |
| `MICROSOFT_CLIENT_ID` | Application/client ID |
| `MICROSOFT_CLIENT_SECRET` | Client secret **value** |
| `AUTH_SECRET` | Random secret of at least 32 characters; generate with `openssl rand -base64 48` |
| `RELAY_ALLOWED_MICROSOFT_ACCOUNTS` | Exact permitted account email(s), comma-separated |

7. Open **Client portal → Connections → Connect Outlook**, sign into the allowed Microsoft account, consent, then click **Read / refresh inbox**.

Never paste secrets into the questionnaire or commit `.env.local`. Never prefix these variables with `NEXT_PUBLIC_`. Work accounts require an Exchange Online mailbox. Some personal account aliases differ from the primary Graph profile address; allow the primary account address if needed. Wrong state, expired session, missing config and unlisted accounts fail closed.

## Persistent workspace storage

1. Use a dedicated Relay Supabase project. Apply `supabase/migrations/20261117120430_relay_workspaces.sql` through the migration tool or SQL editor once.
2. Add **server-only** Vercel environment variables `SUPABASE_URL` (project URL) and `SUPABASE_SECRET_KEY` (a secret API key, `sb_secret_...`). Legacy service-role JWTs also work. Never use `NEXT_PUBLIC_` or expose these credentials to the browser.
3. Redeploy. Sign in with an allowed Microsoft account, save a brief, reload, and confirm it persists.

The API uses the stable Microsoft Graph profile ID from the encrypted session, never a client-submitted company, email or owner ID. Older sessions acquire the ID from Microsoft's `/me` endpoint. An email alias change does not create another workspace. Each ID owns one workspace in this pilot; shared teams, delegated membership and enterprise tenant administration are not implemented.

The `relay_workspaces` table enables RLS and revokes access from browser roles (`anon` and `authenticated`). There are intentionally no browser RLS policies. Only the server role has SELECT/INSERT/UPDATE; every server read/update filters by the verified owner. The server secret bypasses RLS, so these owner checks are essential. Tokens, secrets and email contents are not stored in the workspace table. App-level tests cover owner spoofing and cross-account isolation.

An opaque account-bound workspace ID also rejects stale-tab saves after switching Microsoft accounts in the same browser. Saves use a revision number and a conditional database UPDATE; stale tabs receive HTTP 409 instead of overwriting newer work. Reload to reconcile a conflict. Storage errors do not report success or fall back to ephemeral files. Sign-out clears this browser's session, not the database; reconnecting restores the workspace. Session expiry remains one hour and background processing remains disabled.

### Local preview

When no database is configured, `next dev` on loopback can save `.local-data/workspace.json` (single-user development only, excluded from Git). Do not expose this server publicly. Production builds require Microsoft authentication and the database even on loopback. Existing local preview files are not automatically uploaded to a person's cloud workspace.

## Verification

`npm run build` runs the production build and TypeScript checks. `npm test` checks OAuth boundaries and workspace authorization, cross-account isolation, alias-stable ownership, revision conflicts, database failure handling and persistence across Next server restarts using a local mock PostgREST server. It does not contact Microsoft or prove real database RLS. Before declaring production ready, apply the migration, inspect database access controls, and test save/reload through the live Vercel portal.
