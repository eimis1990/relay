# Relay

Next.js website and client portal for managed email operations.

## Run locally

Node 22.13+; `npm ci`, then `npm run dev`. Open http://localhost:3000. Copy `.env.example` to `.env.local` to configure Outlook. No secrets are required to preview the site.

## Current functionality

- Responsive website and four-step client brief with persistent, authenticated Supabase storage on Vercel and a localhost-only preview fallback.
- Portal overview, connection setup screens and help.
- Microsoft OAuth authorization-code flow with PKCE/state, delegated `User.Read` and `Mail.Read`, server-side account allowlist, authenticated-encryption session cookies and optional encrypted offline authorization.
- Read latest 25 messages from the connected user's own Inbox (sender, subject, date, preview and safe Outlook link). Only GET requests to Microsoft Graph; reading does not mark messages read.
- Sessions expire after at most one hour. Reconnect to continue. Disconnect clears local browser session; revoke the application's consent in Microsoft to remove permission entirely.

**Not implemented:** automatic sorting/labeling, background polling, attachment analysis, shared mailboxes, Google OAuth/Sheets writes, Telegram pairing, billing and shared organization membership. Account setup details are preferences, not connected accounts.

## Deploy to Vercel

Import `eimis1990/relay`. Framework: **Next.js**. Root directory: repository root. Build: `npm run build`. Keep the default Next.js output setting. Select Node 22 or newer supported by Vercel.

The website can deploy without secrets. Microsoft connections stay disabled until all variables below are configured. Hosted workspace saves require the database configuration below. The same verified Microsoft session opens the private workspace. Missing database configuration and failed authentication stop saves; Vercel never uses a local JSON file.

### Microsoft setup (one-time)

1. In Microsoft Entra admin center, open **App registrations → New registration**.
2. Name the app **Relay**. To support both business and personal Outlook, choose **Accounts in any organizational directory and personal Microsoft accounts**. This code uses the `common` endpoint. Choose a Web application, not SPA.
3. Add exact Web redirect URI: `https://YOUR-PROJECT.vercel.app/api/outlook/callback`. Optionally add `http://localhost:3000/api/outlook/callback` for development. Use a stable Vercel URL, not changing preview URLs.
4. Under **API permissions**, add Microsoft Graph **delegated** `User.Read` and `Mail.Read`. For optional category sync, add delegated `Mail.ReadWrite` and `MailboxSettings.ReadWrite`. Do not add application-wide or sending permissions. Your organization's administrator may need to approve consent.
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

## Review inbox pilot

Open `/#review` after connecting Outlook. Load the latest 25 emails, select up to five, choose English or Lithuanian and optionally specify business priority rules. Confirm the data-sharing checkbox, then run Review selected. Results include company/evidence, topic, priority/reason, summary, next action, quoted deadlines and uncertainty flags. Filters, source text, local priority/topic corrections and Mark checked support human review.

Set `OPENAI_API_KEY` as a sensitive Production variable in Vercel and redeploy. `OPENAI_REVIEW_MODEL` defaults to `gpt-5.4-mini`; changing it requires a model supporting Responses structured outputs and `reasoning.effort: none`. API billing is separate from ChatGPT. Each manual run sends at most five emails and has a 6,000 output-token cap. This owner-allowlisted pilot has no distributed usage quota; add tenant budgets/rate controls before opening access to customers.

The server fetches selected messages only from the signed-in account's Inbox using Graph GET and plain-text preference. Only the first 8,000 characters per email, sender details, subject, date and priority rules are sent to OpenAI. Attachments and link destinations are not fetched. HTML responses fall back to preview text. Truncation, previews and attachments are clearly flagged. Model output is schema-validated, bound to requested message IDs, and company/deadline evidence must occur verbatim in the source. Email content is untrusted data; the model has no tools or write capability.

OpenAI requests use `store:false`; this opts out of Responses application-state storage, not OpenAI's separate abuse-monitoring retention. Real reviews and email snapshots are encrypted in Supabase, isolated by the signed-in Microsoft subject; demo results remain temporary. Relay does not send, move, delete, or mark messages read. Optional category sync adds matching labels to Outlook after separate Microsoft consent.

Try furniture demo reviews three explicitly fictional examples through the same AI service without reading or writing Outlook. It requires a signed-in account and the configured API key. It is not evidence of live-mailbox processing. `tests/review-boundaries.mjs` runs the real API routes with mocked Microsoft/OpenAI transports, including unsafe links, truncation, hallucinated evidence, duplicates, provider failures and auth boundaries. A real provider/mailbox smoke test is also required before demonstrating the live workflow.

References: [Microsoft message GET](https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0), [OpenAI Structured Outputs](https://developers.openai.com/api/docs/guides/structured-outputs).

## Saved reviews and automatic inbox processing

Outlook review snapshots are encrypted with AES-256-GCM using the server's `AUTH_SECRET`, with owner/message-specific authenticated context. Supabase stores the snapshot, immutable Microsoft message ID, review timestamp/language, checked state, corrections, and revision. The backend derives ownership only from the verified Microsoft subject in the signed session. Browser roles have no direct table or scheduler access. Keep `AUTH_SECRET` stable: rotation requires re-encryption of existing snapshots and tokens (and reinstalling the scheduler token).

`GET /api/outlook/history` loads 50 reviews per page with an ID cursor. `PATCH` updates priority/topic/checked using an optimistic revision check. Re-reviewing the same message does not overwrite existing human corrections. Demo results remain temporary. Existing reviews from the old memory-only pilot cannot be recovered; review those emails again to save them.

Automatic processing is opt-in through Review inbox. The extra Microsoft consent requests `offline_access` alongside the existing `User.Read` and `Mail.Read` scopes; it does not request sending or mailbox modification. An encrypted refresh token stays server-side. The activation callback verifies the same Microsoft subject as the signed-in user. Signing out clears the browser session but does not stop background processing. Pause stops new AI dispatches; a call already running may finish. Remove background connection clears the stored refresh token, keeps saved history, and stops processing. Provider revocation or an invalid grant pauses processing with a reconnect message.

The Supabase migration installs `pg_cron` and `pg_net`. After consent, a service-role-only RPC installs a five-minute job for the fixed production `/api/cron/review` URL. Its scoped bearer key is derived from `AUTH_SECRET` and stored in Supabase Vault. The privileged implementation lives in `relay_private`, with a security-invoker wrapper exposed only to the service role. The HTTP route authenticates the key using constant-time comparison. No Vercel paid cron plan is required for this scheduler.

Each run acquires an expiring mailbox lease, refreshes and rotates Microsoft's token, consumes one delta page, durably enqueues IDs, and then checkpoints the exact next/delta link. All Graph requests use immutable IDs and GET. Unfiltered delta avoids Graph's 5,000-message filtered-sync cap; only messages received after activation enter the queue. Large existing mailboxes can take multiple runs to finish the initial metadata scan. No attachments are fetched. Moved/deleted messages that are no longer in Inbox are skipped. A unique owner/message key prevents duplicate saved reviews. Queue failures retry up to three times with increasing delay; failed jobs can be retried from Relay. Expired delta cursors restart the metadata scan, with existing queue/history keys preventing re-review.

Runs process at most five emails and 300 review attempts per account per UTC day. The scheduler handles one eligible mailbox per invocation, oldest-checked first: this is a private pilot, not yet a multi-client high-throughput service. Review costs use the configured OpenAI API account. A crash after AI completion but before database commit can result in another billable AI attempt; exactly-once external billing is not guaranteed. Results persist before a job is marked done. Budget counts are conservative and include failed/interrupted attempts. Review snapshots and queue records are retained until an explicit retention/deletion workflow is added; do not promise automatic expiry.

Operational checks: the UI exposes enabled/paused state, pending and failed counts, last successful check, and sanitized errors. `cron.job_run_details` shows scheduler dispatch; `net._http_response` shows HTTP completion. A successful cron SQL execution alone does not prove the Vercel worker succeeded. Browser sessions still require Microsoft sign-in after expiry to view history; the background worker does not depend on that session.

Tests: `tests/history-automation.mjs` uses mocked Microsoft/OpenAI/PostgREST transports behind a real production Next server. It checks persistence across server restart, owner isolation, CAS corrections, encrypted snapshots, duplicate avoidance, leases, pause/resume, token rotation/revocation, retry limits, daily cap, cursor destination validation, cron authentication, and explicit offline consent. Deployment verification must additionally inspect database grants and a real scheduled HTTP completion.

## Workflow assistant and portal

The public landing page is unchanged. Portal routes are `#portal`, `#review`, `#workflow`, `#connections`, `#intake`, and `#support`. The responsive review inbox retains encrypted history, corrections, checked state, and fictional/manual review flows. Motion respects reduced-motion preferences.

Assistant & rules stores an encrypted, owner-scoped configuration in `relay_workflows`, guarded by optimistic revisions and server-only database access. The assistant calls the existing server-side OpenAI integration with structured output and `store:false`; it returns a proposal only. Users must use the proposal and save before future reviews change. Chat is temporary. Browser speech recognition is optional and may send audio to the browser provider; the transcript is editable and never automatically submitted. Voice availability varies by browser.

Label rules match literal case-insensitive substrings in sender, subject, body, AI-established company, or topic. All matching labels are attached to the saved Relay review. The first matching rule with an explicit priority overrides AI priority. With optional Outlook category sync enabled, matching labels also become Outlook categories. Semantic priorities go in reviewer instructions. Existing reviews are not relabeled retroactively; their Apply current labels action explicitly queues the current matching rules for Outlook.

Continuous mode uses the existing five-minute worker. Daily mode queues incoming mail and processes messages received by the configured local cutoff, starting on the first scheduler run at/after that time. Emails arriving later wait until the next local day; backlogs drain in bounded five-message batches. IANA time zones account for daylight-saving changes. The existing 300-attempt UTC daily cap remains in effect. Polling metadata still occurs between review times. A worker already running can finish with the previous settings.

Apply `20261119060723_relay_workflow_preferences.sql` before deploying the worker change. Existing accounts retain their current language/instructions and continuous scheduling until they save a workflow. Tests cover encrypted persistence, account isolation, revision conflicts, no implicit proposal saves, label matching, cutoff deferral, and winter/summer time-zone behavior.

### Outlook category sync

Apply `20261119064530_relay_outlook_categories.sql` before deploying. In Assistant & rules, the user opts in to category sync and completes Microsoft consent for delegated `Mail.ReadWrite`, `MailboxSettings.ReadWrite`, and `offline_access`. The callback checks both granted scopes and the existing Microsoft subject. Category consent preserves saved review settings and the activation date. A normal read-only background reconnect pauses category syncing.

Review persistence atomically records encrypted category payloads and a pending sync state. The existing leased worker syncs up to five pending updates per check, independent of the AI daily cap. It reuses category names case-insensitively, creates missing colors, and PATCHes only the categories property with existing values preserved and an If-Match version check. Existing colors and categories are never removed or renamed. A missing message is skipped; transient errors retry with backoff up to three attempts, separately from AI reviews. Permission errors stop category writes while allowing read-only reviews to continue. The UI shows pending/failed counts, last sync, pause/resume, and retry controls.

For an older saved review, Apply current labels reads the owner-isolated snapshot and current rules server-side; no arbitrary message IDs or category payloads are accepted from the browser. This does not invoke AI again. Pausing automatic processing stops category dispatch too; an already in-flight request may finish. Removing background access clears the refresh token and both category flags. Revoking permissions completely is done in Microsoft account settings.

Verification: `npm run build` and `npm test`. The history/automation integration suite covers opt-in scopes, missing-grant rejection, owner isolation, atomic outbox, additive updates, ETag conflicts, deduplication, throttling, pause, deleted emails, and independent permission failure. A live Microsoft consent and mailbox check are still required when first activating.
