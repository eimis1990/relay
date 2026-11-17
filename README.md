# Relay

Next.js website and client portal for managed email operations.

## Run locally

Node 22.13+; `npm ci`, then `npm run dev`. Open http://localhost:3000. Copy `.env.example` to `.env.local` to configure Outlook. No secrets are required to preview the site.

## Current functionality

- Responsive website and four-step client brief with local draft persistence.
- Portal overview, connection setup screens and help.
- Microsoft OAuth authorization-code flow with PKCE/state, delegated `User.Read` and `Mail.Read`, server-side account allowlist, authenticated-encryption session cookies and no refresh token.
- Read latest 25 messages from the connected user's own Inbox (sender, subject, date, preview and safe Outlook link). Only GET requests to Microsoft Graph; reading does not mark messages read.
- Sessions expire after at most one hour. Reconnect to continue. Disconnect clears local browser session; revoke the application's consent in Microsoft to remove permission entirely.

**Not implemented:** automatic sorting/labeling, background polling, full-body reader, attachments, shared mailboxes, Google OAuth/Sheets writes, Telegram pairing, billing, hosted questionnaire storage, multi-client organization membership. Account setup details are preferences, not connected accounts.

## Deploy to Vercel

Import `eimis1990/relay`. Framework: **Next.js**. Root directory: repository root. Build: `npm run build`. Keep the default Next.js output setting. Select Node 22 or newer supported by Vercel.

The website can deploy without secrets. Microsoft connections stay disabled until all variables below are configured. Questionnaire saves deliberately fail closed in production until hosted auth/storage are implemented; local JSON files are not a Vercel database. The Outlook session and inbox reader are independent of that prototype storage and can work on Vercel after configuration.

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

## Local data

`.local-data/workspace.json` is single-user development data only, excluded from Git. The API validates schema and same-origin writes and atomically replaces the file. Do not expose the dev server publicly. On Vercel, add a database and genuine tenant-scoped authentication before accepting client questionnaires. The Microsoft allowlist is an owner pilot safeguard, not a complete B2B authorization model.

## Verification

`npm run build` runs the production build and TypeScript checks. `npm test` exercises authorization boundaries against a local production server without contacting Microsoft; it does not prove a real Microsoft login. Real end-to-end OAuth requires the configured Microsoft registration and user consent.
