# Boardera Decode API — MS_EmailRFQTriage Demo

Email-driven RFQ triage, powered by the [Boardera Decode API](https://boardera.io).

A customer emails a circuit board design package. This script picks it up from an Outlook folder, runs it through the Decode API, and within minutes delivers to your quoting team:

- **Executive Summary PDF** — board specs, costing, exceptions at a glance
- **BOM Report Excel** — cleansed, costed BOM with sourcing exceptions highlighted
- **DFM Report PDF** — SpeedDFM fabrication/assembly violations with rendered image tiles
- **Raw export JSON** — the full structured Decode payload (specs, BOM, costing, DFM, pricing)

The results arrive as a **forward of the original customer email** — subject
`Decode results: <original subject>`, the analysis summary up top, and the
customer's message and design package carried below it — so every result is
tied to the request it answers and threads with it in Outlook.

No portal, no upload form — the inbox *is* the interface.

> **Who receives the reports?** The generated reports include internal
> pricing, cost, and margin data, so by default they go to the mailbox you
> configure in `OUTLOOK_REPLY_TO` (typically your quoting team) — **not**
> back to the customer who sent the RFQ. Set `OUTLOOK_REPLY_TO=sender` only
> if senders should see that data (e.g. an internal-only intake address).

## How it works

```
Outlook folder (unread)
   └─ download attachments (a .zip/.tar archive is forwarded as-is — Decode extracts it server-side)
        └─ Decode API: createProject → upload → analyzeProject (poll) → exportProduct (poll)
             │           → getDFM (poll retrieveDFM) → dfmEntryImage × N (poll retrieveAsset)
             │           → getProjectPrices
             ├─ save raw export JSON       (_id + meta/specs/costing.scenarios + dfm + prices)
             ├─ generate BOM .xlsx         (specs.projectBom × costing.scenarios[0].pricingLineItems)
             ├─ generate Exec Summary .pdf (Puppeteer render)
             └─ generate DFM .pdf          (Puppeteer render: summary + by-rule + image gallery)
        └─ forward the original email to OUTLOOK_REPLY_TO (quoting team mailbox — or the
           sender), results summary on top, report files attached, subject "Decode results: …"
        └─ mark source message as read
```

## What you need

### 1. A Boardera Decode API key

Register at **<https://boardera.io/register>**. Your key is the base64-encoded
`{ _id, apiKey }` string shown in your Boardera console — it is sent as the
`X-Boardera-Key` header on every request. Each processed design package
consumes tokens from your plan; a typical package runs on the order of
1,000 tokens (see the token calculator on the [pricing page](https://boardera.io/pricing)).

Every API key is tied to a manufacturer — analysis capabilities, costing,
and pricing in the results all reflect that manufacturer's rates and rule
sets, so there is nothing to select in this tool.

> This demo's code is free and MIT-licensed. Use of the Decode API itself is
> governed by your Boardera plan and terms — the license on this repository
> grants no rights to the API.

### 2. Register an Azure AD application (for Outlook access)

The demo reads and sends mail as *you*, via Microsoft Graph with device-code
sign-in. Register your own (free) app:

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it (e.g. `decode-email-demo`). Choose the supported account types that match your mailbox (your org only, or include personal Microsoft accounts). No redirect URI needed.
3. Under **Authentication** → **Advanced settings**, set **Allow public client flows** to **Yes** (this enables the device-code flow).
4. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**, add: `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `offline_access`. Grant admin consent if your tenant requires it.
5. Copy the **Directory (tenant) ID** and **Application (client) ID** into your `.env`. For personal Microsoft accounts, set `AZURE_TENANT_ID=consumers` (or `common`).

### 3. An Outlook rule on the watched mailbox

- Source: mail addressed to your chosen alias (e.g. `you+rfq@yourdomain.com`)
- Action: **move to folder** named `Decode Demo Inbox` (or whatever you set in `OUTLOOK_WATCHED_FOLDER`)
- Create that folder before enabling the rule

### 4. A design package to test with

This repo includes a ready-made sample: **`sample-package/Room Controller.zip`**
(gerbers, drill files, BOM, and supporting docs). Or use your own — either a
single `.zip`, or the files as separate attachments; they are treated as one
project. If an email contains an archive, only the first archive is processed
and any other attachments are dropped (the archive is assumed to be the full
package — the log notes what was dropped). Optionally include `Qty: 25` in
the subject or body.

## Setup

```sh
cp .env.example .env
# Edit .env: AZURE_TENANT_ID, AZURE_CLIENT_ID, OUTLOOK_USER_EMAIL,
# OUTLOOK_WATCHED_FOLDER, OUTLOOK_REPLY_TO, DECODE_API_KEY

npm install
npm run auth
# Follow the device-code instructions printed in the terminal. Sign in once
# with the OUTLOOK_USER_EMAIL account at https://microsoft.com/devicelogin.
# A token-cache.json file is created so subsequent runs are non-interactive.
```

To verify your API key and endpoint before the first full run:

```sh
npx tsx scripts/test-decode.ts
```

### Configuration reference

All configuration is via `.env` (see `.env.example` for inline documentation):

| Variable | Required | Purpose |
| --- | --- | --- |
| `AZURE_TENANT_ID` | yes | Directory (tenant) ID of your Azure AD app registration; `consumers` or `common` for personal Microsoft accounts |
| `AZURE_CLIENT_ID` | yes | Application (client) ID of the app registration |
| `OUTLOOK_USER_EMAIL` | yes | Mailbox to read from and send as (the account you sign in with) |
| `OUTLOOK_WATCHED_FOLDER` | yes | Display name of the Outlook folder your inbox rule moves RFQ mail into |
| `OUTLOOK_REPLY_TO` | yes | Where result/error replies go — your quoting team's mailbox, or the literal value `sender` to reply to the RFQ sender (the reports contain internal pricing; see the note above) |
| `OUTLOOK_REPLY_BCC` | no | Blind-copy every outgoing reply to this address; empty = no BCC |
| `DECODE_API_KEY` | yes | Your Boardera Decode API key from <https://boardera.io/register> |
| `DECODE_API_ENDPOINT` | no | Override the API endpoint (defaults to production) |
| `DEFAULT_QUANTITY` | no | Quantity used when the email doesn't specify one (default 10) |
| `POLL_INTERVAL_MS` | no | Polling interval for `npm start` (default 30000) |
| `DFM_IMAGE_LIMIT` | no | Max per-violation image tiles in the DFM report; 0 disables (default 20) |

No other configuration is read from the environment or hardcoded — the repo
contains no Boardera- or Microsoft-account-specific values.

## Running the demo

```sh
# Process every unread message in the watched folder (oldest first) and exit.
# --dry-run skips the outbound reply so you can verify the generated reports
# under tmp/<messageId>/out/ first. Messages are still marked read — mark one
# unread in Outlook to process it again for real.
npm run run-once -- --dry-run

# Same, but actually send the replies.
npm run run-once

# Poll forever (every POLL_INTERVAL_MS) and process unread messages as they arrive.
npm start

# Poll without sending replies.
npm start -- --dry-run
```

A successful end-to-end run takes ~3–5 minutes; the Decode `analyzeProject`
step dominates, with SpeedDFM + per-violation tile rendering adding up to
~1 minute. Lower `DFM_IMAGE_LIMIT` (or set it to `0`) to trim the
tile-rendering time.

## File map

```
src/
  config.ts                env loading, paths
  index.ts                 CLI: auth | run-once | start  [--dry-run]
  pipeline.ts              orchestrator (email → Decode → reports → reply)
  outlook/
    auth.ts                MSAL Node device-code, token cache in token-cache.json
    graph.ts               graphFetch / graphJson helper
    inbox.ts               folder resolve, listUnreadMessages, markMessageRead
    attachments.ts         downloadAttachmentsToTemp, in/out dir helpers
    send.ts                forwardMessageWith (forward original + attach reports)
  decode/
    client.ts              GraphQL POST with X-Boardera-Key
    pipeline.ts            classify files → create → upload → analyze → exportProduct
                           → getDFM + dfmEntryImage tiles → getProjectPrices
    types.ts
  reports/
    bom-excel.ts           BOM report generator, writes .xlsx
    exec-pdf.ts            executive summary, renders via Puppeteer
    dfm-pdf.ts             DFM report: summary + by-rule + image gallery + table
  utils/
    zip.ts                 zipDirectory (re-zips the originals for the reply)
    quantity.ts            extract `Qty: N` from subject/body, strip HTML
sample-package/            a shareable test design package (Room Controller)
tmp/{messageId}/in/        attachments downloaded from the email
tmp/{messageId}/out/       generated reports + original.zip
tmp/{messageId}/out/dfm-images/   downloaded per-violation PNG tiles (embedded in the DFM PDF)
token-cache.json           MSAL refresh tokens (gitignored)
folder-cache.json          cached Outlook folder id (gitignored)
```

## Quantity parsing

The reply uses `DEFAULT_QUANTITY` (10) unless the email's subject or body contains a recognized quantity phrase, in either shape (case-insensitive):

1. **Keyword before the number** — `quantity`, `qty`, or `boards`, then an optional `:`, `=`, the word `of`, or just a space:

   ```
   qty 25    qty: 25    qty=25    quantity 25    quantity: 25
   quantity of 250    boards 10    boards: 10
   ```

2. **Unit after the number** — a number followed by a pieces unit (`pc`, `pcs`, `piece`, `pieces`):

   ```
   100 pcs    100pcs    100 pieces    100pc
   ```

First match wins; subject takes priority over body. It won't fire inside another word or on unrelated numbers (e.g. `keyboards`, `antiquity`, `10 PCB`, `0402`).

## Exceptions & warnings (tunable)

All three surfaces (BOM Excel, Exec PDF, email reply) share one classifier in
**`src/reports/bomAnalysis.ts`** — `classifyRow()` returns `{ exceptions, warnings }`,
and `filterRealBom()` drops the empty DNP placeholder rows the API emits (no part
number, no placement, `quantityPerBoard: 0`) so they never inflate counts. Thresholds
live at the top of that file (`LONG_LEAD_DAYS = 21`, `LOW_STOCK_THRESHOLD = 100`).

**Exceptions (red, counted):**

- **Customer supplied** — `projectBom[].customerSupplied` is true. Also excluded from the sourcing checks below (the customer provides the part).
- **EOL / Obsolete** — `lifeCycleStatus` matches `/eol|obsolete|end[\s_-]?of[\s_-]?life/i`
- **Unfindable** — no distributor resolved (sourced, non-customer-supplied rows only)
- **No stock** — `stock` is 0 or null
- **Insufficient stock** — `stock < orderQty`
- **Long lead** — `leadTime > LONG_LEAD_DAYS` **and** stock does not cover the order (if stock ≥ orderQty, lead time is irrelevant and is **not** flagged)
- **Backordered** / **Uncertain lead** — from `additionalDetails[].is_backordered` / `is_uncertain_lead`
- **Substitution** — `projectBom[].substitutions`, or appears in `validationDetails.assemblyWithSubstitutions`

**Warnings (amber, not counted as exceptions):**

- **Low stock** — `stock < LOW_STOCK_THRESHOLD` but still covers the order quantity

Sourcing-derived flags only fire when the API actually returned sourcing data for the
row. Lifecycle-unknown is **not** flagged. Sourcing values come from the top-level
`specs` rollup, falling back to `additionalDetails[].availability` when the rollup is empty.

## Behaviour on failure

If a message fails at any step (Decode error, malformed package, etc.) the
script forwards the original to `OUTLOOK_REPLY_TO` with subject
`Decode error: <original subject>`, an error summary, and fix guidance on
top; it logs the error and marks the message read so the queue advances. To
retry, mark the message unread in Outlook (or re-send it).

## Verification checklist

1. `npm install` && `npm run typecheck` clean
2. `npm run auth` completes the device-code flow once
3. `npm run run-once -- --dry-run` after sending one test email produces under `tmp/<id>/out/`:
   - `<project>_Analysis_Report.pdf`
   - `<project>_BOM_Report.xlsx`
   - `<project>_DFM_Report.pdf` (when SpeedDFM is available for the project)
   - `<project>_export.json`
   - `<project>_original_files.zip`
   - `dfm-images/violation-NN.png` (the tiles embedded in the DFM PDF)
4. Drop `--dry-run`, re-send a test, verify the results forward arrives at `OUTLOOK_REPLY_TO` — subject `Decode results: …`, the four report files attached, the original message and package below the summary — and the source message is marked read
5. Send a malformed package (no gerbers) — confirm the script emails an error reply and continues

## Troubleshooting

- **`Missing required env var`** — copy `.env.example` to `.env` and fill in the listed variable.
- **Device-code sign-in fails / AADSTS errors** — confirm "Allow public client flows" is enabled on the app registration and the delegated Graph permissions are granted. For personal accounts use `AZURE_TENANT_ID=consumers`.
- **`Decode HTTP 401/403`** — run `npx tsx scripts/test-decode.ts`; check the key is the full base64 string from your console and your plan is active.
- **Watched folder not found** — `OUTLOOK_WATCHED_FOLDER` must match the folder's display name exactly; delete `folder-cache.json` after renaming folders.
- **The tool's own replies appear in the watched folder** — happens when the inbox rule also matches the `OUTLOOK_REPLY_TO` address (e.g. the quoting mailbox is in the same account). They are recognized by their `Decode results:` / `Decode error:` subject prefix, skipped, and marked read — but prefer tightening the rule so replies don't enter the queue at all.
- **Puppeteer/Chromium download blocked** — first `npm install` downloads Chromium (~150 MB). Set `PUPPETEER_SKIP_DOWNLOAD=true` and point `PUPPETEER_EXECUTABLE_PATH` at a local Chrome instead.
- **Reply fails on large report files** — report attachments use Graph's simple upload, capped at ~3 MB per file. Lower `DFM_IMAGE_LIMIT` (the DFM PDF is usually the biggest file) or extend `src/outlook/send.ts` to the attachment upload-session flow.

## Limits and known constraints

- **Single-user.** Reads from one mailbox and sends every result to the one configured `OUTLOOK_REPLY_TO` (or per-sender with `OUTLOOK_REPLY_TO=sender`). Multi-tenant or app-only auth would require swapping device code for client-credentials and using the `/users/{id}/...` endpoints.
- **Polls every 30 s by default.** Configurable via `POLL_INTERVAL_MS`. Microsoft Graph also supports change notifications (webhooks) — overkill for a local demo.
- **Simple attachment uploads only.** Report files over ~3 MB each need the upload-session flow — see Troubleshooting above.

## License

[MIT](LICENSE) — free to use, modify, and redistribute, with no warranty.
The Boardera Decode API is a separate commercial service; get a key at
<https://boardera.io/register>.
