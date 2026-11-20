# WPM RACER

The static frontend is hosted on GitHub Pages, the API and authorization remain in Google Apps Script, and the bound Google Sheet remains the database. A Cloudflare Worker provides the browser-facing API proxy.

## Request flow

1. `site/index.html` loads `site/config.js` and sends JSON `POST` requests to the configured Worker `/api` endpoint. Session tokens and login passwords are only in the HTTPS request body, never in URLs.
2. The Worker answers CORS `OPTIONS` preflights for exact origins listed in `ALLOWED_ORIGINS`. It validates the request, adds the server-only `WPM_PROXY_SECRET`, and forwards the request to Apps Script with redirect following enabled.
3. Apps Script `doPost` checks the proxy secret and passes the original request to the existing `apiRequest` dispatcher. `ContentService` responses redirect to a one-time `script.googleusercontent.com` URL; the Worker follows that redirect and returns JSON with GitHub Pages CORS headers.
4. `GET /health` is a proxy-only health check. Authenticated API operations use `POST` so tokens are not placed in query strings.

The Worker forwards every existing action: `login`, `logout`, `bootstrap`, `leaderboard`, `createRace`, `joinRace`, `beginRace`, `raceState`, `liveProgress`, `submitRace`, `adminOverview`, `adminExport`, `adminCreateStudent`, `adminAddPassage`, and `adminSetAccuracy`.

## Deploy the Apps Script backend

Generate one fresh random secret locally and keep it available for the Apps Script and Worker setup below. For PowerShell:

```powershell
$bytes = New-Object byte[] 32
$rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$rng.GetBytes($bytes)
$secret = [BitConverter]::ToString($bytes).Replace('-', '').ToLowerInvariant()
$rng.Dispose()
$secret
```

1. Keep using the existing bound spreadsheet. In its Apps Script project, replace `Code.gs` with this repository's `Code.gs` and update `appsscript.json`. The frontend is no longer served by Apps Script; `doGet` returns a small JSON service response.
2. Run `setupDatabase` only if you still need to initialize missing tabs. Existing rows are not overwritten.
3. In Apps Script **Project Settings → Script properties**, set:
   - `WPM_PROXY_SECRET`: a fresh random secret that will also be configured as a Cloudflare Worker secret.
   - `WPM_ADMIN_USERNAME`, `WPM_ADMIN_PASSWORD` (at least 12 characters), and optionally `WPM_ADMIN_NAME` if bootstrapping the teacher account.
4. Run `provisionAdminAccount` once. It stores a password hash and removes the temporary admin password property. Remove the username/name properties afterward if desired.
5. Deploy as a **Web app**, executing as the spreadsheet owner and accessible to **Anyone**. The Worker needs an unauthenticated HTTPS route to reach Apps Script; `doPost` rejects calls that do not carry the proxy secret. If Workspace policy prohibits public web apps, this deployment topology will not work as-is.
6. Preserve access to the existing Sheets tabs and keep spreadsheet edit access limited to trusted teachers.

## Deploy the Cloudflare Worker

1. Edit `worker/wrangler.toml`:
   - Set `APPS_SCRIPT_URL` to the deployed `https://script.google.com/macros/s/.../exec` URL.
   - Set `ALLOWED_ORIGINS` to the exact GitHub Pages **origin**, such as `https://OWNER.github.io` (no repository path), or your custom domain. For local testing, add the exact local origin too.
2. Install and authenticate Wrangler, deploy the Worker, and then set the same secret value as a Worker secret when prompted:

   ```powershell
   cd .\worker
   npm install --global wrangler
   wrangler login
   wrangler deploy
   wrangler secret put WPM_PROXY_SECRET
   ```

3. Set `site/config.js` `apiUrl` to the deployed Worker URL plus `/api`, for example `https://wpm-racer-api.YOUR-SUBDOMAIN.workers.dev/api`.
4. Confirm the proxy health endpoint returns JSON at `https://YOUR-WORKER.workers.dev/health`.

Do not put `WPM_PROXY_SECRET` in `wrangler.toml`, `site/config.js`, GitHub Pages, or any committed file. The Apps Script URL is a deployment address, not an authentication credential.

## Deploy the GitHub Pages frontend

1. Push this repository to GitHub. The included `.github/workflows/pages.yml` publishes only the `site/` directory on pushes to `main`.
2. In repository **Settings → Pages**, choose **GitHub Actions** as the build/deploy source. If your default branch is not `main`, update the workflow trigger.
3. Open the published Pages URL. The frontend and static assets are served by GitHub Pages; no Apps Script iframe/runtime is used.

`site/config.js` is public configuration and should contain only the Worker URL. CORS is a browser-origin policy, not authentication: non-browser clients can spoof an `Origin` header. The Apps Script proxy secret protects the backend from direct requests, while student sessions and server-side role checks continue to authorize application operations.

The redirect behavior follows Google's [Apps Script Content Service guidance](https://developers.google.com/apps-script/guides/content); Cloudflare's [Workers Fetch API](https://developers.cloudflare.com/workers/runtime-apis/fetch/) is configured to follow upstream redirects.

## Data, scoring, and platform limits

All existing database tabs and API behavior are retained. The app still provides student sign-in, student dashboards, passage/free-writing/accuracy/daily races, private room codes, polling-based competitor updates, race results, classroom/weekly/personal-best/daily leaderboards, achievements, teacher account/passage management, and CSV export.

Passwords use salted PBKDF2-HMAC-SHA-256 (60,000 iterations) in Apps Script. Sessions use opaque tokens, store only token hashes, expire after eight hours, and can be revoked at logout. Official results are calculated server-side from typed text and the stored passage; live progress is display-only. The local word list is intentionally a starter list and can be expanded in the `Words` sheet.

The Worker necessarily processes API request bodies, including login credentials and session tokens, in transit; it does not log or persist them. Cloudflare is therefore part of the trust boundary. If your policy does not permit that, use an API proxy you operate instead. Apps Script and Sheets have latency, quotas, and concurrency limits; the Worker polls no live connection, and browser-based typing cannot be made cheat-proof. These protections are not suitable for high-stakes assessment.

To enable winner-only `ko.mp4` playback, set `WPM_WINNER_VIDEO_URL` in Apps Script Script Properties to a URL the browser can play. The video itself is not included in this repository; if the property is unset, the winner screen omits it.

## Validation

`Code.gs` runs in Apps Script's V8 runtime. The static frontend and Worker are plain JavaScript with no application dependencies; the Worker can be syntax-checked with `node --check .\src\index.js` from the `worker` folder. Browser/API behavior should be verified after configuring the Apps Script deployment, Worker secret/origin, and GitHub Pages URL.
