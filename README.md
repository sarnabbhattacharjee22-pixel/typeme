# WPM RACER

The frontend is hosted on GitHub Pages. Google Apps Script and a bound Google
Sheet provide the backend and database. A hidden Apps Script HTML page bridges
the Pages site to `google.script.run`; a Cloudflare Worker is not required.

## Deploy the Apps Script backend

1. Open the bound spreadsheet and select **Extensions → Apps Script**.
2. Replace `Code.gs` with this project's `Code.gs`, and update
   `appsscript.json`. Create an HTML file named `Bridge` and copy in
   `Bridge.html`.
3. In **Project Settings → Script properties**, set `WPM_ALLOWED_ORIGIN` to
   the exact GitHub Pages origin, for example
   `https://sarnabbhattacharjee22-pixel.github.io`. Do not include the repository
   path or a trailing slash.
4. If the sheets are not initialized, run `setupDatabase` from the editor.
   Existing rows are not overwritten. These setup functions require the signed-in
   editor to be the effective script owner.
5. To create a teacher account, set `WPM_ADMIN_USERNAME` and
   `WPM_ADMIN_PASSWORD` (at least 12 characters) in Script Properties, and
   optionally `WPM_ADMIN_NAME`. Run `provisionAdminAccount` from the editor.
   The temporary password property is removed after successful provisioning.
   Do not run this again if the teacher account already exists.
6. Deploy or update the web app to a **new version**, executing as the
   spreadsheet owner and accessible to **Anyone**. Copy its `/exec` URL.
   If Workspace policy prohibits public web apps, this setup will not work
   as-is.

Keep spreadsheet edit access limited to trusted teachers. Student accounts
are created from the teacher portal, not by inserting passwords into Sheets.

## Deploy the GitHub Pages frontend

1. Set `appsScriptUrl` in `site/config.js` to the Apps Script web app `/exec`
   URL. This URL is public configuration, not a credential.
2. Push the project to GitHub. The `.github/workflows/pages.yml` workflow
   publishes `site/` on pushes to `main`.
3. In **Settings → Pages**, choose **GitHub Actions**. If your branch is not
   named `main`, update the workflow trigger. Disable any other Pages workflow
   that publishes a different folder.
4. Open the published site and sign in. The bridge checks the configured origin
   and requires browsers to allow the embedded Apps Script page. If the bridge
   cannot load, check the deployment version, access settings, allowed origin,
   and browser restrictions.

When migrating from the Worker-based version, deploy the updated Apps Script
first, then publish the updated frontend. The legacy `doPost` adapter remains
available for the old Worker while the transition is underway; it still requires
`WPM_PROXY_SECRET`. After verifying the new site, you can delete the old
Cloudflare Worker and its secret. The existing Worker files are retained only
as legacy deployment material; they are not used by the new frontend.

## Features and limits

The backend supports `login`, `logout`, `bootstrap`, `leaderboard`,
`createRace`, `joinRace`, `beginRace`, `raceState`, `liveProgress`,
`submitRace`, `adminOverview`, `adminExport`, `adminCreateStudent`,
`adminAddPassage`, and `adminSetAccuracy`. Student sessions and server-side role
checks authorize operations; the allowed origin is not authentication.

Passwords use salted PBKDF2-HMAC-SHA-256 (60,000 iterations) in Apps Script.
Existing password hashes are preserved. Removing the Worker does not make that
password check faster, and Apps Script and Sheets still have latency, quotas,
and concurrency limits. Race scoring is server-side, but browser-based typing
cannot be made cheat-proof; this app is not suitable for high-stakes assessment.

To enable winner-only video playback, set `WPM_WINNER_VIDEO_URL` in Apps Script
Script Properties to a URL the browser can play. The video itself is not included.
