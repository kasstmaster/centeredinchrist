# Google Apps Script staff/admin backend setup

The website remains static-hostable. Authentication and password updates are handled by a Google Apps Script Web App connected to the Google Sheet.

## 1. Paste the backend code

1. Open your Google Apps Script project that is connected to the Google Sheet.
2. Replace the current test code with the full contents of `apps-script-backend.js` from this repository.
3. `apps-script-backend.js` includes the current spreadsheet ID as a fallback. If you need to point the backend at a different spreadsheet, set `SHEET_ID` as a Script Property or update `FALLBACK_SHEET_ID` inside Apps Script only.
4. Optional for stream status: add Script Properties `YOUTUBE_API_KEY` and `YOUTUBE_CHANNEL_ID`. If omitted, the stream status safely falls back to offline.

## 2. Deploy as a Web App

1. Click **Deploy** > **New deployment**.
2. Select **Web app**.
3. Set **Execute as** to **Me**.
4. Set **Who has access** to **Anyone**.
5. Click **Deploy**.
6. Copy the `/exec` Web App URL. It looks like `https://script.google.com/macros/s/AKfycb.../exec`.

## 3. Paste the Web App URL into the website

In `index.html`, replace:

```js
const STAFF_BACKEND_URL = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";
```

with the copied `/exec` Web App URL.

The Web App URL is not a password or Sheet ID. It is safe to place in the static website because all password checks and Sheet access happen inside Apps Script.

## 4. Test the Web App JSON response

After each Apps Script change, deploy a new Web App version and test that the deployed `/exec` URL returns JSON before testing the login form:

```sh
curl -L "YOUR_WEB_APP_EXEC_URL?action=ping"
```

Expected response shape:

```json
{"ok":true,"service":"centeredinchrist-staff-auth","action":"ping","timestamp":"..."}
```

You can also test the same URL-encoded POST format used by the website:

```sh
curl -L -X POST --data 'action=ping' "YOUR_WEB_APP_EXEC_URL"
```

The website intentionally sends a simple request without custom headers so Apps Script can handle it without a CORS preflight.
