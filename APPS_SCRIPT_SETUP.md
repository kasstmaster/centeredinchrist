# Google Apps Script staff/admin backend setup

The website remains static-hostable. Authentication and password updates are handled by a Google Apps Script Web App connected to the Google Sheet.

## 1. Paste the backend code

1. Open your Google Apps Script project that is connected to the Google Sheet.
2. Replace the current test code with the full contents of `apps-script-backend.js` from this repository.
3. Make sure `SHEET_ID` is available as a Script Property named `SHEET_ID`. If you used a hard-coded constant in the test code instead, replace `PASTE_YOUR_EXISTING_SHEET_ID_HERE` in `apps-script-backend.js` with your real sheet ID inside Apps Script only.
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
