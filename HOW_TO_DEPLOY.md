# SMARTIS — How to share it with your team (deploy a link)

Your dashboard is now a single self-contained file: **`index.html`**.
It has no backend and no build step, so hosting it = putting one file online.
Once hosted you get a URL like `https://smartis-deeprock.netlify.app` that anyone
on your team can open — no install, no Google account.

> **File to deploy:** `index.html` (identical to `SMARTIS Deeprock Ghana v4.html`,
> just renamed so the web address is clean). Always deploy `index.html`.

---

## Option A — Netlify Drop (easiest, ~2 minutes, free)

1. Go to **https://app.netlify.com/drop**
2. Drag the **`index.html`** file onto the page (or the whole `SMART` folder).
3. Wait ~10 seconds. Netlify gives you a live URL (e.g. `random-name-123.netlify.app`).
4. (Optional) Create a free Netlify account to rename the site to something like
   `smartis-deeprock.netlify.app` and to keep the link permanent.
5. Share that URL with your team.

**To update later:** drag the new `index.html` onto the same site (Deploys → drag & drop),
or just re-drop at netlify.com/drop for a fresh link.

---

## Option B — Cloudflare Pages (free, custom domain friendly)

1. Sign in at **https://dash.cloudflare.com** → **Workers & Pages** → **Create** → **Pages**.
2. Choose **Upload assets** (no Git needed).
3. Upload `index.html`.
4. Click **Deploy**. You get a `*.pages.dev` URL to share.
5. (Optional) Add your own domain under the project's **Custom domains** tab.

---

## Option C — GitHub Pages (free, good if you already use GitHub)

1. Create a new repository at **https://github.com/new** (e.g. `smartis`), set it **Public**.
2. Click **Add file → Upload files**, upload `index.html`, then **Commit**.
3. Go to the repo's **Settings → Pages**.
4. Under **Source**, choose **Deploy from a branch** → branch **main** → folder **/(root)** → **Save**.
5. After ~1 minute your site is live at
   `https://<your-username>.github.io/smartis/`.
6. Share that URL.

**To update later:** upload a new `index.html` to the repo (it redeploys automatically).

---

## How your team uses it

1. Open the shared link.
2. On the welcome screen, either:
   - **Upload a file** — in Google Sheets: **File → Download → Microsoft Excel (.xlsx)**,
     then drag that file onto the dashboard. (Recommended — no setup.)
   - **USE DEMO DATA** — to explore with sample data.
3. The uploaded data is remembered in that person's browser, so a refresh keeps it.

### Important notes
- **Data is private to each viewer's browser.** Uploading a file does **not** push data to
  the server or to other viewers — everyone who should see the numbers uploads the same
  `.xlsx` (or you send it to them). If you want everyone to see one always-current dataset
  automatically, use the **Advanced → live Google Sheet** option instead of file upload.
- The dashboard loads one small library (SheetJS) from a public CDN, so viewers need
  internet access. Everything else runs locally in the browser.
- Your Excel file's tab names must match exactly:
  `RC_Rigs, DD_Rigs, Compressors, LightVehicles, WorkOrders, Spares, KPITargets, Technicians`.
  (The `SMARTIS Demo Data Template.xlsx` already has them correct — build from that.)

---

## Files in this folder

| File | What it is |
|------|-----------|
| `index.html` | **Deploy this.** The finished dashboard with file-upload support. |
| `SMARTIS Deeprock Ghana v4.html` | Same as index.html (your working copy). |
| `SMARTIS Deeprock Ghana v4.BACKUP.html` | The version *before* upload was added (safety backup). |
| `SMARTIS Demo Data Template.xlsx` | Template with the 8 correctly-named tabs + demo data. |
| `SMARTIS_v4_Source_Package.zip` | Older distribution bundle (Apps-Script based) — superseded. |
