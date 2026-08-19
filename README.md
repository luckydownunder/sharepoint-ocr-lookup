# SharePoint OCR Lookup

Web app: upload a PDF or photo, detect a name, then search Microsoft 365 SharePoint. Next step (not built yet): upload the file to a matching SharePoint subsite.

## Run locally

```powershell
cd $HOME\sharepoint-ocr-lookup
copy .env.example .env
npm install
npm run dev
```

Open http://localhost:5173

## Azure app registration (required for SharePoint)

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**.
2. Name it `SharePoint OCR Lookup`.
3. Supported account types: **Accounts in this organizational directory only**.
4. Redirect URI: platform **Single-page application (SPA)** → `http://localhost:5173`
5. After it is created, copy:
   - **Application (client) ID** → `VITE_AZURE_CLIENT_ID` in `.env`
   - **Directory (tenant) ID** → `VITE_AZURE_TENANT_ID` in `.env`
6. **API permissions** → Microsoft Graph → **Delegated**:
   - `User.Read`
   - `Files.Read.All`
   - `Sites.Read.All`
7. Click **Grant admin consent** if your tenant requires it (IT often must do this).

Optional: set `VITE_SHAREPOINT_SITE_URL` in `.env` to limit search to one site, for example `https://yourtenant.sharepoint.com/sites/YourSite`.

Restart `npm run dev` after changing `.env`.

### AADSTS50011 (redirect URI does not match)

The sign-in popup sends `http://localhost:5173`. That exact value must exist on the app as a **SPA** redirect URI (not Web, not `https`, no trailing slash).

1. Open [App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and select this app (client ID in `.env`).
2. **Authentication** → **Add a platform** → **Single-page application**.
3. Add `http://localhost:5173` and save.
4. If it was already listed under **Web**, remove it from Web and add it under **SPA**. MSAL popup login requires SPA.
5. Try **Sign in with Microsoft** again. No app restart is needed for this Azure change.

## How it works

1. You upload a PDF or image. A first-page preview is shown.
2. Digital PDFs use the embedded text. Scanned PDFs and photos use [Tesseract.js](https://tesseract.projectnaptha.com/) OCR in the browser.
3. The detected name is shown large so you can confirm or edit it.
4. After you sign in, Microsoft Graph Search lists matching SharePoint files. Upload-to-subsite is not implemented yet.

## Notes

- OCR is strongest on clear, high-contrast photos. Handwriting is unreliable.
- Graph Search only returns items your signed-in account can already open.
- Production hosting needs a new SPA redirect URI for that HTTPS origin, plus IT approval for the Graph permissions.
