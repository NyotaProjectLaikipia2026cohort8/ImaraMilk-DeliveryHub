# Imara Milk Delivery Hub — PWA package

## Included files
- `index.html` — your uploaded `index(6).html`, updated with PWA metadata, mobile safe-area support, a clean Google Maps container, corrected Apps Script JSONP transport, and field normalization for the existing backend responses.
- `Code.gs` — your uploaded `code(20261113-185032).gs`, preserving the existing functions, sheet names, column headers, and record structures. Script locks were added around database setup and the customer/order/message write transactions. The existing `savePrimaryLocation()` lock remains.
- `manifest.webmanifest`, `sw.js`, and `icons/` — install metadata, app-shell caching, and Android/iOS icons.

## Deploy to GitHub Pages
1. Unzip this package.
2. Upload `index.html`, `manifest.webmanifest`, `sw.js`, and the complete `icons` folder to the same GitHub Pages root directory.
3. Ensure GitHub Pages is served over HTTPS.

## Deploy the Apps Script backend
1. Open the existing Apps Script project connected to the Imara Milk spreadsheet.
2. Replace its `Code.gs` with the included `Code.gs` and save.
3. Run `setupImaraDatabase()` only if you have not already initialized the existing sheets. It preserves existing columns and appends any missing columns; it does not rename or remove columns.
4. Deploy > Manage deployments > Edit the existing Web App deployment, select a new version, and deploy.
5. Confirm that the `/exec` URL in `CONFIG.API_BASE` in `index.html` is the URL for this exact Apps Script project. If it differs, update only that value before publishing.

## Important limitations / checks
- The frontend uses JSONP GET requests because GitHub Pages cannot directly read Apps Script cross-origin JSON POST responses without CORS support. The included backend already supports JSONP through `doGet`.
- The existing backend does not implement file uploads or attachment storage. The UI now reports that limitation instead of pretending an attachment was saved.
- The existing `summarizeConversation()` function returns saved messages, not an AI-generated summary; the UI label was changed to Conversation details.
- Offline mode caches the app shell only. Live orders, messages, Google Maps tiles, and Google Sheets data still require an internet connection. Do not interpret cached shell availability as offline database synchronization.
- The current Apps Script API has no user authentication/authorization layer. Before exposing this operations dashboard or write endpoints to the public, add server-side authentication and role checks. Do not treat an obscure `/exec` URL as security.
- Restrict the existing Google Maps browser API key by HTTP referrer to your GitHub Pages domain and enable only the APIs you need. The key is visible in a browser app by design.
- iPhone/iPad: open the HTTPS site in Safari, tap Share, then Add to Home Screen. The `apple-touch-icon.png` is provided for the home-screen icon.

## Existing database schema preserved
`Customers`, `UserLocations`, `Orders`, `OrderStatusHistory`, `Messages`, and `Staff` remain as defined in the source `Code.gs`. No new spreadsheet tabs or columns are required for the PWA shell.
