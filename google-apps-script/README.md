# Portfolio contact and admin backend

The portfolio's admin login and dashboard are rendered inside the portfolio HTML. This Apps Script project provides the protected login, request listing, Sheets storage, and email actions; it does not need a separate `Admin.html` file.

## Deploy

1. Replace the Apps Script project's `Code.gs` with this folder's `Code.gs`.
2. In **Project Settings → Script Properties**, set:
   - `ADMIN_EMAIL`: the admin sign-in email.
   - `ADMIN_PASSWORD`: a new, private password. Do not reuse a password shared in chat.
   - `PORTFOLIO_ORIGIN`: the exact origin where the portfolio is hosted, such as `https://portfolio.example`. Do not include a path or trailing slash. The local `file://` preview is not a valid origin for authenticated use.
3. Save, then use **Deploy → Manage deployments → Edit → New version → Deploy**. Keep **Execute as** set to your account and enable access for visitors who need to submit the public contact form.
4. Keep the deployed web-app URL ending in `/exec` in the portfolio's `APPS_SCRIPT_URL`.
5. Authorize the requested Google Sheets and Gmail permissions when prompted. On the first contact submission the script creates a spreadsheet named `Mujtaba Portfolio Contact Requests` and a `Requests` tab. You can optionally set `SHEET_ID` to an existing spreadsheet ID.
6. Open the hosted portfolio, triple-click the footer copyright or `Mujtaba` wordmark, and sign in. The dashboard lists requests and can send replies through the Google account that authorized the script.

The `PORTFOLIO_ORIGIN` allowlist ensures admin session tokens are only posted back to the portfolio's trusted origin. The password and session token are not stored in the portfolio HTML.
