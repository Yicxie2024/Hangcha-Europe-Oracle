Hangcha Sales Assistant v1.5.0 - Teams invoice upload workspace
================================================================

What this release changes
-------------------------
* Adds a prominent "☁ 上传发票" button to the existing main toolbar.
* Opens a SEPARATE Teams invoice workspace. No new columns in the order tree.
* Previews eligible local Sales Invoice PDFs and supports search, sorting,
  checkboxes, batch uploads, and a right-click menu.
* Uses the WINDOWS COMPUTER'S CURRENT DATE (not PDF Invoice date) to locate
  the folder: configured channel Documents root -> 10月 -> 30.10 -> PDFs.
* Does not require or create a nested 发票 folder.
* Supports optional one-time manual folder link override, reviewed each upload.
* Keeps local per-employee upload history and failed attempts in SQLite.
* Never uploads void invoices, credit notes, or unverifiable PDFs automatically.
* Never replaces a remote file with the same name.
* Does not modify or depend on Excel writes, archives or payment statuses.

Upgrade an existing installation
--------------------------------
1. Close the app. Back up its installation directory.
2. Copy into the existing app directory:
     gui.py             (replace)
     gui_teams.py       (add)
     teams_graph.py     (add)
     teams_store.py     (add)
     requirements.txt   (replace)
3. Run 01_INSTALL.bat once to install additional msal + requests dependencies.
4. Restart with 10_START_GUI.bat or the desktop shortcut.
5. In the main app first run the PDF scanner. Then click "☁ 上传发票".

Microsoft 365 prerequisites (IT/admin)
-------------------------------------
The app intentionally contains NO company login or client secret.
Ask company Microsoft 365 IT to register and approve a single-tenant
Microsoft Entra public desktop client for Hangcha Sales Assistant.
* Redirect URI: http://localhost (Mobile and desktop applications platform)
* Delegated Microsoft Graph permissions: Files.ReadWrite, User.Read
* Consent/admin approval as required by tenant policies.
* The user must have actual write access to the private channel's SharePoint
  folder. App registration does not grant any folder access by itself.
* Provide the user with Tenant ID and Application (Client) ID.
No client secret, company password or long-lived tokens are stored by the app.
You can also use the organization directory-domain for Tenant ID.

First-time folder setup
-----------------------
1. Enter company Tenant ID and approved Client ID, click Microsoft 365 login.
2. Paste the SHAREPOINT FOLDER link (:f: format) you previously copied
   from a DATE folder (e.g. 29.10), or a MONTH folder (e.g. 10月).
3. Click "验证目标根目录". App resolves the common month-root through
   Microsoft Graph IDs. It does NOT guess based on team/channel names.
4. Click "检查云端 / 预览待上传" to locate today's month/date folder and
   query today's remote file names.
5. Select eligible invoices; check exact target and full file list;
   explicitly confirm. Nothing is uploaded before confirmation.

Important known limitations
---------------------------
* Real Microsoft Graph authentication and private-channel upload require
  testing with your company's approved account and permissions. Offline
  tests cannot verify access to the actual SharePoint folder.
* Initial upload history is empty; historical PDFs may be shown as pending
  even if previously uploaded by other means. Check the cloud destination
  before selecting old invoices. The app checks duplicates only within the
  selected day's folder, not across every historic day folder.
* The current date folder and month folder must already exist. No fallback
  or automatic creation. A new year requires folder-root re-verification.
* Full pending list requires an up-to-date local PDF scan.
* Upload status is independent of the Excel invoice status and PDF archive.
* Microsoft 365 sign-in and every upload require user action; no background
  uploading or unattended Teams write is enabled.
