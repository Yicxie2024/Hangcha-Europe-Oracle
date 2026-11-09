Hangcha Sales Assistant v1.5.1 | Windows Teams Drag-and-Drop Patch
===================================================================

This patch adds a second way to transfer sales invoice PDFs into the Teams
channel folder without Microsoft Entra / Graph permissions.

Upgrade an installed v1.5.0
--------------------------
1. Close Hangcha Sales Assistant.
2. Back up your application folder (e.g. C:\Hangcha Assistant).
3. Copy these four files into that application folder, replacing existing ones:
      gui_teams.py     REPLACE
      teams_drag.py    ADD
      teams_store.py   REPLACE
      requirements.txt REPLACE
4. Run 01_INSTALL.bat ONCE to install tkinterdnd2 0.7.x into the app's .venv.
5. Start Hangcha Assistant normally via 10_START_GUI.bat / desktop shortcut.
6. Scan local invoice PDFs first, then click the existing Teams invoice button.

Daily drag-to-Teams workflow
----------------------------
1. Open the correct Teams channel's Shared -> Documents -> current MONTH ->
   current DATE folder, e.g. '10月' -> '30.10'.
2. Open the Teams upload workspace in Hangcha Assistant (default mode is
   '拖拽到 Teams（无需 Entra）').
3. Check the invoices to transfer. Click '准备勾选发票（需确认）' and review the
   document list and the expected folder name.
4. HOLD the mouse button on the green drag area and DRAG to the files list
   of the open Teams DATE folder. Release the mouse to hand the files to Teams.
5. WAIT for Teams to finish uploading and verify the new PDFs are present.
6. Optionally click '已在 Teams 核实 → 标记完成' to keep them out of the future
   pending upload list. This is a MANUAL CONFIRMATION, not a Graph API check.

Windows native file drag support
--------------------------------
* Depends on tkinterdnd2 >= 0.7.0 (included in updated requirements.txt).
* Uses native CF_HDROP file drag/copy, not plain text and not moving originals.
* A special dragging surface is used instead of the checkbox/table row, to
  avoid conflicts with current sorting, checkbox selection, and right-click.
* Run the app normally (NOT as Administrator). Windows integrity levels can
  prevent drag/drop between an elevated and a normal Teams window.
* Teams still controls upload and may show conflicts or upload messages.
* The app cannot prove that Teams received or saved a file; check in Teams.

Fallback
--------
'备用：在资源管理器中拖拽' creates throwaway copies of the explicitly selected
PDFs in the current Windows user's LOCALAPPDATA and opens that folder. In
Explorer press Ctrl+A, then drag into the Teams DATE folder. The copies are
removed when the upload workspace closes (provided cleanup succeeds).
The original PDFs are never renamed, deleted, or moved.

Important business rules preserved
---------------------------------
* Does not alter Excel, PDF classification, order matching or archived records.
* VOID ('作废') invoices are not eligible. Duplicate invoice numbers, invalid
  PDFs and unreviewed files are blocked. Files tagged '有误' receive an extra
  confirmation.
* Desktop computer's date is for the RECOMMENDED folder only. The actual
  Teams drop location must be verified by the user.
* Graph mode is still available in a separate advanced option if the company
  later supplies Microsoft 365 authorization.
* Manual upload confirmation is stored locally in the existing employee-
  isolated teams_upload_history.sqlite3 database. It is labeled separately
  from an API-verified upload.

Validation
----------
109 offline Python tests and GUI start/mode-switch smoke tests passed.
Actual drag/drop from Windows 11 to a live company Teams folder has NOT been
verified here and must be tried with a single non-sensitive test PDF first.
