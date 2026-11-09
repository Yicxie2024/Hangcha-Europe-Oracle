Hangcha Sales Assistant v1.5.2 -- Teams upload button requires completed scan

PATCH FOR v1.5.1 (not a full installer)

1. Close Hangcha Sales Assistant and its Teams upload window.
2. Back up your current gui.py in your existing app folder.
3. Copy gui.py from this ZIP into that same app folder, replacing gui.py.
   Example from user's recent install: C:\Hangcha Assistant\gui.py
4. Reopen 10_START_GUI.bat. No Python/requirements re-install needed.

Changes (GUI ONLY):
- Teams Upload button starts disabled; only a fully completed and current PDF scan enables it.
- Scan in progress, cancelled, failed, with inaccessible source directory, or with changes
  to the source list/employee/Excel setting: upload button remains disabled until rescan.
- Direct invocation of the upload action also checks readiness.
- Before rescanning, an already open Teams window must be closed with user confirmation
  to avoid stale results. An active Teams operation prevents starting a new scan.
- No changes to PDF classification, Excel writing, invoices, Teams drag/drop or Graph.
- No automatic uploads or other mutations are enabled by this patch.

Tested: original 109 tests + 7 new gated-upload tests, 116 tests passed
         (Linux Xvfb/Tkinter; test on your Windows 11 computer as well).
