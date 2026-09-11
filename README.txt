POST DECOM TOOL
===============

Launch:
  Double-click "Launch Decommissioning Tool.cmd".

Files:
  Decommissioning GUI.ps1       Main graphical application.
  Service deletion.ps1          Standalone service-deletion worker.
  IIS site and pool deletion.ps1
                                Standalone IIS worker.
  Folder deletion.ps1           Standalone exact-folder-deletion worker.
  Local Admin Cleanup.ps1       Standalone local-Administrators cleanup worker.
  Scan inventory.ps1            Read-only combined discovery worker used by SCAN.
  ODBC System DSN Cleanup.ps1   Discovers and removes confirmed 32/64-bit System DSNs.

GUI workflow:
  1. Enter the service servers and application/IIS servers.
  2. Enter one instance token. Service discovery searches for that literal,
     case-insensitive token in the account-name portion of each Log On As
     value. For example, IM35880 matches both Olympus\APIM35880 and
     APIM35880@olympus.gaia.kosmos.
  3. Click "Choose Client Folders..." to list M:\OBOL\Clients across both
     server lists. The picker displays each full UNC path; check the exact
     server/folder rows to target.
     UNC paths containing the current IIS instance token are highlighted in
     green using a case-insensitive comparison. Paths that do not contain the
     token are highlighted in red.
     "Open Selected Paths" opens checked rows (or highlighted rows when
     nothing is checked) in separate File Explorer windows.
  4. Click "SCAN" for a read-only combined inventory. Results appear inside
     the main window, directly above Background activity, with separate sortable
     grid tabs for matching services, IIS sites, application pools, 32-bit and
     64-bit ODBC System DSNs, local Administrators principals, and every client folder. Green folder rows
     contain the instance token; blocked application pools appear in red. SCAN
     never changes configuration.
     After a scan returns results, click "Export Scan Results to Excel..." to create an
     Excel .xlsx workbook. Services, IIS Sites, Application Pools, ODBC Data
     Sources, Local Admins, and Client Folders are saved as separate named
     worksheets with filters, frozen header rows, scan metadata, and the
     category-specific details shown by the tool.
     When name-only service or application-pool candidates are found, SCAN
     opens a review window with separate Services and Application Pools tabs.
     Candidates are unchecked by default. Only checked rows are added to the
     normal read-only Scan Results grids and to the corresponding action's
     later confirmation and exact target manifest.
  5. Run "What If (All)" and review the activity log.
  6. Run the required destructive action button. Delete Services, IIS Pool and
     Site Decom, Local Admin Cleanup, and ODBC DSN Cleanup reuse the latest
     successful Scan Results instead of running a second discovery pass. The
     scan must still match the current inputs and be no more than 30 minutes old.

Background activity text log:
  - Each normal application session automatically creates one UTF-8 text file
    under the tool's Logs folder. The name is:
      BackgroundActivity_yyyyMMdd_HHmmss_<user>_<session-id>.txt
  - Every line displayed in Background activity is appended to that file
    immediately, including scan output, What If output, warnings, errors, and
    operation-report locations.
  - "Clear log" clears only the on-screen display. The session text file and
    its existing entries are retained.
  - Preview/render modes used by automated UI tests do not create session logs.

Safety:
  - Destructive GUI confirmation defaults to No.
  - Scan results and Background activity are separated by a movable splitter,
    allowing either section to be resized.
  - Background actions show the current server and step progress.
  - Every destructive or WhatIf worker run writes a timestamped CSV operation
    report under the Logs folder. Reports contain one row per confirmed target
    with server, category, action, outcome, exit code, and message.
  - The GUI passes -Force to workers only after confirmation.
  - WhatIf always takes precedence over Force.
  - Protected Windows service accounts are excluded from token matching.
    A service whose name contains the token but whose Log On As does not is
    shown separately as an opt-in candidate; it is never included merely by
    running Delete Services. Checked candidates are re-discovered by exact
    server/name and the name-only rule is revalidated before deletion.
  - "Delete Services" uses the services and GCS entries already displayed by
    the latest valid SCAN. It shows every server, service name, display name,
    state, and Log On As account before confirmation. SCAN also discovers lines
    containing the instance token in these files on every service server:
      C:\OBOL\Utilities\GCSserviceStartRestarter\MonitoredServices.txt
      C:\OBOL\Utilities\GCSserviceStartRestarter\ServicestoRestart.txt
  - The confirmation window shows both services and GCS list entries. Only
    listed targets enter the exact confirmation manifest, and each token match
    is revalidated before modification.
  - GCS list files are replaced atomically and receive timestamped .bak files.
    Comment lines beginning with # or ; are preserved. If a confirmed service
    fails to stop or delete, GCS list cleanup is skipped on that server.
  - "IIS Pool and Site Decom" uses the latest valid SCAN and shows every
    matching site, internally retained child application, eligible pool, and
    blocked pool. The intentionally hidden child-application records remain in
    the exact IIS removal plan without restoring a separate applications tab.
  - Only IIS rows marked DELETE in that confirmation window are placed in the
    exact target manifest; live pool dependencies are checked again before
    removal.
  - Shared IIS app pools are blocked.
  - An application pool whose name matches the IIS token but whose configured
    pool identity does not is shown separately as an opt-in candidate. It is
    included only when checked in the latest valid scan, and the normal live
    consumer/dependency protection still applies.
  - Default Web Site is never targeted.
  - Matching non-root IIS applications beneath preserved non-default sites
    are removed before application pools; parent sites remain.
  - Removing an IIS application changes IIS configuration only. Its physical
    files remain for the separate Folder Deletion action.
  - IIS pools are listed independently when their names match the token.
  - A matching pool is blocked when any consumer outside the planned site and
    child-application removals, including Default Web Site, still references it.
  - Folder names must be exact leaf names; paths and traversal are rejected.
  - Folder discovery uses the combined service and application server lists.
  - Folder deletion affects only the manually checked server/folder rows and
    shows the complete target list in a final confirmation window.
  - "Local Admin Cleanup" searches the combined service and application
    server lists. It matches the instance token as a literal,
    case-insensitive substring and targets group principals only.
  - The built-in Administrators group is resolved by SID on each server, so
    localized Windows group names are supported.
  - Local-admin targets are shown with their exact ADSI paths before removal.
    The cleanup button uses the latest valid SCAN rather than repeating discovery.
    Read-only SCAN does not create a separate LocalAdmins_Audit CSV; cleanup and
    WhatIf cleanup runs continue to create the audit report.
  - "ODBC DSN Cleanup" searches both 64-bit and 32-bit machine registry views
    during SCAN across the combined server lists. The cleanup button uses those
    cached results, shows architecture, driver, data server, database, and
    registry path, then deletes only exact confirmed targets after live revalidation.
    Each membership is revalidated before and after removal, and every run
    writes a timestamped CSV audit report under the Logs folder.

IIS match modes:
  Token     Delimiter-bounded match (recommended).
  Exact     Whole-name match.
  Contains  Broad substring match.
