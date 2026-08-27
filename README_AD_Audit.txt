===============================================================================
 ACTIVE DIRECTORY ACCOUNT AUDIT  -  QUICK READING GUIDE
 File: AD-Account-Audit-GUI.ps1
===============================================================================

PURPOSE
-------
Audit a list of AD accounts to know whether they are still in use
(active / inactive / disabled / expired / never logged on / not found) and
produce a usable Excel report (.xlsx).


-------------------------------------------------------------------------------
1. HOW TO RUN
-------------------------------------------------------------------------------
STA mode is required (GUI):

    powershell.exe -STA -File .\AD-Account-Audit-GUI.ps1

Prerequisites: ActiveDirectory module (RSAT) installed, and AD read rights.
Stay on powershell.exe (5.1). On pwsh 7, remember the -STA switch too.


-------------------------------------------------------------------------------
2. USAGE IN 4 STEPS
-------------------------------------------------------------------------------
1) "Account list (.txt)"  -> Browse button: pick the text file
   (one identifier per line: SamAccountName, UPN, UPN prefix, mail,
    mail prefix, mailNickname, DisplayName or Name - case-insensitive).
2) "Output CSV/XLSX"      -> Browse button: choose where to save the report
   (auto-filled in the same folder, as .xlsx).
3) Adjust thresholds if needed (see section 4) and tick options.
4) "Run Audit". The .xlsx report is generated and saved automatically at the
   end. Use "Export XLSX" to re-export on demand.

The .xlsx opens with a double-click: frozen header, auto-filter, rows coloured
by status. No more delimiter issues.


-------------------------------------------------------------------------------
3. HOW TO READ THE KEY COLUMNS
-------------------------------------------------------------------------------
InputValue    The exact value read from the file (useful for tracing).

ResolvedBy    HOW the account was matched:
              - SamAccountName / UPN / UPNPrefix / mail / mailPrefix /
                mailNickname / DisplayName / Name  = the key that matched.
              - "Ambiguous (xxx)"  = several accounts match this entry.
              - NotFound / Error   = no account matched.

MatchCount    Number of AD accounts found for this entry:
              - 0  = not found
              - 1  = unique (normal case)
              - >1 = DUPLICATE: several accounts share this identifier.
                     Each account gets its own row, same InputValue.
              --> Excel filter "MatchCount > 1" = all duplicates.

AuditStatus   The audit verdict (see section 5).

Other useful columns: LastLogon (last logon), DaysSinceLogon (days since),
PasswordLastSet, AccountExpirationDate, Enabled, LockedOut, Department, Title,
DistinguishedName (OU location).


-------------------------------------------------------------------------------
4. OPTIONS AND THRESHOLDS
-------------------------------------------------------------------------------
Inactive warning (days)   "Warning" threshold (default 90).
Inactive critical (days)  "Critical" threshold (default 180). Must be > warning.

Accurate last logon (query all DCs - slower)
   Unchecked by default. In normal mode, LastLogon comes from
   lastLogonTimestamp, replicated only every ~14 days: an account may look
   inactive when it is not. This option queries EVERY domain controller for
   the true last logon: far more accurate, but noticeably slower.
   Tick it before making a real disable decision.


-------------------------------------------------------------------------------
5. AUDIT STATUSES AND COLOURS
-------------------------------------------------------------------------------
Active             Active account, recent logon.               (white)
Inactive > 90d     Inactive beyond the "warning" threshold.    (light yellow)
Inactive > 180d    Inactive beyond the "critical" threshold.   (salmon)
Never Logged On    Enabled but no recorded logon.              (yellow)
Expired            Account expiration date has passed.          (orange)
Disabled           Account disabled in AD.                      (grey)
Not Found          No account found in the domain.              (red)

Precedence: a disabled account is marked "Disabled" even if it is also
inactive. Read the status from the strongest to the weakest.


-------------------------------------------------------------------------------
6. READING FOR DECISIONS (cheat sheet)
-------------------------------------------------------------------------------
DELETE / cleanup candidates:
   - Disabled + inactive for a long time.
   - Not Found = account already gone (confirm in ADUC).

REVIEW / disable candidates:
   - Enabled + "Inactive > 180d".
   - Never Logged On + old WhenCreated (caution: may be a SERVICE account,
     which never logs on interactively).
   - Expired: already expired, decide.

DUPLICATES (MatchCount > 1):
   - Compare the rows with the same InputValue. Classic case: one active
     account + one disabled/inactive -> keep the active one, handle the other.

KEEP:
   - Active with a recent logon.


-------------------------------------------------------------------------------
7. USEFUL EXCEL FILTERS
-------------------------------------------------------------------------------
- AuditStatus = "Not Found"        -> accounts absent from the domain.
- AuditStatus contains "Inactive"  -> all inactive accounts.
- MatchCount > 1                   -> all duplicates.
- Enabled = True AND DaysSinceLogon > 180 -> enabled but dormant.
- Sort by LastLogon (ISO format: text sort = chronological sort).


-------------------------------------------------------------------------------
8. LIMITS TO KNOW (important)
-------------------------------------------------------------------------------
- SINGLE DOMAIN: the script queries the session's domain. Accounts from another
  domain would wrongly show as "Not Found" (not applicable here, single
  domain confirmed).
- LastLogon has ~14-day latency in normal mode (see option in section 4).
- Service accounts: often "Never Logged On" while actually in use.
  Do not disable them on this criterion alone.
- "Not Found" = the account does not exist in the domain (likely deleted/
  offboarded), OR the supplied identifier is not covered by the index
  (e.g. full DN, objectGUID, secondary SMTP alias).
- The total row count may exceed the number of file entries
  (normal: one ambiguous entry = several rows).
- A fallback CSV (";" delimiter + BOM) is generated automatically if the
  .xlsx write fails.


-------------------------------------------------------------------------------
9. TROUBLESHOOTING
-------------------------------------------------------------------------------
- Window frozen for a few seconds at the end: normal (.xlsx being written).
- Scrolling bar ("Loading / Indexing"): one-time directory load.
- Many unexpected "Not Found": check the FORMAT of the identifiers in the file
  (the ResolvedBy column shows which key the others matched on).
- Always confirm any edge case in ADUC before taking action.

===============================================================================
