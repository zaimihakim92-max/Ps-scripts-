#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Active Directory account audit tool with a WinForms GUI.

.DESCRIPTION
    Reads a list of account identifiers (one per line) from a text file, resolves
    each entry against Active Directory, classifies whether the account is still in
    use, shows the result in a live grid, and exports a full CSV report.

    Key improvements over the original script:
      * Full WinForms GUI: input/output pickers, options, live progress bar,
        colour-coded results grid, summary counters and one-click CSV export.
      * Non-blocking: the audit runs in a background runspace, so the window stays
        responsive and shows real-time progress.
      * "No user missed" resolution: every entry is tried by -Identity first, then by
        an attribute filter (SamAccountName / UPN / mail / Name / DisplayName). Each
        row records how it was resolved (ResolvedBy) and keeps the original InputValue.
        Not-found and ambiguous entries are reported, never silently skipped.
      * O(n) collection (synchronized list) instead of array "+=".
      * Optional accurate last-logon (queries every DC for the true lastLogon instead
        of the replicated, ~14-day-latent LastLogonDate).
      * De-duplicated input, cancellable run, configurable inactivity thresholds.

.NOTES
    WinForms needs STA. Launch with:
        powershell.exe -STA -File .\AD-Account-Audit-GUI.ps1
        pwsh.exe       -STA -File .\AD-Account-Audit-GUI.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Shared, thread-safe state between the UI thread and the worker runspace
# ---------------------------------------------------------------------------
$sync            = [hashtable]::Synchronized(@{})
$sync.Results    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$sync.Processed  = 0
$sync.Total      = 0
$sync.Done       = $false
$sync.Cancel     = $false
$sync.Error      = $null
$sync.Status     = $null

# Column order used everywhere (grid + CSV)
$script:Columns = @(
    'InputValue','ResolvedBy','MatchCount','SamAccountName','UserPrincipalName','DisplayName',
    'Enabled','LockedOut','Expired','AccountExpirationDate','LastLogon',
    'DaysSinceLogon','PasswordLastSet','DaysSincePassword','PasswordNeverExpires',
    'WhenCreated','WhenChanged','Department','Title','Description',
    'DistinguishedName','AuditStatus'
)

# ---------------------------------------------------------------------------
# Worker (runs inside the background runspace)
# ---------------------------------------------------------------------------
$worker = {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $props = @(
            'Enabled','LastLogonDate','PasswordLastSet','AccountExpirationDate',
            'WhenCreated','WhenChanged','LockedOut','Department','Title',
            'Description','UserPrincipalName','PasswordNeverExpires','mail',
            'mailNickname','DisplayName','SamAccountName','Name'
        )

        # DC list is only needed for the accurate last-logon option
        $dcList = @()
        if ($config.AccurateLastLogon) {
            try { $dcList = @((Get-ADDomainController -Filter * -ErrorAction Stop).HostName) } catch { $dcList = @() }
        }

        function Get-TrueLastLogon {
            param($Sam, $DCs)
            $best = $null
            foreach ($dc in $DCs) {
                try {
                    $ll = (Get-ADUser -Identity $Sam -Server $dc -Properties lastLogon -ErrorAction Stop).lastLogon
                    if ($ll) {
                        $dt = [DateTime]::FromFileTime([int64]$ll)
                        if (-not $best -or $dt -gt $best) { $best = $dt }
                    }
                } catch { }
            }
            return $best
        }

        # Make free-text AD fields (Description, Title, etc.) CSV-safe: strip the
        # line breaks / tabs that AD often stores, which otherwise shift columns.
        function Clean-Text {
            param($Value)
            if ($null -eq $Value) { return '' }
            $s = [string]$Value -replace "[\r\n\t]+", ' '
            $s = $s -replace '\s{2,}', ' '
            return $s.Trim()
        }

        # ------------------------------------------------------------------
        # 1) Load the whole directory ONCE (fast, single query) and build a
        #    set of per-attribute indexes so any identifier format resolves:
        #    SamAccountName, full UPN, UPN prefix (before @), full mail,
        #    mail prefix, mailNickname, DisplayName, Name (CN).
        #    This is why "abbie.briggs" (a UPN/mail prefix, SAM = "briggs")
        #    now resolves instead of being reported as Not Found.
        # ------------------------------------------------------------------
        $sync.Status = 'Loading directory (one-time)...'
        $allUsers = Get-ADUser -Filter * -Properties $props -ResultPageSize 2000 -ErrorAction Stop

        $sync.Status = 'Indexing directory...'
        $indexNames = @('SamAccountName','UPN','UPNPrefix','mail','mailPrefix','mailNickname','DisplayName','Name')
        $idx = @{}   # per-attribute : key(lower) -> user
        $amb = @{}   # per-attribute : key(lower) -> $true when >1 distinct user shares it
        foreach ($n in $indexNames) { $idx[$n] = @{}; $amb[$n] = @{} }

        function Add-Key {
            param($Name, $Key, $User)
            if ([string]::IsNullOrWhiteSpace($Key)) { return }
            $k = ([string]$Key).Trim().ToLowerInvariant()
            $h = $idx[$Name]
            if ($h.ContainsKey($k)) {
                $existing = $h[$k]
                if ($amb[$Name].ContainsKey($k)) {
                    # already a collision list: add this user if not already present
                    $list = $amb[$Name][$k]
                    $seen = $false
                    foreach ($x in $list) { if ($x.DistinguishedName -eq $User.DistinguishedName) { $seen = $true; break } }
                    if (-not $seen) { [void]$list.Add($User) }
                }
                elseif ($existing.DistinguishedName -ne $User.DistinguishedName) {
                    # first collision on this key: start a list holding BOTH accounts
                    $list = New-Object System.Collections.ArrayList
                    [void]$list.Add($existing)
                    [void]$list.Add($User)
                    $amb[$Name][$k] = $list
                }
            } else {
                $h[$k] = $User
            }
        }

        foreach ($u in $allUsers) {
            Add-Key 'SamAccountName' $u.SamAccountName $u
            Add-Key 'DisplayName'    $u.DisplayName    $u
            Add-Key 'Name'           $u.Name           $u
            Add-Key 'mailNickname'   $u.mailNickname   $u
            if ($u.UserPrincipalName) {
                Add-Key 'UPN'       $u.UserPrincipalName $u
                Add-Key 'UPNPrefix' ($u.UserPrincipalName -split '@')[0] $u
            }
            if ($u.mail) {
                Add-Key 'mail'       $u.mail $u
                Add-Key 'mailPrefix' ($u.mail -split '@')[0] $u
            }
        }

        # Look up one input value against the indexes, in priority order.
        # Returns ALL matching accounts (a list), so duplicates are surfaced.
        function Resolve-FromIndex {
            param($Value)
            $k = ([string]$Value).Trim().ToLowerInvariant()
            if ($k -eq '') { return [PSCustomObject]@{ Users = @(); By = 'NotFound' } }
            foreach ($n in $indexNames) {
                if ($amb[$n].ContainsKey($k)) { return [PSCustomObject]@{ Users = @($amb[$n][$k]); By = $n } }
                if ($idx[$n].ContainsKey($k)) { return [PSCustomObject]@{ Users = @($idx[$n][$k]); By = $n } }
            }
            return [PSCustomObject]@{ Users = @(); By = 'NotFound' }
        }

        # Build a full result row for one resolved account.
        function New-ResultRow {
            param($InputVal, $By, $Count, $U, $Now, $Config, $DcList)

            $lastLogon = $U.LastLogonDate
            if ($Config.AccurateLastLogon -and $DcList.Count -gt 0) {
                $realLL = Get-TrueLastLogon -Sam $U.SamAccountName -DCs $DcList
                if ($realLL) { $lastLogon = $realLL }
            }
            $daysLogon = $null
            if ($lastLogon)         { $daysLogon = [math]::Round(($Now - $lastLogon).TotalDays, 0) }
            $daysPwd = $null
            if ($U.PasswordLastSet) { $daysPwd   = [math]::Round(($Now - $U.PasswordLastSet).TotalDays, 0) }

            $expired = $false
            if ($U.AccountExpirationDate -and $U.AccountExpirationDate -lt $Now) { $expired = $true }

            $status = 'Active'
            if     (-not $U.Enabled)                 { $status = 'Disabled' }
            elseif ($expired)                        { $status = 'Expired' }
            elseif (-not $lastLogon)                 { $status = 'Never Logged On' }
            elseif ($daysLogon -gt $Config.CritDays) { $status = "Inactive > $($Config.CritDays)d" }
            elseif ($daysLogon -gt $Config.WarnDays) { $status = "Inactive > $($Config.WarnDays)d" }

            return [PSCustomObject]@{
                InputValue            = $InputVal
                ResolvedBy            = $By
                MatchCount            = $Count
                SamAccountName        = $U.SamAccountName
                UserPrincipalName     = $U.UserPrincipalName
                DisplayName           = (Clean-Text $U.Name)
                Enabled               = $U.Enabled
                LockedOut             = $U.LockedOut
                Expired               = $expired
                AccountExpirationDate = $U.AccountExpirationDate
                LastLogon             = $lastLogon
                DaysSinceLogon        = $daysLogon
                PasswordLastSet       = $U.PasswordLastSet
                DaysSincePassword     = $daysPwd
                PasswordNeverExpires  = $U.PasswordNeverExpires
                WhenCreated           = $U.WhenCreated
                WhenChanged           = $U.WhenChanged
                Department            = (Clean-Text $U.Department)
                Title                 = (Clean-Text $U.Title)
                Description           = (Clean-Text $U.Description)
                DistinguishedName     = (Clean-Text $U.DistinguishedName)
                AuditStatus           = $status
            }
        }

        # ------------------------------------------------------------------
        # 2) Match every input line against the index (in-memory, O(1) each).
        #    One input can yield several rows when several accounts match
        #    (duplicates) - each is emitted with its real status + MatchCount.
        # ------------------------------------------------------------------
        $sync.Status = $null   # switch back to processed/total display
        foreach ($acc in $accounts) {
            if ($sync.Cancel) { break }
            $now = Get-Date

            try   { $res = Resolve-FromIndex -Value $acc }
            catch { $res = [PSCustomObject]@{ Users = @(); By = 'Error' } }

            $matched = @($res.Users)
            if ($matched.Count -gt 0) {
                $by = if ($matched.Count -gt 1) { "Ambiguous ($($res.By))" } else { $res.By }
                foreach ($u in $matched) {
                    [void]$sync.Results.Add(
                        (New-ResultRow -InputVal $acc -By $by -Count $matched.Count -U $u -Now $now -Config $config -DcList $dcList)
                    )
                }
            }
            else {
                [void]$sync.Results.Add([PSCustomObject]@{
                    InputValue            = $acc
                    ResolvedBy            = $res.By
                    MatchCount            = 0
                    SamAccountName        = $acc
                    UserPrincipalName     = ''
                    DisplayName           = ''
                    Enabled               = ''
                    LockedOut             = ''
                    Expired               = ''
                    AccountExpirationDate = ''
                    LastLogon             = ''
                    DaysSinceLogon        = ''
                    PasswordLastSet       = ''
                    DaysSincePassword     = ''
                    PasswordNeverExpires  = ''
                    WhenCreated           = ''
                    WhenChanged           = ''
                    Department            = ''
                    Title                 = ''
                    Description           = ''
                    DistinguishedName     = ''
                    AuditStatus           = 'Not Found'
                })
            }
            $sync.Processed++
        }
    }
    catch {
        $sync.Error = $_.Exception.Message
    }
    finally {
        $sync.Done = $true
    }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
$form                = New-Object System.Windows.Forms.Form
$form.Text           = 'Active Directory Account Audit'
$form.Size           = New-Object System.Drawing.Size(1180, 740)
$form.StartPosition  = 'CenterScreen'
$form.MinimumSize    = New-Object System.Drawing.Size(960, 620)

# ----- Top panel (inputs & options) ----------------------------------------
$panelTop        = New-Object System.Windows.Forms.Panel
$panelTop.Dock   = 'Top'
$panelTop.Height = 170

# Row 1 : input file
$lblInput          = New-Object System.Windows.Forms.Label
$lblInput.Text     = 'Account list (.txt):'
$lblInput.Location = New-Object System.Drawing.Point(12, 15)
$lblInput.Size     = New-Object System.Drawing.Size(140, 20)

$txtInput          = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(155, 12)
$txtInput.Size     = New-Object System.Drawing.Size(600, 22)
$txtInput.ReadOnly = $true

$btnInput          = New-Object System.Windows.Forms.Button
$btnInput.Text     = 'Browse...'
$btnInput.Location = New-Object System.Drawing.Point(765, 11)
$btnInput.Size     = New-Object System.Drawing.Size(100, 24)

# Row 2 : output CSV
$lblOutput          = New-Object System.Windows.Forms.Label
$lblOutput.Text     = 'Output CSV:'
$lblOutput.Location = New-Object System.Drawing.Point(12, 48)
$lblOutput.Size     = New-Object System.Drawing.Size(140, 20)

$txtOutput          = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(155, 45)
$txtOutput.Size     = New-Object System.Drawing.Size(600, 22)
$txtOutput.ReadOnly = $true

$btnOutput          = New-Object System.Windows.Forms.Button
$btnOutput.Text     = 'Browse...'
$btnOutput.Location = New-Object System.Drawing.Point(765, 44)
$btnOutput.Size     = New-Object System.Drawing.Size(100, 24)

# Row 3 : thresholds & options
$lblWarn          = New-Object System.Windows.Forms.Label
$lblWarn.Text     = 'Inactive warning (days):'
$lblWarn.Location = New-Object System.Drawing.Point(12, 85)
$lblWarn.Size     = New-Object System.Drawing.Size(160, 20)

$numWarn          = New-Object System.Windows.Forms.NumericUpDown
$numWarn.Location = New-Object System.Drawing.Point(175, 82)
$numWarn.Size     = New-Object System.Drawing.Size(70, 22)
$numWarn.Minimum  = 1
$numWarn.Maximum  = 3650
$numWarn.Value    = 90

$lblCrit          = New-Object System.Windows.Forms.Label
$lblCrit.Text     = 'Inactive critical (days):'
$lblCrit.Location = New-Object System.Drawing.Point(265, 85)
$lblCrit.Size     = New-Object System.Drawing.Size(160, 20)

$numCrit          = New-Object System.Windows.Forms.NumericUpDown
$numCrit.Location = New-Object System.Drawing.Point(425, 82)
$numCrit.Size     = New-Object System.Drawing.Size(70, 22)
$numCrit.Minimum  = 1
$numCrit.Maximum  = 3650
$numCrit.Value    = 180

$chkAccurate          = New-Object System.Windows.Forms.CheckBox
$chkAccurate.Text     = 'Accurate last logon (query all DCs - slower)'
$chkAccurate.Location = New-Object System.Drawing.Point(525, 83)
$chkAccurate.Size     = New-Object System.Drawing.Size(360, 22)

# Row 4 : action buttons
$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Audit'
$btnRun.Location = New-Object System.Drawing.Point(12, 122)
$btnRun.Size     = New-Object System.Drawing.Size(130, 32)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = 'Flat'

$btnCancel          = New-Object System.Windows.Forms.Button
$btnCancel.Text     = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(150, 122)
$btnCancel.Size     = New-Object System.Drawing.Size(110, 32)
$btnCancel.Enabled  = $false

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export XLSX'
$btnExport.Location = New-Object System.Drawing.Point(268, 122)
$btnExport.Size     = New-Object System.Drawing.Size(130, 32)
$btnExport.Enabled  = $false

$panelTop.Controls.AddRange(@(
    $lblInput, $txtInput, $btnInput,
    $lblOutput, $txtOutput, $btnOutput,
    $lblWarn, $numWarn, $lblCrit, $numCrit, $chkAccurate,
    $btnRun, $btnCancel, $btnExport
))

# ----- Bottom panel (progress & summary) -----------------------------------
$panelBottom        = New-Object System.Windows.Forms.Panel
$panelBottom.Dock   = 'Bottom'
$panelBottom.Height = 92

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(12, 10)
$progress.Size     = New-Object System.Drawing.Size(900, 20)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 38)
$lblStatus.Size     = New-Object System.Drawing.Size(900, 20)
$lblStatus.Text     = 'Select an input file and an output path, then click Run Audit.'

$lblSummary           = New-Object System.Windows.Forms.Label
$lblSummary.Location  = New-Object System.Drawing.Point(12, 62)
$lblSummary.Size      = New-Object System.Drawing.Size(900, 22)
$lblSummary.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$panelBottom.Controls.AddRange(@($progress, $lblStatus, $lblSummary))

# ----- Results grid ---------------------------------------------------------
$dgv                      = New-Object System.Windows.Forms.DataGridView
$dgv.Dock                 = 'Fill'
$dgv.ReadOnly             = $true
$dgv.AllowUserToAddRows   = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.RowHeadersVisible    = $false
$dgv.SelectionMode        = 'FullRowSelect'
$dgv.AutoSizeColumnsMode  = 'DisplayedCells'
$dgv.AllowUserToOrderColumns = $true

# Add docked-fill control LAST so it respects the top/bottom panels
$form.Controls.Add($dgv)
$form.Controls.Add($panelBottom)
$form.Controls.Add($panelTop)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Set-RowColors {
    if ($dgv.Rows.Count -gt 6000) { return }  # skip colouring on very large sets
    foreach ($r in $dgv.Rows) {
        $s = [string]$r.Cells['AuditStatus'].Value
        $c =
            if     ($s -eq 'Not Found')          { [System.Drawing.Color]::FromArgb(255, 205, 210) }
            elseif ($s -eq 'Ambiguous')          { [System.Drawing.Color]::FromArgb(255, 224, 178) }
            elseif ($s -eq 'Disabled')           { [System.Drawing.Color]::FromArgb(224, 224, 224) }
            elseif ($s -eq 'Expired')            { [System.Drawing.Color]::FromArgb(255, 204, 153) }
            elseif ($s -eq 'Never Logged On')    { [System.Drawing.Color]::FromArgb(255, 241, 181) }
            elseif ($s -eq $script:critLabel)    { [System.Drawing.Color]::FromArgb(255, 179, 179) }
            elseif ($s -eq $script:warnLabel)    { [System.Drawing.Color]::FromArgb(255, 249, 196) }
            else                                 { [System.Drawing.Color]::White }
        $r.DefaultCellStyle.BackColor = $c
    }
}

# ---------------------------------------------------------------------------
# Native .xlsx writer (Open XML = a ZIP of XML parts). No third-party module,
# no Excel/COM. Frozen header, autofilter, column widths, colour by status.
# ---------------------------------------------------------------------------
function Get-ColLetter {
    param([int]$Index)   # 0-based
    $n = $Index + 1
    $s = ''
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = ([char]([int][char]'A' + $r)) + $s
        $n = [math]::Floor(($n - 1) / 26)
    }
    return $s
}

function ConvertTo-XmlText {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x1F]', '')   # drop control chars
    $s = $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&apos;')
    return $s
}

function Export-AuditXlsx {
    param($Rows, $Path)

    $cols    = $script:Columns
    $nCols   = $cols.Count
    $lastCol = Get-ColLetter ($nCols - 1)
    $lastRow = $Rows.Count + 1

    $numeric = @{ 'DaysSinceLogon' = $true; 'DaysSincePassword' = $true; 'MatchCount' = $true }
    $dateCol = @{ 'LastLogon'=$true;'PasswordLastSet'=$true;'WhenCreated'=$true;'WhenChanged'=$true;'AccountExpirationDate'=$true }

    function Get-StatusStyle {
        param($Status)
        switch -Regex ("$Status") {
            '^Not Found$'       { return 2 }
            '^Ambiguous'        { return 3 }
            '^Disabled$'        { return 4 }
            '^Expired$'         { return 5 }
            '^Never Logged On$' { return 6 }
            default {
                if ("$Status" -eq $script:critLabel) { return 7 }
                if ("$Status" -eq $script:warnLabel) { return 8 }
                return 0
            }
        }
    }

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
'@
    $relsRoot = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@
    $workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="AD_Account_Audit" sheetId="1" r:id="rId1"/></sheets></workbook>
'@
    $wbRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@
    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="10"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCDD2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFE0B2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE0E0E0"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC99"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF1B5"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFB3B3"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
'@

    $enc = New-Object System.Text.UTF8Encoding($false)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $fs  = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($p in @(
                @{ n='[Content_Types].xml';        t=$contentTypes },
                @{ n='_rels/.rels';                 t=$relsRoot },
                @{ n='xl/workbook.xml';             t=$workbook },
                @{ n='xl/_rels/workbook.xml.rels';  t=$wbRels },
                @{ n='xl/styles.xml';               t=$styles }
            )) {
                $e = $zip.CreateEntry($p.n)
                $sw = New-Object System.IO.StreamWriter($e.Open(), $enc)
                $sw.Write($p.t.Trim()); $sw.Flush(); $sw.Dispose()
            }

            $e  = $zip.CreateEntry('xl/worksheets/sheet1.xml')
            $w  = New-Object System.IO.StreamWriter($e.Open(), $enc)
            $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
            $w.Write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
            $w.Write('<sheetFormatPr defaultRowHeight="15"/>')
            $w.Write('<sheetData>')

            $w.Write('<row r="1">')
            for ($j = 0; $j -lt $nCols; $j++) {
                $ref = (Get-ColLetter $j) + '1'
                $w.Write("<c r=""$ref"" s=""1"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XmlText $cols[$j])</t></is></c>")
            }
            $w.Write('</row>')

            $rn = 2
            foreach ($r in $Rows) {
                $sIdx = Get-StatusStyle $r.AuditStatus
                $w.Write("<row r=""$rn"">")
                for ($j = 0; $j -lt $nCols; $j++) {
                    $col = $cols[$j]
                    $ref = (Get-ColLetter $j) + $rn
                    $val = $r.$col
                    if ($numeric.ContainsKey($col) -and $null -ne $val -and "$val" -ne '') {
                        $w.Write("<c r=""$ref"" s=""$sIdx""><v>$([int]$val)</v></c>")
                    }
                    elseif ($dateCol.ContainsKey($col) -and $val -is [datetime]) {
                        $w.Write("<c r=""$ref"" s=""$sIdx"" t=""inlineStr""><is><t>$($val.ToString('yyyy-MM-dd HH:mm:ss'))</t></is></c>")
                    }
                    else {
                        $txt = ConvertTo-XmlText ([string]$val)
                        if ($txt -eq '') { $w.Write("<c r=""$ref"" s=""$sIdx""/>") }
                        else { $w.Write("<c r=""$ref"" s=""$sIdx"" t=""inlineStr""><is><t xml:space=""preserve"">$txt</t></is></c>") }
                    }
                }
                $w.Write('</row>')
                $rn++
            }

            $w.Write('</sheetData>')
            $w.Write("<autoFilter ref=""A1:$lastCol$lastRow""/>")
            $w.Write('</worksheet>')
            $w.Flush(); $w.Dispose()
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

function Export-CsvFallback {
    param($Rows, $Path)
    $lines = $Rows | Select-Object $script:Columns | ConvertTo-Csv -Delimiter ';' -NoTypeInformation
    $content = "sep=;`r`n" + ($lines -join "`r`n")
    [System.IO.File]::WriteAllText($Path, $content, (New-Object System.Text.UTF8Encoding($true)))  # UTF-8 BOM
}

function Export-Report {
    if ($sync.Results.Count -eq 0) { return }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text)) { return }
    $path = $txtOutput.Text
    $rows = @($sync.Results)
    $lblStatus.Text = "Writing $([System.IO.Path]::GetFileName($path)) ..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        if ($path.ToLowerInvariant().EndsWith('.xlsx')) { Export-AuditXlsx -Rows $rows -Path $path }
        else { Export-CsvFallback -Rows $rows -Path $path }
        $lblStatus.Text = "$($lblStatus.Text)  |  Saved: $path"
    }
    catch {
        try {
            $csv = [System.IO.Path]::ChangeExtension($path, '.csv')
            Export-CsvFallback -Rows $rows -Path $csv
            [System.Windows.Forms.MessageBox]::Show("XLSX export failed - saved a CSV instead:`n$csv`n`n$($_.Exception.Message)", 'Export fallback', 'OK', 'Warning') | Out-Null
            $lblStatus.Text = "Saved (CSV fallback): $csv"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", 'Export error', 'OK', 'Error') | Out-Null
        }
    }
}

function Complete-Audit {
    $errMsg = $sync.Error

    # Tear down the runspace
    try { if ($script:ps) { $script:ps.EndInvoke($script:handle) } } catch { }
    try { if ($script:ps) { $script:ps.Dispose() } }              catch { }
    try { if ($script:rs) { $script:rs.Close(); $script:rs.Dispose() } } catch { }
    $script:ps = $null; $script:rs = $null; $script:handle = $null

    $rows = @($sync.Results)

    # Build a DataTable and bind it (fast, virtualized rendering)
    $dt = New-Object System.Data.DataTable
    foreach ($c in $script:Columns) { [void]$dt.Columns.Add($c) }
    foreach ($r in $rows) {
        $dr = $dt.NewRow()
        foreach ($c in $script:Columns) {
            $val = $r.$c
            $dr[$c] = if ($null -eq $val -or "$val" -eq '') { [DBNull]::Value } else { [string]$val }
        }
        [void]$dt.Rows.Add($dr)
    }
    $dgv.DataSource = $dt
    Set-RowColors

    # Summary counters
    $total    = $rows.Count
    $active   = @($rows | Where-Object { $_.AuditStatus -eq 'Active' }).Count
    $warn     = @($rows | Where-Object { $_.AuditStatus -eq $script:warnLabel }).Count
    $crit     = @($rows | Where-Object { $_.AuditStatus -eq $script:critLabel }).Count
    $never    = @($rows | Where-Object { $_.AuditStatus -eq 'Never Logged On' }).Count
    $expired  = @($rows | Where-Object { $_.AuditStatus -eq 'Expired' }).Count
    $disabled = @($rows | Where-Object { $_.AuditStatus -eq 'Disabled' }).Count
    $missing  = @($rows | Where-Object { $_.AuditStatus -eq 'Not Found' }).Count
    $dupRows  = @($rows | Where-Object { try { [int]$_.MatchCount -gt 1 } catch { $false } }).Count

    $lblSummary.Text = "Total $total   |   Active $active   |   $($script:warnLabel) $warn   |   $($script:critLabel) $crit   |   Never $never   |   Expired $expired   |   Disabled $disabled   |   Not found $missing   |   Duplicate matches $dupRows"

    $progress.Style = 'Blocks'
    $progress.Value = 100
    if ($errMsg)          { $lblStatus.Text = "Completed with error: $errMsg" }
    elseif ($sync.Cancel) { $lblStatus.Text = "Cancelled. $total accounts processed. ($($script:dupCount) duplicates removed)" }
    else                  { $lblStatus.Text = "Audit complete. $total accounts. ($($script:dupCount) duplicates removed)" }

    $btnRun.Enabled = $true; $btnInput.Enabled = $true; $btnOutput.Enabled = $true
    $btnCancel.Enabled = $false
    $btnExport.Enabled = ($total -gt 0)

    if ($total -gt 0 -and -not $sync.Cancel) { Export-Report }  # auto-save on success
}

# ---------------------------------------------------------------------------
# Timer : polls worker progress and refreshes the UI
# ---------------------------------------------------------------------------
$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = 300
$timer.Add_Tick({
    $p = $sync.Processed
    $t = $sync.Total
    if ($sync.Status) {
        $progress.Style = 'Marquee'
        $lblStatus.Text = $sync.Status
    } else {
        $progress.Style = 'Blocks'
        if ($t -gt 0) { $progress.Value = [math]::Min(100, [int](($p / $t) * 100)) }
        $lblStatus.Text = "Matching $p / $t ..."
    }
    if ($sync.Done) {
        $timer.Stop()
        Complete-Audit
    }
})

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------
$btnInput.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = 'Select the .txt file containing account identifiers'
    $ofd.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtInput.Text = $ofd.FileName
        if ([string]::IsNullOrWhiteSpace($txtOutput.Text)) {
            $dir = [System.IO.Path]::GetDirectoryName($ofd.FileName)
            $txtOutput.Text = [System.IO.Path]::Combine($dir, 'AD_Account_Audit.xlsx')
        }
    }
})

$btnOutput.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title    = 'Choose where to save the CSV report'
    $sfd.Filter   = 'Excel Workbook (*.xlsx)|*.xlsx|CSV (semicolon) (*.csv)|*.csv'
    $sfd.FileName = 'AD_Account_Audit.xlsx'
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutput.Text = $sfd.FileName
    }
})

$btnExport.Add_Click({ Export-Report })

$btnCancel.Add_Click({
    $sync.Cancel = $true
    $lblStatus.Text = 'Cancelling...'
    $btnCancel.Enabled = $false
})

$btnRun.Add_Click({
    if (-not (Test-Path -LiteralPath $txtInput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid input file.', 'Missing input', 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please choose an output CSV path.', 'Missing output', 'OK', 'Warning') | Out-Null
        return
    }
    if ([int]$numCrit.Value -le [int]$numWarn.Value) {
        [System.Windows.Forms.MessageBox]::Show('Critical threshold must be greater than the warning threshold.', 'Invalid thresholds', 'OK', 'Warning') | Out-Null
        return
    }

    # Read, trim, drop blanks, de-duplicate (case-insensitive)
    $raw    = Get-Content -LiteralPath $txtInput.Text
    $clean  = @($raw | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne '' })
    $unique = @($clean | Sort-Object -Unique)
    $script:dupCount = $clean.Count - $unique.Count

    if ($unique.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('The input file contains no account identifiers.', 'Empty file', 'OK', 'Warning') | Out-Null
        return
    }

    # Reset shared state
    $sync.Results.Clear()
    $sync.Processed = 0
    $sync.Total     = $unique.Count
    $sync.Done      = $false
    $sync.Cancel    = $false
    $sync.Error     = $null
    $sync.Status    = 'Starting...'

    $config = @{
        WarnDays          = [int]$numWarn.Value
        CritDays          = [int]$numCrit.Value
        AccurateLastLogon = $chkAccurate.Checked
    }
    $script:warnLabel = "Inactive > $($config.WarnDays)d"
    $script:critLabel = "Inactive > $($config.CritDays)d"

    # Clear grid
    $dgv.DataSource = $null
    $dgv.Columns.Clear()

    # Spin up the worker runspace
    $script:rs = [runspacefactory]::CreateRunspace()
    $script:rs.ApartmentState = 'MTA'
    $script:rs.ThreadOptions  = 'ReuseThread'
    $script:rs.Open()
    $script:rs.SessionStateProxy.SetVariable('sync',     $sync)
    $script:rs.SessionStateProxy.SetVariable('accounts', $unique)
    $script:rs.SessionStateProxy.SetVariable('config',   $config)

    $script:ps          = [powershell]::Create()
    $script:ps.Runspace = $script:rs
    [void]$script:ps.AddScript($worker.ToString())
    $script:handle = $script:ps.BeginInvoke()

    $btnRun.Enabled = $false; $btnInput.Enabled = $false; $btnOutput.Enabled = $false; $btnExport.Enabled = $false
    $btnCancel.Enabled = $true
    $progress.Value = 0
    $lblSummary.Text = ''
    $lblStatus.Text  = "Starting audit of $($unique.Count) accounts... ($($script:dupCount) duplicates removed)"
    $timer.Start()
})

$form.Add_FormClosing({
    $sync.Cancel = $true
    try { if ($script:ps) { $script:ps.Stop(); $script:ps.Dispose() } } catch { }
    try { if ($script:rs) { $script:rs.Close(); $script:rs.Dispose() } } catch { }
})

# ---------------------------------------------------------------------------
# Reliable layout: reposition stretchy controls to the real window width.
# (Replaces anchoring, which misbehaves for right-aligned controls inside
#  docked panels created programmatically.)
# ---------------------------------------------------------------------------
$LayoutControls = {
    $margin = 12
    $btnW   = 100
    $labelW = 155

    $w    = $panelTop.ClientSize.Width
    $btnX = $w - $margin - $btnW
    if ($btnX -lt ($labelW + 120)) { $btnX = $labelW + 120 }  # keep on-screen if tiny
    $txtW = $btnX - $labelW - 8

    $btnInput.Left  = $btnX
    $btnOutput.Left = $btnX
    $txtInput.Width  = $txtW
    $txtOutput.Width = $txtW

    $bw = $panelBottom.ClientSize.Width - (2 * $margin)
    if ($bw -lt 200) { $bw = 200 }
    $progress.Width   = $bw
    $lblStatus.Width  = $bw
    $lblSummary.Width = $bw
}

$form.Add_Shown($LayoutControls)
$form.Add_SizeChanged($LayoutControls)

[void]$form.ShowDialog()
