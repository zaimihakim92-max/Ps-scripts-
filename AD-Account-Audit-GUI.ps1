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
    'InputValue','ResolvedBy','SamAccountName','UserPrincipalName','DisplayName',
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
                if ($h[$k].DistinguishedName -ne $User.DistinguishedName) { $amb[$Name][$k] = $true }
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
        function Resolve-FromIndex {
            param($Value)
            $k = ([string]$Value).Trim().ToLowerInvariant()
            if ($k -eq '') { return [PSCustomObject]@{ User = $null; By = 'NotFound' } }
            foreach ($n in $indexNames) {
                if ($amb[$n].ContainsKey($k)) { return [PSCustomObject]@{ User = $null;        By = "Ambiguous ($n)" } }
                if ($idx[$n].ContainsKey($k)) { return [PSCustomObject]@{ User = $idx[$n][$k];  By = $n } }
            }
            return [PSCustomObject]@{ User = $null; By = 'NotFound' }
        }

        # ------------------------------------------------------------------
        # 2) Match every input line against the index (in-memory, O(1) each)
        # ------------------------------------------------------------------
        $sync.Status = $null   # switch back to processed/total display
        foreach ($acc in $accounts) {
            if ($sync.Cancel) { break }
            $now = Get-Date

            try   { $res = Resolve-FromIndex -Value $acc }
            catch { $res = [PSCustomObject]@{ User = $null; By = 'Error' } }

            if ($res.User) {
                $u = $res.User

                $lastLogon = $u.LastLogonDate
                if ($config.AccurateLastLogon -and $dcList.Count -gt 0) {
                    $realLL = Get-TrueLastLogon -Sam $u.SamAccountName -DCs $dcList
                    if ($realLL) { $lastLogon = $realLL }
                }

                $daysLogon = $null
                if ($lastLogon)         { $daysLogon = [math]::Round(($now - $lastLogon).TotalDays, 0) }
                $daysPwd = $null
                if ($u.PasswordLastSet) { $daysPwd   = [math]::Round(($now - $u.PasswordLastSet).TotalDays, 0) }

                $expired = $false
                if ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt $now) { $expired = $true }

                $status = 'Active'
                if     (-not $u.Enabled)                   { $status = 'Disabled' }
                elseif ($expired)                          { $status = 'Expired' }
                elseif (-not $lastLogon)                   { $status = 'Never Logged On' }
                elseif ($daysLogon -gt $config.CritDays)   { $status = "Inactive > $($config.CritDays)d" }
                elseif ($daysLogon -gt $config.WarnDays)   { $status = "Inactive > $($config.WarnDays)d" }

                $row = [PSCustomObject]@{
                    InputValue            = $acc
                    ResolvedBy            = $res.By
                    SamAccountName        = $u.SamAccountName
                    UserPrincipalName     = $u.UserPrincipalName
                    DisplayName           = (Clean-Text $u.Name)
                    Enabled               = $u.Enabled
                    LockedOut             = $u.LockedOut
                    Expired               = $expired
                    AccountExpirationDate = $u.AccountExpirationDate
                    LastLogon             = $lastLogon
                    DaysSinceLogon        = $daysLogon
                    PasswordLastSet       = $u.PasswordLastSet
                    DaysSincePassword     = $daysPwd
                    PasswordNeverExpires  = $u.PasswordNeverExpires
                    WhenCreated           = $u.WhenCreated
                    WhenChanged           = $u.WhenChanged
                    Department            = (Clean-Text $u.Department)
                    Title                 = (Clean-Text $u.Title)
                    Description           = (Clean-Text $u.Description)
                    DistinguishedName     = (Clean-Text $u.DistinguishedName)
                    AuditStatus           = $status
                }
            }
            else {
                $status = if ($res.By -like 'Ambiguous*') { 'Ambiguous' } else { 'Not Found' }
                $row = [PSCustomObject]@{
                    InputValue            = $acc
                    ResolvedBy            = $res.By
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
                    AuditStatus           = $status
                }
            }

            [void]$sync.Results.Add($row)
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
$btnExport.Text     = 'Export CSV'
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

function Export-Report {
    if ($sync.Results.Count -eq 0) { return }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text)) { return }
    try {
        @($sync.Results) |
            Select-Object $script:Columns |
            Export-Csv -Path $txtOutput.Text -Delimiter ';' -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "$($lblStatus.Text)  |  Saved: $($txtOutput.Text)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", 'Export error', 'OK', 'Error') | Out-Null
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
    $missing  = @($rows | Where-Object { $_.AuditStatus -eq 'Not Found' -or $_.AuditStatus -eq 'Ambiguous' }).Count

    $lblSummary.Text = "Total $total   |   Active $active   |   $($script:warnLabel) $warn   |   $($script:critLabel) $crit   |   Never $never   |   Expired $expired   |   Disabled $disabled   |   Not found $missing"

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
            $txtOutput.Text = [System.IO.Path]::Combine($dir, 'AD_Account_Audit.csv')
        }
    }
})

$btnOutput.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title    = 'Choose where to save the CSV report'
    $sfd.Filter   = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = 'AD_Account_Audit.csv'
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
