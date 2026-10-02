<#
.SYNOPSIS
    SharePoint tenant site audit - GUI. Module: PnP.PowerShell.

.DESCRIPTION
    Lists every site collection in the tenant (like the SharePoint admin center) via
    Get-PnPTenantSite, adding the identifiers the console does not show in the list: SiteId
    and GroupId. A "Type" column is derived from the template; an optional "Teams" detection
    flags M365-group sites that have a Team. Live filter (title/URL) and CSV export with a
    column chooser (pick and reorder columns before export).

    Connection: reuses an active PnP session, otherwise connect to the ADMIN site
    (https://YOURTENANT-admin.sharepoint.com), which Get-PnPTenantSite requires.

.PREREQUISITES
    - PnP.PowerShell module, SharePoint administration rights, an Entra app (Client ID) for PnP.
    - Teams detection additionally needs the Graph permission Group.Read.All on the PnP app.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Sites = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Functions
# ==================================================================
function Get-GProp {
    param($obj, [string]$Pascal)
    if ($null -eq $obj) { return $null }
    $camel = $Pascal.Substring(0,1).ToLower() + $Pascal.Substring(1)
    foreach ($name in @($Pascal, $camel)) {
        $p = $obj.PSObject.Properties[$name]
        if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne "") { return $p.Value }
    }
    $ap = $obj.PSObject.Properties['AdditionalProperties']
    if ($ap -and $ap.Value) { foreach ($key in @($camel, $Pascal)) { try { if ($ap.Value.ContainsKey($key)) { return $ap.Value[$key] } } catch { } } }
    return $null
}
function Clean-Guid { param([string]$g) if ($g -and $g -ne '00000000-0000-0000-0000-000000000000') { return $g } return "" }
function Test-PnPConnected { try { return [bool](Get-PnPConnection -ErrorAction Stop) } catch { return $false } }

$script:TemplateMap = @{
    'GROUP#0'                  = "Team site (M365 group)"
    'STS#3'                    = "Team site (no M365 group)"
    'STS#0'                    = "Team site (classic)"
    'SITEPAGEPUBLISHING#0'     = "Communication site"
    'TEAMCHANNEL#0'            = "Teams private channel"
    'TEAMCHANNEL#1'            = "Teams private/shared channel"
    'SPSPERS#10'               = "OneDrive for Business"
    'REDIRECTSITE#0'           = "Redirect site"
    'APPCATALOG#0'             = "App catalog"
    'EHS#1'                    = "Root / tenant config site"
    'BDR#0'                    = "Document Center"
    'DEV#0'                    = "Developer site"
    'OFFILE#1'                 = "Records Center"
    'SRCHCEN#0'                = "Search Center"
    'SRCHCENTERLITE#0'         = "Basic Search Center"
    'BLANKINTERNETCONTAINER#0' = "Publishing Portal"
    'ENTERWIKI#0'              = "Enterprise Wiki"
    'PROJECTSITE#0'            = "Project site"
    'PWA#0'                    = "Project Web App"
    'COMMUNITY#0'              = "Community site"
    'COMMUNITYPORTAL#0'        = "Community portal"
}
function Get-SiteType {
    param([string]$code)
    if ([string]::IsNullOrWhiteSpace($code)) { return "" }
    if ($script:TemplateMap.ContainsKey($code)) { return $script:TemplateMap[$code] }
    return "Unknown ($code)"
}

# Export column catalogue: property key on the site object + display label
$script:ColumnCatalog = @(
    [PSCustomObject]@{ Key='Title';    Label='Title' }
    [PSCustomObject]@{ Key='Url';      Label='URL' }
    [PSCustomObject]@{ Key='Template'; Label='Template' }
    [PSCustomObject]@{ Key='Type';     Label='Type' }
    [PSCustomObject]@{ Key='Teams';    Label='Teams' }
    [PSCustomObject]@{ Key='SiteId';   Label='SiteId' }
    [PSCustomObject]@{ Key='GroupId';  Label='GroupId' }
    [PSCustomObject]@{ Key='Storage';  Label='Storage (MB)' }
    [PSCustomObject]@{ Key='Lock';     Label='Lock state' }
    [PSCustomObject]@{ Key='Sharing';  Label='Sharing' }
    [PSCustomObject]@{ Key='Owner';    Label='Owner' }
)

# ==================================================================
# Form
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint tenant site audit"
$form.Size = New-Object System.Drawing.Size(1010, 780)
$form.MinimumSize = New-Object System.Drawing.Size(900, 640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connect (admin)"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Status: checking..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(790,18); $lblConn.Anchor="Top,Left,Right"

Add-Lbl $form "Admin URL:" 15 50 | Out-Null
$txtUrl = New-Object System.Windows.Forms.TextBox; $txtUrl.Location=New-Object System.Drawing.Point(110,47); $txtUrl.Size=New-Object System.Drawing.Size(470,23); $txtUrl.Anchor="Top,Left,Right"; $form.Controls.Add($txtUrl)
Add-Lbl $form "Client ID:" 595 50 | Out-Null
$txtCid = New-Object System.Windows.Forms.TextBox; $txtCid.Location=New-Object System.Drawing.Point(665,47); $txtCid.Size=New-Object System.Drawing.Size(320,23); $txtCid.Anchor="Top,Right"; $form.Controls.Add($txtCid)

$chkOneDrive = New-Object System.Windows.Forms.CheckBox; $chkOneDrive.Text="Include OneDrive"; $chkOneDrive.Location=New-Object System.Drawing.Point(15,80); $chkOneDrive.AutoSize=$true; $form.Controls.Add($chkOneDrive)
$chkDetailed = New-Object System.Windows.Forms.CheckBox; $chkDetailed.Text="Full details"; $chkDetailed.Location=New-Object System.Drawing.Point(165,80); $chkDetailed.AutoSize=$true; $chkDetailed.Checked=$true; $form.Controls.Add($chkDetailed)
$btnList = New-Object System.Windows.Forms.Button; $btnList.Text="List sites"; $btnList.Location=New-Object System.Drawing.Point(300,76); $btnList.Size=New-Object System.Drawing.Size(180,28); $form.Controls.Add($btnList)
$btnTeams = New-Object System.Windows.Forms.Button; $btnTeams.Text="Detect Teams"; $btnTeams.Location=New-Object System.Drawing.Point(490,76); $btnTeams.Size=New-Object System.Drawing.Size(170,28); $btnTeams.Enabled=$false; $form.Controls.Add($btnTeams)

Add-Lbl $form "Filter:" 15 117 | Out-Null
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilter.Location=New-Object System.Drawing.Point(70,114); $txtFilter.Size=New-Object System.Drawing.Size(400,23); $form.Controls.Add($txtFilter)
$lblCount = Add-Lbl $form "0 sites" 485 117
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text="Export CSV..."; $btnExport.Location=New-Object System.Drawing.Point(820,112); $btnExport.Size=New-Object System.Drawing.Size(165,28); $btnExport.Anchor="Top,Right"; $btnExport.Enabled=$false; $form.Controls.Add($btnExport)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(15,150); $grid.Size=New-Object System.Drawing.Size(970,440); $grid.Anchor="Top,Bottom,Left,Right"
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.ReadOnly=$true; $grid.RowHeadersVisible=$false
$grid.SelectionMode="FullRowSelect"; $grid.AutoSizeColumnsMode="Fill"; $grid.ColumnHeadersHeightSizeMode="AutoSize"
$null=$grid.Columns.Add("Title","Title")
$null=$grid.Columns.Add("Url","URL")
$null=$grid.Columns.Add("Template","Template")
$null=$grid.Columns.Add("Type","Type")
$null=$grid.Columns.Add("Teams","Teams")
$null=$grid.Columns.Add("SiteId","SiteId")
$null=$grid.Columns.Add("GroupId","GroupId")
$null=$grid.Columns.Add("Storage","Storage (MB)")
$null=$grid.Columns.Add("Lock","Lock state")
$null=$grid.Columns.Add("Sharing","Sharing")
$null=$grid.Columns.Add("Owner","Owner")
$grid.Columns["Title"].FillWeight=12; $grid.Columns["Url"].FillWeight=18; $grid.Columns["Template"].FillWeight=7
$grid.Columns["Type"].FillWeight=12; $grid.Columns["Teams"].FillWeight=5
$grid.Columns["SiteId"].FillWeight=13; $grid.Columns["GroupId"].FillWeight=13; $grid.Columns["Storage"].FillWeight=7
$grid.Columns["Lock"].FillWeight=5; $grid.Columns["Sharing"].FillWeight=8; $grid.Columns["Owner"].FillWeight=10
$form.Controls.Add($grid)

Add-Lbl $form "Log:" 15 596 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,616); $txtLog.Size=New-Object System.Drawing.Size(970,110); $txtLog.Anchor="Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# UI helpers
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

function Update-Grid {
    $grid.Rows.Clear()
    $f = $txtFilter.Text.Trim().ToLower()
    $n = 0
    foreach ($s in $script:Sites) {
        if ($f) { if (-not ("$($s.Title) $($s.Url)").ToLower().Contains($f)) { continue } }
        [void]$grid.Rows.Add($s.Title,$s.Url,$s.Template,$s.Type,$s.Teams,$s.SiteId,$s.GroupId,$s.Storage,$s.Lock,$s.Sharing,$s.Owner)
        $n++
    }
    $lblCount.Text = "$n / $($script:Sites.Count) sites"
}

function Show-ColumnChooser {
    # Column selection + reorder dialog. Returns an ordered array of property keys, or $null if cancelled.
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Choose export columns"; $dlg.Size = New-Object System.Drawing.Size(360, 430); $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.Font = $form.Font

    Add-Lbl $dlg "Check the columns to export. Use Up/Down to reorder." 12 10 | Out-Null

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(12, 36); $clb.Size = New-Object System.Drawing.Size(230, 320); $clb.CheckOnClick = $true
    $clb.DisplayMember = "Label"
    foreach ($c in $script:ColumnCatalog) { [void]$clb.Items.Add($c, $true) }
    $dlg.Controls.Add($clb)

    $btnUp = New-Object System.Windows.Forms.Button; $btnUp.Text="Up"; $btnUp.Location=New-Object System.Drawing.Point(252,36); $btnUp.Size=New-Object System.Drawing.Size(80,28); $dlg.Controls.Add($btnUp)
    $btnDn = New-Object System.Windows.Forms.Button; $btnDn.Text="Down"; $btnDn.Location=New-Object System.Drawing.Point(252,70); $btnDn.Size=New-Object System.Drawing.Size(80,28); $dlg.Controls.Add($btnDn)
    $btnAll = New-Object System.Windows.Forms.Button; $btnAll.Text="All"; $btnAll.Location=New-Object System.Drawing.Point(252,120); $btnAll.Size=New-Object System.Drawing.Size(80,28); $dlg.Controls.Add($btnAll)
    $btnNone = New-Object System.Windows.Forms.Button; $btnNone.Text="None"; $btnNone.Location=New-Object System.Drawing.Point(252,154); $btnNone.Size=New-Object System.Drawing.Size(80,28); $dlg.Controls.Add($btnNone)

    $btnUp.Add_Click({ $i=$clb.SelectedIndex; if ($i -gt 0) { $ck=$clb.GetItemChecked($i); $it=$clb.Items[$i]; $clb.Items.RemoveAt($i); $clb.Items.Insert($i-1,$it); $clb.SetItemChecked($i-1,$ck); $clb.SelectedIndex=$i-1 } })
    $btnDn.Add_Click({ $i=$clb.SelectedIndex; if ($i -ge 0 -and $i -lt $clb.Items.Count-1) { $ck=$clb.GetItemChecked($i); $it=$clb.Items[$i]; $clb.Items.RemoveAt($i); $clb.Items.Insert($i+1,$it); $clb.SetItemChecked($i+1,$ck); $clb.SelectedIndex=$i+1 } })
    $btnAll.Add_Click({ for ($i=0;$i -lt $clb.Items.Count;$i++){ $clb.SetItemChecked($i,$true) } })
    $btnNone.Add_Click({ for ($i=0;$i -lt $clb.Items.Count;$i++){ $clb.SetItemChecked($i,$false) } })

    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text="OK"; $btnOk.Location=New-Object System.Drawing.Point(160,366); $btnOk.Size=New-Object System.Drawing.Size(85,30); $btnOk.DialogResult="OK"; $dlg.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text="Cancel"; $btnCancel.Location=New-Object System.Drawing.Point(252,366); $btnCancel.Size=New-Object System.Drawing.Size(85,30); $btnCancel.DialogResult="Cancel"; $dlg.Controls.Add($btnCancel)
    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -ne "OK") { return $null }
    $keys = @()
    for ($i=0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { $keys += $clb.Items[$i].Key } }
    return ,$keys
}

# ==================================================================
# Events
# ==================================================================
$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUrl.Text)) { [System.Windows.Forms.MessageBox]::Show("Enter the SharePoint admin site URL.","Connection","OK","Warning")|Out-Null; return }
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { [System.Windows.Forms.MessageBox]::Show("PnP.PowerShell module not found.","Prerequisite","OK","Error")|Out-Null; return }
    $btnConnect.Enabled=$false; Set-Conn $false "Connecting..."
    try {
        if ([string]::IsNullOrWhiteSpace($txtCid.Text)) { Connect-PnPOnline -Url $txtUrl.Text.Trim() -Interactive -ErrorAction Stop }
        else { Connect-PnPOnline -Url $txtUrl.Text.Trim() -Interactive -ClientId $txtCid.Text.Trim() -ErrorAction Stop }
        if (Test-PnPConnected) { $c=Get-PnPConnection; Set-Conn $true "Connected: $($c.Url)"; Write-Log "Connected: $($txtUrl.Text.Trim())" }
    } catch { Set-Conn $false "Connection failed."; [System.Windows.Forms.MessageBox]::Show("Failed:`n$($_.Exception.Message)","Connection","OK","Error")|Out-Null }
    $btnConnect.Enabled=$true
})

$btnList.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Not connected. Connect to the admin site.","Connection","OK","Warning")|Out-Null; return }
    $btnList.Enabled=$false; Write-Log "Loading tenant sites..."
    try {
        $opts=@{}
        if ($chkOneDrive.Checked) { $opts['IncludeOneDriveSites']=$true }
        if ($chkDetailed.Checked) { $opts['Detailed']=$true }
        $sites = @(Get-PnPTenantSite @opts -ErrorAction Stop)
        $script:Sites.Clear()
        foreach ($s in $sites) {
            $tpl = [string](Get-GProp $s 'Template')
            $script:Sites.Add([PSCustomObject]@{
                Title    = [string](Get-GProp $s 'Title')
                Url      = [string](Get-GProp $s 'Url')
                Template = $tpl
                Type     = Get-SiteType $tpl
                Teams    = ""
                SiteId   = Clean-Guid ([string](Get-GProp $s 'SiteId'))
                GroupId  = Clean-Guid ([string](Get-GProp $s 'GroupId'))
                Storage  = "$([string](Get-GProp $s 'StorageUsageCurrent')) / $([string](Get-GProp $s 'StorageQuota'))"
                Lock     = [string](Get-GProp $s 'LockState')
                Sharing  = [string](Get-GProp $s 'SharingCapability')
                Owner    = [string](Get-GProp $s 'Owner')
            }) | Out-Null
        }
        Update-Grid
        $btnExport.Enabled = ($script:Sites.Count -gt 0)
        $btnTeams.Enabled  = ($script:Sites.Count -gt 0)
        Write-Log "$($script:Sites.Count) site(s) loaded."
    } catch { Write-Log "ERROR: $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Failed:`n$($_.Exception.Message)`n`n(Get-PnPTenantSite requires a connection to the admin site.)","List","OK","Error")|Out-Null }
    $btnList.Enabled=$true
})

$btnTeams.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Not connected.","Teams","OK","Warning")|Out-Null; return }
    if ($script:Sites.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("List the sites first.","Teams","OK","Warning")|Out-Null; return }
    $btnTeams.Enabled=$false
    Write-Log "Retrieving M365 groups (HasTeam property)..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $groups = @(Get-PnPMicrosoft365Group -ErrorAction Stop)
    } catch {
        Write-Log "Teams ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Could not list M365 groups:`n$($_.Exception.Message)`n`nThe Entra app used by PnP needs the Graph permission 'Group.Read.All' (admin consent).","Teams","OK","Error")|Out-Null
        $btnTeams.Enabled=$true; return
    }
    $withTeam = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($g in $groups) { if ([bool](Get-GProp $g 'HasTeam')) { $gid=[string](Get-GProp $g 'Id'); if ($gid) { [void]$withTeam.Add($gid) } } }
    Write-Log "M365 groups: $($groups.Count)  |  with Teams: $($withTeam.Count)"
    $n=0
    foreach ($s in $script:Sites) {
        if ($s.GroupId) { $s.Teams = if ($withTeam.Contains($s.GroupId)) { 'Yes' } else { 'No' }; $n++ }
    }
    Update-Grid
    Write-Log "Teams set for $n group-connected site(s)."
    $btnTeams.Enabled=$true
})

$txtFilter.Add_TextChanged({ if ($script:Sites.Count -gt 0) { Update-Grid } })

$btnExport.Add_Click({
    if ($script:Sites.Count -eq 0) { return }
    $keys = Show-ColumnChooser
    if ($null -eq $keys) { return }
    if (@($keys).Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select at least one column.","Export","OK","Warning")|Out-Null; return }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="CSV file (*.csv)|*.csv"; $dlg.FileName="Tenant_Sites.csv"
    if ($dlg.ShowDialog() -ne "OK") { return }
    try {
        $script:Sites | Select-Object $keys | Export-Csv -Path $dlg.FileName -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-Log "Exported ($($keys -join ', ')): $($dlg.FileName)"
        [System.Windows.Forms.MessageBox]::Show("Export complete:`n$($dlg.FileName)","Export","OK","Information")|Out-Null
    } catch { [System.Windows.Forms.MessageBox]::Show("Failed:`n$($_.Exception.Message)","Export","OK","Error")|Out-Null }
})

# ==================================================================
# Adopt an active PnP connection
# ==================================================================
if (Get-Module -ListAvailable -Name PnP.PowerShell) {
    try {
        if (Test-PnPConnected) {
            $c = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($c) { $txtUrl.Text = $c.Url; Set-Conn $true "Reusing active PnP connection: $($c.Url)"; Write-Log "PnP connection detected: $($c.Url)" }
        } else { Set-Conn $false "Not connected. Enter admin URL + Client ID, then Connect." }
    } catch { Set-Conn $false "Not connected." }
} else { Set-Conn $false "PnP.PowerShell module required." }

[void]$form.ShowDialog()
