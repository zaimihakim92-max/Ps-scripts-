<#
.SYNOPSIS
    Audit / inventaire des sites SharePoint du tenant - Interface graphique. Module : PnP.PowerShell.

.DESCRIPTION
    Liste tous les sites du tenant (comme la console d'administration SharePoint) via
    Get-PnPTenantSite, en ajoutant les identifiants que la console ne montre pas dans la liste :
    SiteId et GroupId. Colonnes : titre, URL, modele, SiteId, GroupId, stockage, verrou, partage,
    proprietaire. Filtre en direct (titre/URL) et export CSV.

    Connexion : reutilise une session PnP active, sinon bouton de connexion au SITE D'ADMINISTRATION
    (https://VOTRETENANT-admin.sharepoint.com), car Get-PnPTenantSite exige une connexion admin.

.PREREQUIS
    - Module PnP.PowerShell, role d'administration SharePoint, et une app Entra (Client ID) pour PnP.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Sites = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Fonctions
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
    'GROUP#0'                  = "Site d'equipe (groupe M365)"
    'STS#3'                    = "Site d'equipe (sans groupe M365)"
    'STS#0'                    = "Site d'equipe (classique)"
    'SITEPAGEPUBLISHING#0'     = "Site de communication"
    'TEAMCHANNEL#0'            = "Canal Teams prive"
    'TEAMCHANNEL#1'            = "Canal Teams prive/partage"
    'SPSPERS#10'               = "OneDrive Entreprise"
    'REDIRECTSITE#0'           = "Site de redirection"
    'APPCATALOG#0'             = "Catalogue d'applications"
    'EHS#1'                    = "Site racine / config tenant"
    'BDR#0'                    = "Document Center"
    'DEV#0'                    = "Site developpeur"
    'OFFILE#1'                 = "Records Center"
    'SRCHCEN#0'                = "Centre de recherche"
    'SRCHCENTERLITE#0'         = "Centre de recherche (basique)"
    'BLANKINTERNETCONTAINER#0' = "Publishing Portal"
    'ENTERWIKI#0'              = "Wiki d'entreprise"
    'PROJECTSITE#0'            = "Site de projet"
    'PWA#0'                    = "Project Web App"
    'COMMUNITY#0'              = "Site de communaute"
    'COMMUNITYPORTAL#0'        = "Portail de communaute"
}
function Get-SiteType {
    param([string]$code)
    if ([string]::IsNullOrWhiteSpace($code)) { return "" }
    if ($script:TemplateMap.ContainsKey($code)) { return $script:TemplateMap[$code] }
    return "Inconnu ($code)"
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Audit sites SharePoint (tenant)"
$form.Size = New-Object System.Drawing.Size(1010, 780)
$form.MinimumSize = New-Object System.Drawing.Size(900, 640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connexion admin"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Etat : verification..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(790,18); $lblConn.Anchor="Top,Left,Right"

Add-Lbl $form "URL admin :" 15 50 | Out-Null
$txtUrl = New-Object System.Windows.Forms.TextBox; $txtUrl.Location=New-Object System.Drawing.Point(110,47); $txtUrl.Size=New-Object System.Drawing.Size(470,23); $txtUrl.Anchor="Top,Left,Right"; $form.Controls.Add($txtUrl)
Add-Lbl $form "Client ID :" 595 50 | Out-Null
$txtCid = New-Object System.Windows.Forms.TextBox; $txtCid.Location=New-Object System.Drawing.Point(665,47); $txtCid.Size=New-Object System.Drawing.Size(320,23); $txtCid.Anchor="Top,Right"; $form.Controls.Add($txtCid)

$chkOneDrive = New-Object System.Windows.Forms.CheckBox; $chkOneDrive.Text="Inclure OneDrive"; $chkOneDrive.Location=New-Object System.Drawing.Point(15,80); $chkOneDrive.AutoSize=$true; $form.Controls.Add($chkOneDrive)
$chkDetailed = New-Object System.Windows.Forms.CheckBox; $chkDetailed.Text="Details complets"; $chkDetailed.Location=New-Object System.Drawing.Point(165,80); $chkDetailed.AutoSize=$true; $chkDetailed.Checked=$true; $form.Controls.Add($chkDetailed)
$btnList = New-Object System.Windows.Forms.Button; $btnList.Text="Lister les sites"; $btnList.Location=New-Object System.Drawing.Point(320,76); $btnList.Size=New-Object System.Drawing.Size(180,28); $form.Controls.Add($btnList)
$btnTeams = New-Object System.Windows.Forms.Button; $btnTeams.Text="Detecter Teams"; $btnTeams.Location=New-Object System.Drawing.Point(510,76); $btnTeams.Size=New-Object System.Drawing.Size(170,28); $btnTeams.Enabled=$false; $form.Controls.Add($btnTeams)

Add-Lbl $form "Filtrer :" 15 117 | Out-Null
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilter.Location=New-Object System.Drawing.Point(75,114); $txtFilter.Size=New-Object System.Drawing.Size(400,23); $form.Controls.Add($txtFilter)
$lblCount = Add-Lbl $form "0 site" 490 117
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text="Exporter CSV"; $btnExport.Location=New-Object System.Drawing.Point(820,112); $btnExport.Size=New-Object System.Drawing.Size(165,28); $btnExport.Anchor="Top,Right"; $btnExport.Enabled=$false; $form.Controls.Add($btnExport)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(15,150); $grid.Size=New-Object System.Drawing.Size(970,440); $grid.Anchor="Top,Bottom,Left,Right"
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.ReadOnly=$true; $grid.RowHeadersVisible=$false
$grid.SelectionMode="FullRowSelect"; $grid.AutoSizeColumnsMode="Fill"; $grid.ColumnHeadersHeightSizeMode="AutoSize"
$null=$grid.Columns.Add("Title","Titre")
$null=$grid.Columns.Add("Url","URL")
$null=$grid.Columns.Add("Template","Modele")
$null=$grid.Columns.Add("Type","Type")
$null=$grid.Columns.Add("Teams","Teams")
$null=$grid.Columns.Add("SiteId","SiteId")
$null=$grid.Columns.Add("GroupId","GroupId")
$null=$grid.Columns.Add("Storage","Stockage (Mo)")
$null=$grid.Columns.Add("Lock","Verrou")
$null=$grid.Columns.Add("Sharing","Partage")
$null=$grid.Columns.Add("Owner","Proprietaire")
$grid.Columns["Title"].FillWeight=12; $grid.Columns["Url"].FillWeight=18; $grid.Columns["Template"].FillWeight=7
$grid.Columns["Type"].FillWeight=12; $grid.Columns["Teams"].FillWeight=5
$grid.Columns["SiteId"].FillWeight=13; $grid.Columns["GroupId"].FillWeight=13; $grid.Columns["Storage"].FillWeight=7
$grid.Columns["Lock"].FillWeight=5; $grid.Columns["Sharing"].FillWeight=8; $grid.Columns["Owner"].FillWeight=10
$form.Controls.Add($grid)

Add-Lbl $form "Journal :" 15 596 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,616); $txtLog.Size=New-Object System.Drawing.Size(970,110); $txtLog.Anchor="Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# Helpers UI
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

function Update-Grid {
    $grid.Rows.Clear()
    $f = $txtFilter.Text.Trim().ToLower()
    $n = 0
    foreach ($s in $script:Sites) {
        if ($f) {
            $hay = ("$($s.Title) $($s.Url)").ToLower()
            if (-not $hay.Contains($f)) { continue }
        }
        [void]$grid.Rows.Add($s.Title,$s.Url,$s.Template,$s.Type,$s.Teams,$s.SiteId,$s.GroupId,$s.Storage,$s.Lock,$s.Sharing,$s.Owner)
        $n++
    }
    $lblCount.Text = "$n / $($script:Sites.Count) sites"
}

# ==================================================================
# Evenements
# ==================================================================
$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUrl.Text)) { [System.Windows.Forms.MessageBox]::Show("Renseignez l'URL du site d'administration SharePoint.","Connexion","OK","Warning")|Out-Null; return }
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { [System.Windows.Forms.MessageBox]::Show("Module PnP.PowerShell introuvable.","Prerequis","OK","Error")|Out-Null; return }
    $btnConnect.Enabled=$false; Set-Conn $false "Connexion en cours..."
    try {
        if ([string]::IsNullOrWhiteSpace($txtCid.Text)) { Connect-PnPOnline -Url $txtUrl.Text.Trim() -Interactive -ErrorAction Stop }
        else { Connect-PnPOnline -Url $txtUrl.Text.Trim() -Interactive -ClientId $txtCid.Text.Trim() -ErrorAction Stop }
        if (Test-PnPConnected) { $c=Get-PnPConnection; Set-Conn $true "Connecte : $($c.Url)"; Write-Log "Connecte : $($txtUrl.Text.Trim())" }
    } catch { Set-Conn $false "Echec de connexion."; [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Connexion","OK","Error")|Out-Null }
    $btnConnect.Enabled=$true
})

$btnList.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte. Connectez-vous au site d'administration.","Connexion","OK","Warning")|Out-Null; return }
    $btnList.Enabled=$false; Write-Log "Chargement des sites du tenant..."
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
        $btnTeams.Enabled = ($script:Sites.Count -gt 0)
        Write-Log "$($script:Sites.Count) site(s) charge(s)."
    } catch { Write-Log "ERREUR : $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)`n`n(Get-PnPTenantSite exige une connexion au site d'administration.)","Liste","OK","Error")|Out-Null }
    $btnList.Enabled=$true
})

$btnTeams.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte.","Teams","OK","Warning")|Out-Null; return }
    $targets = @($script:Sites | Where-Object { $_.Template -eq 'GROUP#0' -and $_.GroupId })
    if ($targets.Count -eq 0) { Write-Log "Aucun site a groupe M365 a analyser."; return }
    $btnTeams.Enabled=$false; Write-Log "Detection Teams sur $($targets.Count) site(s) a groupe..."
    $n=0
    foreach ($s in $targets) {
        $n++
        try {
            $grp = Get-PnPMicrosoft365Group -Identity $s.GroupId -ErrorAction Stop
            $ht = Get-GProp $grp 'HasTeam'
            if ($null -eq $ht -or "$ht" -eq "") { $rpo = Get-GProp $grp 'ResourceProvisioningOptions'; $ht = ($rpo -and (@($rpo) -contains 'Team')) }
            $s.Teams = if ([bool]$ht) { 'Oui' } else { 'Non' }
        } catch { $s.Teams = '?' }
        if ($n % 10 -eq 0) { Write-Log "  $n / $($targets.Count)..." }
        [System.Windows.Forms.Application]::DoEvents()
    }
    Update-Grid
    Write-Log "Detection Teams terminee ($($targets.Count) site(s) a groupe)."
    $btnTeams.Enabled=$true
})

$txtFilter.Add_TextChanged({ if ($script:Sites.Count -gt 0) { Update-Grid } })

$btnExport.Add_Click({
    if ($script:Sites.Count -eq 0) { return }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="Fichier CSV (*.csv)|*.csv"; $dlg.FileName="Sites_Tenant.csv"
    if ($dlg.ShowDialog() -ne "OK") { return }
    try {
        $script:Sites | Select-Object Title,Url,Template,Type,Teams,SiteId,GroupId,Storage,Lock,Sharing,Owner |
            Export-Csv -Path $dlg.FileName -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-Log "Export : $($dlg.FileName)"
        [System.Windows.Forms.MessageBox]::Show("Export termine :`n$($dlg.FileName)","Export","OK","Information")|Out-Null
    } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Export","OK","Error")|Out-Null }
})

# ==================================================================
# Adopter une connexion PnP active
# ==================================================================
if (Get-Module -ListAvailable -Name PnP.PowerShell) {
    try {
        if (Test-PnPConnected) {
            $c = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($c) { $txtUrl.Text = $c.Url; Set-Conn $true "Connexion PnP active reutilisee : $($c.Url)"; Write-Log "Connexion PnP detectee : $($c.Url)" }
        } else { Set-Conn $false "Non connecte. Renseignez l'URL admin + Client ID puis Connexion." }
    } catch { Set-Conn $false "Non connecte." }
} else { Set-Conn $false "Module PnP.PowerShell requis." }

[void]$form.ShowDialog()
