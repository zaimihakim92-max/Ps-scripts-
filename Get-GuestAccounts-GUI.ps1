<#
.SYNOPSIS
    Repertorier les comptes invites (Guest / B2B) et recuperer leur e-mail - Interface graphique.
    Module : Az (Connect-AzAccount / Get-AzADUser).

.DESCRIPTION
    Source : un fichier .txt ET/OU une zone de collage (un identifiant par ligne : e-mail, UPN,
    ou nom). Les deux sources sont fusionnees et dedoublonnees.

    L'outil charge une fois tous les comptes ou userType = Guest (filtre cote serveur), les indexe
    (mail, otherMails, UPN encode "..._domaine#EXT#@..."), puis resout chaque identifiant vers son
    compte invite et affiche son e-mail, son UPN invite et ses adresses secondaires.

    Bouton "Lister tous les invites" pour un inventaire, et export CSV.

.PREREQUIS
    - Module Az (Az.Accounts, Az.Resources). Connexion : bouton, ou session Connect-AzAccount deja active.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Guests  = $null
$script:ByMail  = @{}
$script:ByOther = @{}
$script:GuestIdx = @()
$script:Results = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Fonctions Az
# ==================================================================
function ConvertTo-Ascii {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $norm = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) { if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) } }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}
function Normalize-Text { param([string]$s) if (-not $s) { return "" } return (ConvertTo-Ascii $s).ToLower().Trim() }

function Get-GProp {
    # Lit une propriete Graph, qu'elle soit forte (.Mail) ou dans AdditionalProperties (['mail']).
    param($obj, [string]$Pascal)
    $v = $obj.$Pascal
    if ($null -ne $v -and "$v" -ne "") { return $v }
    $ap = $obj.PSObject.Properties['AdditionalProperties']
    if ($ap -and $obj.AdditionalProperties) {
        $camel = $Pascal.Substring(0,1).ToLower() + $Pascal.Substring(1)
        foreach ($key in @($camel, $Pascal)) { if ($obj.AdditionalProperties.ContainsKey($key)) { return $obj.AdditionalProperties[$key] } }
    }
    return $null
}
function Test-AzConnected { try { return [bool](Get-AzContext -ErrorAction Stop) } catch { return $false } }

function Ensure-Guests {
    if ($null -ne $script:Guests) { return $true }
    if (-not (Test-AzConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a Azure.","Connexion","OK","Warning")|Out-Null; return $false }
    Write-Log "Chargement des comptes invites (Guest)..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        try {
            $g = Get-AzADUser -Filter "userType eq 'Guest'" -Select 'displayName','userPrincipalName','mail','otherMails','accountEnabled','id' -ErrorAction Stop
        } catch {
            Write-Log "  (-Select non pris en charge par cette version d'Az : chargement standard)"
            $g = Get-AzADUser -Filter "userType eq 'Guest'" -ErrorAction Stop
        }
        $flat = New-Object System.Collections.Generic.List[object]
        foreach ($u in @($g)) {
            $om = Get-GProp $u 'OtherMails'
            $flat.Add([PSCustomObject]@{
                DisplayName       = [string](Get-GProp $u 'DisplayName')
                UserPrincipalName = [string](Get-GProp $u 'UserPrincipalName')
                Mail              = [string](Get-GProp $u 'Mail')
                OtherMails        = @($om | Where-Object { $_ })
                AccountEnabled    = (Get-GProp $u 'AccountEnabled')
                Id                = [string](Get-GProp $u 'Id')
            }) | Out-Null
        }
        $script:Guests = $flat
        $script:ByMail = @{}; $script:ByOther = @{}
        $idx = New-Object System.Collections.Generic.List[object]
        foreach ($u in $script:Guests) {
            if ($u.Mail) { $script:ByMail[$u.Mail.ToLower()] = $u }
            if ($u.OtherMails) { foreach ($o in $u.OtherMails) { if ($o) { $script:ByOther[$o.ToLower()] = $u } } }
            $idx.Add([PSCustomObject]@{ G = $u; ND = (Normalize-Text $u.DisplayName) }) | Out-Null
        }
        $script:GuestIdx = $idx
        Write-Log "$($script:Guests.Count) compte(s) invite(s) charge(s)."
        return $true
    } catch { Write-Log "ERREUR chargement invites : $($_.Exception.Message)"; return $false }
}

function Resolve-GuestByMail {
    param([string]$Id)
    $k = $Id.Trim().ToLower(); if (-not $k) { return $null }
    if ($script:ByMail.ContainsKey($k))  { return $script:ByMail[$k] }
    if ($script:ByOther.ContainsKey($k)) { return $script:ByOther[$k] }
    $enc = ($k -replace '@', '_')
    foreach ($u in $script:Guests) {
        if ($u.UserPrincipalName) {
            $upn = $u.UserPrincipalName.ToLower()
            if ($upn.Contains('#ext#') -and $upn.StartsWith($enc)) { return $u }
        }
    }
    return $null
}

function Resolve-GuestByName {
    # Correspondance insensible a l'ordre : tous les mots de l'entree doivent figurer dans le DisplayName.
    param([string]$Id)
    $tokens = @((Normalize-Text $Id) -split '[;,\s]+' | Where-Object { $_ })
    if ($tokens.Count -eq 0) { return @() }
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($e in $script:GuestIdx) {
        $ok = $true
        foreach ($t in $tokens) { if ($e.ND -notlike "*$t*") { $ok = $false; break } }
        if ($ok) { $hits.Add($e.G) | Out-Null }
    }
    return , $hits.ToArray()
}

function Get-InputIds {
    $raw = @()
    if (-not [string]::IsNullOrWhiteSpace($txtFile.Text) -and (Test-Path -LiteralPath $txtFile.Text)) { $raw += Get-Content -LiteralPath $txtFile.Text -ErrorAction SilentlyContinue }
    if (-not [string]::IsNullOrWhiteSpace($txtPaste.Text)) { $raw += ($txtPaste.Text -split "`r?`n") }
    return @($raw | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Comptes invites (Guest) - recherche et e-mails"
$form.Size = New-Object System.Drawing.Size(900, 780)
$form.MinimumSize = New-Object System.Drawing.Size(820, 680)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connexion Azure"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Etat : verification..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(680,18); $lblConn.Anchor="Top,Left,Right"

$gbSrc = New-Object System.Windows.Forms.GroupBox
$gbSrc.Text="Source (un identifiant par ligne : e-mail, UPN ou nom)"; $gbSrc.Location=New-Object System.Drawing.Point(15,50); $gbSrc.Size=New-Object System.Drawing.Size(855,150); $gbSrc.Anchor="Top,Left,Right"
$form.Controls.Add($gbSrc)

Add-Lbl $gbSrc "Fichier .txt :" 14 26 | Out-Null
$txtFile = New-Object System.Windows.Forms.TextBox; $txtFile.Location=New-Object System.Drawing.Point(110,23); $txtFile.Size=New-Object System.Drawing.Size(625,23); $txtFile.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtFile)
$btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text="Parcourir..."; $btnBrowse.Location=New-Object System.Drawing.Point(745,22); $btnBrowse.Size=New-Object System.Drawing.Size(95,25); $btnBrowse.Anchor="Top,Right"; $gbSrc.Controls.Add($btnBrowse)
Add-Lbl $gbSrc "Ou coller :" 14 56 | Out-Null
$txtPaste = New-Object System.Windows.Forms.TextBox; $txtPaste.Location=New-Object System.Drawing.Point(110,53); $txtPaste.Size=New-Object System.Drawing.Size(730,80); $txtPaste.Multiline=$true; $txtPaste.ScrollBars="Vertical"; $txtPaste.WordWrap=$false; $txtPaste.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtPaste)

$btnResolve = New-Object System.Windows.Forms.Button; $btnResolve.Text="Resoudre depuis la liste"; $btnResolve.Location=New-Object System.Drawing.Point(15,210); $btnResolve.Size=New-Object System.Drawing.Size(220,28); $form.Controls.Add($btnResolve)
$btnListAll = New-Object System.Windows.Forms.Button; $btnListAll.Text="Lister tous les invites"; $btnListAll.Location=New-Object System.Drawing.Point(245,210); $btnListAll.Size=New-Object System.Drawing.Size(200,28); $form.Controls.Add($btnListAll)
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text="Exporter CSV"; $btnExport.Location=New-Object System.Drawing.Point(655,210); $btnExport.Size=New-Object System.Drawing.Size(215,28); $btnExport.Anchor="Top,Right"; $btnExport.Enabled=$false; $form.Controls.Add($btnExport)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(15,248); $grid.Size=New-Object System.Drawing.Size(855,270); $grid.Anchor="Top,Left,Right"
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.ReadOnly=$true; $grid.RowHeadersVisible=$false
$grid.SelectionMode="FullRowSelect"; $grid.AutoSizeColumnsMode="Fill"; $grid.ColumnHeadersHeightSizeMode="AutoSize"
$null=$grid.Columns.Add("Input","Entree")
$null=$grid.Columns.Add("Display","Invite (nom)")
$null=$grid.Columns.Add("Upn","UPN invite")
$null=$grid.Columns.Add("Mail","E-mail")
$null=$grid.Columns.Add("Other","Adresses secondaires")
$null=$grid.Columns.Add("Enabled","Actif")
$null=$grid.Columns.Add("Statut","Statut")
$grid.Columns["Input"].FillWeight=16; $grid.Columns["Display"].FillWeight=15; $grid.Columns["Upn"].FillWeight=22; $grid.Columns["Mail"].FillWeight=17; $grid.Columns["Other"].FillWeight=17; $grid.Columns["Enabled"].FillWeight=6; $grid.Columns["Statut"].FillWeight=9
$form.Controls.Add($grid)

Add-Lbl $form "Journal :" 15 526 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,546); $txtLog.Size=New-Object System.Drawing.Size(855,160); $txtLog.Anchor="Top,Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# Helpers UI
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

function Add-ResultRow {
    param([string]$input,$guest,[string]$statut)
    $disp=""; $upn=""; $mail=""; $other=""; $en=""
    if ($guest) { $disp=$guest.DisplayName; $upn=$guest.UserPrincipalName; $mail=$guest.Mail; if ($guest.OtherMails) { $other=($guest.OtherMails -join '; ') }; $en=[string]$guest.AccountEnabled }
    $i=$grid.Rows.Add($input,$disp,$upn,$mail,$other,$en,$statut)
    if ($statut -eq "INTROUVABLE") { $grid.Rows[$i].DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,224,150) }
    elseif ($statut -like "AMBIGU*") { $grid.Rows[$i].DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,235,200) }
    $script:Results.Add([PSCustomObject]@{ Entree=$input; DisplayName=$disp; UserPrincipalName=$upn; Mail=$mail; OtherMails=$other; AccountEnabled=$en; Statut=$statut })|Out-Null
}

# ==================================================================
# Evenements
# ==================================================================
$btnBrowse.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Fichiers texte (*.txt)|*.txt|Tous (*.*)|*.*"; if($dlg.ShowDialog() -eq "OK"){$txtFile.Text=$dlg.FileName} })

$btnConnect.Add_Click({
    if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) { [System.Windows.Forms.MessageBox]::Show("Module Az introuvable.`nInstall-Module Az -Scope CurrentUser","Prerequis","OK","Error")|Out-Null; return }
    $btnConnect.Enabled=$false; Set-Conn $false "Connexion en cours..."
    try { Connect-AzAccount -ErrorAction Stop | Out-Null } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Connexion","OK","Error")|Out-Null }
    $ctx=Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connecte : $($ctx.Account.Id)"; $script:Guests=$null } else { Set-Conn $false "Non connecte." }
    $btnConnect.Enabled=$true
})

$btnResolve.Add_Click({
    if (-not (Ensure-Guests)) { return }
    $ids = Get-InputIds
    if ($ids.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Aucun identifiant (fichier et/ou collage).","Source vide","OK","Warning")|Out-Null; return }
    $grid.Rows.Clear(); $script:Results.Clear()
    $found=0
    foreach ($id in $ids) {
        if ($id -match '@') {
            $g = Resolve-GuestByMail $id
            if ($g) { Add-ResultRow $id $g "TROUVE"; $found++ } else { Add-ResultRow $id $null "INTROUVABLE" }
        } else {
            $hits = @(Resolve-GuestByName $id)
            if ($hits.Count -eq 0) { Add-ResultRow $id $null "INTROUVABLE" }
            elseif ($hits.Count -eq 1) { Add-ResultRow $id $hits[0] "TROUVE"; $found++ }
            else { foreach ($h in $hits) { Add-ResultRow $id $h "AMBIGU ($($hits.Count))" } }
        }
    }
    Write-Log "Resolution : $($ids.Count) identifiant(s), $found invite(s) trouve(s) sans ambiguite."
    $btnExport.Enabled = ($script:Results.Count -gt 0)
})

$btnListAll.Add_Click({
    if (-not (Ensure-Guests)) { return }
    $grid.Rows.Clear(); $script:Results.Clear()
    foreach ($g in $script:Guests) { Add-ResultRow "(inventaire)" $g "TROUVE" }
    Write-Log "Inventaire : $($script:Guests.Count) compte(s) invite(s)."
    $btnExport.Enabled = ($script:Results.Count -gt 0)
})

$btnExport.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="Fichier CSV (*.csv)|*.csv"; $dlg.FileName="Comptes_Invites.csv"
    if ($dlg.ShowDialog() -ne "OK") { return }
    try {
        $script:Results | Select-Object Entree,DisplayName,UserPrincipalName,Mail,OtherMails,AccountEnabled,Statut |
            Export-Csv -Path $dlg.FileName -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-Log "Export : $($dlg.FileName)"
        [System.Windows.Forms.MessageBox]::Show("Export termine :`n$($dlg.FileName)","Export","OK","Information")|Out-Null
    } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Export","OK","Error")|Out-Null }
})

# ==================================================================
# Adopter une connexion Az active
# ==================================================================
if (Get-Command Get-AzContext -ErrorAction SilentlyContinue) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connexion Az active reutilisee : $($ctx.Account.Id)"; Write-Log "Connexion Az detectee : $($ctx.Account.Id)" }
    else { Set-Conn $false "Non connecte. Cliquez sur 'Connexion Azure'." }
} else {
    Set-Conn $false "Module Az requis (Install-Module Az)."
}

[void]$form.ShowDialog()
