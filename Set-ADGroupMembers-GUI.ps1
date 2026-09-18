<#
.SYNOPSIS
    Ajout ou retrait en masse d'utilisateurs dans un groupe Active Directory - Interface graphique.

.DESCRIPTION
    Source des utilisateurs : un fichier .txt ET/OU une zone de collage (un identifiant par ligne :
    sAMAccountName, UPN, e-mail ou DistinguishedName). Les deux sources sont fusionnees et dedoublonnees.

    Chaque identifiant est resolu dans AD. Selon l'action choisie (Ajouter / Retirer), l'outil calcule
    le statut (a ajouter / deja membre / a retirer / non membre / introuvable), puis applique.

    Garde-fous : mode simulation coche par defaut (lecture seule), confirmation avant ecriture,
    bouton Arreter, journal complet.

.PREREQUIS
    - Module ActiveDirectory (RSAT) et droits d'ecriture sur le groupe cible.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Group     = $null
$script:MemberDNs = $null
$script:Rows      = New-Object System.Collections.Generic.List[object]
$script:Cancel    = $false

# ==================================================================
# Fonctions AD
# ==================================================================
function Get-SrvParam { if ($txtServer.Text.Trim()) { return @{ Server = $txtServer.Text.Trim() } } else { return @{} } }

function Resolve-Group {
    param([string]$Id)
    $srv = Get-SrvParam
    $q = $Id.Trim(); if (-not $q) { return $null }
    try { return (Get-ADGroup -Identity $q -Properties DistinguishedName @srv -ErrorAction Stop) } catch { }
    $safe = $q.Replace("'", "''")
    try { return (Get-ADGroup -Filter "Name -eq '$safe' -or SamAccountName -eq '$safe'" @srv -ErrorAction Stop | Select-Object -First 1) } catch { }
    return $null
}

function Resolve-User {
    param([string]$Id)
    $srv = Get-SrvParam
    $q = $Id.Trim(); if (-not $q) { return $null }
    $safe = $q.Replace("'", "''")
    $u = Get-ADUser -Filter "SamAccountName -eq '$safe' -or UserPrincipalName -eq '$safe' -or mail -eq '$safe'" -Properties mail, DisplayName @srv -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($u) { return $u }
    if ($q -match '(?i)^CN=') { try { return (Get-ADUser -Identity $q -Properties DisplayName @srv -ErrorAction Stop) } catch { } }
    return $null
}

function Get-InputIds {
    $raw = @()
    if (-not [string]::IsNullOrWhiteSpace($txtFile.Text) -and (Test-Path -LiteralPath $txtFile.Text)) {
        $raw += Get-Content -LiteralPath $txtFile.Text -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($txtPaste.Text)) { $raw += ($txtPaste.Text -split "`r?`n") }
    return @($raw | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Groupe AD - Ajout / Retrait en masse"
$form.Size = New-Object System.Drawing.Size(880, 810)
$form.MinimumSize = New-Object System.Drawing.Size(800, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

# ---------- Groupe + action ----------
$gbTop = New-Object System.Windows.Forms.GroupBox
$gbTop.Text = "Groupe cible et action"; $gbTop.Location = New-Object System.Drawing.Point(15,10); $gbTop.Size = New-Object System.Drawing.Size(835,115); $gbTop.Anchor="Top,Left,Right"
$form.Controls.Add($gbTop)

Add-Lbl $gbTop "Groupe AD :" 14 26 | Out-Null
$txtGroup = New-Object System.Windows.Forms.TextBox; $txtGroup.Location=New-Object System.Drawing.Point(100,23); $txtGroup.Size=New-Object System.Drawing.Size(420,23); $gbTop.Controls.Add($txtGroup)
$btnCheck = New-Object System.Windows.Forms.Button; $btnCheck.Text="Verifier"; $btnCheck.Location=New-Object System.Drawing.Point(528,22); $btnCheck.Size=New-Object System.Drawing.Size(95,25); $gbTop.Controls.Add($btnCheck)
Add-Lbl $gbTop "Serveur AD :" 635 26 | Out-Null
$txtServer = New-Object System.Windows.Forms.TextBox; $txtServer.Location=New-Object System.Drawing.Point(715,23); $txtServer.Size=New-Object System.Drawing.Size(105,23); $txtServer.Anchor="Top,Right"; $gbTop.Controls.Add($txtServer)

$lblGroupInfo = Add-Lbl $gbTop "(non verifie)" 100 51; $lblGroupInfo.ForeColor=[System.Drawing.Color]::Gray; $lblGroupInfo.AutoSize=$false; $lblGroupInfo.Size=New-Object System.Drawing.Size(720,18); $lblGroupInfo.Anchor="Top,Left,Right"

$rbAdd = New-Object System.Windows.Forms.RadioButton; $rbAdd.Text="Ajouter au groupe"; $rbAdd.Location=New-Object System.Drawing.Point(14,80); $rbAdd.AutoSize=$true; $rbAdd.Checked=$true; $gbTop.Controls.Add($rbAdd)
$rbRemove = New-Object System.Windows.Forms.RadioButton; $rbRemove.Text="Retirer du groupe"; $rbRemove.Location=New-Object System.Drawing.Point(200,80); $rbRemove.AutoSize=$true; $gbTop.Controls.Add($rbRemove)

# ---------- Source ----------
$gbSrc = New-Object System.Windows.Forms.GroupBox
$gbSrc.Text = "Source des utilisateurs (un identifiant par ligne : sAMAccountName, UPN, e-mail ou DN)"
$gbSrc.Location = New-Object System.Drawing.Point(15,135); $gbSrc.Size = New-Object System.Drawing.Size(835,150); $gbSrc.Anchor="Top,Left,Right"
$form.Controls.Add($gbSrc)

Add-Lbl $gbSrc "Fichier .txt :" 14 26 | Out-Null
$txtFile = New-Object System.Windows.Forms.TextBox; $txtFile.Location=New-Object System.Drawing.Point(110,23); $txtFile.Size=New-Object System.Drawing.Size(605,23); $txtFile.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtFile)
$btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text="Parcourir..."; $btnBrowse.Location=New-Object System.Drawing.Point(725,22); $btnBrowse.Size=New-Object System.Drawing.Size(95,25); $btnBrowse.Anchor="Top,Right"; $gbSrc.Controls.Add($btnBrowse)

Add-Lbl $gbSrc "Ou coller :" 14 56 | Out-Null
$txtPaste = New-Object System.Windows.Forms.TextBox; $txtPaste.Location=New-Object System.Drawing.Point(110,53); $txtPaste.Size=New-Object System.Drawing.Size(710,80); $txtPaste.Multiline=$true; $txtPaste.ScrollBars="Vertical"; $txtPaste.WordWrap=$false; $txtPaste.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtPaste)

# ---------- Actions ----------
$btnResolve = New-Object System.Windows.Forms.Button; $btnResolve.Text="Resoudre / Analyser"; $btnResolve.Location=New-Object System.Drawing.Point(15,298); $btnResolve.Size=New-Object System.Drawing.Size(200,28); $form.Controls.Add($btnResolve)
$chkSim = New-Object System.Windows.Forms.CheckBox; $chkSim.Text="Mode simulation (aucune modification)"; $chkSim.Location=New-Object System.Drawing.Point(235,302); $chkSim.AutoSize=$true; $chkSim.Checked=$true; $form.Controls.Add($chkSim)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(15,335); $grid.Size=New-Object System.Drawing.Size(835,230); $grid.Anchor="Top,Left,Right"
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.ReadOnly=$true; $grid.RowHeadersVisible=$false
$grid.SelectionMode="FullRowSelect"; $grid.AutoSizeColumnsMode="Fill"; $grid.ColumnHeadersHeightSizeMode="AutoSize"
$null=$grid.Columns.Add("Input","Entree")
$null=$grid.Columns.Add("Sam","Compte")
$null=$grid.Columns.Add("Display","Nom affiche")
$null=$grid.Columns.Add("Statut","Statut")
$null=$grid.Columns.Add("DN","DN")
$grid.Columns["Input"].FillWeight=24; $grid.Columns["Sam"].FillWeight=16; $grid.Columns["Display"].FillWeight=24; $grid.Columns["Statut"].FillWeight=18
$grid.Columns["DN"].Visible=$false
$form.Controls.Add($grid)

$btnExecute = New-Object System.Windows.Forms.Button; $btnExecute.Text="Executer"; $btnExecute.Location=New-Object System.Drawing.Point(15,575); $btnExecute.Size=New-Object System.Drawing.Size(250,30)
$btnExecute.BackColor=[System.Drawing.Color]::FromArgb(0,120,215); $btnExecute.ForeColor=[System.Drawing.Color]::White; $btnExecute.FlatStyle="Flat"; $btnExecute.Enabled=$false; $form.Controls.Add($btnExecute)
$btnStop = New-Object System.Windows.Forms.Button; $btnStop.Text="Arreter"; $btnStop.Location=New-Object System.Drawing.Point(275,575); $btnStop.Size=New-Object System.Drawing.Size(110,30)
$btnStop.BackColor=[System.Drawing.Color]::FromArgb(200,60,60); $btnStop.ForeColor=[System.Drawing.Color]::White; $btnStop.FlatStyle="Flat"; $btnStop.Enabled=$false; $form.Controls.Add($btnStop)
$progress = New-Object System.Windows.Forms.ProgressBar; $progress.Location=New-Object System.Drawing.Point(400,579); $progress.Size=New-Object System.Drawing.Size(450,22); $progress.Anchor="Top,Left,Right"; $form.Controls.Add($progress)

Add-Lbl $form "Journal :" 15 612 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,632); $txtLog.Size=New-Object System.Drawing.Size(835,120); $txtLog.Anchor="Top,Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# Helpers UI
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }

function Set-RowStyle { param($row,[string]$s)
    switch -Wildcard ($s) {
        "A AJOUTER"   { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(213,245,213) }
        "A RETIRER"   { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,224,150) }
        "AJOUTE"      { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(198,226,255) }
        "RETIRE"      { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(198,226,255) }
        "DEJA MEMBRE" { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(235,235,235) }
        "NON MEMBRE"  { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(235,235,235) }
        "INTROUVABLE" { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,190,190); $row.DefaultCellStyle.ForeColor=[System.Drawing.Color]::DarkRed }
        "ERREUR"      { $row.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(255,190,190); $row.DefaultCellStyle.ForeColor=[System.Drawing.Color]::DarkRed }
    }
}

# ==================================================================
# Evenements
# ==================================================================
$btnBrowse.Add_Click({
    $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Fichiers texte (*.txt)|*.txt|Tous (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtFile.Text=$dlg.FileName }
})

$btnCheck.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { [System.Windows.Forms.MessageBox]::Show("Module ActiveDirectory introuvable.","Prerequis","OK","Error")|Out-Null; return }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    $script:Group = Resolve-Group $txtGroup.Text
    if ($script:Group) {
        $lblGroupInfo.ForeColor=[System.Drawing.Color]::ForestGreen
        $lblGroupInfo.Text="OK : $($script:Group.Name)  |  $($script:Group.DistinguishedName)"
        Write-Log "Groupe resolu : $($script:Group.DistinguishedName)"
    } else {
        $lblGroupInfo.ForeColor=[System.Drawing.Color]::Firebrick; $lblGroupInfo.Text="Groupe introuvable."
        Write-Log "Groupe introuvable : $($txtGroup.Text)"
    }
})

$btnResolve.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { [System.Windows.Forms.MessageBox]::Show("Module ActiveDirectory introuvable.","Prerequis","OK","Error")|Out-Null; return }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    if (-not $script:Group) { $script:Group = Resolve-Group $txtGroup.Text }
    if (-not $script:Group) { [System.Windows.Forms.MessageBox]::Show("Verifiez d'abord le groupe cible.","Groupe","OK","Warning")|Out-Null; return }

    $ids = Get-InputIds
    if ($ids.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Aucun identifiant (fichier et/ou collage).","Source vide","OK","Warning")|Out-Null; return }

    $srv = Get-SrvParam
    $script:MemberDNs = $null
    try {
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($m in @(Get-ADGroupMember -Identity $script:Group.DistinguishedName @srv -ErrorAction Stop)) { [void]$set.Add($m.distinguishedName) }
        $script:MemberDNs = $set
        Write-Log "Membres actuels du groupe : $($set.Count)"
    } catch { Write-Log "Enumeration des membres impossible (groupe volumineux ?) - statut determine a l'execution." }

    $isAdd = $rbAdd.Checked
    $grid.Rows.Clear(); $script:Rows.Clear()
    foreach ($id in $ids) {
        $u = Resolve-User $id
        $sam=""; $disp=""; $dn=""; $st=""
        if (-not $u) { $st="INTROUVABLE" }
        else {
            $sam=$u.SamAccountName; $disp=$u.DisplayName; $dn=$u.DistinguishedName
            if ($null -ne $script:MemberDNs) {
                $isMember = $script:MemberDNs.Contains($dn)
                if ($isAdd) { $st = if ($isMember) {"DEJA MEMBRE"} else {"A AJOUTER"} }
                else        { $st = if ($isMember) {"A RETIRER"}  else {"NON MEMBRE"} }
            } else { $st = "A TRAITER" }
        }
        $obj=[PSCustomObject]@{ Input=$id; Sam=$sam; Display=$disp; DN=$dn; Statut=$st }
        $script:Rows.Add($obj)|Out-Null
        $i=$grid.Rows.Add($id,$sam,$disp,$st,$dn); Set-RowStyle $grid.Rows[$i] $st
    }
    $todo = @($script:Rows | Where-Object { $_.Statut -in @("A AJOUTER","A RETIRER","A TRAITER") }).Count
    $nf   = @($script:Rows | Where-Object { $_.Statut -eq "INTROUVABLE" }).Count
    Write-Log "Analyse : $($ids.Count) identifiant(s), $todo a traiter, $nf introuvable(s)."
    $btnExecute.Enabled = ($todo -gt 0)
})

$btnStop.Add_Click({ $script:Cancel=$true; $btnStop.Enabled=$false; Write-Log "Arret demande..." })

$btnExecute.Add_Click({
    if ($chkSim.Checked) { [System.Windows.Forms.MessageBox]::Show("Mode simulation actif : decochez-le pour appliquer.","Execution","OK","Information")|Out-Null; return }
    if (-not $script:Group) { return }
    $isAdd = $rbAdd.Checked
    $targets = @(0..($script:Rows.Count-1) | Where-Object { $script:Rows[$_].Statut -in @("A AJOUTER","A RETIRER","A TRAITER") })
    if ($targets.Count -eq 0) { return }

    $verbe = if ($isAdd) {"AJOUTER"} else {"RETIRER"}
    $recap = "$verbe $($targets.Count) utilisateur(s)`n$(if($isAdd){'au'}else{'du'}) groupe :`n$($script:Group.Name)`n$($script:Group.DistinguishedName)`n`nConfirmer ?"
    if ([System.Windows.Forms.MessageBox]::Show($recap,"Confirmation","YesNo","Warning") -ne "Yes") { Write-Log "Execution annulee."; return }

    $script:Cancel=$false
    $btnExecute.Enabled=$false; $btnResolve.Enabled=$false; $btnStop.Enabled=$true
    $progress.Minimum=0; $progress.Maximum=$targets.Count; $progress.Value=0
    $srv = Get-SrvParam; $gDN=$script:Group.DistinguishedName
    $done=0; $ok=0; $skip=0; $err=0

    foreach ($idx in $targets) {
        if ($script:Cancel) { break }
        $done++
        $row=$script:Rows[$idx]; $gr=$grid.Rows[$idx]
        try {
            if ($isAdd) {
                Add-ADGroupMember -Identity $gDN -Members $row.DN @srv -ErrorAction Stop
                $row.Statut="AJOUTE"; $ok++
            } else {
                Remove-ADGroupMember -Identity $gDN -Members $row.DN -Confirm:$false @srv -ErrorAction Stop
                $row.Statut="RETIRE"; $ok++
            }
        } catch {
            $msg=$_.Exception.Message
            if ($msg -match 'already a member|deja membre')      { $row.Statut="DEJA MEMBRE"; $skip++ }
            elseif ($msg -match 'not a member|pas membre|0x561') { $row.Statut="NON MEMBRE";  $skip++ }
            else { $row.Statut="ERREUR"; $err++; Write-Log "  ! $($row.Input) : $msg" }
        }
        $gr.Cells["Statut"].Value=$row.Statut; Set-RowStyle $gr $row.Statut
        $progress.Value=$done; [System.Windows.Forms.Application]::DoEvents()
    }

    $suffix = if ($script:Cancel) { " [INTERROMPU]" } else { "" }
    Write-Log "Termine$suffix : $ok applique(s), $skip ignore(s), $err erreur(s)."
    $btnStop.Enabled=$false; $btnResolve.Enabled=$true; $btnExecute.Enabled=$true
})

# ==================================================================
[void]$form.ShowDialog()
