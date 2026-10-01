<#
.SYNOPSIS
    Ajout ou retrait en masse de membres d'un groupe Entra (Azure AD / cloud) - Interface graphique.
    Module : Az (Connect-AzAccount / Get-AzADGroup / *-AzADGroupMember).

.DESCRIPTION
    Source : un fichier .txt ET/OU une zone de collage (un identifiant par ligne : UPN, e-mail,
    ObjectId, ou nom affiche). Les deux sources sont fusionnees et dedoublonnees.

    Chaque identifiant est resolu vers un utilisateur Entra. Selon l'action (Ajouter / Retirer),
    l'outil calcule le statut (a ajouter / deja membre / a retirer / non membre / introuvable),
    puis applique.

    Garde-fous : mode simulation coche par defaut (lecture seule), confirmation avant ecriture,
    bouton Arreter, journal. Reutilise une session Az deja active.

.PREREQUIS
    - Module Az (Az.Accounts, Az.Resources) et droits d'ecriture sur le groupe (proprietaire /
      Groups Administrator / role equivalent).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Group     = $null
$script:GroupId   = $null
$script:MemberIds = $null
$script:Rows      = New-Object System.Collections.Generic.List[object]
$script:Cancel    = $false

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

function Test-AzConnected { try { return [bool](Get-AzContext -ErrorAction Stop) } catch { return $false } }

function Test-IsGuid { param([string]$s) return ($s -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') }

function Resolve-Group {
    param([string]$Id)
    $q = $Id.Trim(); if (-not $q) { return $null }
    if (Test-IsGuid $q) { try { return (Get-AzADGroup -ObjectId $q -ErrorAction Stop) } catch { } }
    try { $g = Get-AzADGroup -DisplayName $q -ErrorAction Stop; if ($g) { return ($g | Select-Object -First 1) } } catch { }
    $safe = $q.Replace("'", "''")
    try { $g = Get-AzADGroup -Filter "displayName eq '$safe'" -ErrorAction Stop; if ($g) { return ($g | Select-Object -First 1) } } catch { }
    return $null
}

function Resolve-User {
    param([string]$Id)
    $q = $Id.Trim(); if (-not $q) { return $null }
    $safe = $q.Replace("'", "''")
    if (Test-IsGuid $q) { try { $u = Get-AzADUser -ObjectId $q -ErrorAction SilentlyContinue; if ($u) { return $u } } catch { } }
    if ($q -match '@') {
        $u = Get-AzADUser -Filter "userPrincipalName eq '$safe' or mail eq '$safe'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($u) { return $u }
    } else {
        $u = Get-AzADUser -Filter "displayName eq '$safe'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($u) { return $u }
    }
    return $null
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
$form.Text = "Groupe Entra (cloud) - Ajout / Retrait en masse"
$form.Size = New-Object System.Drawing.Size(880, 810)
$form.MinimumSize = New-Object System.Drawing.Size(800, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connexion Azure"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Etat : verification..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(660,18); $lblConn.Anchor="Top,Left,Right"

# ---------- Groupe + action ----------
$gbTop = New-Object System.Windows.Forms.GroupBox
$gbTop.Text="Groupe cible et action"; $gbTop.Location=New-Object System.Drawing.Point(15,48); $gbTop.Size=New-Object System.Drawing.Size(835,110); $gbTop.Anchor="Top,Left,Right"
$form.Controls.Add($gbTop)

Add-Lbl $gbTop "Groupe (nom ou ObjectId) :" 14 26 | Out-Null
$txtGroup = New-Object System.Windows.Forms.TextBox; $txtGroup.Location=New-Object System.Drawing.Point(190,23); $txtGroup.Size=New-Object System.Drawing.Size(520,23); $txtGroup.Anchor="Top,Left,Right"; $gbTop.Controls.Add($txtGroup)
$btnCheck = New-Object System.Windows.Forms.Button; $btnCheck.Text="Verifier"; $btnCheck.Location=New-Object System.Drawing.Point(720,22); $btnCheck.Size=New-Object System.Drawing.Size(100,25); $btnCheck.Anchor="Top,Right"; $gbTop.Controls.Add($btnCheck)
$lblGroupInfo = Add-Lbl $gbTop "(non verifie)" 190 50; $lblGroupInfo.ForeColor=[System.Drawing.Color]::Gray; $lblGroupInfo.AutoSize=$false; $lblGroupInfo.Size=New-Object System.Drawing.Size(630,18); $lblGroupInfo.Anchor="Top,Left,Right"

$rbAdd = New-Object System.Windows.Forms.RadioButton; $rbAdd.Text="Ajouter au groupe"; $rbAdd.Location=New-Object System.Drawing.Point(14,78); $rbAdd.AutoSize=$true; $rbAdd.Checked=$true; $gbTop.Controls.Add($rbAdd)
$rbRemove = New-Object System.Windows.Forms.RadioButton; $rbRemove.Text="Retirer du groupe"; $rbRemove.Location=New-Object System.Drawing.Point(200,78); $rbRemove.AutoSize=$true; $gbTop.Controls.Add($rbRemove)

# ---------- Source ----------
$gbSrc = New-Object System.Windows.Forms.GroupBox
$gbSrc.Text="Source (un identifiant par ligne : UPN, e-mail, ObjectId ou nom)"; $gbSrc.Location=New-Object System.Drawing.Point(15,168); $gbSrc.Size=New-Object System.Drawing.Size(835,150); $gbSrc.Anchor="Top,Left,Right"
$form.Controls.Add($gbSrc)

Add-Lbl $gbSrc "Fichier .txt :" 14 26 | Out-Null
$txtFile = New-Object System.Windows.Forms.TextBox; $txtFile.Location=New-Object System.Drawing.Point(110,23); $txtFile.Size=New-Object System.Drawing.Size(605,23); $txtFile.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtFile)
$btnBrowse = New-Object System.Windows.Forms.Button; $btnBrowse.Text="Parcourir..."; $btnBrowse.Location=New-Object System.Drawing.Point(725,22); $btnBrowse.Size=New-Object System.Drawing.Size(95,25); $btnBrowse.Anchor="Top,Right"; $gbSrc.Controls.Add($btnBrowse)
Add-Lbl $gbSrc "Ou coller :" 14 56 | Out-Null
$txtPaste = New-Object System.Windows.Forms.TextBox; $txtPaste.Location=New-Object System.Drawing.Point(110,53); $txtPaste.Size=New-Object System.Drawing.Size(710,80); $txtPaste.Multiline=$true; $txtPaste.ScrollBars="Vertical"; $txtPaste.WordWrap=$false; $txtPaste.Anchor="Top,Left,Right"; $gbSrc.Controls.Add($txtPaste)

# ---------- Actions ----------
$btnResolve = New-Object System.Windows.Forms.Button; $btnResolve.Text="Resoudre / Analyser"; $btnResolve.Location=New-Object System.Drawing.Point(15,330); $btnResolve.Size=New-Object System.Drawing.Size(200,28); $form.Controls.Add($btnResolve)
$chkSim = New-Object System.Windows.Forms.CheckBox; $chkSim.Text="Mode simulation (aucune modification)"; $chkSim.Location=New-Object System.Drawing.Point(235,334); $chkSim.AutoSize=$true; $chkSim.Checked=$true; $form.Controls.Add($chkSim)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location=New-Object System.Drawing.Point(15,365); $grid.Size=New-Object System.Drawing.Size(835,200); $grid.Anchor="Top,Left,Right"
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.ReadOnly=$true; $grid.RowHeadersVisible=$false
$grid.SelectionMode="FullRowSelect"; $grid.AutoSizeColumnsMode="Fill"; $grid.ColumnHeadersHeightSizeMode="AutoSize"
$null=$grid.Columns.Add("Input","Entree")
$null=$grid.Columns.Add("Upn","UPN")
$null=$grid.Columns.Add("Display","Nom affiche")
$null=$grid.Columns.Add("Statut","Statut")
$null=$grid.Columns.Add("Oid","ObjectId")
$grid.Columns["Input"].FillWeight=26; $grid.Columns["Upn"].FillWeight=30; $grid.Columns["Display"].FillWeight=26; $grid.Columns["Statut"].FillWeight=18
$grid.Columns["Oid"].Visible=$false
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
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

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
$btnBrowse.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Fichiers texte (*.txt)|*.txt|Tous (*.*)|*.*"; if($dlg.ShowDialog() -eq "OK"){$txtFile.Text=$dlg.FileName} })

$btnConnect.Add_Click({
    if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) { [System.Windows.Forms.MessageBox]::Show("Module Az introuvable.`nInstall-Module Az -Scope CurrentUser","Prerequis","OK","Error")|Out-Null; return }
    $btnConnect.Enabled=$false; Set-Conn $false "Connexion en cours..."
    try { Connect-AzAccount -ErrorAction Stop | Out-Null } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Connexion","OK","Error")|Out-Null }
    $ctx=Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connecte : $($ctx.Account.Id)" } else { Set-Conn $false "Non connecte." }
    $btnConnect.Enabled=$true
})

$btnCheck.Add_Click({
    if (-not (Test-AzConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a Azure.","Connexion","OK","Warning")|Out-Null; return }
    $script:Group = Resolve-Group $txtGroup.Text
    if ($script:Group) {
        $script:GroupId = [string](Get-GProp $script:Group 'Id')
        $lblGroupInfo.ForeColor=[System.Drawing.Color]::ForestGreen
        $lblGroupInfo.Text="OK : $([string](Get-GProp $script:Group 'DisplayName'))  |  ObjectId $($script:GroupId)"
        Write-Log "Groupe resolu : $([string](Get-GProp $script:Group 'DisplayName')) ($($script:GroupId))"
    } else { $lblGroupInfo.ForeColor=[System.Drawing.Color]::Firebrick; $lblGroupInfo.Text="Groupe introuvable."; Write-Log "Groupe introuvable : $($txtGroup.Text)" }
})

$btnResolve.Add_Click({
    if (-not (Test-AzConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a Azure.","Connexion","OK","Warning")|Out-Null; return }
    if (-not $script:Group) { $script:Group = Resolve-Group $txtGroup.Text; if ($script:Group) { $script:GroupId=[string](Get-GProp $script:Group 'Id') } }
    if (-not $script:Group) { [System.Windows.Forms.MessageBox]::Show("Verifiez d'abord le groupe cible.","Groupe","OK","Warning")|Out-Null; return }

    $ids = Get-InputIds
    if ($ids.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Aucun identifiant (fichier et/ou collage).","Source vide","OK","Warning")|Out-Null; return }

    $script:MemberIds = $null
    try {
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($m in @(Get-AzADGroupMember -GroupObjectId $script:GroupId -ErrorAction Stop)) { $mid=[string](Get-GProp $m 'Id'); if ($mid) { [void]$set.Add($mid) } }
        $script:MemberIds = $set
        Write-Log "Membres actuels du groupe : $($set.Count)"
    } catch { Write-Log "Enumeration des membres impossible : statut determine a l'execution." }

    $isAdd = $rbAdd.Checked
    $grid.Rows.Clear(); $script:Rows.Clear()
    foreach ($id in $ids) {
        $u = Resolve-User $id
        $upn=""; $disp=""; $oid=""; $st=""
        if (-not $u) { $st="INTROUVABLE" }
        else {
            $upn=[string](Get-GProp $u 'UserPrincipalName'); $disp=[string](Get-GProp $u 'DisplayName'); $oid=[string](Get-GProp $u 'Id')
            if ($null -ne $script:MemberIds) {
                $isMember = $script:MemberIds.Contains($oid)
                if ($isAdd) { $st = if ($isMember) {"DEJA MEMBRE"} else {"A AJOUTER"} }
                else        { $st = if ($isMember) {"A RETIRER"}  else {"NON MEMBRE"} }
            } else { $st = "A TRAITER" }
        }
        $obj=[PSCustomObject]@{ Input=$id; Upn=$upn; Display=$disp; Oid=$oid; Statut=$st }
        $script:Rows.Add($obj)|Out-Null
        $i=$grid.Rows.Add($id,$upn,$disp,$st,$oid); Set-RowStyle $grid.Rows[$i] $st
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
    $recap = "$verbe $($targets.Count) membre(s)`n$(if($isAdd){'au'}else{'du'}) groupe Entra :`n$([string](Get-GProp $script:Group 'DisplayName'))`nObjectId $($script:GroupId)`n`nConfirmer ?"
    if ([System.Windows.Forms.MessageBox]::Show($recap,"Confirmation","YesNo","Warning") -ne "Yes") { Write-Log "Execution annulee."; return }

    $script:Cancel=$false
    $btnExecute.Enabled=$false; $btnResolve.Enabled=$false; $btnStop.Enabled=$true
    $progress.Minimum=0; $progress.Maximum=$targets.Count; $progress.Value=0
    $gid=$script:GroupId; $done=0; $ok=0; $skip=0; $err=0

    foreach ($idx in $targets) {
        if ($script:Cancel) { break }
        $done++
        $row=$script:Rows[$idx]; $gr=$grid.Rows[$idx]
        if (-not $row.Oid) { $row.Statut="ERREUR"; $err++; $gr.Cells["Statut"].Value="ERREUR"; Set-RowStyle $gr "ERREUR"; $progress.Value=$done; continue }
        try {
            if ($isAdd) { Add-AzADGroupMember -TargetGroupObjectId $gid -MemberObjectId $row.Oid -ErrorAction Stop; $row.Statut="AJOUTE"; $ok++ }
            else        { Remove-AzADGroupMember -GroupObjectId $gid -MemberObjectId $row.Oid -ErrorAction Stop; $row.Statut="RETIRE"; $ok++ }
        } catch {
            $msg=$_.Exception.Message
            if     ($msg -match 'already exist|added object references already') { $row.Statut="DEJA MEMBRE"; $skip++ }
            elseif ($msg -match 'does not exist|not present|Request_ResourceNotFound') { $row.Statut="NON MEMBRE"; $skip++ }
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
# Adopter une connexion Az active
# ==================================================================
if (Get-Command Get-AzContext -ErrorAction SilentlyContinue) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connexion Az active reutilisee : $($ctx.Account.Id)"; Write-Log "Connexion Az detectee : $($ctx.Account.Id)" }
    else { Set-Conn $false "Non connecte. Cliquez sur 'Connexion Azure'." }
} else { Set-Conn $false "Module Az requis (Install-Module Az)." }

[void]$form.ShowDialog()
