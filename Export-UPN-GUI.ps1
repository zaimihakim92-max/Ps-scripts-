<#
.SYNOPSIS
    Export des UPN (Entra ID) a partir d'une liste d'adresses email - Interface graphique.

.DESCRIPTION
    Pour chaque adresse email, recherche le compte dans Entra ID selon 3 methodes
    (attribut mail, UserPrincipalName, adresses secondaires). Les comptes non trouves
    sont clairement identifies. En cas de migration de domaine (ex : aldautomotive -> ayvens),
    un ou plusieurs DOMAINES ALTERNATIFS peuvent etre essayes : la partie locale de
    l'adresse est reforgee avec le nouveau domaine, puis re-recherchee.

    - Recherche auto des domaines alternatifs pendant le traitement (case a cocher).
    - Bouton "Resoudre les non trouves" pour retenter APRES coup, sans tout relancer.
    - Aucune erreur bloquante : tout est capture et reporte dans la grille.
    - Double-clic sur une ligne = copie de l'UPN dans le presse-papier.

.PREREQUIS
    - Module Az (Az.Accounts, Az.Resources)
    - Connexion via le bouton "Connexion Azure" (ou Connect-AzAccount prealable)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Results = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Fonctions metier
# ==================================================================

function Test-AzReady {
    $r = [PSCustomObject]@{ ModuleOk = $false; Connected = $false; Account = $null }
    if (Get-Module -ListAvailable -Name Az.Accounts) { $r.ModuleOk = $true }
    try {
        $ctx = Get-AzContext -ErrorAction Stop
        if ($ctx) { $r.Connected = $true; $r.Account = $ctx.Account.Id }
    } catch { }
    return $r
}

function Find-EntraUser {
    # Recherche multi-attributs sur une adresse donnee. Retourne l'objet user ou $null.
    param([Parameter(Mandatory)][string]$Email)
    $safe = $Email.Replace("'", "''")
    try {
        $u = Get-AzADUser -Mail $Email -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }

        $u = Get-AzADUser -UserPrincipalName $Email -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }

        $u = Get-AzADUser -Filter "otherMails/any(c:c eq '$safe')" -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }
    } catch { }
    return $null
}

function Resolve-EmailToUser {
    # Essaie l'adresse d'origine puis, si absente, la reforge avec chaque domaine alternatif.
    param(
        [Parameter(Mandatory)][string]$Email,
        [string[]]$AltDomains = @()
    )
    $u = Find-EntraUser -Email $Email
    if ($u) {
        return [PSCustomObject]@{ User = $u; MatchedAddress = $Email; MatchedBy = "Adresse d'origine" }
    }

    if ($AltDomains.Count -gt 0 -and $Email -match '^(.+?)@(.+)$') {
        $local = $Matches[1]
        foreach ($d in $AltDomains) {
            $d = $d.Trim().TrimStart('@')
            if ([string]::IsNullOrWhiteSpace($d)) { continue }
            $forged = "$local@$d"
            if ($forged -ieq $Email) { continue }
            $u2 = Find-EntraUser -Email $forged
            if ($u2) {
                return [PSCustomObject]@{ User = $u2; MatchedAddress = $forged; MatchedBy = "Domaine alternatif ($d)" }
            }
        }
    }
    return [PSCustomObject]@{ User = $null; MatchedAddress = $null; MatchedBy = $null }
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "Export UPN Entra ID"
$form.Size          = New-Object System.Drawing.Size(920, 720)
$form.MinimumSize   = New-Object System.Drawing.Size(820, 600)
$form.StartPosition = "CenterScreen"
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Fichier source ---
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Fichier .txt :"; $lblInput.Location = New-Object System.Drawing.Point(15, 15); $lblInput.AutoSize = $true
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(150, 12); $txtInput.Size = New-Object System.Drawing.Size(625, 23)
$txtInput.Anchor = "Top,Left,Right"; $txtInput.Text = "C:\Temp\emails.txt"
$form.Controls.Add($txtInput)

$btnInput = New-Object System.Windows.Forms.Button
$btnInput.Text = "Parcourir..."; $btnInput.Location = New-Object System.Drawing.Point(785, 11); $btnInput.Size = New-Object System.Drawing.Size(105, 25)
$btnInput.Anchor = "Top,Right"
$form.Controls.Add($btnInput)

# --- Fichier export ---
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Export CSV :"; $lblOutput.Location = New-Object System.Drawing.Point(15, 48); $lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(150, 45); $txtOutput.Size = New-Object System.Drawing.Size(625, 23)
$txtOutput.Anchor = "Top,Left,Right"; $txtOutput.Text = "C:\Temp\Resultat_UPN.csv"
$form.Controls.Add($txtOutput)

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = "Parcourir..."; $btnOutput.Location = New-Object System.Drawing.Point(785, 44); $btnOutput.Size = New-Object System.Drawing.Size(105, 25)
$btnOutput.Anchor = "Top,Right"
$form.Controls.Add($btnOutput)

# --- Domaines alternatifs ---
$lblAlt = New-Object System.Windows.Forms.Label
$lblAlt.Text = "Domaines alt. :"; $lblAlt.Location = New-Object System.Drawing.Point(15, 81); $lblAlt.AutoSize = $true
$form.Controls.Add($lblAlt)

$txtAlt = New-Object System.Windows.Forms.TextBox
$txtAlt.Location = New-Object System.Drawing.Point(150, 78); $txtAlt.Size = New-Object System.Drawing.Size(740, 23)
$txtAlt.Anchor = "Top,Left,Right"
$txtAlt.Text = "ayvens.com"
$form.Controls.Add($txtAlt)

$lblAltHint = New-Object System.Windows.Forms.Label
$lblAltHint.Text = "Separez plusieurs domaines par une virgule. Ex : jean.dupont@aldautomotive.com --> jean.dupont@ayvens.com"
$lblAltHint.Location = New-Object System.Drawing.Point(150, 103); $lblAltHint.AutoSize = $true
$lblAltHint.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblAltHint)

$chkAutoAlt = New-Object System.Windows.Forms.CheckBox
$chkAutoAlt.Text = "Essayer les domaines alternatifs pour les comptes non trouves"
$chkAutoAlt.Location = New-Object System.Drawing.Point(150, 126); $chkAutoAlt.AutoSize = $true; $chkAutoAlt.Checked = $true
$form.Controls.Add($chkAutoAlt)

# --- Statut Azure + connexion ---
$lblAz = New-Object System.Windows.Forms.Label
$lblAz.Location = New-Object System.Drawing.Point(15, 156); $lblAz.Size = New-Object System.Drawing.Size(600, 20)
$lblAz.Anchor = "Top,Left,Right"; $lblAz.Text = "Etat Azure : verification..."
$form.Controls.Add($lblAz)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connexion Azure"; $btnConnect.Location = New-Object System.Drawing.Point(765, 152); $btnConnect.Size = New-Object System.Drawing.Size(125, 25)
$btnConnect.Anchor = "Top,Right"
$form.Controls.Add($btnConnect)

# --- Lancer + progression ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Lancer le traitement"; $btnRun.Location = New-Object System.Drawing.Point(15, 188); $btnRun.Size = New-Object System.Drawing.Size(250, 30)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215); $btnRun.ForeColor = [System.Drawing.Color]::White; $btnRun.FlatStyle = "Flat"
$form.Controls.Add($btnRun)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(280, 192); $progress.Size = New-Object System.Drawing.Size(610, 22)
$progress.Anchor = "Top,Left,Right"
$form.Controls.Add($progress)

# --- Grille ---
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 232); $grid.Size = New-Object System.Drawing.Size(875, 375)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.ReadOnly = $true
$grid.RowHeadersVisible = $false; $grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"; $grid.ColumnHeadersHeightSizeMode = "AutoSize"

$null = $grid.Columns.Add("Ligne", "#")
$null = $grid.Columns.Add("Mail", "Adresse d'origine")
$null = $grid.Columns.Add("Statut", "Statut")
$null = $grid.Columns.Add("AdresseResolue", "Adresse resolue")
$null = $grid.Columns.Add("UPN", "UserPrincipalName")
$null = $grid.Columns.Add("DisplayName", "Nom affiche")
$null = $grid.Columns.Add("ObjectId", "ObjectId")
$null = $grid.Columns.Add("Detail", "Detail")

$grid.Columns["Ligne"].FillWeight = 6
$grid.Columns["Mail"].FillWeight = 19
$grid.Columns["Statut"].FillWeight = 11
$grid.Columns["AdresseResolue"].FillWeight = 19
$grid.Columns["UPN"].FillWeight = 20
$grid.Columns["DisplayName"].FillWeight = 14
$grid.Columns["ObjectId"].Visible = $false      # exporte dans le CSV, masque a l'ecran
$grid.Columns["Detail"].FillWeight = 18
$form.Controls.Add($grid)

# --- Recap + actions ---
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Location = New-Object System.Drawing.Point(15, 620); $lblSummary.Size = New-Object System.Drawing.Size(500, 40)
$lblSummary.Anchor = "Bottom,Left"
$lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSummary.Text = "En attente..."
$form.Controls.Add($lblSummary)

$btnResolve = New-Object System.Windows.Forms.Button
$btnResolve.Text = "Resoudre les non trouves"; $btnResolve.Location = New-Object System.Drawing.Point(555, 620); $btnResolve.Size = New-Object System.Drawing.Size(180, 30)
$btnResolve.Anchor = "Bottom,Right"; $btnResolve.Enabled = $false
$form.Controls.Add($btnResolve)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exporter en CSV"; $btnExport.Location = New-Object System.Drawing.Point(745, 620); $btnExport.Size = New-Object System.Drawing.Size(145, 30)
$btnExport.Anchor = "Bottom,Right"; $btnExport.Enabled = $false
$form.Controls.Add($btnExport)

# ==================================================================
# Helpers UI
# ==================================================================
function Update-AzStatus {
    $az = Test-AzReady
    if (-not $az.ModuleOk) {
        $lblAz.Text = "Etat Azure : module Az.Accounts introuvable."; $lblAz.ForeColor = [System.Drawing.Color]::Firebrick
    } elseif ($az.Connected) {
        $lblAz.Text = "Etat Azure : connecte ($($az.Account))"; $lblAz.ForeColor = [System.Drawing.Color]::ForestGreen
    } else {
        $lblAz.Text = "Etat Azure : non connecte - cliquez sur 'Connexion Azure'."; $lblAz.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    return $az
}

function Get-AltDomains {
    if ([string]::IsNullOrWhiteSpace($txtAlt.Text)) { return @() }
    return @($txtAlt.Text -split '[,;\s]+' | ForEach-Object { $_.Trim().TrimStart('@') } |
             Where-Object { $_ } | Select-Object -Unique)
}

function Set-RowStyle {
    param($row, [string]$statut)
    switch ($statut) {
        "TROUVE"       { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(213,245,213); $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black }
        "TROUVE (ALT)" { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(198,226,255); $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0,40,110) }
        "NON TROUVE"   { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,224,150); $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(120,60,0) }
        "ERREUR"       { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,190,190); $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
    }
}

function Write-GridRow {
    param([int]$rowIndex, $res)
    $r = $grid.Rows[$rowIndex]
    $r.Cells["Ligne"].Value          = $res.Ligne
    $r.Cells["Mail"].Value           = $res.Mail
    $r.Cells["Statut"].Value         = $res.Statut
    $r.Cells["AdresseResolue"].Value = $res.AdresseResolue
    $r.Cells["UPN"].Value            = $res.UserPrincipalName
    $r.Cells["DisplayName"].Value    = $res.DisplayName
    $r.Cells["ObjectId"].Value       = $res.ObjectId
    $r.Cells["Detail"].Value         = $res.Detail
    Set-RowStyle -row $r -statut $res.Statut
}

function Update-Summary {
    $total = $script:Results.Count
    $alt   = @($script:Results | Where-Object { $_.Statut -eq "TROUVE (ALT)" }).Count
    $ok    = @($script:Results | Where-Object { $_.Statut -eq "TROUVE" }).Count + $alt
    $nf    = @($script:Results | Where-Object { $_.Statut -eq "NON TROUVE" }).Count
    $err   = @($script:Results | Where-Object { $_.Statut -eq "ERREUR" }).Count
    $lblSummary.Text = "Total : $total   -   Trouves : $ok (dont $alt via domaine alt.)   -   Non trouves : $nf   -   Erreurs : $err"
    $btnResolve.Enabled = ($nf -gt 0)
    $btnExport.Enabled  = ($total -gt 0)
}

# ==================================================================
# Evenements
# ==================================================================
$btnInput.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtInput.Text = $dlg.FileName }
})

$btnOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Fichier CSV (*.csv)|*.csv"; $dlg.FileName = "Resultat_UPN.csv"
    if ($dlg.ShowDialog() -eq "OK") { $txtOutput.Text = $dlg.FileName }
})

$btnConnect.Add_Click({
    $btnConnect.Enabled = $false
    $lblAz.Text = "Connexion en cours..."; $lblAz.ForeColor = [System.Drawing.Color]::DarkSlateGray
    [System.Windows.Forms.Application]::DoEvents()
    try { Connect-AzAccount -ErrorAction Stop | Out-Null }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Echec de la connexion :`n$($_.Exception.Message)", "Connexion Azure", "OK", "Warning") | Out-Null
    }
    Update-AzStatus | Out-Null
    $btnConnect.Enabled = $true
})

# Double-clic : copie de l'UPN de la ligne dans le presse-papier
$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $upn = $grid.Rows[$e.RowIndex].Cells["UPN"].Value
    if ($upn) {
        [System.Windows.Forms.Clipboard]::SetText([string]$upn)
        $lblSummary.Text = "UPN copie : $upn"
    }
})

$btnRun.Add_Click({
    $az = Update-AzStatus
    if (-not $az.ModuleOk) {
        [System.Windows.Forms.MessageBox]::Show("Le module Az.Accounts n'est pas installe.`nInstall-Module Az -Scope CurrentUser", "Prerequis", "OK", "Error") | Out-Null; return
    }
    if (-not $az.Connected) {
        [System.Windows.Forms.MessageBox]::Show("Vous n'etes pas connecte a Azure. Cliquez sur 'Connexion Azure'.", "Prerequis", "OK", "Warning") | Out-Null; return
    }
    if (-not (Test-Path -LiteralPath $txtInput.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Fichier introuvable :`n$($txtInput.Text)", "Fichier source", "OK", "Error") | Out-Null; return
    }

    $emails = @(Get-Content -LiteralPath $txtInput.Text -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)

    if ($emails.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Le fichier ne contient aucune adresse exploitable.", "Fichier vide", "OK", "Warning") | Out-Null; return
    }

    $grid.Rows.Clear(); $script:Results.Clear()
    $btnRun.Enabled = $false; $btnExport.Enabled = $false; $btnResolve.Enabled = $false
    $progress.Minimum = 0; $progress.Maximum = $emails.Count; $progress.Value = 0

    $altDomains = if ($chkAutoAlt.Checked) { Get-AltDomains } else { @() }
    $index = 0; $total = $emails.Count

    foreach ($email in $emails) {
        $index++
        $statut = ""; $resolue = ""; $upn = ""; $display = ""; $oid = ""; $detail = ""
        try {
            $res = Resolve-EmailToUser -Email $email -AltDomains $altDomains
            if ($res.User) {
                $u = $res.User
                $resolue = $res.MatchedAddress; $upn = $u.UserPrincipalName; $display = $u.DisplayName; $oid = $u.Id
                if ($res.MatchedBy -like "Domaine alternatif*") { $statut = "TROUVE (ALT)" } else { $statut = "TROUVE" }
                $detail = "Resolu via : $($res.MatchedBy)"
            } else {
                $statut = "NON TROUVE"
                $detail = if ($altDomains.Count -gt 0) { "Introuvable (origine + domaines alternatifs)" } else { "Aucun compte correspondant dans Entra ID" }
            }
        } catch {
            $statut = "ERREUR"; $detail = $_.Exception.Message
        }

        $obj = [PSCustomObject]@{
            Ligne = $index; Mail = $email; Statut = $statut; AdresseResolue = $resolue
            UserPrincipalName = $upn; DisplayName = $display; ObjectId = $oid; Detail = $detail
        }
        $script:Results.Add($obj) | Out-Null
        $rowIdx = $grid.Rows.Add()
        Write-GridRow -rowIndex $rowIdx -res $obj

        $progress.Value = $index
        $lblSummary.Text = "Traitement $index / $total  -  $email"
        [System.Windows.Forms.Application]::DoEvents()
    }

    Update-Summary
    $btnRun.Enabled = $true
})

$btnResolve.Add_Click({
    $altDomains = Get-AltDomains
    if ($altDomains.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Renseignez au moins un domaine alternatif.", "Domaines alternatifs", "OK", "Warning") | Out-Null; return
    }
    $targets = @(0..($script:Results.Count - 1) | Where-Object { $script:Results[$_].Statut -eq "NON TROUVE" })
    if ($targets.Count -eq 0) { return }

    $btnResolve.Enabled = $false; $btnRun.Enabled = $false
    $progress.Minimum = 0; $progress.Maximum = $targets.Count; $progress.Value = 0
    $done = 0

    foreach ($i in $targets) {
        $done++
        $obj = $script:Results[$i]
        try {
            $res = Resolve-EmailToUser -Email $obj.Mail -AltDomains $altDomains
            if ($res.User -and $res.MatchedBy -like "Domaine alternatif*") {
                $u = $res.User
                $obj.Statut = "TROUVE (ALT)"; $obj.AdresseResolue = $res.MatchedAddress
                $obj.UserPrincipalName = $u.UserPrincipalName; $obj.DisplayName = $u.DisplayName; $obj.ObjectId = $u.Id
                $obj.Detail = "Resolu via : $($res.MatchedBy)"
                Write-GridRow -rowIndex $i -res $obj
            }
        } catch {
            $obj.Statut = "ERREUR"; $obj.Detail = $_.Exception.Message
            Write-GridRow -rowIndex $i -res $obj
        }
        $progress.Value = $done
        $lblSummary.Text = "Resolution $done / $($targets.Count)  -  $($obj.Mail)"
        [System.Windows.Forms.Application]::DoEvents()
    }

    Update-Summary
    $btnRun.Enabled = $true
})

$btnExport.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    $path = $txtOutput.Text
    try {
        $dir = Split-Path -Parent $path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $script:Results |
            Select-Object Ligne, Mail, Statut, AdresseResolue, UserPrincipalName, DisplayName, ObjectId, Detail |
            Export-Csv -Path $path -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Export termine :`n$path", "Export CSV", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Impossible d'ecrire le fichier :`n$($_.Exception.Message)", "Export CSV", "OK", "Error") | Out-Null
    }
})

# ==================================================================
Update-AzStatus | Out-Null
[void]$form.ShowDialog()
