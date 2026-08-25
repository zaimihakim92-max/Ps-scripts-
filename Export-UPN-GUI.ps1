<#
.SYNOPSIS
    Export des UPN (Entra ID) a partir d'une liste d'adresses email - Interface graphique.

.DESCRIPTION
    Pour chaque adresse email, recherche le compte correspondant dans Entra ID selon
    3 methodes successives (attribut mail, UserPrincipalName, adresses secondaires),
    affiche le resultat dans une grille avec code couleur et identifie clairement les
    comptes NON TROUVES. Export CSV optionnel.

    Aucune erreur bloquante : tout est capture et reporte dans la grille et la barre
    de statut. Le module et la connexion Azure sont verifies avant traitement.

.PREREQUIS
    - Module Az (Az.Accounts, Az.Resources)
    - La connexion se fait via le bouton "Connexion Azure" (ou Connect-AzAccount prealable)
#>

# ------------------------------------------------------------------
# Assemblies WinForms
# ------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# Etat global
# ------------------------------------------------------------------
$script:Results = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Fonctions
# ==================================================================

function Test-AzReady {
    # Verifie module + connexion, sans jamais lever d'erreur.
    $r = [PSCustomObject]@{ ModuleOk = $false; Connected = $false; Account = $null }

    if (Get-Module -ListAvailable -Name Az.Accounts) { $r.ModuleOk = $true }

    try {
        $ctx = Get-AzContext -ErrorAction Stop
        if ($ctx) { $r.Connected = $true; $r.Account = $ctx.Account.Id }
    } catch { }

    return $r
}

function Find-EntraUser {
    # Recherche multi-attributs. Retourne l'objet utilisateur ou $null.
    param([Parameter(Mandatory)][string]$Email)

    $safe = $Email.Replace("'", "''")

    try {
        # 1) Attribut mail (correspondance exacte)
        $u = Get-AzADUser -Mail $Email -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }

        # 2) UserPrincipalName
        $u = Get-AzADUser -UserPrincipalName $Email -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }

        # 3) Adresses secondaires (otherMails / alias)
        $u = Get-AzADUser -Filter "otherMails/any(c:c eq '$safe')" -ErrorAction SilentlyContinue
        if ($u) { return ($u | Select-Object -First 1) }
    } catch { }

    return $null
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "Export UPN Entra ID"
$form.Size          = New-Object System.Drawing.Size(900, 640)
$form.MinimumSize   = New-Object System.Drawing.Size(780, 540)
$form.StartPosition = "CenterScreen"
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Fichier source ---
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Fichier .txt :"
$lblInput.Location = New-Object System.Drawing.Point(15, 15)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(115, 12)
$txtInput.Size = New-Object System.Drawing.Size(650, 23)
$txtInput.Anchor = "Top,Left,Right"
$txtInput.Text = "C:\Temp\emails.txt"
$form.Controls.Add($txtInput)

$btnInput = New-Object System.Windows.Forms.Button
$btnInput.Text = "Parcourir..."
$btnInput.Location = New-Object System.Drawing.Point(775, 11)
$btnInput.Size = New-Object System.Drawing.Size(95, 25)
$btnInput.Anchor = "Top,Right"
$form.Controls.Add($btnInput)

# --- Fichier export ---
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Export CSV :"
$lblOutput.Location = New-Object System.Drawing.Point(15, 48)
$lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(115, 45)
$txtOutput.Size = New-Object System.Drawing.Size(650, 23)
$txtOutput.Anchor = "Top,Left,Right"
$txtOutput.Text = "C:\Temp\Resultat_UPN.csv"
$form.Controls.Add($txtOutput)

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = "Parcourir..."
$btnOutput.Location = New-Object System.Drawing.Point(775, 44)
$btnOutput.Size = New-Object System.Drawing.Size(95, 25)
$btnOutput.Anchor = "Top,Right"
$form.Controls.Add($btnOutput)

# --- Statut Azure + connexion ---
$lblAz = New-Object System.Windows.Forms.Label
$lblAz.Location = New-Object System.Drawing.Point(15, 82)
$lblAz.Size = New-Object System.Drawing.Size(650, 20)
$lblAz.Anchor = "Top,Left,Right"
$lblAz.Text = "Etat Azure : verification..."
$form.Controls.Add($lblAz)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connexion Azure"
$btnConnect.Location = New-Object System.Drawing.Point(745, 78)
$btnConnect.Size = New-Object System.Drawing.Size(125, 25)
$btnConnect.Anchor = "Top,Right"
$form.Controls.Add($btnConnect)

# --- Lancer + progression ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Lancer le traitement"
$btnRun.Location = New-Object System.Drawing.Point(15, 115)
$btnRun.Size = New-Object System.Drawing.Size(250, 30)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = "Flat"
$form.Controls.Add($btnRun)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(280, 119)
$progress.Size = New-Object System.Drawing.Size(590, 22)
$progress.Anchor = "Top,Left,Right"
$form.Controls.Add($progress)

# --- Grille resultats ---
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 160)
$grid.Size = New-Object System.Drawing.Size(855, 355)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$grid.ColumnHeadersHeightSizeMode = "AutoSize"

$null = $grid.Columns.Add("Ligne", "#")
$null = $grid.Columns.Add("Mail", "Adresse email")
$null = $grid.Columns.Add("Statut", "Statut")
$null = $grid.Columns.Add("UPN", "UserPrincipalName")
$null = $grid.Columns.Add("DisplayName", "Nom affiche")
$null = $grid.Columns.Add("ObjectId", "ObjectId")
$null = $grid.Columns.Add("Detail", "Detail")

$grid.Columns["Ligne"].FillWeight = 6
$grid.Columns["Mail"].FillWeight = 22
$grid.Columns["Statut"].FillWeight = 12
$grid.Columns["UPN"].FillWeight = 22
$grid.Columns["DisplayName"].FillWeight = 16
$grid.Columns["ObjectId"].FillWeight = 0
$grid.Columns["ObjectId"].Visible = $false   # visible dans le CSV, masque a l'ecran
$grid.Columns["Detail"].FillWeight = 22
$form.Controls.Add($grid)

# --- Recapitulatif + export ---
$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Location = New-Object System.Drawing.Point(15, 528)
$lblSummary.Size = New-Object System.Drawing.Size(600, 30)
$lblSummary.Anchor = "Bottom,Left"
$lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSummary.Text = "En attente..."
$form.Controls.Add($lblSummary)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exporter en CSV"
$btnExport.Location = New-Object System.Drawing.Point(720, 525)
$btnExport.Size = New-Object System.Drawing.Size(150, 30)
$btnExport.Anchor = "Bottom,Right"
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

# ==================================================================
# Rafraichissement de l'etat Azure
# ==================================================================
function Update-AzStatus {
    $az = Test-AzReady
    if (-not $az.ModuleOk) {
        $lblAz.Text = "Etat Azure : module Az.Accounts introuvable."
        $lblAz.ForeColor = [System.Drawing.Color]::Firebrick
    }
    elseif ($az.Connected) {
        $lblAz.Text = "Etat Azure : connecte ($($az.Account))"
        $lblAz.ForeColor = [System.Drawing.Color]::ForestGreen
    }
    else {
        $lblAz.Text = "Etat Azure : non connecte - cliquez sur 'Connexion Azure'."
        $lblAz.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    return $az
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
    $dlg.Filter = "Fichier CSV (*.csv)|*.csv"
    $dlg.FileName = "Resultat_UPN.csv"
    if ($dlg.ShowDialog() -eq "OK") { $txtOutput.Text = $dlg.FileName }
})

$btnConnect.Add_Click({
    $btnConnect.Enabled = $false
    $lblAz.Text = "Connexion en cours..."
    $lblAz.ForeColor = [System.Drawing.Color]::DarkSlateGray
    [System.Windows.Forms.Application]::DoEvents()
    try {
        Connect-AzAccount -ErrorAction Stop | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Echec de la connexion :`n$($_.Exception.Message)",
            "Connexion Azure", "OK", "Warning") | Out-Null
    }
    Update-AzStatus | Out-Null
    $btnConnect.Enabled = $true
})

$btnRun.Add_Click({

    # --- Verifications prealables ---
    $az = Update-AzStatus
    if (-not $az.ModuleOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "Le module Az.Accounts n'est pas installe.`nInstall-Module Az -Scope CurrentUser",
            "Prerequis", "OK", "Error") | Out-Null
        return
    }
    if (-not $az.Connected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Vous n'etes pas connecte a Azure. Cliquez sur 'Connexion Azure'.",
            "Prerequis", "OK", "Warning") | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $txtInput.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Fichier introuvable :`n$($txtInput.Text)",
            "Fichier source", "OK", "Error") | Out-Null
        return
    }

    # --- Lecture + nettoyage des emails ---
    $emails = @(Get-Content -LiteralPath $txtInput.Text -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)

    if ($emails.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Le fichier ne contient aucune adresse exploitable.",
            "Fichier vide", "OK", "Warning") | Out-Null
        return
    }

    # --- Preparation UI ---
    $grid.Rows.Clear()
    $script:Results.Clear()
    $btnRun.Enabled = $false
    $btnExport.Enabled = $false
    $progress.Minimum = 0
    $progress.Maximum = $emails.Count
    $progress.Value = 0

    $found = 0; $notFound = 0; $errors = 0
    $index = 0
    $total = $emails.Count

    foreach ($email in $emails) {
        $index++
        $statut = ""; $upn = ""; $display = ""; $oid = ""; $detail = ""

        try {
            $user = Find-EntraUser -Email $email
            if ($user) {
                $found++
                $statut  = "TROUVE"
                $upn     = $user.UserPrincipalName
                $display = $user.DisplayName
                $oid     = $user.Id
                $detail  = "Compte trouve dans Entra ID"
            }
            else {
                $notFound++
                $statut = "NON TROUVE"
                $detail = "Aucun compte correspondant dans Entra ID"
            }
        }
        catch {
            $errors++
            $statut = "ERREUR"
            $detail = $_.Exception.Message
        }

        # Ajout ligne grille
        $rowIdx = $grid.Rows.Add($index, $email, $statut, $upn, $display, $oid, $detail)
        $row = $grid.Rows[$rowIdx]
        switch ($statut) {
            "TROUVE"     { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(213, 245, 213) }
            "NON TROUVE" { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 224, 150)
                           $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(120, 60, 0) }
            "ERREUR"     { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 190, 190)
                           $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed }
        }

        # Stockage pour export
        $script:Results.Add([PSCustomObject]@{
            Ligne             = $index
            Mail              = $email
            Statut            = $statut
            UserPrincipalName = $upn
            DisplayName       = $display
            ObjectId          = $oid
            Detail            = $detail
        }) | Out-Null

        # Progression
        $progress.Value = $index
        $lblSummary.Text = "Traitement $index / $total  -  $email"
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblSummary.Text = "Termine  |  Total : $total   -   Trouves : $found   -   Non trouves : $notFound   -   Erreurs : $errors"
    $btnRun.Enabled = $true
    $btnExport.Enabled = $true
})

$btnExport.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    $path = $txtOutput.Text
    try {
        $dir = Split-Path -Parent $path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $script:Results |
            Export-Csv -Path $path -Delimiter ";" -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show(
            "Export termine :`n$path",
            "Export CSV", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Impossible d'ecrire le fichier :`n$($_.Exception.Message)",
            "Export CSV", "OK", "Error") | Out-Null
    }
})

# ==================================================================
# Affichage
# ==================================================================
Update-AzStatus | Out-Null
[void]$form.ShowDialog()
