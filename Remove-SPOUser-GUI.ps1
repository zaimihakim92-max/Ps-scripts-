<#
.SYNOPSIS
    Rechercher un utilisateur et le SUPPRIMER entierement d'un site SharePoint Online.
    Interface graphique. Module : PnP.PowerShell uniquement.

.DESCRIPTION
    Remove-PnPUser retire l'utilisateur de la User Information List du site, donc de TOUS
    ses groupes et de ses attributions de permissions sur ce site. C'est le moyen de purger
    une entree perimee (mismatch d'identite) : il suffit ensuite de re-ajouter l'utilisateur
    (People Picker, partage, ou Add-PnPGroupMember) pour que SharePoint re-resolve l'identite
    courante dans l'annuaire.

    Avant suppression, l'outil affiche et journalise les groupes du site auxquels l'utilisateur
    appartient, pour que tu saches quoi re-ajouter.

    Reutilise automatiquement une connexion PnP deja active dans la session.

.PREREQUIS
    - Module PnP.PowerShell
    - Une connexion active (Connect-PnPOnline dans la session) OU URL + Client ID via le bouton.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==================================================================
# Fonctions PnP
# ==================================================================
function Test-PnPConnected { try { return [bool](Get-PnPWeb -ErrorAction Stop) } catch { return $false } }

function Get-MatchingUsers {
    param([Parameter(Mandatory)][string]$Term)
    $t = $Term.Trim().ToLower()
    $all = @(Get-PnPUser -ErrorAction Stop)
    $matches = @($all | Where-Object {
        ($_.LoginName -and $_.LoginName.ToLower().Contains($t)) -or
        ($_.Email     -and $_.Email.ToLower().Contains($t))     -or
        ($_.Title     -and $_.Title.ToLower().Contains($t))
    })
    $dupEmails = @($matches | Where-Object { $_.Email } |
        Group-Object { $_.Email.ToLower() } | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($u in $matches) {
        $tags = @()
        if ($u.IsSiteAdmin) { $tags += "ADMIN" }
        if ([string]::IsNullOrWhiteSpace($u.Email)) { $tags += "SANS EMAIL" }
        if ($u.Email -and $dupEmails -contains $u.Email.ToLower()) { $tags += "DOUBLON" }
        if ($tags.Count -eq 0) { $tags += "OK" }
        $out.Add([PSCustomObject]@{
            Id = $u.Id; Title = $u.Title; LoginName = $u.LoginName; Email = $u.Email
            Type = $u.PrincipalType; Diag = ($tags -join ", ")
        }) | Out-Null
    }
    return $out
}

function Get-UserGroups {
    param([Parameter(Mandatory)][string]$LoginName)
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($g in @(Get-PnPGroup -ErrorAction SilentlyContinue)) {
        $members = @(Get-PnPGroupMember -Group $g -ErrorAction SilentlyContinue)
        if ($members | Where-Object { $_.LoginName -and ($_.LoginName -ieq $LoginName) }) { $found.Add($g.Title) | Out-Null }
    }
    return ,$found.ToArray()
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "SharePoint - Supprimer un utilisateur du site"
$form.Size = New-Object System.Drawing.Size(900, 660)
$form.MinimumSize = New-Object System.Drawing.Size(800, 560)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Text = "URL du site :"; $lblSite.Location = New-Object System.Drawing.Point(15, 15); $lblSite.AutoSize = $true
$form.Controls.Add($lblSite)

$txtSite = New-Object System.Windows.Forms.TextBox
$txtSite.Location = New-Object System.Drawing.Point(150, 12); $txtSite.Size = New-Object System.Drawing.Size(719, 23); $txtSite.Anchor = "Top,Left,Right"
$form.Controls.Add($txtSite)

$lblCid = New-Object System.Windows.Forms.Label
$lblCid.Text = "Client ID (app) :"; $lblCid.Location = New-Object System.Drawing.Point(15, 48); $lblCid.AutoSize = $true
$form.Controls.Add($lblCid)

$txtCid = New-Object System.Windows.Forms.TextBox
$txtCid.Location = New-Object System.Drawing.Point(150, 45); $txtCid.Size = New-Object System.Drawing.Size(585, 23); $txtCid.Anchor = "Top,Left,Right"
$form.Controls.Add($txtCid)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connexion"; $btnConnect.Location = New-Object System.Drawing.Point(744, 44); $btnConnect.Size = New-Object System.Drawing.Size(125, 25); $btnConnect.Anchor = "Top,Right"
$form.Controls.Add($btnConnect)

$lblConn = New-Object System.Windows.Forms.Label
$lblConn.Location = New-Object System.Drawing.Point(150, 71); $lblConn.Size = New-Object System.Drawing.Size(719, 18); $lblConn.Anchor = "Top,Left,Right"; $lblConn.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblConn)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Utilisateur :"; $lblSearch.Location = New-Object System.Drawing.Point(15, 101); $lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(150, 98); $txtSearch.Size = New-Object System.Drawing.Size(585, 23); $txtSearch.Anchor = "Top,Left,Right"
$form.Controls.Add($txtSearch)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Rechercher"; $btnSearch.Location = New-Object System.Drawing.Point(744, 97); $btnSearch.Size = New-Object System.Drawing.Size(125, 25); $btnSearch.Anchor = "Top,Right"; $btnSearch.Enabled = $false
$form.Controls.Add($btnSearch)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 130); $grid.Size = New-Object System.Drawing.Size(854, 200); $grid.Anchor = "Top,Left,Right"
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.ReadOnly = $true
$grid.RowHeadersVisible = $false; $grid.SelectionMode = "FullRowSelect"; $grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"; $grid.ColumnHeadersHeightSizeMode = "AutoSize"
$null = $grid.Columns.Add("Id", "Id")
$null = $grid.Columns.Add("Title", "Nom affiche")
$null = $grid.Columns.Add("LoginName", "LoginName (claim)")
$null = $grid.Columns.Add("Email", "Email")
$null = $grid.Columns.Add("Type", "Type")
$null = $grid.Columns.Add("Diag", "Diagnostic")
$grid.Columns["Id"].FillWeight = 6; $grid.Columns["Title"].FillWeight = 18; $grid.Columns["LoginName"].FillWeight = 34
$grid.Columns["Email"].FillWeight = 20; $grid.Columns["Type"].FillWeight = 9; $grid.Columns["Diag"].FillWeight = 13
$form.Controls.Add($grid)

$btnShowGroups = New-Object System.Windows.Forms.Button
$btnShowGroups.Text = "Voir les groupes"; $btnShowGroups.Location = New-Object System.Drawing.Point(15, 338); $btnShowGroups.Size = New-Object System.Drawing.Size(200, 28); $btnShowGroups.Anchor = "Top,Left"
$form.Controls.Add($btnShowGroups)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Supprimer du site"; $btnDelete.Location = New-Object System.Drawing.Point(669, 338); $btnDelete.Size = New-Object System.Drawing.Size(200, 28); $btnDelete.Anchor = "Top,Right"
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60); $btnDelete.ForeColor = [System.Drawing.Color]::White; $btnDelete.FlatStyle = "Flat"
$form.Controls.Add($btnDelete)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Journal :"; $lblLog.Location = New-Object System.Drawing.Point(15, 374); $lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 394); $txtLog.Size = New-Object System.Drawing.Size(854, 220); $txtLog.Anchor = "Top,Bottom,Left,Right"
$txtLog.Multiline = $true; $txtLog.ScrollBars = "Vertical"; $txtLog.ReadOnly = $true; $txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

# ==================================================================
# Helpers UI
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-ConnState { param([bool]$c, [string]$msg) $lblConn.ForeColor = if ($c) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::Gray }; $lblConn.Text = $msg; $btnSearch.Enabled = $c }

function Invoke-Search {
    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) { return }
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a un site.", "Recherche", "OK", "Warning") | Out-Null; return }
    $grid.Rows.Clear()
    try {
        $res = Get-MatchingUsers -Term $txtSearch.Text
        foreach ($e in $res) {
            $i = $grid.Rows.Add($e.Id, $e.Title, $e.LoginName, $e.Email, $e.Type, $e.Diag)
            if ($e.Diag -like "*DOUBLON*") { $grid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,224,150) }
            elseif ($e.Diag -like "*SANS EMAIL*") { $grid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,235,200) }
        }
        Write-Log "Recherche '$($txtSearch.Text)' : $($res.Count) entree(s)."
    } catch { Write-Log "ERREUR recherche : $($_.Exception.Message)" }
}

# ==================================================================
# Evenements
# ==================================================================
$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text)) { [System.Windows.Forms.MessageBox]::Show("Renseignez l'URL du site.", "Connexion", "OK", "Warning") | Out-Null; return }
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) { [System.Windows.Forms.MessageBox]::Show("Module PnP.PowerShell introuvable.", "Prerequis", "OK", "Error") | Out-Null; return }
    $btnConnect.Enabled = $false; Set-ConnState $false "Connexion en cours..."
    try {
        if ([string]::IsNullOrWhiteSpace($txtCid.Text)) { Connect-PnPOnline -Url $txtSite.Text.Trim() -Interactive -ErrorAction Stop }
        else { Connect-PnPOnline -Url $txtSite.Text.Trim() -Interactive -ClientId $txtCid.Text.Trim() -ErrorAction Stop }
        if (Test-PnPConnected) { $w = Get-PnPWeb; Set-ConnState $true "Connecte : $($w.Title)"; Write-Log "Connecte : $($txtSite.Text.Trim())" }
    } catch { Set-ConnState $false "Echec de connexion."; [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)", "Connexion", "OK", "Error") | Out-Null }
    $btnConnect.Enabled = $true
})

$btnSearch.Add_Click({ Invoke-Search })

$btnShowGroups.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $login = [string]$grid.SelectedRows[0].Cells["LoginName"].Value
    if (-not $login) { return }
    Write-Log "Groupes de : $login"
    try { $g = Get-UserGroups -LoginName $login; if ($g.Count -eq 0) { Write-Log "  (aucun groupe de site)" } else { $g | ForEach-Object { Write-Log "  - $_" } } }
    catch { Write-Log "  ERREUR : $($_.Exception.Message)" }
})

$btnDelete.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte.", "Suppression", "OK", "Warning") | Out-Null; return }
    if ($grid.SelectedRows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Selectionnez une entree.", "Suppression", "OK", "Warning") | Out-Null; return }
    $row   = $grid.SelectedRows[0]
    $id    = [int]$row.Cells["Id"].Value
    $login = [string]$row.Cells["LoginName"].Value
    $email = [string]$row.Cells["Email"].Value

    # Journalise les groupes avant suppression (pour savoir quoi re-ajouter)
    $groups = @()
    try { $groups = Get-UserGroups -LoginName $login } catch { }
    Write-Log "Avant suppression - groupes de '$login' ($($groups.Count)) :"
    if ($groups.Count -eq 0) { Write-Log "  (aucun)" } else { $groups | ForEach-Object { Write-Log "  - $_" } }

    $recap = "Supprimer completement cet utilisateur du site ?`n`n" +
             "Id     : $id`n" +
             "Login  : $login`n" +
             "Email  : $email`n" +
             "Groupes actuels : $($groups.Count)`n`n" +
             "Il sera retire de tous ses groupes et permissions sur ce site.`n" +
             "Tu pourras le re-ajouter ensuite (l'identite courante sera re-resolue)."
    if ([System.Windows.Forms.MessageBox]::Show($recap, "Confirmation", "YesNo", "Warning") -ne "Yes") { Write-Log "Suppression annulee."; return }

    try {
        Remove-PnPUser -Identity $id -Force -ErrorAction Stop
        Write-Log "SUPPRIME du site : Id $id - $login"
        [System.Windows.Forms.MessageBox]::Show("Utilisateur supprime du site.`nTu peux maintenant le re-ajouter.", "Termine", "OK", "Information") | Out-Null
        Invoke-Search
    } catch {
        Write-Log "ERREUR suppression : $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)", "Suppression", "OK", "Error") | Out-Null
    }
})

# ==================================================================
# Adopter une connexion PnP active dans la session
# ==================================================================
$adopted = $false
if (Get-Module -ListAvailable -Name PnP.PowerShell) {
    try {
        if (Test-PnPConnected) {
            $url = $null; try { $url = (Get-PnPConnection -ErrorAction Stop).Url } catch { }
            $w = Get-PnPWeb -ErrorAction SilentlyContinue
            if (-not $url -and $w) { $url = $w.Url }
            if ($url) { $txtSite.Text = $url }
            Set-ConnState $true ("Connexion active reutilisee : " + $(if ($w) { $w.Title } else { $url }))
            Write-Log "Connexion PnP active detectee : $url"
            $adopted = $true
        }
    } catch { }
}
if (-not $adopted) { Set-ConnState $false "Aucune connexion active. Connectez-vous avant, ou via le bouton Connexion." }

[void]$form.ShowDialog()
