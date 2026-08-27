#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM + AD - Dernier utilisateur connecte (interface graphique)

.DESCRIPTION
    Version GUI de sccm_ad_lastloggedin.
      - Selection de la liste des machines (.txt) via bouton
      - Resolution SCCM (Get-CMResource) + AD (Get-ADUser)
      - Affichage avance dans un DataGridView :
            * tri par clic sur l'en-tete
            * redimensionnement des colonnes
            * reordonnancement des colonnes par glisser-deposer
            * filtre texte en temps reel
            * coloration des lignes en erreur
      - Export CSV avec choix des colonnes ET de leur ordre

.NOTES
    A lancer en mode STA (par defaut avec powershell.exe -File .\sccm_ad_lastloggedin_gui.ps1).
    Sous PowerShell 7 : powershell.exe -STA -File .\...  (Windows PowerShell recommande).
    Doit tourner dans une session ou Get-CMResource est disponible
    (console Configuration Manager / lecteur de site charge).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==================================================================
#  Verification des dependances
# ==================================================================
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        "Le module ActiveDirectory est introuvable.`n`n$($_.Exception.Message)",
        "Dependance manquante", 'OK', 'Error')
    return
}

if (-not (Get-Command Get-CMResource -ErrorAction SilentlyContinue)) {
    [void][System.Windows.Forms.MessageBox]::Show(
        "La commande Get-CMResource est introuvable.`n`n" +
        "Lancez ce script depuis la session PowerShell de la console " +
        "Configuration Manager (ou importez le module ConfigurationManager " +
        "et placez-vous sur le lecteur du site).",
        "SCCM non disponible", 'OK', 'Error')
    return
}

# ==================================================================
#  Configuration
# ==================================================================
# Ordre par defaut des colonnes (= ordre initial d'affichage et d'export)
$script:ColumnOrder = @(
    'machine','deviceId','lastLogonUser','displayName','mail',
    'userPrincipalName','objectType','userType','id',
    'isUser','isGroup','isGuest','status'
)

# Colonnes prises en compte par le filtre texte
$script:TextFilterColumns = @(
    'machine','deviceId','lastLogonUser','displayName',
    'mail','userPrincipalName','status'
)

# Libelles d'en-tete plus lisibles
$script:Headers = @{
    machine           = 'Machine'
    deviceId          = 'ResourceID'
    lastLogonUser     = 'Dernier logon (SCCM)'
    displayName       = 'Nom affiche'
    mail              = 'Email'
    userPrincipalName = 'UPN'
    objectType        = 'objectType'
    userType          = 'userType'
    id                = 'ObjectGUID (AD)'
    isUser            = 'isUser'
    isGroup           = 'isGroup'
    isGuest           = 'isGuest'
    status            = 'Statut'
}

$script:Computers = @()

# ==================================================================
#  Fonctions
# ==================================================================
function New-ResultTable {
    $dt = New-Object System.Data.DataTable
    foreach ($c in $script:ColumnOrder) {
        $type = if ($c -in 'isUser','isGroup','isGuest') { [bool] } else { [string] }
        [void]$dt.Columns.Add($c, $type)
    }
    return ,$dt
}

function Update-RowColors {
    param([System.Windows.Forms.DataGridView]$Grid)
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $st = [string]$row.Cells['status'].Value
        if     ($st -like 'Machine introuvable*') {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,224,224)
        }
        elseif ($st -like 'Utilisateur AD introuvable*' -or $st -like 'Aucun dernier*') {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,245,204)
        }
        else {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
        }
    }
}

function Invoke-Processing {
    param(
        [string[]]$Computers,
        [System.Data.DataTable]$Table,
        [System.Windows.Forms.ToolStripProgressBar]$Progress,
        [System.Windows.Forms.ToolStripStatusLabel]$Status
    )

    $Table.Rows.Clear()

    # --- Construction de l'index SCCM ---
    $Status.Text = "Chargement des ressources SCCM..."
    [System.Windows.Forms.Application]::DoEvents()

    $CmIndex = @{}
    Get-CMResource -ResourceType System -Fast |
        Where-Object Name |
        ForEach-Object { $CmIndex[$_.Name.ToLower()] = $_ }

    $Progress.Minimum = 0
    $Progress.Maximum = [Math]::Max(1, $Computers.Count)
    $Progress.Value   = 0

    $i = 0
    foreach ($Computer in $Computers) {
        $i++
        $Progress.Value = $i
        $Status.Text = "Traitement $i / $($Computers.Count) : $Computer"
        [System.Windows.Forms.Application]::DoEvents()

        $Resource = $CmIndex[$Computer.ToLower()]

        # --- Machine absente de SCCM ---
        if (-not $Resource) {
            $row = $Table.NewRow()
            $row['machine']           = ''
            $row['deviceId']          = $Computer
            $row['lastLogonUser']     = ''
            $row['displayName']       = ''
            $row['mail']              = ''
            $row['userPrincipalName'] = ''
            $row['objectType']        = 'user'
            $row['userType']          = ''
            $row['id']                = ''
            $row['isUser']            = $true
            $row['isGroup']           = $false
            $row['isGuest']           = $false
            $row['status']            = 'Machine introuvable SCCM'
            [void]$Table.Rows.Add($row)
            continue
        }

        # --- Resolution du dernier utilisateur ---
        $LastLogon   = $Resource.LastLogonUserName
        $DisplayName = ''
        $Mail        = ''
        $UPN         = ''
        $Guid        = ''
        $UserType    = ''
        $State       = 'OK'

        if ($LastLogon) {
            $SamAccount = $LastLogon.Split('\')[-1]
            try {
                $ADUser = Get-ADUser -Identity $SamAccount `
                    -Properties DisplayName,Mail,UserPrincipalName -ErrorAction Stop

                $DisplayName = $ADUser.DisplayName
                $Mail        = $ADUser.Mail
                $UPN         = $ADUser.UserPrincipalName
                $Guid        = $ADUser.ObjectGUID
                $UserType    = 'Member'

                # Si mail vide, on force avec l'UPN
                if ([string]::IsNullOrWhiteSpace($Mail)) { $Mail = $UPN }
            }
            catch {
                $State = "Utilisateur AD introuvable : $SamAccount"
            }
        }
        else {
            $State = 'Aucun dernier utilisateur SCCM'
        }

        $row = $Table.NewRow()
        $row['machine']           = [string]$Resource.Name
        $row['deviceId']          = [string]$Resource.ResourceID
        $row['lastLogonUser']     = [string]$LastLogon
        $row['displayName']       = [string]$DisplayName
        $row['mail']              = [string]$Mail
        $row['userPrincipalName'] = [string]$UPN
        $row['objectType']        = 'user'
        $row['userType']          = [string]$UserType
        $row['id']                = [string]$Guid
        $row['isUser']            = $true
        $row['isGroup']           = $false
        $row['isGuest']           = $false
        $row['status']            = $State
        [void]$Table.Rows.Add($row)
    }

    $Status.Text  = "Termine : $($Table.Rows.Count) ligne(s)."
    $Progress.Value = $Progress.Maximum
}

function Show-ExportDialog {
    param([System.Windows.Forms.DataGridView]$Grid)

    # Colonnes dans l'ordre d'affichage courant du grid
    $cols = $Grid.Columns | Sort-Object DisplayIndex

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Colonnes et ordre d'export CSV"
    $dlg.Size            = New-Object System.Drawing.Size(370,430)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = "Cochez les colonnes a exporter, ordonnez-les avec Monter/Descendre :"
    $lbl.Location = New-Object System.Drawing.Point(12,8)
    $lbl.Size     = New-Object System.Drawing.Size(340,32)
    $dlg.Controls.Add($lbl)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location    = New-Object System.Drawing.Point(12,44)
    $clb.Size        = New-Object System.Drawing.Size(230,338)
    $clb.CheckOnClick = $true
    foreach ($c in $cols) { [void]$clb.Items.Add($c.Name, $c.Visible) }
    $dlg.Controls.Add($clb)

    function New-DlgButton([string]$Text,[int]$Top) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text     = $Text
        $b.Location = New-Object System.Drawing.Point(252,$Top)
        $b.Size     = New-Object System.Drawing.Size(100,28)
        $dlg.Controls.Add($b)
        return $b
    }

    $btnUp   = New-DlgButton "Monter"        44
    $btnDown = New-DlgButton "Descendre"     78
    $btnAll  = New-DlgButton "Tout cocher"   122
    $btnNone = New-DlgButton "Tout decocher" 156

    $btnOk = New-DlgButton "Exporter..." 320
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-DlgButton "Annuler" 354
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    $moveItem = {
        param([int]$delta)
        $i = $clb.SelectedIndex
        if ($i -lt 0) { return }
        $j = $i + $delta
        if ($j -lt 0 -or $j -ge $clb.Items.Count) { return }
        $item    = $clb.Items[$i]
        $checked = $clb.GetItemChecked($i)
        $clb.Items.RemoveAt($i)
        $clb.Items.Insert($j, $item)
        $clb.SetItemChecked($j, $checked)
        $clb.SelectedIndex = $j
    }
    $btnUp.Add_Click(  { & $moveItem (-1) })
    $btnDown.Add_Click({ & $moveItem (1)  })
    $btnAll.Add_Click( { for ($k=0; $k -lt $clb.Items.Count; $k++) { $clb.SetItemChecked($k,$true) } })
    $btnNone.Add_Click({ for ($k=0; $k -lt $clb.Items.Count; $k++) { $clb.SetItemChecked($k,$false) } })

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $ordered = @()
    for ($k=0; $k -lt $clb.Items.Count; $k++) {
        if ($clb.GetItemChecked($k)) { $ordered += [string]$clb.Items[$k] }
    }
    return ,$ordered
}

# ==================================================================
#  Interface
# ==================================================================
$Table   = New-ResultTable
$Binding = New-Object System.Windows.Forms.BindingSource
$Binding.DataSource = $Table

$Form = New-Object System.Windows.Forms.Form
$Form.Text          = "SCCM + AD - Dernier utilisateur connecte"
$Form.Size          = New-Object System.Drawing.Size(1150,720)
$Form.StartPosition = 'CenterScreen'
$Form.MinimumSize   = New-Object System.Drawing.Size(820,520)

# --- Barre haute ---
$Top = New-Object System.Windows.Forms.Panel
$Top.Height = 82
$Top.Dock   = 'Top'

$btnFile = New-Object System.Windows.Forms.Button
$btnFile.Text     = "Liste TXT..."
$btnFile.Location = New-Object System.Drawing.Point(10,10)
$btnFile.Size     = New-Object System.Drawing.Size(110,28)

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text     = "Aucun fichier selectionne"
$lblFile.Location = New-Object System.Drawing.Point(130,15)
$lblFile.AutoSize = $true

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text    = "Lancer"
$btnRun.Size    = New-Object System.Drawing.Size(110,28)
$btnRun.Anchor  = 'Top,Right'
$btnRun.Enabled = $false

$lblFilterTitle = New-Object System.Windows.Forms.Label
$lblFilterTitle.Text     = "Filtre :"
$lblFilterTitle.Location = New-Object System.Drawing.Point(10,50)
$lblFilterTitle.AutoSize = $true

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(60,47)
$txtFilter.Size     = New-Object System.Drawing.Size(320,24)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text    = "Export CSV..."
$btnExport.Size    = New-Object System.Drawing.Size(120,28)
$btnExport.Anchor  = 'Top,Right'
$btnExport.Enabled = $false

$Top.Controls.AddRange(@($btnFile,$lblFile,$btnRun,$lblFilterTitle,$txtFilter,$btnExport))

# --- Grille ---
$Grid = New-Object System.Windows.Forms.DataGridView
$Grid.Dock                      = 'Fill'
$Grid.DataSource                = $Binding
$Grid.AllowUserToAddRows        = $false
$Grid.AllowUserToDeleteRows     = $false
$Grid.ReadOnly                  = $true
$Grid.AllowUserToOrderColumns   = $true    # reordonnancement par glisser-deposer
$Grid.AllowUserToResizeColumns  = $true    # redimensionnement des colonnes
$Grid.AllowUserToResizeRows     = $false
$Grid.AutoSizeColumnsMode       = 'AllCells'
$Grid.SelectionMode             = 'FullRowSelect'
$Grid.MultiSelect               = $true
$Grid.RowHeadersVisible         = $false
$Grid.ColumnHeadersHeightSizeMode = 'AutoSize'

$Grid.Add_DataBindingComplete({
    foreach ($c in $Grid.Columns) {
        if ($script:Headers.ContainsKey($c.Name)) { $c.HeaderText = $script:Headers[$c.Name] }
    }
})

# --- Barre d'etat ---
$Status    = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text      = "Pret."
$lblStatus.Spring    = $true
$lblStatus.TextAlign = 'MiddleLeft'
$Progress  = New-Object System.Windows.Forms.ToolStripProgressBar
$Progress.Size = New-Object System.Drawing.Size(240,16)
[void]$Status.Items.Add($lblStatus)
[void]$Status.Items.Add($Progress)

# Ordre d'ajout : Fill d'abord, puis Top / Bottom ramenes au premier plan
$Form.Controls.Add($Grid)
$Form.Controls.Add($Top);    $Top.BringToFront()
$Form.Controls.Add($Status); $Status.BringToFront()

# ==================================================================
#  Evenements
# ==================================================================
$Form.Add_Shown({
    $btnRun.Left    = $Top.ClientSize.Width - $btnRun.Width - 10
    $btnRun.Top     = 10
    $btnExport.Left = $Top.ClientSize.Width - $btnExport.Width - 10
    $btnExport.Top  = 46
})

$btnFile.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichier texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    $ofd.Title  = "Selectionner la liste des machines"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Computers = Get-Content -LiteralPath $ofd.FileName |
            ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $name = [System.IO.Path]::GetFileName($ofd.FileName)
        $lblFile.Text  = "$name  -  $($script:Computers.Count) machine(s)"
        $btnRun.Enabled = ($script:Computers.Count -gt 0)
    }
})

$btnRun.Add_Click({
    if (-not $script:Computers -or $script:Computers.Count -eq 0) { return }
    $btnFile.Enabled = $false; $btnRun.Enabled = $false; $btnExport.Enabled = $false
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Invoke-Processing -Computers $script:Computers -Table $Table -Progress $Progress -Status $lblStatus
        Update-RowColors -Grid $Grid
        $btnExport.Enabled = ($Table.Rows.Count -gt 0)
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Erreur pendant le traitement :`n$($_.Exception.Message)",
            "Erreur", 'OK', 'Error')
        $lblStatus.Text = "Erreur."
    }
    finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnFile.Enabled = $true; $btnRun.Enabled = $true
    }
})

$txtFilter.Add_TextChanged({
    $q = $txtFilter.Text.Replace("'","''")
    if ([string]::IsNullOrWhiteSpace($q)) {
        $Binding.Filter = $null
    }
    else {
        $parts = foreach ($c in $script:TextFilterColumns) { "$c LIKE '%$q%'" }
        $Binding.Filter = ($parts -join ' OR ')
    }
    Update-RowColors -Grid $Grid
})

$btnExport.Add_Click({
    if ($Table.Rows.Count -eq 0) { return }

    $ordered = Show-ExportDialog -Grid $Grid
    if (-not $ordered) { return }
    if ($ordered.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Aucune colonne selectionnee.","Export",'OK','Warning')
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter           = "CSV (*.csv)|*.csv"
    $sfd.Title            = "Enregistrer l'export CSV"
    $sfd.FileName         = "SCCM_LastUser_Email.csv"
    $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    # Lignes actuellement affichees (respecte le filtre et le tri), colonnes dans l'ordre choisi
    $rows = foreach ($r in $Grid.Rows) {
        if ($r.IsNewRow) { continue }
        $o = [ordered]@{}
        foreach ($col in $ordered) { $o[$col] = $r.Cells[$col].Value }
        [PSCustomObject]$o
    }

    try {
        $rows | Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Export : $($sfd.FileName)"
        $ask = [System.Windows.Forms.MessageBox]::Show(
            "Export termine ($($rows.Count) ligne(s)).`n`nOuvrir le dossier ?",
            "Export CSV", 'YesNo', 'Information')
        if ($ask -eq 'Yes') { Start-Process explorer.exe "/select,`"$($sfd.FileName)`"" }
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Erreur d'export :`n$($_.Exception.Message)","Erreur",'OK','Error')
    }
})

# ==================================================================
#  Lancement
# ==================================================================
[void]$Form.ShowDialog()
$Form.Dispose()
