Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module ActiveDirectory

[System.Windows.Forms.Application]::EnableVisualStyles()

#================================================
# THEME / COULEURS
#================================================

$script:clrHeader     = [System.Drawing.Color]::FromArgb(32,41,64)
$script:clrAccent     = [System.Drawing.Color]::FromArgb(0,122,204)
$script:clrBg         = [System.Drawing.Color]::FromArgb(245,246,248)
$script:clrGridHead   = [System.Drawing.Color]::FromArgb(45,55,80)
$script:clrAlt        = [System.Drawing.Color]::FromArgb(238,242,247)
$script:clrGreen      = [System.Drawing.Color]::FromArgb(46,160,67)
$script:clrRed        = [System.Drawing.Color]::FromArgb(200,60,60)
$script:clrAmber      = [System.Drawing.Color]::FromArgb(232,151,0)
$script:clrEditHint   = [System.Drawing.Color]::FromArgb(255,249,196)
$script:clrStatusOk   = [System.Drawing.Color]::FromArgb(120,220,140)
$script:clrStatusBad  = [System.Drawing.Color]::FromArgb(255,120,120)

#================================================
# VARIABLES GLOBALES
#================================================

$script:CurrentUser  = $null      # SamAccountName resolu
$script:InfoOrig     = @{}        # valeurs d'origine de la fiche (pour diff a l'edition)

# Champs courants editables directement dans la fiche : libelle -> attribut LDAP
$script:InfoEditable = [ordered]@{
    "Email"       = "mail"
    "Description" = "description"
    "Societe"     = "company"
    "Departement" = "department"
    "Fonction"    = "title"
}

# Attributs jamais ecrits par l'editeur complet (systeme / construits par AD)
$script:ReadOnlyAttrs = @(
    'DistinguishedName','CanonicalName','CN','Name','ObjectClass','ObjectCategory',
    'ObjectGUID','ObjectSID','SID','SamAccountType','PrimaryGroup','primaryGroupID',
    'whenCreated','whenChanged','Created','Modified','createTimeStamp','modifyTimeStamp',
    'uSNCreated','uSNChanged','LastLogonDate','lastLogon','lastLogonTimestamp',
    'PasswordLastSet','pwdLastSet','LastBadPasswordAttempt','badPasswordTime',
    'BadLogonCount','badPwdCount','logonCount','AccountLockoutTime','LockedOut',
    'Enabled','Deleted','isDeleted','MemberOf','memberof','TokenGroups','sIDHistory',
    'nTSecurityDescriptor','msDS-User-Account-Control-Computed','userAccountControl',
    'PropertyNames','PropertyCount','AddedProperties','RemovedProperties',
    'ModifiedProperties','ProtectedFromAccidentalDeletion'
)

#================================================
# HELPERS DE STYLE
#================================================

function Set-FlatButton
{
    param($Button, $Back, $Fore = [System.Drawing.Color]::White)
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $Back
    $Button.ForeColor = $Fore
    $Button.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $Button.Cursor = 'Hand'
}

function Set-GridStyle
{
    param($Grid)
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.BackgroundColor = [System.Drawing.Color]::White
    $Grid.BorderStyle = 'None'
    $Grid.RowHeadersVisible = $false
    $Grid.GridColor = [System.Drawing.Color]::FromArgb(225,228,232)
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:clrGridHead
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $Grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $Grid.ColumnHeadersHeight = 32
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:clrAlt
    $Grid.DefaultCellStyle.SelectionBackColor = $script:clrAccent
    $Grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $Grid.RowTemplate.Height = 26
    $Grid.CellBorderStyle = 'SingleHorizontal'
}

#================================================
# HELPERS METIER
#================================================

function Format-AttrValue
{
    param($Value)
    if ($null -eq $Value) { return "" }
    if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string]))
    {
        return (($Value | ForEach-Object { "$_" }) -join " ; ")
    }
    return "$Value"
}

#------------------------------------------------
# MOTEUR DE RECHERCHE AD
# Prefixe d'abord (indexe, rapide) sur les principaux attributs,
# puis repli en "contient" si aucun resultat. Limite a 30 lignes.
#------------------------------------------------
function Invoke-SmartSearch
{
    param([string]$Query)

    $Query = $Query.Trim()
    if ($Query.Length -lt 2) { return @() }

    $q = $Query -replace "'", "''"
    $props = 'DisplayName','Mail','SamAccountName','Department','Title','Enabled','UserPrincipalName','GivenName','Surname'

    # 1) recherche par prefixe (rapide car indexee)
    $filterPrefix = "SamAccountName -like '$q*' -or GivenName -like '$q*' -or Surname -like '$q*' -or DisplayName -like '$q*' -or Mail -like '$q*' -or UserPrincipalName -like '$q*'"
    $results = @(Get-ADUser -Filter $filterPrefix -Properties $props -ErrorAction SilentlyContinue |
                 Sort-Object DisplayName | Select-Object -First 30)

    # 2) repli "contient"
    if ($results.Count -eq 0)
    {
        $w = "*$q*"
        $filterWild = "Name -like '$w' -or DisplayName -like '$w' -or SamAccountName -like '$w' -or Mail -like '$w'"
        $results = @(Get-ADUser -Filter $filterWild -Properties $props -ErrorAction SilentlyContinue |
                     Sort-Object DisplayName | Select-Object -First 30)
    }

    return $results
}

#------------------------------------------------
# AFFICHAGE de la fiche utilisateur (grille + groupes + statut)
#------------------------------------------------
function Show-UserDetails
{
    param([string]$Sam)

    try
    {
        $ADUser = Get-ADUser -Identity $Sam -Properties *

        $table = New-Object System.Data.DataTable
        [void]$table.Columns.Add("Attribut")
        [void]$table.Columns.Add("Valeur")

        $Manager = ""
        if ($ADUser.Manager)
        {
            try { $Manager = (Get-ADUser $ADUser.Manager -Properties DisplayName).DisplayName } catch {}
        }

        $rows = [ordered]@{
            "Nom"                              = $ADUser.Name
            "Login"                            = $ADUser.SamAccountName
            "Email"                            = $ADUser.Mail
            "UPN"                              = $ADUser.UserPrincipalName
            "Description"                      = $ADUser.Description
            "Societe"                          = $ADUser.Company
            "Departement"                      = $ADUser.Department
            "Fonction"                         = $ADUser.Title
            "Manager"                          = $Manager
            "Compte actif"                     = $(if($ADUser.Enabled){"Oui"}else{"Non"})
            "Compte verrouille"                = $(if($ADUser.LockedOut){"Oui"}else{"Non"})
            "Date creation"                    = $ADUser.WhenCreated
            "Derniere modification"            = $ADUser.WhenChanged
            "Derniere connexion"               = $ADUser.LastLogonDate
            "Dernier changement mot de passe"  = $ADUser.PasswordLastSet
            "Dernier mauvais mot de passe"     = $ADUser.LastBadPasswordAttempt
            "Nb erreurs authentification"      = $ADUser.BadLogonCount
            "Compte expire le"                 = $ADUser.AccountExpirationDate
            "DN"                               = $ADUser.DistinguishedName
        }

        $script:InfoOrig = @{}
        foreach ($k in $rows.Keys)
        {
            $v = Format-AttrValue $rows[$k]
            [void]$table.Rows.Add($k, $v)
            $script:InfoOrig[$k] = $v
        }

        $gridUser.ReadOnly = $true
        $gridUser.DataSource = $null
        $gridUser.DataSource = $table
        if ($gridUser.Columns.Count -ge 2)
        {
            $gridUser.Columns["Attribut"].FillWeight = 38
            $gridUser.Columns["Valeur"].FillWeight   = 62
        }

        $listGroups.Items.Clear()
        $Groups = Get-ADPrincipalGroupMembership $Sam | Sort-Object Name
        foreach ($Group in $Groups) { [void]$listGroups.Items.Add($Group.Name) }

        $script:CurrentUser = $ADUser.SamAccountName
        $btnEdit.Text = "Editer"
        Set-FlatButton $btnEdit $script:clrGreen

        $status.Text = "  Utilisateur : $($ADUser.Name)   |   login : $($ADUser.SamAccountName)   |   $($Groups.Count) groupes"
        $status.ForeColor = if ($ADUser.LockedOut) { $script:clrStatusBad } else { $script:clrStatusOk }
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show("Utilisateur introuvable","Erreur")
    }
}

#------------------------------------------------
# EDITEUR D'ATTRIBUTS COMPLET (integralite + intelligent)
#------------------------------------------------
function Show-AttributeEditor
{
    param([string]$Sam)

    if ([string]::IsNullOrWhiteSpace($Sam))
    {
        [System.Windows.Forms.MessageBox]::Show("Recherchez d'abord un utilisateur.","Info")
        return
    }

    $adu = Get-ADUser -Identity $Sam -Properties *

    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "Editeur d'attributs - $Sam"
    $editForm.Size = New-Object System.Drawing.Size(960,730)
    $editForm.StartPosition = "CenterParent"
    $editForm.BackColor = $script:clrBg
    $editForm.Font = New-Object System.Drawing.Font("Segoe UI",9)

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtrer :"
    $lblFilter.Location = New-Object System.Drawing.Point(15,18)
    $lblFilter.AutoSize = $true
    $editForm.Controls.Add($lblFilter)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(75,15)
    $txtFilter.Size = New-Object System.Drawing.Size(300,25)
    $editForm.Controls.Add($txtFilter)

    $chkEditableOnly = New-Object System.Windows.Forms.CheckBox
    $chkEditableOnly.Text = "Modifiables uniquement"
    $chkEditableOnly.Location = New-Object System.Drawing.Point(400,17)
    $chkEditableOnly.AutoSize = $true
    $editForm.Controls.Add($chkEditableOnly)

    $chkNonEmpty = New-Object System.Windows.Forms.CheckBox
    $chkNonEmpty.Text = "Masquer les vides"
    $chkNonEmpty.Location = New-Object System.Drawing.Point(600,17)
    $chkNonEmpty.AutoSize = $true
    $editForm.Controls.Add($chkNonEmpty)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(15,50)
    $grid.Size = New-Object System.Drawing.Size(915,580)
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.SelectionMode = "CellSelect"
    Set-GridStyle $grid
    $editForm.Controls.Add($grid)

    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add("Attribut")
    [void]$dt.Columns.Add("Valeur")
    [void]$dt.Columns.Add("Type")
    [void]$dt.Columns.Add("Modifiable")

    $orig = @{}
    foreach ($p in ($adu.PropertyNames | Sort-Object))
    {
        $raw     = $adu.$p
        $val     = Format-AttrValue $raw
        $isMulti = ($raw -is [System.Collections.IEnumerable]) -and ($raw -isnot [string])
        $editable = (-not ($script:ReadOnlyAttrs -contains $p)) -and (-not $isMulti)
        if ($isMulti)           { $type = "Multi-valeur" }
        elseif ($null -ne $raw) { $type = $raw.GetType().Name }
        else                    { $type = "(vide)" }
        [void]$dt.Rows.Add($p, $val, $type, $(if ($editable) { "Oui" } else { "Non" }))
        $orig[$p] = $val
    }

    $bs = New-Object System.Windows.Forms.BindingSource
    $bs.DataSource = $dt
    $grid.DataSource = $bs

    $grid.Columns["Attribut"].ReadOnly     = $true
    $grid.Columns["Type"].ReadOnly         = $true
    $grid.Columns["Modifiable"].ReadOnly   = $true
    $grid.Columns["Attribut"].FillWeight   = 28
    $grid.Columns["Valeur"].FillWeight     = 46
    $grid.Columns["Type"].FillWeight       = 14
    $grid.Columns["Modifiable"].FillWeight = 12

    $grid.Add_CellBeginEdit({
        param($s,$e)
        if ("$($grid.Rows[$e.RowIndex].Cells['Modifiable'].Value)" -eq "Non") { $e.Cancel = $true }
    })
    $grid.Add_DataBindingComplete({
        foreach ($row in $grid.Rows)
        {
            if ("$($row.Cells['Modifiable'].Value)" -eq "Non")
            {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    })

    $applyFilter = {
        $parts = @()
        $f = ($txtFilter.Text -replace "'","''").Trim()
        if ($f)                       { $parts += "(Attribut LIKE '%$f%' OR Valeur LIKE '%$f%')" }
        if ($chkEditableOnly.Checked) { $parts += "Modifiable = 'Oui'" }
        if ($chkNonEmpty.Checked)     { $parts += "Valeur <> ''" }
        $bs.Filter = ($parts -join " AND ")
    }
    $txtFilter.Add_TextChanged($applyFilter)
    $chkEditableOnly.Add_CheckedChanged($applyFilter)
    $chkNonEmpty.Add_CheckedChanged($applyFilter)

    $btnSaveAttr = New-Object System.Windows.Forms.Button
    $btnSaveAttr.Text = "Enregistrer les modifications"
    $btnSaveAttr.Location = New-Object System.Drawing.Point(15,645)
    $btnSaveAttr.Size = New-Object System.Drawing.Size(240,34)
    Set-FlatButton $btnSaveAttr $script:clrGreen
    $editForm.Controls.Add($btnSaveAttr)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"
    $btnClose.Location = New-Object System.Drawing.Point(265,645)
    $btnClose.Size = New-Object System.Drawing.Size(120,34)
    Set-FlatButton $btnClose ([System.Drawing.Color]::FromArgb(120,130,150))
    $btnClose.Add_Click({ $editForm.Close() })
    $editForm.Controls.Add($btnClose)

    $lblHelp = New-Object System.Windows.Forms.Label
    $lblHelp.Location = New-Object System.Drawing.Point(400,653)
    $lblHelp.Size = New-Object System.Drawing.Size(540,25)
    $lblHelp.Text = "Vider une cellule efface l'attribut (-Clear)."
    $lblHelp.ForeColor = [System.Drawing.Color]::Gray
    $editForm.Controls.Add($lblHelp)

    $btnSaveAttr.Add_Click({
        $grid.EndEdit()
        $applied = 0
        $errors  = @()
        foreach ($r in $dt.Rows)
        {
            if ("$($r['Modifiable'])" -ne "Oui") { continue }
            $attr = [string]$r["Attribut"]
            $new  = [string]$r["Valeur"]
            $old  = [string]$orig[$attr]
            if ($new -eq $old) { continue }
            try
            {
                if ([string]::IsNullOrWhiteSpace($new)) { Set-ADUser -Identity $Sam -Clear $attr -ErrorAction Stop }
                else { Set-ADUser -Identity $Sam -Replace @{ $attr = $new } -ErrorAction Stop }
                $orig[$attr] = $new
                $applied++
            }
            catch { $errors += "$attr : $($_.Exception.Message)" }
        }
        $msg = "$applied attribut(s) modifie(s)."
        if ($errors.Count -gt 0) { $msg += "`r`n`r`nEchecs :`r`n" + ($errors -join "`r`n") }
        [System.Windows.Forms.MessageBox]::Show($msg,"Resultat")
        Show-UserDetails $Sam
    })

    [void]$editForm.ShowDialog()
}

#================================================
# FORMULAIRE PRINCIPAL
#================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Audit Active Directory"
$form.Size = New-Object System.Drawing.Size(1200,740)
$form.StartPosition = "CenterScreen"
$form.BackColor = $script:clrBg
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)
$form.MinimumSize = New-Object System.Drawing.Size(1000,650)

#------------------------------------------------
# BANDEAU (titre + recherche)
#------------------------------------------------
$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 68
$header.BackColor = $script:clrHeader
$form.Controls.Add($header)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "AUDIT ACTIVE DIRECTORY"
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20,18)
$lblTitle.AutoSize = $true
$header.Controls.Add($lblTitle)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(430,20)
$txtUser.Size = New-Object System.Drawing.Size(340,28)
$txtUser.Font = New-Object System.Drawing.Font("Segoe UI",11)
$header.Controls.Add($txtUser)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Nom, login, mail, UPN..."
$lblHint.ForeColor = [System.Drawing.Color]::FromArgb(150,160,180)
$lblHint.Font = New-Object System.Drawing.Font("Segoe UI",8)
$lblHint.Location = New-Object System.Drawing.Point(432,50)
$lblHint.AutoSize = $true
$header.Controls.Add($lblHint)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Rechercher"
$btnSearch.Location = New-Object System.Drawing.Point(785,19)
$btnSearch.Size = New-Object System.Drawing.Size(120,30)
Set-FlatButton $btnSearch $script:clrAccent
$header.Controls.Add($btnSearch)

#------------------------------------------------
# RESULTATS DE RECHERCHE (moteur temps reel)
#------------------------------------------------
$lblResults = New-Object System.Windows.Forms.Label
$lblResults.Text = "Resultats"
$lblResults.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblResults.Location = New-Object System.Drawing.Point(20,76)
$lblResults.AutoSize = $true
$form.Controls.Add($lblResults)

$lvResults = New-Object System.Windows.Forms.ListView
$lvResults.Location = New-Object System.Drawing.Point(20,98)
$lvResults.Size = New-Object System.Drawing.Size(1150,140)
$lvResults.View = "Details"
$lvResults.FullRowSelect = $true
$lvResults.MultiSelect = $false
$lvResults.HideSelection = $false
$lvResults.GridLines = $false
$lvResults.Font = New-Object System.Drawing.Font("Segoe UI",9)
$lvResults.BackColor = [System.Drawing.Color]::White
[void]$lvResults.Columns.Add("Nom",300)
[void]$lvResults.Columns.Add("Login",150)
[void]$lvResults.Columns.Add("Email",300)
[void]$lvResults.Columns.Add("Departement",220)
[void]$lvResults.Columns.Add("Actif",70)
$form.Controls.Add($lvResults)

#------------------------------------------------
# BARRE D'ACTIONS
#------------------------------------------------
$btnUnlock = New-Object System.Windows.Forms.Button
$btnUnlock.Text = "Deverrouiller"
$btnUnlock.Location = New-Object System.Drawing.Point(20,250)
$btnUnlock.Size = New-Object System.Drawing.Size(130,34)
Set-FlatButton $btnUnlock $script:clrAmber
$form.Controls.Add($btnUnlock)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Reset Password"
$btnReset.Location = New-Object System.Drawing.Point(160,250)
$btnReset.Size = New-Object System.Drawing.Size(140,34)
Set-FlatButton $btnReset $script:clrRed
$form.Controls.Add($btnReset)

$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = "Editer"
$btnEdit.Location = New-Object System.Drawing.Point(310,250)
$btnEdit.Size = New-Object System.Drawing.Size(110,34)
Set-FlatButton $btnEdit $script:clrGreen
$form.Controls.Add($btnEdit)

$btnEditor = New-Object System.Windows.Forms.Button
$btnEditor.Text = "Editeur d'attributs"
$btnEditor.Location = New-Object System.Drawing.Point(430,250)
$btnEditor.Size = New-Object System.Drawing.Size(170,34)
Set-FlatButton $btnEditor $script:clrAccent
$form.Controls.Add($btnEditor)

$chkPwd = New-Object System.Windows.Forms.CheckBox
$chkPwd.Text = "Changer au prochain logon"
$chkPwd.Location = New-Object System.Drawing.Point(620,258)
$chkPwd.AutoSize = $true
$chkPwd.Checked = $true
$form.Controls.Add($chkPwd)

#------------------------------------------------
# FICHE (colonne gauche) + GROUPES (colonne droite)
#------------------------------------------------
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Informations"
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblInfo.Location = New-Object System.Drawing.Point(20,296)
$lblInfo.AutoSize = $true
$form.Controls.Add($lblInfo)

$gridUser = New-Object System.Windows.Forms.DataGridView
$gridUser.Location = New-Object System.Drawing.Point(20,318)
$gridUser.Size = New-Object System.Drawing.Size(720,318)
$gridUser.Anchor = "Top,Left,Bottom"
$gridUser.ReadOnly = $true
$gridUser.AllowUserToAddRows = $false
$gridUser.AutoSizeColumnsMode = "Fill"
$gridUser.SelectionMode = "CellSelect"
Set-GridStyle $gridUser
$form.Controls.Add($gridUser)

$lblGroups = New-Object System.Windows.Forms.Label
$lblGroups.Text = "Groupes Active Directory"
$lblGroups.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblGroups.Location = New-Object System.Drawing.Point(760,296)
$lblGroups.AutoSize = $true
$form.Controls.Add($lblGroups)

$listGroups = New-Object System.Windows.Forms.ListBox
$listGroups.Location = New-Object System.Drawing.Point(760,318)
$listGroups.Size = New-Object System.Drawing.Size(410,318)
$listGroups.Anchor = "Top,Left,Right,Bottom"
$listGroups.Font = New-Object System.Drawing.Font("Segoe UI",9)
$listGroups.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($listGroups)

#------------------------------------------------
# BARRE DE STATUT
#------------------------------------------------
$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = "Bottom"
$statusBar.Height = 28
$statusBar.BackColor = $script:clrHeader
$form.Controls.Add($statusBar)

$status = New-Object System.Windows.Forms.Label
$status.Dock = "Fill"
$status.TextAlign = "MiddleLeft"
$status.ForeColor = [System.Drawing.Color]::White
$status.Font = New-Object System.Drawing.Font("Segoe UI",9)
$status.Text = "  Pret"
$statusBar.Controls.Add($status)

#================================================
# MOTEUR TEMPS REEL (debounce via Timer)
#================================================

$searchTimer = New-Object System.Windows.Forms.Timer
$searchTimer.Interval = 350

function Update-Suggestions
{
    $q = $txtUser.Text.Trim()
    $lvResults.Items.Clear()

    if ($q.Length -lt 2)
    {
        $status.Text = "  Saisissez au moins 2 caracteres..."
        $status.ForeColor = [System.Drawing.Color]::White
        return
    }

    $matches = @(Invoke-SmartSearch $q)

    if ($matches.Count -eq 0)
    {
        $status.Text = "  Aucun resultat pour : $q"
        $status.ForeColor = $script:clrStatusBad
        return
    }

    foreach ($m in $matches)
    {
        $nom = if ($m.DisplayName) { $m.DisplayName } else { $m.Name }
        $it = New-Object System.Windows.Forms.ListViewItem($nom)
        [void]$it.SubItems.Add([string]$m.SamAccountName)
        [void]$it.SubItems.Add([string]$m.Mail)
        [void]$it.SubItems.Add([string]$m.Department)
        [void]$it.SubItems.Add($(if ($m.Enabled) { "Oui" } else { "Non" }))
        $it.Tag = $m.SamAccountName
        if (-not $m.Enabled) { $it.ForeColor = [System.Drawing.Color]::Gray }
        [void]$lvResults.Items.Add($it)
    }

    $status.Text = "  $($matches.Count) resultat(s) - double-cliquez pour ouvrir"
    $status.ForeColor = $script:clrStatusOk
}

$searchTimer.Add_Tick({
    $searchTimer.Stop()
    Update-Suggestions
})

$txtUser.Add_TextChanged({
    $searchTimer.Stop()
    $searchTimer.Start()
})

# ouvre l'element selectionne (ou le premier)
function Open-Selected
{
    $sam = $null
    if ($lvResults.SelectedItems.Count -gt 0) { $sam = $lvResults.SelectedItems[0].Tag }
    elseif ($lvResults.Items.Count -gt 0)     { $sam = $lvResults.Items[0].Tag }
    if ($sam) { Show-UserDetails $sam }
}

$lvResults.Add_DoubleClick({ Open-Selected })
$lvResults.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { Open-Selected } })

$txtUser.Add_KeyDown({
    if ($_.KeyCode -eq "Enter")
    {
        $searchTimer.Stop()
        Update-Suggestions
        Open-Selected
    }
})

$btnSearch.Add_Click({
    $searchTimer.Stop()
    Update-Suggestions
    Open-Selected
})

#================================================
# EDITION INLINE DE LA FICHE (bouton Editer)
#================================================

$gridUser.Add_CellBeginEdit({
    param($s,$e)
    $label = "$($gridUser.Rows[$e.RowIndex].Cells['Attribut'].Value)"
    if (-not $script:InfoEditable.Contains($label)) { $e.Cancel = $true }
})

$btnEdit.Add_Click({

    if (-not $script:CurrentUser) { return }

    if ($gridUser.ReadOnly)
    {
        # --- passage en mode edition ---
        $gridUser.ReadOnly = $false
        foreach ($row in $gridUser.Rows)
        {
            $label = "$($row.Cells['Attribut'].Value)"
            if ($script:InfoEditable.Contains($label))
            {
                $row.Cells['Valeur'].Style.BackColor = $script:clrEditHint
            }
        }
        $btnEdit.Text = "Enregistrer"
        Set-FlatButton $btnEdit $script:clrAccent
        $status.Text = "  Mode edition : Email / Description / Societe / Departement / Fonction"
        $status.ForeColor = [System.Drawing.Color]::White
    }
    else
    {
        # --- enregistrement ---
        $gridUser.EndEdit()
        $applied = 0
        $errors  = @()

        foreach ($row in $gridUser.Rows)
        {
            $label = "$($row.Cells['Attribut'].Value)"
            if (-not $script:InfoEditable.Contains($label)) { continue }

            $new = "$($row.Cells['Valeur'].Value)"
            $old = "$($script:InfoOrig[$label])"
            if ($new -eq $old) { continue }

            $attr = $script:InfoEditable[$label]
            try
            {
                if ([string]::IsNullOrWhiteSpace($new)) { Set-ADUser -Identity $script:CurrentUser -Clear $attr -ErrorAction Stop }
                else { Set-ADUser -Identity $script:CurrentUser -Replace @{ $attr = $new } -ErrorAction Stop }
                $applied++
            }
            catch { $errors += "$label : $($_.Exception.Message)" }
        }

        if ($errors.Count -gt 0)
        {
            [System.Windows.Forms.MessageBox]::Show(
                "$applied champ(s) modifie(s).`r`n`r`nEchecs :`r`n" + ($errors -join "`r`n"), "Resultat")
        }
        elseif ($applied -gt 0)
        {
            [System.Windows.Forms.MessageBox]::Show("$applied champ(s) modifie(s) avec succes.","Succes")
        }

        Show-UserDetails $script:CurrentUser
    }
})

#================================================
# UNLOCK
#================================================

$btnUnlock.Add_Click({
    $target = if ($script:CurrentUser) { $script:CurrentUser } else { $null }
    if (-not $target) { [System.Windows.Forms.MessageBox]::Show("Selectionnez un utilisateur.","Info"); return }
    try
    {
        Unlock-ADAccount -Identity $target
        [System.Windows.Forms.MessageBox]::Show("Compte deverrouille","Succes")
        Show-UserDetails $target
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Erreur") }
})

#================================================
# RESET PASSWORD
#================================================

$btnReset.Add_Click({

    $User = if ($script:CurrentUser) { $script:CurrentUser } else { $null }
    if (-not $User) { [System.Windows.Forms.MessageBox]::Show("Selectionnez un utilisateur.","Info"); return }

    try
    {
        $pwdForm = New-Object System.Windows.Forms.Form
        $pwdForm.Text = "Nouveau mot de passe"
        $pwdForm.Size = New-Object System.Drawing.Size(450,180)
        $pwdForm.StartPosition = "CenterParent"
        $pwdForm.BackColor = $script:clrBg
        $pwdForm.Font = New-Object System.Drawing.Font("Segoe UI",9)

        $lblPwd = New-Object System.Windows.Forms.Label
        $lblPwd.Text = "Nouveau mot de passe :"
        $lblPwd.Location = New-Object System.Drawing.Point(20,25)
        $lblPwd.AutoSize = $true
        $pwdForm.Controls.Add($lblPwd)

        $txtPwd = New-Object System.Windows.Forms.TextBox
        $txtPwd.Location = New-Object System.Drawing.Point(180,22)
        $txtPwd.Size = New-Object System.Drawing.Size(220,25)
        $txtPwd.UseSystemPasswordChar = $true
        $pwdForm.Controls.Add($txtPwd)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Valider"
        $btnOK.Location = New-Object System.Drawing.Point(180,72)
        $btnOK.Size = New-Object System.Drawing.Size(120,32)
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        Set-FlatButton $btnOK $script:clrGreen
        $pwdForm.Controls.Add($btnOK)
        $pwdForm.AcceptButton = $btnOK

        if ($pwdForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $SecurePassword = ConvertTo-SecureString $txtPwd.Text -AsPlainText -Force
            Set-ADAccountPassword -Identity $User -Reset -NewPassword $SecurePassword
            Set-ADUser -Identity $User -ChangePasswordAtLogon $chkPwd.Checked
            [System.Windows.Forms.MessageBox]::Show("Mot de passe reinitialise avec succes","Succes")
            $status.Text = "  Password reset effectue pour $User"
            $status.ForeColor = $script:clrStatusOk
        }
    }
    catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Erreur") }
})

#================================================
# EDITEUR D'ATTRIBUTS
#================================================

$btnEditor.Add_Click({ Show-AttributeEditor $script:CurrentUser })

#================================================
# DEMARRAGE
#================================================

$txtUser.Select()
[void]$form.ShowDialog()
