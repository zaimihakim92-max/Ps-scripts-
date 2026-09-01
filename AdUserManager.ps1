Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module ActiveDirectory

#================================================
# VARIABLES GLOBALES
#================================================

# Utilisateur actuellement resolu (SamAccountName).
# Toutes les actions (Unlock / Reset / Editeur) agissent dessus,
# et non plus sur le texte brut de la zone de recherche.
$script:CurrentUser = $null

# Attributs consideres comme lecture seule / construits par AD :
# l'editeur ne tentera jamais de les ecrire.
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
    'ModifiedProperties','DistinguishedName','ProtectedFromAccidentalDeletion'
)

#================================================
# FONCTIONS UTILITAIRES
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
# RECHERCHE DYNAMIQUE
# Cherche sur SamAccountName / UPN / Mail / EmployeeID / proxyAddresses
# puis, si rien, en "contient" sur Nom / Prenom / DisplayName / Login / Mail.
#------------------------------------------------
function Invoke-SmartSearch
{
    param([string]$Query)

    $Query = $Query.Trim()
    if ([string]::IsNullOrWhiteSpace($Query)) { return @() }

    # Echappement des apostrophes pour le filtre AD
    $q = $Query -replace "'", "''"

    $props = 'DisplayName','Mail','SamAccountName','Department','Title','Enabled'

    # 1) Correspondance exacte (recherche precise)
    $filterExact = "SamAccountName -eq '$q' -or UserPrincipalName -eq '$q' -or Mail -eq '$q' -or EmployeeID -eq '$q' -or proxyAddresses -eq 'smtp:$q'"
    $results = @(Get-ADUser -Filter $filterExact -Properties $props -ErrorAction SilentlyContinue)

    # 2) Sinon recherche approximative (contient)
    if ($results.Count -eq 0)
    {
        $w = "*$q*"
        $filterWild = "Name -like '$w' -or DisplayName -like '$w' -or GivenName -like '$w' -or Surname -like '$w' -or SamAccountName -like '$w' -or Mail -like '$w'"
        $results = @(Get-ADUser -Filter $filterWild -Properties $props -ErrorAction SilentlyContinue |
                     Sort-Object DisplayName |
                     Select-Object -First 100)
    }

    return $results
}

#------------------------------------------------
# SELECTION quand plusieurs resultats
# Retourne le SamAccountName choisi, ou $null si annule.
#------------------------------------------------
function Select-FromMatches
{
    param($Matches)

    $pick = New-Object System.Windows.Forms.Form
    $pick.Text = "Plusieurs resultats - selectionnez un utilisateur"
    $pick.Size = New-Object System.Drawing.Size(760,460)
    $pick.StartPosition = "CenterParent"

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location = New-Object System.Drawing.Point(10,10)
    $g.Size = New-Object System.Drawing.Size(725,360)
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AutoSizeColumnsMode = "Fill"
    $g.SelectionMode = "FullRowSelect"
    $g.MultiSelect = $false

    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add("Nom")
    [void]$dt.Columns.Add("Login")
    [void]$dt.Columns.Add("Mail")
    [void]$dt.Columns.Add("Departement")
    [void]$dt.Columns.Add("Actif")

    foreach ($m in $Matches)
    {
        [void]$dt.Rows.Add(
            $m.DisplayName,
            $m.SamAccountName,
            $m.Mail,
            $m.Department,
            $(if ($m.Enabled) { "Oui" } else { "Non" })
        )
    }

    $g.DataSource = $dt
    $pick.Controls.Add($g)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Selectionner"
    $btnOK.Location = New-Object System.Drawing.Point(535,385)
    $btnOK.Size = New-Object System.Drawing.Size(120,30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pick.Controls.Add($btnOK)
    $pick.AcceptButton = $btnOK

    # Double-clic = selection directe
    $g.Add_CellDoubleClick({ $btnOK.PerformClick() })

    $script:PickedSam = $null
    if ($pick.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if ($g.SelectedRows.Count -gt 0)
        {
            $script:PickedSam = $g.SelectedRows[0].Cells["Login"].Value
        }
    }

    return $script:PickedSam
}

#------------------------------------------------
# AFFICHAGE des details d'un utilisateur (grille + groupes + statut)
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

        $table.Rows.Add("Nom",$ADUser.Name)
        $table.Rows.Add("Login",$ADUser.SamAccountName)
        $table.Rows.Add("Email",$ADUser.Mail)
        $table.Rows.Add("UPN",$ADUser.UserPrincipalName)
        $table.Rows.Add("Description",$ADUser.Description)
        $table.Rows.Add("Societe",$ADUser.Company)
        $table.Rows.Add("Departement",$ADUser.Department)
        $table.Rows.Add("Fonction",$ADUser.Title)
        $table.Rows.Add("Manager",$Manager)
        $table.Rows.Add("Compte actif",$(if($ADUser.Enabled){"Oui"}else{"Non"}))
        $table.Rows.Add("Compte verrouille",$(if($ADUser.LockedOut){"Oui"}else{"Non"}))
        $table.Rows.Add("Date creation",$ADUser.WhenCreated)
        $table.Rows.Add("Derniere modification",$ADUser.WhenChanged)
        $table.Rows.Add("Derniere connexion",$ADUser.LastLogonDate)
        $table.Rows.Add("Dernier changement mot de passe",$ADUser.PasswordLastSet)
        $table.Rows.Add("Dernier mauvais mot de passe",$ADUser.LastBadPasswordAttempt)
        $table.Rows.Add("Nb erreurs authentification",$ADUser.BadLogonCount)
        $table.Rows.Add("Compte expire le",$ADUser.AccountExpirationDate)
        $table.Rows.Add("DN",$ADUser.DistinguishedName)

        $gridUser.DataSource = $null
        $gridUser.DataSource = $table

        $listGroups.Items.Clear()
        $Groups = Get-ADPrincipalGroupMembership $Sam | Sort-Object Name
        foreach ($Group in $Groups) { [void]$listGroups.Items.Add($Group.Name) }

        $script:CurrentUser = $ADUser.SamAccountName

        $status.Text = "OK Utilisateur : $($ADUser.Name)  -  $($Groups.Count) groupes  -  login : $($ADUser.SamAccountName)"
        if ($ADUser.LockedOut) { $status.ForeColor = "Red" } else { $status.ForeColor = "Green" }
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show("Utilisateur introuvable","Erreur")
    }
}

#------------------------------------------------
# EDITEUR D'ATTRIBUTS INTELLIGENT
# - visualise l'integralite des attributs (Get-ADUser -Properties *)
# - filtre en direct (nom d'attribut ou valeur)
# - affiche le type et si l'attribut est modifiable
# - n'autorise l'edition que sur les attributs mono-valeur non systeme
# - applique via Set-ADUser (-Replace) ou -Clear si vide, attribut par attribut
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
    $editForm.Size = New-Object System.Drawing.Size(950,720)
    $editForm.StartPosition = "CenterParent"

    # --- barre de filtre ---
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

    # --- grille ---
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(15,50)
    $grid.Size = New-Object System.Drawing.Size(905,570)
    $grid.AllowUserToAddRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.SelectionMode = "CellSelect"
    $editForm.Controls.Add($grid)

    # DataTable source
    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add("Attribut")
    [void]$dt.Columns.Add("Valeur")
    [void]$dt.Columns.Add("Type")
    [void]$dt.Columns.Add("Modifiable")

    # Valeurs d'origine pour calculer les modifications au moment de l'enregistrement
    $orig = @{}

    foreach ($p in ($adu.PropertyNames | Sort-Object))
    {
        $raw   = $adu.$p
        $val   = Format-AttrValue $raw
        $isMulti = ($raw -is [System.Collections.IEnumerable]) -and ($raw -isnot [string])
        $editable = (-not ($script:ReadOnlyAttrs -contains $p)) -and (-not $isMulti)

        if ($isMulti)              { $type = "Multi-valeur" }
        elseif ($null -ne $raw)    { $type = $raw.GetType().Name }
        else                       { $type = "(vide)" }

        [void]$dt.Rows.Add($p, $val, $type, $(if ($editable) { "Oui" } else { "Non" }))
        $orig[$p] = $val
    }

    $bs = New-Object System.Windows.Forms.BindingSource
    $bs.DataSource = $dt
    $grid.DataSource = $bs

    # Colonnes en lecture seule sauf "Valeur"
    $grid.Columns["Attribut"].ReadOnly   = $true
    $grid.Columns["Type"].ReadOnly       = $true
    $grid.Columns["Modifiable"].ReadOnly = $true
    $grid.Columns["Attribut"].FillWeight = 28
    $grid.Columns["Valeur"].FillWeight   = 46
    $grid.Columns["Type"].FillWeight     = 14
    $grid.Columns["Modifiable"].FillWeight = 12

    # Bloque l'edition d'une cellule Valeur si la ligne n'est pas modifiable
    $grid.Add_CellBeginEdit({
        param($s,$e)
        $row = $grid.Rows[$e.RowIndex]
        if ("$($row.Cells['Modifiable'].Value)" -eq "Non") { $e.Cancel = $true }
    })

    # Grise les lignes non modifiables
    $grid.Add_DataBindingComplete({
        foreach ($row in $grid.Rows)
        {
            if ("$($row.Cells['Modifiable'].Value)" -eq "Non")
            {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    })

    # Filtre en direct
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

    # --- boutons ---
    $btnSaveAttr = New-Object System.Windows.Forms.Button
    $btnSaveAttr.Text = "Enregistrer les modifications"
    $btnSaveAttr.Location = New-Object System.Drawing.Point(15,635)
    $btnSaveAttr.Size = New-Object System.Drawing.Size(230,32)
    $editForm.Controls.Add($btnSaveAttr)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"
    $btnClose.Location = New-Object System.Drawing.Point(255,635)
    $btnClose.Size = New-Object System.Drawing.Size(120,32)
    $btnClose.Add_Click({ $editForm.Close() })
    $editForm.Controls.Add($btnClose)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(390,642)
    $lblInfo.Size = New-Object System.Drawing.Size(540,25)
    $lblInfo.Text = "Astuce : vider une cellule effacera l'attribut (-Clear)."
    $lblInfo.ForeColor = "Gray"
    $editForm.Controls.Add($lblInfo)

    $btnSaveAttr.Add_Click({

        # Valide l'edition en cours de cellule
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
                if ([string]::IsNullOrWhiteSpace($new))
                {
                    Set-ADUser -Identity $Sam -Clear $attr -ErrorAction Stop
                }
                else
                {
                    Set-ADUser -Identity $Sam -Replace @{ $attr = $new } -ErrorAction Stop
                }
                $orig[$attr] = $new
                $applied++
            }
            catch
            {
                $errors += "$attr : $($_.Exception.Message)"
            }
        }

        $msg = "$applied attribut(s) modifie(s)."
        if ($errors.Count -gt 0)
        {
            $msg += "`r`n`r`nEchecs :`r`n" + ($errors -join "`r`n")
        }
        [System.Windows.Forms.MessageBox]::Show($msg,"Resultat")

        # Rafraichit la vue principale
        Show-UserDetails $Sam
    })

    [void]$editForm.ShowDialog()
}

#================================================
# FORMULAIRE
#================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Audit Active Directory"
$form.Size = New-Object System.Drawing.Size(1120,760)
$form.StartPosition = "CenterScreen"

#================================================
# RECHERCHE
#================================================

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Recherche :"
$lblUser.Location = New-Object System.Drawing.Point(20,20)
$lblUser.AutoSize = $true
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(100,18)
$txtUser.Size = New-Object System.Drawing.Size(250,25)
$form.Controls.Add($txtUser)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Rechercher"
$btnSearch.Location = New-Object System.Drawing.Point(360,15)
$btnSearch.Size = New-Object System.Drawing.Size(110,30)
$form.Controls.Add($btnSearch)

$btnUnlock = New-Object System.Windows.Forms.Button
$btnUnlock.Text = "Deverrouiller"
$btnUnlock.Location = New-Object System.Drawing.Point(475,15)
$btnUnlock.Size = New-Object System.Drawing.Size(110,30)
$form.Controls.Add($btnUnlock)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Reset Password"
$btnReset.Location = New-Object System.Drawing.Point(590,15)
$btnReset.Size = New-Object System.Drawing.Size(120,30)
$form.Controls.Add($btnReset)

$btnEditor = New-Object System.Windows.Forms.Button
$btnEditor.Text = "Editeur d'attributs"
$btnEditor.Location = New-Object System.Drawing.Point(715,15)
$btnEditor.Size = New-Object System.Drawing.Size(150,30)
$form.Controls.Add($btnEditor)

$chkPwd = New-Object System.Windows.Forms.CheckBox
$chkPwd.Text = "Changer au prochain logon"
$chkPwd.Location = New-Object System.Drawing.Point(880,20)
$chkPwd.AutoSize = $true
$chkPwd.Checked = $true
$form.Controls.Add($chkPwd)

#================================================
# GRID INFOS
#================================================

$gridUser = New-Object System.Windows.Forms.DataGridView
$gridUser.Location = New-Object System.Drawing.Point(20,60)
$gridUser.Size = New-Object System.Drawing.Size(1060,320)
$gridUser.ReadOnly = $true
$gridUser.AllowUserToAddRows = $false
$gridUser.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($gridUser)

#================================================
# GROUPES
#================================================

$lblGroups = New-Object System.Windows.Forms.Label
$lblGroups.Text = "Groupes Active Directory"
$lblGroups.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$lblGroups.Location = New-Object System.Drawing.Point(20,390)
$lblGroups.AutoSize = $true
$form.Controls.Add($lblGroups)

$listGroups = New-Object System.Windows.Forms.ListBox
$listGroups.Location = New-Object System.Drawing.Point(20,420)
$listGroups.Size = New-Object System.Drawing.Size(1060,220)
$form.Controls.Add($listGroups)

#================================================
# STATUS
#================================================

$status = New-Object System.Windows.Forms.Label
$status.Location = New-Object System.Drawing.Point(20,660)
$status.Size = New-Object System.Drawing.Size(1040,25)
$form.Controls.Add($status)

#================================================
# EVENEMENTS
#================================================

# --- RECHERCHE (dynamique + selection si plusieurs resultats) ---
$btnSearch.Add_Click({

    $query = $txtUser.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($query)) { return }

    $matches = @(Invoke-SmartSearch $query)

    if ($matches.Count -eq 0)
    {
        $status.Text = "Aucun resultat pour : $query"
        $status.ForeColor = "Red"
        $gridUser.DataSource = $null
        $listGroups.Items.Clear()
        return
    }

    if ($matches.Count -eq 1)
    {
        $sam = $matches[0].SamAccountName
    }
    else
    {
        $sam = Select-FromMatches $matches
        if (-not $sam) { return }
    }

    Show-UserDetails $sam
})

# --- UNLOCK ---
$btnUnlock.Add_Click({

    $target = if ($script:CurrentUser) { $script:CurrentUser } else { $txtUser.Text.Trim() }
    if ([string]::IsNullOrWhiteSpace($target)) { return }

    try
    {
        Unlock-ADAccount -Identity $target
        [System.Windows.Forms.MessageBox]::Show("Compte deverrouille","Succes")
        Show-UserDetails $target
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Erreur")
    }
})

# --- RESET PASSWORD ---
$btnReset.Add_Click({

    $User = if ($script:CurrentUser) { $script:CurrentUser } else { $txtUser.Text.Trim() }
    if ([string]::IsNullOrWhiteSpace($User)) { return }

    try
    {
        $pwdForm = New-Object System.Windows.Forms.Form
        $pwdForm.Text = "Nouveau mot de passe"
        $pwdForm.Size = New-Object System.Drawing.Size(450,180)
        $pwdForm.StartPosition = "CenterParent"

        $lblPwd = New-Object System.Windows.Forms.Label
        $lblPwd.Text = "Nouveau mot de passe :"
        $lblPwd.Location = New-Object System.Drawing.Point(20,25)
        $lblPwd.AutoSize = $true
        $pwdForm.Controls.Add($lblPwd)

        $txtPwd = New-Object System.Windows.Forms.TextBox
        $txtPwd.Location = New-Object System.Drawing.Point(180,20)
        $txtPwd.Size = New-Object System.Drawing.Size(220,25)
        $txtPwd.UseSystemPasswordChar = $true
        $pwdForm.Controls.Add($txtPwd)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Valider"
        $btnOK.Location = New-Object System.Drawing.Point(180,70)
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $pwdForm.Controls.Add($btnOK)

        $pwdForm.AcceptButton = $btnOK

        if ($pwdForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $SecurePassword = ConvertTo-SecureString `
                $txtPwd.Text `
                -AsPlainText `
                -Force

            Set-ADAccountPassword `
                -Identity $User `
                -Reset `
                -NewPassword $SecurePassword

            Set-ADUser `
                -Identity $User `
                -ChangePasswordAtLogon $chkPwd.Checked

            [System.Windows.Forms.MessageBox]::Show("Mot de passe reinitialise avec succes","Succes")

            $status.Text = "OK Password reset effectue pour $User"
            $status.ForeColor = "Blue"
        }
    }
    catch
    {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Erreur")
    }
})

# --- EDITEUR D'ATTRIBUTS ---
$btnEditor.Add_Click({
    Show-AttributeEditor $script:CurrentUser
})

# --- TOUCHE ENTREE ---
$txtUser.Add_KeyDown({
    if ($_.KeyCode -eq "Enter")
    {
        $btnSearch.PerformClick()
    }
})

#================================================
# DEMARRAGE
#================================================

[void]$form.ShowDialog()
