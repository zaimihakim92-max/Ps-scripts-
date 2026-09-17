<#
.SYNOPSIS
    Diagnostic et remediation des mismatch d'identite utilisateur sur UN site SharePoint Online.
    Interface graphique. Module : PnP.PowerShell uniquement.

.DESCRIPTION
    Chaque site collection garde une copie en cache de l'identite dans la User Information
    List (UIL). Apres une migration/recreation de compte ou un changement d'UPN, cette entree
    peut ne plus correspondre : permissions accrochees a l'ancienne entree, doublons, orphelins.

    L'outil, en PnP pur :
      1. DIAGNOSTIC  : recherche toutes les entrees UIL correspondant a un utilisateur,
                       et signale doublons (meme email, login different), entrees sans email,
                       et donne un indice d'anciennete (Id le plus ancien = souvent le perime).
      2. CAPTURE     : liste les groupes SharePoint de l'entree perimee choisie.
      3. REMEDIATION : transfere ces appartenances vers le login courant, gere le flag
                       administrateur de collection, puis purge l'entree perimee.

    Garde-fous :
      - Mode SIMULATION coche par defaut : "Analyser" ne fait que lire et afficher le plan.
      - "Executer" est verrouille tant qu'aucun plan n'a ete analyse, et demande une
        confirmation explicite avant toute ecriture.
      - Hors perimetre : permissions directes au niveau element/liste (non transferees ;
        signalees le cas echeant a verifier manuellement).

.PREREQUIS
    - Module PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser)
    - Une application Entra enregistree (Client ID) autorisee pour PnP (depuis fin 2024).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Etat global ---
$script:Entries      = New-Object System.Collections.Generic.List[object]
$script:StaleUser    = $null      # entree UIL a purger (objet)
$script:TargetLogin  = ""         # claim/login cible (courant)
$script:PlanGroups   = @()        # groupes a reappliquer
$script:PlanSiteAdmin= $false     # transferer le flag admin de collection ?
$script:Analyzed     = $false

# ==================================================================
# Fonctions PnP
# ==================================================================
function Test-PnPConnected {
    try { $w = Get-PnPWeb -ErrorAction Stop; return [bool]$w } catch { return $false }
}

function Get-MatchingUsers {
    # Retourne les entrees UIL correspondant au terme (login/email/titre), avec diagnostic.
    param([Parameter(Mandatory)][string]$Term)
    $t = $Term.Trim().ToLower()
    $all = @(Get-PnPUser -ErrorAction Stop)
    $matches = @($all | Where-Object {
        ($_.LoginName -and $_.LoginName.ToLower().Contains($t)) -or
        ($_.Email     -and $_.Email.ToLower().Contains($t))     -or
        ($_.Title     -and $_.Title.ToLower().Contains($t))
    })

    # Groupes de doublons : meme email (non vide) partage par plusieurs LoginName distincts
    $dupEmails = @($matches | Where-Object { $_.Email } |
        Group-Object { $_.Email.ToLower() } | Where-Object { $_.Count -gt 1 } |
        ForEach-Object { $_.Name })

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($u in $matches) {
        $tags = @()
        if ($u.IsSiteAdmin) { $tags += "ADMIN" }
        if ([string]::IsNullOrWhiteSpace($u.Email)) { $tags += "SANS EMAIL" }
        if ($u.Email -and $dupEmails -contains $u.Email.ToLower()) { $tags += "DOUBLON" }
        if ($tags.Count -eq 0) { $tags += "OK" }

        $result.Add([PSCustomObject]@{
            Id        = $u.Id
            Title     = $u.Title
            LoginName = $u.LoginName
            Email     = $u.Email
            Type      = $u.PrincipalType
            Admin     = [bool]$u.IsSiteAdmin
            Diag      = ($tags -join ", ")
        }) | Out-Null
    }
    return $result
}

function Get-UserGroups {
    # Titres des groupes SharePoint du site contenant ce login.
    param([Parameter(Mandatory)][string]$LoginName)
    $found = New-Object System.Collections.Generic.List[string]
    $groups = @(Get-PnPGroup -ErrorAction SilentlyContinue)
    foreach ($g in $groups) {
        $members = @(Get-PnPGroupMember -Group $g -ErrorAction SilentlyContinue)
        if ($members | Where-Object { $_.LoginName -and ($_.LoginName -ieq $LoginName) }) {
            $found.Add($g.Title) | Out-Null
        }
    }
    return ,$found.ToArray()
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "SharePoint - Remediation mismatch utilisateur (site unique)"
$form.Size          = New-Object System.Drawing.Size(960, 900)
$form.MinimumSize   = New-Object System.Drawing.Size(860, 720)
$form.StartPosition = "CenterScreen"
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Connexion ---
$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Text = "URL du site :"; $lblSite.Location = New-Object System.Drawing.Point(15, 15); $lblSite.AutoSize = $true
$form.Controls.Add($lblSite)

$txtSite = New-Object System.Windows.Forms.TextBox
$txtSite.Location = New-Object System.Drawing.Point(150, 12); $txtSite.Size = New-Object System.Drawing.Size(779, 23)
$txtSite.Anchor = "Top,Left,Right"; $txtSite.Text = "https://contoso.sharepoint.com/sites/MonSite"
$form.Controls.Add($txtSite)

$lblCid = New-Object System.Windows.Forms.Label
$lblCid.Text = "Client ID (app) :"; $lblCid.Location = New-Object System.Drawing.Point(15, 48); $lblCid.AutoSize = $true
$form.Controls.Add($lblCid)

$txtCid = New-Object System.Windows.Forms.TextBox
$txtCid.Location = New-Object System.Drawing.Point(150, 45); $txtCid.Size = New-Object System.Drawing.Size(645, 23)
$txtCid.Anchor = "Top,Left,Right"
$form.Controls.Add($txtCid)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connexion"; $btnConnect.Location = New-Object System.Drawing.Point(804, 44); $btnConnect.Size = New-Object System.Drawing.Size(125, 25)
$btnConnect.Anchor = "Top,Right"
$form.Controls.Add($btnConnect)

$lblConn = New-Object System.Windows.Forms.Label
$lblConn.Location = New-Object System.Drawing.Point(150, 71); $lblConn.Size = New-Object System.Drawing.Size(779, 18)
$lblConn.Anchor = "Top,Left,Right"; $lblConn.ForeColor = [System.Drawing.Color]::Gray
$lblConn.Text = "Client ID requis (votre app Entra PnP). Connexion interactive."
$form.Controls.Add($lblConn)

# --- Recherche ---
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Utilisateur :"; $lblSearch.Location = New-Object System.Drawing.Point(15, 101); $lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(150, 98); $txtSearch.Size = New-Object System.Drawing.Size(645, 23)
$txtSearch.Anchor = "Top,Left,Right"
$form.Controls.Add($txtSearch)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Rechercher"; $btnSearch.Location = New-Object System.Drawing.Point(804, 97); $btnSearch.Size = New-Object System.Drawing.Size(125, 25)
$btnSearch.Anchor = "Top,Right"; $btnSearch.Enabled = $false
$form.Controls.Add($btnSearch)

# --- Grille des entrees UIL ---
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 130); $grid.Size = New-Object System.Drawing.Size(914, 200)
$grid.Anchor = "Top,Left,Right"
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.ReadOnly = $true
$grid.RowHeadersVisible = $false; $grid.SelectionMode = "FullRowSelect"; $grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"; $grid.ColumnHeadersHeightSizeMode = "AutoSize"
$null = $grid.Columns.Add("Id", "Id")
$null = $grid.Columns.Add("Title", "Nom affiche")
$null = $grid.Columns.Add("LoginName", "LoginName (claim)")
$null = $grid.Columns.Add("Email", "Email")
$null = $grid.Columns.Add("Type", "Type")
$null = $grid.Columns.Add("Diag", "Diagnostic")
$grid.Columns["Id"].FillWeight = 6
$grid.Columns["Title"].FillWeight = 18
$grid.Columns["LoginName"].FillWeight = 34
$grid.Columns["Email"].FillWeight = 20
$grid.Columns["Type"].FillWeight = 9
$grid.Columns["Diag"].FillWeight = 13
$form.Controls.Add($grid)

$btnShowGroups = New-Object System.Windows.Forms.Button
$btnShowGroups.Text = "Voir les groupes de l'entree selectionnee"; $btnShowGroups.Location = New-Object System.Drawing.Point(15, 338); $btnShowGroups.Size = New-Object System.Drawing.Size(300, 26)
$btnShowGroups.Anchor = "Top,Left"
$form.Controls.Add($btnShowGroups)

# --- Zone remediation ---
$lblStale = New-Object System.Windows.Forms.Label
$lblStale.Text = "Entree perimee :"; $lblStale.Location = New-Object System.Drawing.Point(15, 378); $lblStale.AutoSize = $true
$form.Controls.Add($lblStale)

$lblStaleVal = New-Object System.Windows.Forms.Label
$lblStaleVal.Location = New-Object System.Drawing.Point(150, 378); $lblStaleVal.Size = New-Object System.Drawing.Size(490, 20)
$lblStaleVal.Anchor = "Top,Left,Right"; $lblStaleVal.ForeColor = [System.Drawing.Color]::Firebrick
$lblStaleVal.Text = "(aucune)"
$form.Controls.Add($lblStaleVal)

$btnSetStale = New-Object System.Windows.Forms.Button
$btnSetStale.Text = "Definir depuis la selection"; $btnSetStale.Location = New-Object System.Drawing.Point(719, 374); $btnSetStale.Size = New-Object System.Drawing.Size(210, 26)
$btnSetStale.Anchor = "Top,Right"
$form.Controls.Add($btnSetStale)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Login cible (courant) :"; $lblTarget.Location = New-Object System.Drawing.Point(15, 410); $lblTarget.AutoSize = $true
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(180, 407); $txtTarget.Size = New-Object System.Drawing.Size(530, 23)
$txtTarget.Anchor = "Top,Left,Right"
$form.Controls.Add($txtTarget)

$btnUseTarget = New-Object System.Windows.Forms.Button
$btnUseTarget.Text = "Cible = selection"; $btnUseTarget.Location = New-Object System.Drawing.Point(719, 406); $btnUseTarget.Size = New-Object System.Drawing.Size(210, 26)
$btnUseTarget.Anchor = "Top,Right"
$form.Controls.Add($btnUseTarget)

$lblTargetHint = New-Object System.Windows.Forms.Label
$lblTargetHint.Text = "UPN/email du compte courant, ou son claim i:0#.f|membership|... - il sera ré-assuré dans le site."
$lblTargetHint.Location = New-Object System.Drawing.Point(180, 432); $lblTargetHint.AutoSize = $true; $lblTargetHint.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblTargetHint)

$chkSimulate = New-Object System.Windows.Forms.CheckBox
$chkSimulate.Text = "Mode simulation (aucune modification)"; $chkSimulate.Location = New-Object System.Drawing.Point(15, 458); $chkSimulate.AutoSize = $true; $chkSimulate.Checked = $true
$form.Controls.Add($chkSimulate)

$btnAnalyze = New-Object System.Windows.Forms.Button
$btnAnalyze.Text = "Analyser le plan"; $btnAnalyze.Location = New-Object System.Drawing.Point(15, 484); $btnAnalyze.Size = New-Object System.Drawing.Size(250, 30)
$form.Controls.Add($btnAnalyze)

$btnExecute = New-Object System.Windows.Forms.Button
$btnExecute.Text = "Executer la remediation"; $btnExecute.Location = New-Object System.Drawing.Point(280, 484); $btnExecute.Size = New-Object System.Drawing.Size(280, 30)
$btnExecute.BackColor = [System.Drawing.Color]::FromArgb(200, 60, 60); $btnExecute.ForeColor = [System.Drawing.Color]::White; $btnExecute.FlatStyle = "Flat"
$btnExecute.Enabled = $false
$form.Controls.Add($btnExecute)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(575, 488); $progress.Size = New-Object System.Drawing.Size(354, 22)
$progress.Anchor = "Top,Right"
$form.Controls.Add($progress)

# --- Journal ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Journal :"; $lblLog.Location = New-Object System.Drawing.Point(15, 522); $lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 542); $txtLog.Size = New-Object System.Drawing.Size(914, 305)
$txtLog.Anchor = "Top,Bottom,Left,Right"
$txtLog.Multiline = $true; $txtLog.ScrollBars = "Vertical"; $txtLog.ReadOnly = $true; $txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

# ==================================================================
# Helpers UI
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }

function Set-ConnState {
    param([bool]$connected, [string]$msg)
    if ($connected) { $lblConn.ForeColor = [System.Drawing.Color]::ForestGreen } else { $lblConn.ForeColor = [System.Drawing.Color]::Gray }
    $lblConn.Text = $msg
    $btnSearch.Enabled = $connected
}

function Reset-Plan {
    $script:Analyzed = $false; $script:PlanGroups = @(); $script:PlanSiteAdmin = $false
    $btnExecute.Enabled = $false
}

# ==================================================================
# Evenements
# ==================================================================
$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Renseignez l'URL du site.", "Connexion", "OK", "Warning") | Out-Null; return
    }
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        [System.Windows.Forms.MessageBox]::Show("Module PnP.PowerShell introuvable.`nInstall-Module PnP.PowerShell -Scope CurrentUser", "Prerequis", "OK", "Error") | Out-Null; return
    }
    $btnConnect.Enabled = $false
    Set-ConnState $false "Connexion en cours..."
    try {
        if ([string]::IsNullOrWhiteSpace($txtCid.Text)) {
            Connect-PnPOnline -Url $txtSite.Text.Trim() -Interactive -ErrorAction Stop
        } else {
            Connect-PnPOnline -Url $txtSite.Text.Trim() -Interactive -ClientId $txtCid.Text.Trim() -ErrorAction Stop
        }
        if (Test-PnPConnected) {
            $w = Get-PnPWeb
            Set-ConnState $true "Connecte : $($w.Title)"
            Write-Log "Connecte au site : $($txtSite.Text.Trim())"
        } else { Set-ConnState $false "Connexion non confirmee." }
    } catch {
        Set-ConnState $false "Echec de connexion."
        [System.Windows.Forms.MessageBox]::Show("Echec de la connexion :`n$($_.Exception.Message)", "Connexion", "OK", "Error") | Out-Null
    }
    $btnConnect.Enabled = $true
})

$btnSearch.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Saisissez un email, un UPN ou un nom.", "Recherche", "OK", "Warning") | Out-Null; return
    }
    if (-not (Test-PnPConnected)) {
        [System.Windows.Forms.MessageBox]::Show("Non connecte a un site.", "Recherche", "OK", "Warning") | Out-Null; return
    }
    $grid.Rows.Clear(); $script:Entries.Clear(); Reset-Plan
    try {
        $res = Get-MatchingUsers -Term $txtSearch.Text
        foreach ($e in $res) {
            $script:Entries.Add($e) | Out-Null
            $i = $grid.Rows.Add($e.Id, $e.Title, $e.LoginName, $e.Email, $e.Type, $e.Diag)
            if ($e.Diag -like "*DOUBLON*") { $grid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,224,150) }
            elseif ($e.Diag -like "*SANS EMAIL*") { $grid.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,235,200) }
        }
        $dup = @($res | Where-Object { $_.Diag -like "*DOUBLON*" }).Count
        Write-Log "Recherche '$($txtSearch.Text)' : $($res.Count) entree(s) trouvee(s), dont $dup en doublon."
        if ($dup -gt 0) { Write-Log "  Indice : sur un doublon, l'Id le plus ancien (plus petit) est souvent l'entree perimee." }
    } catch {
        Write-Log "ERREUR recherche : $($_.Exception.Message)"
    }
})

$btnShowGroups.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $login = [string]$grid.SelectedRows[0].Cells["LoginName"].Value
    if (-not $login) { return }
    Write-Log "Groupes SharePoint de : $login"
    try {
        $g = Get-UserGroups -LoginName $login
        if ($g.Count -eq 0) { Write-Log "  (aucun groupe de site)" }
        else { $g | ForEach-Object { Write-Log "  - $_" } }
    } catch { Write-Log "  ERREUR : $($_.Exception.Message)" }
})

$btnSetStale.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $row = $grid.SelectedRows[0]
    $script:StaleUser = [PSCustomObject]@{
        Id        = [int]$row.Cells["Id"].Value
        Title     = [string]$row.Cells["Title"].Value
        LoginName = [string]$row.Cells["LoginName"].Value
        Admin     = ($script:Entries | Where-Object { $_.Id -eq [int]$row.Cells["Id"].Value } | Select-Object -First 1).Admin
    }
    $lblStaleVal.Text = "Id $($script:StaleUser.Id) - $($script:StaleUser.LoginName)"
    Reset-Plan
    Write-Log "Entree perimee definie : Id $($script:StaleUser.Id) - $($script:StaleUser.LoginName)"
})

$btnUseTarget.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $txtTarget.Text = [string]$grid.SelectedRows[0].Cells["LoginName"].Value
    Reset-Plan
})

$btnAnalyze.Add_Click({
    if (-not (Test-PnPConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte.", "Analyse", "OK", "Warning") | Out-Null; return }
    if (-not $script:StaleUser) { [System.Windows.Forms.MessageBox]::Show("Definissez l'entree perimee a purger.", "Analyse", "OK", "Warning") | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($txtTarget.Text)) { [System.Windows.Forms.MessageBox]::Show("Renseignez le login cible (courant).", "Analyse", "OK", "Warning") | Out-Null; return }

    $script:TargetLogin = $txtTarget.Text.Trim()
    if ($script:TargetLogin -ieq $script:StaleUser.LoginName) {
        [System.Windows.Forms.MessageBox]::Show("Le login cible est identique a l'entree perimee.", "Analyse", "OK", "Warning") | Out-Null; return
    }

    Write-Log "----- ANALYSE (lecture seule) -----"
    try {
        $script:PlanGroups = Get-UserGroups -LoginName $script:StaleUser.LoginName
        $script:PlanSiteAdmin = [bool]$script:StaleUser.Admin
        Write-Log "Entree perimee : Id $($script:StaleUser.Id) - $($script:StaleUser.LoginName)"
        Write-Log "Login cible    : $($script:TargetLogin)"
        Write-Log "Groupes a transferer ($($script:PlanGroups.Count)) :"
        if ($script:PlanGroups.Count -eq 0) { Write-Log "  (aucun groupe de site)" } else { $script:PlanGroups | ForEach-Object { Write-Log "  + $_" } }
        if ($script:PlanSiteAdmin) { Write-Log "Administrateur de collection : sera ajoute a la cible." }
        Write-Log "Puis : suppression de l'entree perimee (Remove-PnPUser Id $($script:StaleUser.Id))."
        Write-Log "Note : la cible sera ré-assurée (New-PnPUser). Si le compte n'existe pas dans l'annuaire, l'execution echouera proprement."
        Write-Log "Rappel perimetre : les permissions directes au niveau element/liste ne sont PAS transferees."
        $script:Analyzed = $true
        $btnExecute.Enabled = $true
        Write-Log "Plan pret. Decochez 'Mode simulation' puis 'Executer' pour appliquer."
    } catch {
        Write-Log "ERREUR analyse : $($_.Exception.Message)"; Reset-Plan
    }
})

$btnExecute.Add_Click({
    if (-not $script:Analyzed) { [System.Windows.Forms.MessageBox]::Show("Analysez d'abord le plan.", "Execution", "OK", "Warning") | Out-Null; return }
    if ($chkSimulate.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Mode simulation actif : decochez-le pour appliquer reellement.", "Execution", "OK", "Information") | Out-Null; return
    }

    $recap = "Confirmer la remediation ?`n`n" +
             "Site   : $($txtSite.Text.Trim())`n" +
             "Purger : Id $($script:StaleUser.Id) - $($script:StaleUser.LoginName)`n" +
             "Cible  : $($script:TargetLogin)`n" +
             "Groupes transferes : $($script:PlanGroups.Count)`n" +
             "Admin de collection : $(if($script:PlanSiteAdmin){'oui'}else{'non'})`n`n" +
             "Cette action modifie les permissions du site."
    if ([System.Windows.Forms.MessageBox]::Show($recap, "Confirmation", "YesNo", "Warning") -ne "Yes") {
        Write-Log "Execution annulee par l'utilisateur."; return
    }

    $btnExecute.Enabled = $false; $btnAnalyze.Enabled = $false
    $steps = $script:PlanGroups.Count + 2 + ([int]$script:PlanSiteAdmin)
    $progress.Minimum = 0; $progress.Maximum = [Math]::Max(1,$steps); $progress.Value = 0
    $done = 0
    Write-Log "===== EXECUTION ====="

    try {
        # 1) Ré-assurer la cible
        Write-Log "Ré-assurance de la cible : $($script:TargetLogin)"
        $t = New-PnPUser -LoginName $script:TargetLogin -ErrorAction Stop
        Write-Log "  OK - Id cible : $($t.Id)"
        $done++; $progress.Value = $done

        # 2) Transferer les groupes
        foreach ($grp in $script:PlanGroups) {
            try {
                Add-PnPGroupMember -Group $grp -LoginName $script:TargetLogin -ErrorAction Stop
                Write-Log "  + Ajoute au groupe : $grp"
            } catch {
                Write-Log "  ! Groupe '$grp' : $($_.Exception.Message)"
            }
            $done++; $progress.Value = $done
        }

        # 3) Admin de collection
        if ($script:PlanSiteAdmin) {
            try { Set-PnPSiteCollectionAdmin -Owners $script:TargetLogin -ErrorAction Stop; Write-Log "  + Admin de collection ajoute a la cible." }
            catch { Write-Log "  ! Admin de collection : $($_.Exception.Message)" }
            $done++; $progress.Value = $done
        }

        # 4) Purger l'entree perimee
        Write-Log "Suppression de l'entree perimee : Id $($script:StaleUser.Id)"
        Remove-PnPUser -Identity $script:StaleUser.Id -Force -ErrorAction Stop
        Write-Log "  OK - entree perimee supprimee."
        $done++; $progress.Value = $progress.Maximum

        Write-Log "===== TERMINE ====="
        [System.Windows.Forms.MessageBox]::Show("Remediation appliquee. Relancez une recherche pour verifier.", "Termine", "OK", "Information") | Out-Null
        Reset-Plan
    } catch {
        Write-Log "ERREUR execution : $($_.Exception.Message)"
        Write-Log "La cible n'a peut-etre pas ete ré-assurée. Aucune entree perimee supprimee si l'erreur est survenue avant l'etape de purge."
        [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)", "Execution", "OK", "Error") | Out-Null
    }
    $btnAnalyze.Enabled = $true
})

# ==================================================================
# Adopter une connexion PnP deja active dans la session, le cas echeant
# ==================================================================
$adopted = $false
if (Get-Module -ListAvailable -Name PnP.PowerShell) {
    try {
        if (Test-PnPConnected) {
            $url = $null
            try { $url = (Get-PnPConnection -ErrorAction Stop).Url } catch { }
            $w = Get-PnPWeb -ErrorAction SilentlyContinue
            if (-not $url -and $w) { $url = $w.Url }
            if ($url) { $txtSite.Text = $url }
            $title = if ($w) { $w.Title } else { $url }
            Set-ConnState $true "Connexion active reutilisee : $title"
            Write-Log "Connexion PnP active detectee dans la session : $url"
            $adopted = $true
        }
    } catch { }
}
if (-not $adopted) {
    Set-ConnState $false "Aucune connexion active. Connectez-vous avant, ou via le bouton Connexion (URL + Client ID)."
}

[void]$form.ShowDialog()
