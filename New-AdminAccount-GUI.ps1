<#
.SYNOPSIS
    Creation d'un compte administrateur "A-" dans Active Directory - Interface graphique.

.DESCRIPTION
    Automatise en une seule operation la creation d'un compte admin selon une convention "A-" :
      - Nom d'objet (CN)      : <Prefixe><NOM> <Prenom>        ex. A-DUPONT Jean
      - sAMAccountName        : <Prefixe><NOM>                 ex. A-DUPONT
      - UserPrincipalName     : <sAMAccountName>@<suffixe UPN>
      - Nom d'affichage       : <NOM> <Prenom>
      - Mot de passe          : 18 caracteres, majuscule + minuscule + chiffre
      - Attributs             : extensionAttribute13, employeeType, description, e-mail
    Remplace une procedure manuelle longue (creation, renommage, editeur d'attributs...).

    ANONYMISATION : aucune valeur d'environnement n'est codee en dur. Le domaine (suffixe UPN),
    l'OU cible et le serveur AD sont des champs a renseigner ; ils sont sauvegardes localement
    dans %AppData%\ACreator\config.json (jamais dans ce fichier).

.PREREQUIS
    - Module ActiveDirectory (RSAT) et droits delegues sur l'OU cible.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:CfgPath     = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'ACreator\config.json'
$script:LastSummary = ""

# ==================================================================
# Fonctions
# ==================================================================
function ConvertTo-Ascii {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $norm = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $norm.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function New-CompliantPassword {
    param([int]$Length = 18)
    $U = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; $L = 'abcdefghijkmnpqrstuvwxyz'; $D = '23456789'   # sans caracteres ambigus
    $all = $U + $L + $D
    $chars = @($U[(Get-Random -Maximum $U.Length)], $L[(Get-Random -Maximum $L.Length)], $D[(Get-Random -Maximum $D.Length)])
    for ($i = $chars.Count; $i -lt $Length; $i++) { $chars += $all[(Get-Random -Maximum $all.Length)] }
    return (-join ($chars | Sort-Object { Get-Random }))
}

function Convert-OUPathToDN {
    # Accepte un DN (renvoye tel quel) OU un chemin canonique "domaine.fqdn/OU1/OU2/..." -> DN
    param([string]$Path)
    $p = $Path.Trim().Trim('/')
    if (-not $p) { return "" }
    if ($p -match '(?i)\b(OU|CN|DC)=') { return $p }              # deja un DistinguishedName
    $parts = @($p -split '/')
    $domain = $parts[0]
    $dc = (($domain -split '\.') | ForEach-Object { "DC=$_" }) -join ','
    if ($parts.Count -gt 1) {
        $ouParts = @($parts[1..($parts.Count - 1)])
        [array]::Reverse($ouParts)
        $ou = ($ouParts | ForEach-Object { "OU=$_" }) -join ','
        return "$ou,$dc"
    }
    return $dc
}

# ==================================================================
# Formulaire
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Creation compte administrateur (A-)"
$form.Size = New-Object System.Drawing.Size(810, 810)
$form.MinimumSize = New-Object System.Drawing.Size(720, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ---------- Configuration ----------
$gbCfg = New-Object System.Windows.Forms.GroupBox
$gbCfg.Text = "Configuration (a renseigner une fois - sauvegardee localement)"
$gbCfg.Location = New-Object System.Drawing.Point(15, 10); $gbCfg.Size = New-Object System.Drawing.Size(765, 205); $gbCfg.Anchor = "Top,Left,Right"
$form.Controls.Add($gbCfg)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

Add-Lbl $gbCfg "OU cible (DN) :" 14 26 | Out-Null
$txtOU = New-Object System.Windows.Forms.TextBox; $txtOU.Location=New-Object System.Drawing.Point(140,23); $txtOU.Size=New-Object System.Drawing.Size(610,23); $txtOU.Anchor="Top,Left,Right"; $gbCfg.Controls.Add($txtOU)

Add-Lbl $gbCfg "Suffixe UPN :" 14 56 | Out-Null
$txtSuffix = New-Object System.Windows.Forms.TextBox; $txtSuffix.Location=New-Object System.Drawing.Point(140,53); $txtSuffix.Size=New-Object System.Drawing.Size(180,23); $gbCfg.Controls.Add($txtSuffix)
Add-Lbl $gbCfg "Prefixe :" 335 56 | Out-Null
$txtPrefix = New-Object System.Windows.Forms.TextBox; $txtPrefix.Location=New-Object System.Drawing.Point(410,53); $txtPrefix.Size=New-Object System.Drawing.Size(70,23); $gbCfg.Controls.Add($txtPrefix)
Add-Lbl $gbCfg "Serveur AD :" 500 56 | Out-Null
$txtServer = New-Object System.Windows.Forms.TextBox; $txtServer.Location=New-Object System.Drawing.Point(575,53); $txtServer.Size=New-Object System.Drawing.Size(175,23); $txtServer.Anchor="Top,Right"; $gbCfg.Controls.Add($txtServer)

Add-Lbl $gbCfg "Description :" 14 86 | Out-Null
$txtDesc = New-Object System.Windows.Forms.TextBox; $txtDesc.Location=New-Object System.Drawing.Point(140,83); $txtDesc.Size=New-Object System.Drawing.Size(610,23); $txtDesc.Anchor="Top,Left,Right"; $gbCfg.Controls.Add($txtDesc)

Add-Lbl $gbCfg "extensionAttribute13 :" 14 116 | Out-Null
$txtExt13 = New-Object System.Windows.Forms.TextBox; $txtExt13.Location=New-Object System.Drawing.Point(160,113); $txtExt13.Size=New-Object System.Drawing.Size(100,23); $gbCfg.Controls.Add($txtExt13)
Add-Lbl $gbCfg "employeeType :" 285 116 | Out-Null
$txtEmp = New-Object System.Windows.Forms.TextBox; $txtEmp.Location=New-Object System.Drawing.Point(380,113); $txtEmp.Size=New-Object System.Drawing.Size(100,23); $gbCfg.Controls.Add($txtEmp)

$lblHint = Add-Lbl $gbCfg "Valeurs propres a votre environnement : renseignez-les puis Enregistrer." 14 150
$lblHint.ForeColor = [System.Drawing.Color]::Gray

$btnSaveCfg = New-Object System.Windows.Forms.Button; $btnSaveCfg.Text="Enregistrer la config"; $btnSaveCfg.Location=New-Object System.Drawing.Point(560,145); $btnSaveCfg.Size=New-Object System.Drawing.Size(190,28); $btnSaveCfg.Anchor="Top,Right"; $gbCfg.Controls.Add($btnSaveCfg)

# ---------- Nouvel utilisateur ----------
$gbUser = New-Object System.Windows.Forms.GroupBox
$gbUser.Text = "Nouvel administrateur"
$gbUser.Location = New-Object System.Drawing.Point(15, 225); $gbUser.Size = New-Object System.Drawing.Size(765, 95); $gbUser.Anchor="Top,Left,Right"
$form.Controls.Add($gbUser)

Add-Lbl $gbUser "Prenom :" 14 26 | Out-Null
$txtFirst = New-Object System.Windows.Forms.TextBox; $txtFirst.Location=New-Object System.Drawing.Point(140,23); $txtFirst.Size=New-Object System.Drawing.Size(220,23); $gbUser.Controls.Add($txtFirst)
Add-Lbl $gbUser "Nom :" 390 26 | Out-Null
$txtLast = New-Object System.Windows.Forms.TextBox; $txtLast.Location=New-Object System.Drawing.Point(450,23); $txtLast.Size=New-Object System.Drawing.Size(300,23); $txtLast.Anchor="Top,Left,Right"; $gbUser.Controls.Add($txtLast)

Add-Lbl $gbUser "Mot de passe :" 14 58 | Out-Null
$txtPwd = New-Object System.Windows.Forms.TextBox; $txtPwd.Location=New-Object System.Drawing.Point(140,55); $txtPwd.Size=New-Object System.Drawing.Size(330,23); $gbUser.Controls.Add($txtPwd)
$btnGen = New-Object System.Windows.Forms.Button; $btnGen.Text="Generer"; $btnGen.Location=New-Object System.Drawing.Point(480,54); $btnGen.Size=New-Object System.Drawing.Size(90,25); $gbUser.Controls.Add($btnGen)
$btnCopyPwd = New-Object System.Windows.Forms.Button; $btnCopyPwd.Text="Copier"; $btnCopyPwd.Location=New-Object System.Drawing.Point(575,54); $btnCopyPwd.Size=New-Object System.Drawing.Size(80,25); $gbUser.Controls.Add($btnCopyPwd)
$chkEnabled = New-Object System.Windows.Forms.CheckBox; $chkEnabled.Text="Activer"; $chkEnabled.Location=New-Object System.Drawing.Point(665,56); $chkEnabled.AutoSize=$true; $chkEnabled.Checked=$true; $chkEnabled.Anchor="Top,Right"; $gbUser.Controls.Add($chkEnabled)

# ---------- Valeurs derivees ----------
$gbDer = New-Object System.Windows.Forms.GroupBox
$gbDer.Text = "Valeurs generees (modifiables)"
$gbDer.Location = New-Object System.Drawing.Point(15, 330); $gbDer.Size = New-Object System.Drawing.Size(765, 185); $gbDer.Anchor="Top,Left,Right"
$form.Controls.Add($gbDer)

$btnPreview = New-Object System.Windows.Forms.Button; $btnPreview.Text="Apercu (calculer)"; $btnPreview.Location=New-Object System.Drawing.Point(14,24); $btnPreview.Size=New-Object System.Drawing.Size(170,28); $gbDer.Controls.Add($btnPreview)

Add-Lbl $gbDer "sAMAccountName :" 14 64 | Out-Null
$txtSam = New-Object System.Windows.Forms.TextBox; $txtSam.Location=New-Object System.Drawing.Point(150,61); $txtSam.Size=New-Object System.Drawing.Size(260,23); $gbDer.Controls.Add($txtSam)
Add-Lbl $gbDer "Nom d'affichage :" 430 64 | Out-Null
$txtDisp = New-Object System.Windows.Forms.TextBox; $txtDisp.Location=New-Object System.Drawing.Point(545,61); $txtDisp.Size=New-Object System.Drawing.Size(205,23); $txtDisp.Anchor="Top,Left,Right"; $gbDer.Controls.Add($txtDisp)

Add-Lbl $gbDer "Initiales :" 300 26 | Out-Null
$txtInit = New-Object System.Windows.Forms.TextBox; $txtInit.Location=New-Object System.Drawing.Point(370,23); $txtInit.Size=New-Object System.Drawing.Size(90,23); $gbDer.Controls.Add($txtInit)

Add-Lbl $gbDer "UPN :" 14 94 | Out-Null
$txtUpn = New-Object System.Windows.Forms.TextBox; $txtUpn.Location=New-Object System.Drawing.Point(150,91); $txtUpn.Size=New-Object System.Drawing.Size(600,23); $txtUpn.Anchor="Top,Left,Right"; $gbDer.Controls.Add($txtUpn)
Add-Lbl $gbDer "Nom d'objet (CN) :" 14 124 | Out-Null
$txtCN = New-Object System.Windows.Forms.TextBox; $txtCN.Location=New-Object System.Drawing.Point(150,121); $txtCN.Size=New-Object System.Drawing.Size(600,23); $txtCN.Anchor="Top,Left,Right"; $gbDer.Controls.Add($txtCN)
Add-Lbl $gbDer "E-mail :" 14 154 | Out-Null
$txtMail = New-Object System.Windows.Forms.TextBox; $txtMail.Location=New-Object System.Drawing.Point(150,151); $txtMail.Size=New-Object System.Drawing.Size(600,23); $txtMail.Anchor="Top,Left,Right"; $gbDer.Controls.Add($txtMail)

# ---------- Actions ----------
$btnCreate = New-Object System.Windows.Forms.Button; $btnCreate.Text="Creer le compte administrateur"; $btnCreate.Location=New-Object System.Drawing.Point(15,525); $btnCreate.Size=New-Object System.Drawing.Size(600,34)
$btnCreate.BackColor=[System.Drawing.Color]::FromArgb(0,120,215); $btnCreate.ForeColor=[System.Drawing.Color]::White; $btnCreate.FlatStyle="Flat"; $btnCreate.Anchor="Top,Left,Right"; $form.Controls.Add($btnCreate)
$btnCopySum = New-Object System.Windows.Forms.Button; $btnCopySum.Text="Copier le recap"; $btnCopySum.Location=New-Object System.Drawing.Point(625,525); $btnCopySum.Size=New-Object System.Drawing.Size(155,34); $btnCopySum.Anchor="Top,Right"; $btnCopySum.Enabled=$false; $form.Controls.Add($btnCopySum)

Add-Lbl $form "Journal :" 15 568 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,588); $txtLog.Size=New-Object System.Drawing.Size(765,175); $txtLog.Anchor="Top,Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# Helpers
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)) }

function Update-Derived {
    $first = $txtFirst.Text.Trim(); $last = $txtLast.Text.Trim()
    if (-not $first -or -not $last) { return }
    $prefix = $txtPrefix.Text.Trim(); if (-not $prefix) { $prefix = "A-" }
    $suffix = $txtSuffix.Text.Trim().TrimStart('@')
    $lastUpper = $last.ToUpper()
    $lastAscii = ((ConvertTo-Ascii $last).ToUpper() -replace "[^A-Z0-9-]", "")
    $sam = "$prefix$lastAscii"
    $txtSam.Text  = $sam
    $txtUpn.Text  = if ($suffix) { "$sam@$suffix" } else { $sam }
    $txtCN.Text   = "$prefix$lastUpper $first"
    $txtDisp.Text = "$lastUpper $first"
    $txtMail.Text = $txtUpn.Text
    # Initiales = 1re lettre du prenom + 2 premieres lettres du nom (ex. Ferrero NUTELLA -> FNU)
    $fPart = if ($first.Length -ge 1) { $first.Substring(0,1) } else { "" }
    $lPart = if ($lastUpper.Length -ge 2) { $lastUpper.Substring(0,2) } elseif ($lastUpper.Length -eq 1) { $lastUpper } else { "" }
    $txtInit.Text = ($fPart.ToUpper() + $lPart)
    if ($sam.Length -gt 20) { Write-Log "ATTENTION : sAMAccountName '$sam' depasse 20 caracteres (limite AD). Ajustez-le." }
}

function Load-Cfg {
    if (Test-Path $script:CfgPath) { try { return (Get-Content $script:CfgPath -Raw -ErrorAction Stop | ConvertFrom-Json) } catch { } }
    return $null
}

# ==================================================================
# Evenements
# ==================================================================
$btnSaveCfg.Add_Click({
    try {
        $dir = Split-Path $script:CfgPath; if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [PSCustomObject]@{
            OU=$txtOU.Text; Suffix=$txtSuffix.Text; Prefix=$txtPrefix.Text; Server=$txtServer.Text
            Description=$txtDesc.Text; Ext13=$txtExt13.Text; EmployeeType=$txtEmp.Text
        } | ConvertTo-Json | Set-Content -Path $script:CfgPath -Encoding UTF8
        Write-Log "Configuration enregistree."
    } catch { Write-Log "ERREUR enregistrement config : $($_.Exception.Message)" }
})

$btnGen.Add_Click({ $txtPwd.Text = New-CompliantPassword 18; Write-Log "Mot de passe genere (18 caracteres)." })
$btnCopyPwd.Add_Click({ if ($txtPwd.Text) { [System.Windows.Forms.Clipboard]::SetText($txtPwd.Text); Write-Log "Mot de passe copie." } })
$btnPreview.Add_Click({ Update-Derived })
$txtLast.Add_Leave({ Update-Derived })
$txtFirst.Add_Leave({ Update-Derived })

$btnCopySum.Add_Click({ if ($script:LastSummary) { [System.Windows.Forms.Clipboard]::SetText($script:LastSummary); Write-Log "Recap copie." } })

$btnCreate.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Module ActiveDirectory (RSAT) introuvable.", "Prerequis", "OK", "Error") | Out-Null; return
    }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    if (-not $txtSam.Text.Trim() -or -not $txtCN.Text.Trim()) { Update-Derived }
    $first = $txtFirst.Text.Trim(); $last = $txtLast.Text.Trim()
    $sam = $txtSam.Text.Trim(); $upn = $txtUpn.Text.Trim(); $cn = $txtCN.Text.Trim()
    $disp = $txtDisp.Text.Trim(); $mail = $txtMail.Text.Trim(); $pwd = $txtPwd.Text
    $ouRaw = $txtOU.Text.Trim(); $ou = Convert-OUPathToDN $ouRaw

    if (-not $first -or -not $last) { [System.Windows.Forms.MessageBox]::Show("Prenom et Nom obligatoires.", "Validation", "OK", "Warning") | Out-Null; return }
    if (-not $ou)  { [System.Windows.Forms.MessageBox]::Show("OU cible obligatoire (Configuration).", "Validation", "OK", "Warning") | Out-Null; return }
    if (-not $sam) { [System.Windows.Forms.MessageBox]::Show("sAMAccountName vide - cliquez Apercu.", "Validation", "OK", "Warning") | Out-Null; return }
    if ($sam.Length -gt 20) { [System.Windows.Forms.MessageBox]::Show("sAMAccountName > 20 caracteres. Ajustez-le.", "Validation", "OK", "Warning") | Out-Null; return }
    if ($pwd.Length -lt 18 -or $pwd -notmatch '[A-Z]' -or $pwd -notmatch '[a-z]' -or $pwd -notmatch '[0-9]') {
        [System.Windows.Forms.MessageBox]::Show("Mot de passe non conforme (18+ car., majuscule, minuscule, chiffre).", "Validation", "OK", "Warning") | Out-Null; return
    }

    $srv = $txtServer.Text.Trim()
    $srvParam = @{}; if ($srv) { $srvParam['Server'] = $srv }

    try {
        $exists = Get-ADUser -Filter "SamAccountName -eq '$sam'" @srvParam -ErrorAction SilentlyContinue
        if ($exists) { [System.Windows.Forms.MessageBox]::Show("Un compte '$sam' existe deja.", "Doublon", "OK", "Warning") | Out-Null; return }
    } catch { }

    $recap = "Nom d'objet : $cn`nsAMAccount : $sam`nUPN : $upn`nActiver : $($chkEnabled.Checked)`n`nCreer ce compte dans (DN) :`n$ou ?"
    if ([System.Windows.Forms.MessageBox]::Show($recap, "Confirmation", "YesNo", "Question") -ne "Yes") { return }

    $other = @{}
    if ($txtExt13.Text.Trim()) { $other['extensionAttribute13'] = $txtExt13.Text.Trim() }
    if ($txtEmp.Text.Trim())   { $other['employeeType']         = $txtEmp.Text.Trim() }

    $params = @{
        Name                  = $cn
        GivenName             = $first
        Surname               = $last.ToUpper()
        DisplayName           = $disp
        SamAccountName        = $sam
        UserPrincipalName     = $upn
        Path                  = $ou
        AccountPassword       = (ConvertTo-SecureString $pwd -AsPlainText -Force)
        Enabled               = $chkEnabled.Checked
        ChangePasswordAtLogon = $false
    }
    if ($txtInit.Text.Trim()) { $params['Initials'] = $txtInit.Text.Trim() }
    if ($txtDesc.Text.Trim()) { $params['Description'] = $txtDesc.Text.Trim() }
    if ($mail)                { $params['EmailAddress'] = $mail }
    if ($other.Count)         { $params['OtherAttributes'] = $other }
    if ($srv)                 { $params['Server'] = $srv }

    try {
        New-ADUser @params -ErrorAction Stop
        Write-Log "COMPTE CREE : $sam ($cn)"
        Write-Log "  OU (DN) : $ou"
        foreach ($k in $other.Keys) { Write-Log "  $k = $($other[$k])" }
        $script:LastSummary = "Compte      : $sam`nUPN         : $upn`nE-mail      : $mail`nMot de passe: $pwd`nDescription : $($txtDesc.Text.Trim())"
        $btnCopySum.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Compte cree.`n$sam`nPensez a conserver le mot de passe (bouton 'Copier le recap').`nDisponibilite : quelques dizaines de minutes, ~1 h cote cloud.", "Termine", "OK", "Information") | Out-Null
    } catch {
        Write-Log "ERREUR creation : $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)", "Creation", "OK", "Error") | Out-Null
    }
})

# ==================================================================
# Init : valeurs par defaut neutres + chargement config locale
# ==================================================================
$txtPrefix.Text = "A-"; $txtDesc.Text = "Admin Account for Azure AD"; $txtExt13.Text = "SYNC"; $txtEmp.Text = "YES"
$cfg = Load-Cfg
if ($cfg) {
    if ($cfg.OU)           { $txtOU.Text     = $cfg.OU }
    if ($cfg.Suffix)       { $txtSuffix.Text = $cfg.Suffix }
    if ($cfg.Prefix)       { $txtPrefix.Text = $cfg.Prefix }
    if ($cfg.Server)       { $txtServer.Text = $cfg.Server }
    if ($cfg.Description)  { $txtDesc.Text   = $cfg.Description }
    if ($cfg.Ext13)        { $txtExt13.Text  = $cfg.Ext13 }
    if ($cfg.EmployeeType) { $txtEmp.Text    = $cfg.EmployeeType }
    Write-Log "Configuration locale chargee."
} else {
    Write-Log "Aucune config : renseignez OU cible + suffixe UPN, puis Enregistrer."
}
$txtPwd.Text = New-CompliantPassword 18

[void]$form.ShowDialog()
