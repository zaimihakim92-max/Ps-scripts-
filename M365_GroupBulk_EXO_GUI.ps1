#Requires -Version 5.1
<#
.SYNOPSIS
    Ajout / retrait en masse d'utilisateurs dans un groupe Microsoft 365 via Exchange Online.

.DESCRIPTION
    Interface graphique WinForms. A lancer depuis une console PowerShell ou la connexion Exchange
    Online est deja etablie (Connect-ExchangeOnline effectue, module ExchangeOnlineManagement).

    Cette version gere le groupe et resout les comptes UNIQUEMENT via Exchange Online :
      - groupe            : Get-UnifiedGroup (nom, alias, e-mail ou GUID) ;
      - membres           : Get-UnifiedGroupLinks -LinkType Members ;
      - ajout / retrait    : Add-UnifiedGroupLinks / Remove-UnifiedGroupLinks ;
      - resolution comptes : Get-Recipient (trouve nativement les invites #EXT# par e-mail externe).

    Elle ne depend pas de Microsoft Graph / Az : utile quand Get-AzADGroup ne renvoie rien faute
    de droits annuaire, alors que la gestion Exchange des groupes M365 fonctionne.

    Garde-fous (identiques a la version SCCM/Az) :
      - contexte Exchange detecte, jamais reconnecte a l'aveugle ;
      - groupe cible valide obligatoirement avant toute action ;
      - toute modification de la cible re-arme les verrous ;
      - recherche prealable obligatoire avant execution ;
      - entrees sanitisees (aucun caractere generique) ;
      - resolution stricte : entree ambigue IGNOREE ;
      - confirmation detaillee avant ajout ; saisie de l'ObjectId avant retrait ;
      - journalisation horodatee (GUI + console + fichier) et export CSV.

.NOTES
    Lancer de preference via :  powershell.exe -STA -File .\M365_GroupBulk_EXO_GUI.ps1
    depuis la meme console ou Connect-ExchangeOnline a ete lance.
#>

# ============================================================================
#  0. Pre-requis
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "Le thread courant n'est pas en mode STA. L'interface peut mal se comporter."
    Write-Warning "Relancez avec :  powershell.exe -STA -File .\M365_GroupBulk_EXO_GUI.ps1"
}

# ============================================================================
#  1. Etat global
# ============================================================================
$script:ExoUser        = $null
$script:ExoOrg         = $null
$script:ValidatedGroup = $null      # objet : Id(GUID) / Name / Smtp / Type
$script:SimData        = $null
$script:SimulationDone = $false
$script:IsProcessing   = $false
$script:CancelRequested= $false
$script:LogFile        = $null
$script:BaseTitle      = $null

$script:WILDCARD_CHARS = @('*','?','%','[',']')

# ============================================================================
#  2. Journalisation (GUI + console + fichier)
# ============================================================================
function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO','OK','WARN','ERR','RECH')] [string] $Level = 'INFO'
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Message"

    if ($script:txtLog) {
        $script:txtLog.AppendText($line + "`r`n")
        $script:txtLog.SelectionStart = $script:txtLog.Text.Length
        $script:txtLog.ScrollToCaret()
    }

    $color = switch ($Level) { 'OK'{'Green'} 'WARN'{'Yellow'} 'ERR'{'Red'} 'RECH'{'Cyan'} default{'Gray'} }
    Write-Host $line -ForegroundColor $color

    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
    }
}

function Initialize-LogFile {
    if ($script:LogFile) { return }
    $folder = $script:txtLogFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        $folder = Join-Path ([Environment]::GetFolderPath('Desktop')) 'M365_GroupBulk_Logs'
    }
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $script:LogFile = Join-Path $folder ("M365_GroupBulk_{0}.log" -f $stamp)
    Write-Log "Journal : $($script:LogFile)" 'INFO'
    Write-Log "Operateur : $env:USERNAME  /  Compte Exchange : $($script:ExoUser)" 'INFO'
}

# ============================================================================
#  3. Contexte Exchange Online
# ============================================================================
function Initialize-ExoContext {
    try { $ci = @(Get-ConnectionInformation -ErrorAction Stop) } catch { return $false }
    if (-not $ci -or $ci.Count -eq 0) { return $false }
    $one = $ci | Select-Object -First 1
    $script:ExoUser = $one.UserPrincipalName
    $script:ExoOrg  = $one.Organization
    return $true
}

# ============================================================================
#  4. Validation / sanitisation des entrees
# ============================================================================
function Test-UpnSafe {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    foreach ($c in $script:WILDCARD_CHARS) { if ($Name.Contains($c)) { return $false } }
    return ($Name -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Get-InputUserNames {
    $raw = $script:txtUsers.Text
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $parts = $raw -split "[\r\n,; `t]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

    $seen   = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($p in $parts) {
        $key = $p.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $result.Add($p) }
    }
    return $result
}

# ============================================================================
#  5. Etat des controles (verrous)
# ============================================================================
function Reset-Simulation {
    $script:SimData        = $null
    $script:SimulationDone = $false
    if ($script:grid) { $script:grid.Rows.Clear() }
    Update-ButtonsState
}

function Invalidate-Group {
    $script:ValidatedGroup = $null
    if ($script:lblGroupInfo) {
        $script:lblGroupInfo.Text      = "Groupe non valide."
        $script:lblGroupInfo.ForeColor = [System.Drawing.Color]::DimGray
    }
    Reset-Simulation
}

function Update-ButtonsState {
    $hasGrp   = ($null -ne $script:ValidatedGroup)
    $hasUsers = ((Get-InputUserNames).Count -gt 0)
    $busy     = $script:IsProcessing

    if ($script:btnValidate)    { $script:btnValidate.Enabled    = (-not $busy) }
    if ($script:btnListMembers) { $script:btnListMembers.Enabled = ($hasGrp -and -not $busy) }
    if ($script:btnSimulate)    { $script:btnSimulate.Enabled    = ($hasGrp -and $hasUsers -and -not $busy) }
    if ($script:btnAdd)         { $script:btnAdd.Enabled         = ($hasGrp -and $script:SimulationDone -and -not $busy) }
    if ($script:btnRemove)      { $script:btnRemove.Enabled      = ($hasGrp -and $script:SimulationDone -and -not $busy) }
    if ($script:btnStop)        { $script:btnStop.Enabled        = $busy }
    if ($script:txtGroup)       { $script:txtGroup.Enabled       = (-not $busy) }
    if ($script:txtUsers)       { $script:txtUsers.Enabled       = (-not $busy) }
    if ($script:btnLoadTxt)     { $script:btnLoadTxt.Enabled     = (-not $busy) }
    if ($script:btnClear)       { $script:btnClear.Enabled       = (-not $busy) }
}

function Set-ProcessingState {
    param([bool] $On)
    $script:IsProcessing = $On
    $script:MainForm.Cursor = if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    if ($On) {
        if (-not $script:BaseTitle) { $script:BaseTitle = $script:MainForm.Text }
        $script:MainForm.Text = "*** TRAITEMENT EN COURS - NE PAS FERMER LA FENETRE ***"
    }
    elseif ($script:BaseTitle) {
        $script:MainForm.Text = $script:BaseTitle
    }
    $script:MainForm.Refresh()
    Update-ButtonsState
}

# ============================================================================
#  6. Validation du groupe
# ============================================================================
function Validate-Group {
    $target = $script:txtGroup.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($target)) {
        [System.Windows.Forms.MessageBox]::Show("Saisissez un nom, un e-mail ou un GUID de groupe.","Validation",'OK','Warning') | Out-Null
        return
    }
    foreach ($c in $script:WILDCARD_CHARS) {
        if ($target.Contains($c)) {
            [System.Windows.Forms.MessageBox]::Show("Les caracteres generiques ( * ? % [ ] ) sont interdits dans la cible.","Securite",'OK','Error') | Out-Null
            return
        }
    }

    Invalidate-Group
    Write-Log "Validation du groupe '$target'..." 'INFO'

    $isGuid = ($target -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')
    $grp = $null
    try {
        if ($isGuid -or $target.Contains('@')) {
            $grp = Get-UnifiedGroup -Identity $target -ErrorAction SilentlyContinue
        } else {
            $esc = $target.Replace("'","''")
            $matches = @(Get-UnifiedGroup -Filter "DisplayName -eq '$esc'" -ErrorAction SilentlyContinue)
            if ($matches.Count -gt 1) {
                $ids = ($matches | ForEach-Object { $_.PrimarySmtpAddress }) -join ', '
                Write-Log "Nom ambigu ($($matches.Count) groupes) : $ids. Utilisez l'e-mail ou le GUID." 'ERR'
                $script:lblGroupInfo.Text      = "Nom ambigu ($($matches.Count) groupes). Precisez : $ids"
                $script:lblGroupInfo.ForeColor = [System.Drawing.Color]::Firebrick
                return
            }
            $grp = $matches | Select-Object -First 1
        }
    }
    catch {
        Write-Log "Erreur lors de la resolution : $($_.Exception.Message)" 'ERR'
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la resolution du groupe :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        return
    }

    if (-not $grp) {
        Write-Log "Aucun groupe ne correspond a '$target'." 'ERR'
        $script:lblGroupInfo.Text      = "Aucun groupe trouve."
        $script:lblGroupInfo.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }

    $gid   = [string]$grp.ExternalDirectoryObjectId
    $gname = [string]$grp.DisplayName
    $gsmtp = [string]$grp.PrimarySmtpAddress
    $gtype = ''
    if ($grp.PSObject.Properties['GroupType'] -and $grp.GroupType) { $gtype = ([string[]]$grp.GroupType) -join ',' }

    if ([string]::IsNullOrWhiteSpace($gid)) {
        Write-Log "Le groupe n'expose pas d'ExternalDirectoryObjectId. Impossible de securiser l'operation." 'ERR'
        $script:lblGroupInfo.Text      = "Groupe sans ObjectId exploitable."
        $script:lblGroupInfo.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }

    $script:ValidatedGroup = [pscustomobject]@{
        Id   = $gid
        Name = $gname
        Smtp = $gsmtp
        Type = $gtype
    }

    $info = "VALIDE  |  Nom : $gname  |  E-mail : $gsmtp  |  ObjectId : $gid"
    $script:lblGroupInfo.Text      = $info
    $script:lblGroupInfo.ForeColor = [System.Drawing.Color]::DarkGreen
    Write-Log $info 'OK'
    if ($gtype -match 'Dynamic') {
        Write-Log "Attention : groupe a appartenance DYNAMIQUE. L'ajout/retrait manuel echouera (gere par regle)." 'WARN'
    }

    $script:SimulationDone = $false
    Update-ButtonsState
}

# ============================================================================
#  7. Recherche (aucune modification)
# ============================================================================
function Invoke-Simulation {
    if (-not $script:ValidatedGroup) { return }
    $names = Get-InputUserNames
    if ($names.Count -eq 0) { return }

    if ($names.Count -gt 5000) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Vous avez saisi $($names.Count) entrees. Continuer la recherche ?",
            "Volume important",'YesNo','Warning')
        if ($r -ne 'Yes') { return }
    }

    Initialize-LogFile
    Set-ProcessingState $true
    $script:CancelRequested = $false
    $script:grid.Rows.Clear()
    Write-Log "----- RECHERCHE (aucune modification) -----" 'RECH'
    Write-Log "La fenetre peut afficher (Ne repond pas) pendant le traitement : c'est normal, ne pas forcer la fermeture." 'WARN'
    Write-Log "Cible : $($script:ValidatedGroup.Name) ($($script:ValidatedGroup.Id))  |  $($names.Count) entree(s) unique(s)." 'RECH'

    $data = New-Object System.Collections.Generic.List[object]
    $gid  = $script:ValidatedGroup.Id

    try {
        # Ensembles frais des membres actuels (par GUID et par e-mail).
        $memberGuid = New-Object 'System.Collections.Generic.HashSet[string]'
        $memberSmtp = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($m in @(Get-UnifiedGroupLinks -Identity $gid -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue)) {
            if ($m.ExternalDirectoryObjectId) { [void]$memberGuid.Add(([string]$m.ExternalDirectoryObjectId).ToLowerInvariant()) }
            if ($m.PrimarySmtpAddress)        { [void]$memberSmtp.Add(([string]$m.PrimarySmtpAddress).ToLowerInvariant()) }
        }
        Write-Log "Membres actuels du groupe : $($memberGuid.Count)." 'RECH'

        $script:progress.Minimum = 0
        $script:progress.Maximum = $names.Count
        $script:progress.Value   = 0
        $i = 0

        foreach ($name in $names) {
            if ($script:CancelRequested) { Write-Log "Recherche interrompue par l'utilisateur." 'WARN'; break }
            $i++
            $script:progress.Value = $i

            if (-not (Test-UpnSafe $name)) {
                $data.Add([pscustomobject]@{ Name=$name; Category='INVALIDE'; ObjectId=$null; DisplayName=''; Smtp=''; Detail='Format e-mail invalide' })
                Write-Log "INVALIDE : '$name'" 'WARN'
                Add-GridRow $name 'INVALIDE' '' '' 'Format invalide'
                [System.Windows.Forms.Application]::DoEvents(); continue
            }

            $rcpts = @(Get-Recipient -Identity $name -ErrorAction SilentlyContinue)
            if ($rcpts.Count -eq 0) {
                $data.Add([pscustomobject]@{ Name=$name; Category='NON_TROUVE'; ObjectId=$null; DisplayName=''; Smtp=''; Detail='Aucun destinataire' })
                Write-Log "NON TROUVE : '$name'" 'WARN'
                Add-GridRow $name 'NON_TROUVE' '' '' 'Aucun destinataire'
            }
            elseif ($rcpts.Count -gt 1) {
                $ids = ($rcpts | ForEach-Object { $_.PrimarySmtpAddress }) -join ','
                $data.Add([pscustomobject]@{ Name=$name; Category='AMBIGU'; ObjectId=$null; DisplayName=''; Smtp=''; Detail="Multiples : $ids" })
                Write-Log "AMBIGU : '$name' -> $($rcpts.Count) destinataires ($ids). IGNORE." 'WARN'
                Add-GridRow $name 'AMBIGU' '' '' "Multiples : $ids"
            }
            else {
                $r    = $rcpts[0]
                $uid  = [string]$r.ExternalDirectoryObjectId
                $dn   = [string]$r.DisplayName
                $smtp = [string]$r.PrimarySmtpAddress
                $isMember = ($uid -and $memberGuid.Contains($uid.ToLowerInvariant())) -or `
                            ($smtp -and $memberSmtp.Contains($smtp.ToLowerInvariant()))
                if ($isMember) {
                    $data.Add([pscustomobject]@{ Name=$name; Category='MEMBRE'; ObjectId=$uid; DisplayName=$dn; Smtp=$smtp; Detail='Deja membre' })
                    Add-GridRow $name 'MEMBRE' $uid $dn 'Deja membre'
                } else {
                    $data.Add([pscustomobject]@{ Name=$name; Category='HORS_GROUPE'; ObjectId=$uid; DisplayName=$dn; Smtp=$smtp; Detail='Non membre' })
                    Add-GridRow $name 'HORS_GROUPE' $uid $dn 'Non membre'
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        Write-Log "Erreur de recherche : $($_.Exception.Message)" 'ERR'
    }
    finally {
        $script:progress.Value = 0
    }

    $script:SimData = $data
    $addC = @($data | Where-Object Category -eq 'HORS_GROUPE').Count
    $memC = @($data | Where-Object Category -eq 'MEMBRE').Count
    $nfC  = @($data | Where-Object Category -eq 'NON_TROUVE').Count
    $amC  = @($data | Where-Object Category -eq 'AMBIGU').Count
    $ivC  = @($data | Where-Object Category -eq 'INVALIDE').Count

    Write-Log "Bilan recherche -> a ajouter : $addC | deja membres : $memC | non trouves : $nfC | ambigus : $amC | invalides : $ivC" 'RECH'
    $script:SimulationDone = $true
    Set-ProcessingState $false
}

function Add-GridRow {
    param([string]$Name,[string]$Category,[string]$ObjectId,[string]$DisplayName,[string]$Detail)
    $idx = $script:grid.Rows.Add($Name, $Category, $ObjectId, $DisplayName, $Detail)
    $row = $script:grid.Rows[$idx]
    switch ($Category) {
        'HORS_GROUPE' { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220,245,220) }
        'MEMBRE'      { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220,235,250) }
        'NON_TROUVE'  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,225,225) }
        'AMBIGU'      { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(252,235,200) }
        'INVALIDE'    { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240,220,240) }
    }
}

# ============================================================================
#  8. Confirmation par saisie (pour le retrait)
# ============================================================================
function Confirm-ByTyping {
    param([string]$Expected,[string]$Prompt)

    $margin = 18
    $width  = 520

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Confirmation requise"
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.AutoScaleMode = 'Font'
    $f.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $true
    $l.MaximumSize = New-Object System.Drawing.Size($width, 0)
    $l.Location = New-Object System.Drawing.Point($margin, $margin)
    $l.Text = $Prompt

    $lHint = New-Object System.Windows.Forms.Label
    $lHint.AutoSize = $true
    $lHint.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    $lHint.Text = "Recopiez l'identifiant ci-dessus pour deverrouiller :"

    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $width
    $t.Font = New-Object System.Drawing.Font('Consolas',10)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Confirmer"; $ok.Size = New-Object System.Drawing.Size(110,32)
    $ok.DialogResult = 'OK'; $ok.Enabled = $false
    $ko = New-Object System.Windows.Forms.Button
    $ko.Text = "Annuler"; $ko.Size = New-Object System.Drawing.Size(110,32); $ko.DialogResult = 'Cancel'

    $f.Controls.AddRange(@($l,$lHint,$t,$ok,$ko))
    $f.PerformLayout()

    $y = $l.Bottom + 14
    $lHint.Location = New-Object System.Drawing.Point($margin, $y)
    $y = $lHint.Bottom + 4
    $t.Location = New-Object System.Drawing.Point($margin, $y)
    $y = $t.Bottom + 16
    $ok.Location = New-Object System.Drawing.Point(($margin + $width - $ok.Width - $ko.Width - 10), $y)
    $ko.Location = New-Object System.Drawing.Point(($margin + $width - $ko.Width), $y)

    $f.ClientSize = New-Object System.Drawing.Size(($width + 2*$margin), ($ko.Bottom + $margin))

    $t.Add_TextChanged({ $ok.Enabled = ($t.Text.Trim() -ceq $Expected) })
    $f.Add_Shown({ $t.Focus() })
    $f.AcceptButton = $ok; $f.CancelButton = $ko
    return ($f.ShowDialog($script:MainForm) -eq 'OK')
}

# ============================================================================
#  8bis. Lister les membres d'un groupe (lecture + extraction CSV + retrait)
# ============================================================================
function Show-GroupMembers {
    if (-not $script:ValidatedGroup) { return }
    $gid  = $script:ValidatedGroup.Id
    $name = $script:ValidatedGroup.Name

    if ($script:btnListMembers) { $script:btnListMembers.Enabled = $false }

    $rows      = New-Object System.Collections.Generic.List[object]
    $extraCols = New-Object System.Collections.Generic.List[object]

    $loadData = {
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Write-Log "Lecture des membres de '$name' ($gid)..." 'INFO'
        $rows.Clear()
        $members = @()
        try {
            $members = @(Get-UnifiedGroupLinks -Identity $gid -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue)
        }
        catch { Write-Log "Erreur lecture des membres : $($_.Exception.Message)" 'ERR' }
        finally { $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default }

        foreach ($m in $members) {
            $rows.Add([pscustomobject]@{
                Member      = $m
                DisplayName = [string]$m.DisplayName
                Email       = [string]$m.PrimarySmtpAddress
                ObjectId    = [string]$m.ExternalDirectoryObjectId
                Type        = [string]$m.RecipientType
            })
        }
        Write-Log "Membres charges : $($rows.Count)." 'OK'
    }

    & $loadData

    $w = New-Object System.Windows.Forms.Form
    $w.Text = "Membres de $name ($gid)"
    $w.ClientSize = New-Object System.Drawing.Size(820,560)
    $w.MinimumSize = New-Object System.Drawing.Size(680,440)
    $w.StartPosition = 'CenterParent'
    $w.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.SetBounds(12,10,640,20); $lblInfo.Anchor = 'Top,Left,Right'

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtre (nom/e-mail) :"; $lblFilter.SetBounds(12,36,120,22)
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.SetBounds(135,33,673,24); $txtFilter.Anchor = 'Top,Left,Right'

    $g = New-Object System.Windows.Forms.DataGridView
    $g.SetBounds(12,66,796,444); $g.Anchor = 'Top,Left,Right,Bottom'
    $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $true; $g.AutoSizeColumnsMode = 'Fill'

    $btnRemoveSel = New-Object System.Windows.Forms.Button
    $btnRemoveSel.Text = "Retirer la selection"; $btnRemoveSel.SetBounds(12,518,210,32); $btnRemoveSel.Anchor = 'Bottom,Left'
    $btnRemoveSel.BackColor = [System.Drawing.Color]::FromArgb(245,215,215)
    $btnAddCol = New-Object System.Windows.Forms.Button
    $btnAddCol.Text = "Ajouter une colonne..."; $btnAddCol.SetBounds(430,518,160,32); $btnAddCol.Anchor = 'Bottom,Right'
    $btnCsv = New-Object System.Windows.Forms.Button
    $btnCsv.Text = "Extraire CSV"; $btnCsv.SetBounds(598,518,120,32); $btnCsv.Anchor = 'Bottom,Right'
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"; $btnClose.SetBounds(726,518,82,32); $btnClose.Anchor = 'Bottom,Right'; $btnClose.DialogResult = 'Cancel'

    $w.Controls.AddRange(@($lblInfo,$lblFilter,$txtFilter,$g,$btnRemoveSel,$btnAddCol,$btnCsv,$btnClose))

    $rebuildColumns = {
        $g.Columns.Clear()
        [void]$g.Columns.Add('m1','Nom affiche')
        [void]$g.Columns.Add('m2','E-mail')
        [void]$g.Columns.Add('m3','ObjectId')
        [void]$g.Columns.Add('m4','Type')
        foreach ($ec in $extraCols) { [void]$g.Columns.Add(('x_' + $ec.Prop), $ec.Header) }
    }

    $render = {
        param($flt)
        $g.Rows.Clear()
        $shown = 0
        foreach ($r in $rows) {
            if ($flt) {
                $f2 = [regex]::Escape($flt)
                if (($r.DisplayName -notmatch $f2) -and ($r.Email -notmatch $f2)) { continue }
            }
            $vals = New-Object System.Collections.Generic.List[object]
            $vals.Add($r.DisplayName); $vals.Add($r.Email); $vals.Add($r.ObjectId); $vals.Add($r.Type)
            foreach ($ec in $extraCols) {
                $v = $null
                try { $v = $r.Member.$($ec.Prop) } catch { }
                if ($v -is [array]) { $v = ($v -join '; ') }
                $vals.Add([string]$v)
            }
            [void]$g.Rows.Add($vals.ToArray())
            $shown++
        }
        $lblInfo.Text = "$($rows.Count) membre(s)  -  affiche(s) : $shown"
    }

    & $rebuildColumns
    & $render $null
    $txtFilter.Add_TextChanged({ & $render $txtFilter.Text.Trim() })

    $btnAddCol.Add_Click({
        if ($rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun membre : rien a lister comme propriete.","Colonne",'OK','Information') | Out-Null
            return
        }
        $already = @('DisplayName','PrimarySmtpAddress','ExternalDirectoryObjectId','RecipientType') + ($extraCols | ForEach-Object { $_.Prop })
        $props = $rows[0].Member.PSObject.Properties.Name | Sort-Object -Unique | Where-Object { $_ -notin $already }
        if (-not $props) {
            [System.Windows.Forms.MessageBox]::Show("Aucune autre propriete disponible.","Colonne",'OK','Information') | Out-Null
            return
        }
        $d = New-Object System.Windows.Forms.Form
        $d.Text = "Ajouter une colonne"; $d.ClientSize = New-Object System.Drawing.Size(360,120)
        $d.FormBorderStyle = 'FixedDialog'; $d.StartPosition = 'CenterParent'; $d.MaximizeBox=$false; $d.MinimizeBox=$false
        $dl = New-Object System.Windows.Forms.Label
        $dl.Text = "Propriete a afficher :"; $dl.SetBounds(15,15,330,20)
        $cb = New-Object System.Windows.Forms.ComboBox
        $cb.SetBounds(15,40,330,24); $cb.DropDownStyle = 'DropDownList'
        [void]$cb.Items.AddRange(@($props)); $cb.SelectedIndex = 0
        $dok = New-Object System.Windows.Forms.Button
        $dok.Text = "Ajouter"; $dok.SetBounds(160,80,90,28); $dok.DialogResult='OK'
        $dko = New-Object System.Windows.Forms.Button
        $dko.Text = "Annuler"; $dko.SetBounds(255,80,90,28); $dko.DialogResult='Cancel'
        $d.Controls.AddRange(@($dl,$cb,$dok,$dko)); $d.AcceptButton=$dok; $d.CancelButton=$dko
        if ($d.ShowDialog($w) -eq 'OK' -and $cb.SelectedItem) {
            $p = [string]$cb.SelectedItem
            $extraCols.Add(@{ Prop = $p; Header = $p })
            & $rebuildColumns
            & $render $txtFilter.Text.Trim()
            Write-Log "Colonne ajoutee dans la liste des membres : $p" 'INFO'
        }
    })

    $btnCsv.Add_Click({
        if ($rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun membre a extraire.","Extraction",'OK','Information') | Out-Null
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV (*.csv)|*.csv"
        $sfd.FileName = ("Membres_{0}_{1}.csv" -f $gid, (Get-Date).ToString('yyyyMMdd_HHmmss'))
        if ($sfd.ShowDialog() -eq 'OK') {
            try {
                $out = foreach ($r in $rows) {
                    $o = [ordered]@{ 'Nom affiche'=$r.DisplayName; 'E-mail'=$r.Email; ObjectId=$r.ObjectId; Type=$r.Type }
                    foreach ($ec in $extraCols) {
                        $v = $null
                        try { $v = $r.Member.$($ec.Prop) } catch { }
                        if ($v -is [array]) { $v = ($v -join '; ') }
                        $o[$ec.Header] = [string]$v
                    }
                    [pscustomobject]$o
                }
                $out | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
                Write-Log "Extraction des membres : $($sfd.FileName)" 'OK'
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Extraction impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
            }
        }
    })

    $btnRemoveSel.Add_Click({
        $sel = @($g.SelectedRows)
        if ($sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Selectionnez au moins une ligne.","Retrait",'OK','Information') | Out-Null
            return
        }
        $toRemove = New-Object System.Collections.Generic.List[string]
        foreach ($rowSel in $sel) {
            $oid = [string]$rowSel.Cells['m3'].Value
            if (-not $oid) { $oid = [string]$rowSel.Cells['m2'].Value }   # repli e-mail
            if ($oid) { [void]$toRemove.Add($oid) }
        }
        if ($toRemove.Count -eq 0) { return }

        $prompt = "RETRAIT de $($toRemove.Count) membre(s) du groupe :`n  $name  ($gid)`n`n" +
                  "Operation potentiellement destructrice.`nPour confirmer, saisissez exactement l'ObjectId du groupe :`n`n$gid"
        if (-not (Confirm-ByTyping -Expected $gid -Prompt $prompt)) {
            Write-Log "Retrait (liste des membres) annule (ID non confirme)." 'WARN'
            return
        }

        $w.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnRemoveSel.Enabled = $false; $btnCsv.Enabled = $false; $btnAddCol.Enabled = $false; $btnClose.Enabled = $false
        Write-Log "===== RETRAIT (liste des membres) : $($toRemove.Count) sur '$name' ($gid) =====" 'INFO'
        $okc = 0; $failc = 0
        foreach ($oid in $toRemove) {
            try {
                Remove-UnifiedGroupLinks -Identity $gid -LinkType Members -Links $oid -Confirm:$false -ErrorAction Stop | Out-Null
                $okc++; Write-Log "OK  RETRAIT  $oid" 'OK'
            } catch {
                $failc++; Write-Log "ECHEC  $oid : $($_.Exception.Message)" 'ERR'
            }
        }
        Write-Log "----- Bilan retrait (liste) : reussite $okc / echec $failc -----" $(if($failc -gt 0){'WARN'}else{'OK'})

        $script:SimulationDone = $false
        Update-ButtonsState

        & $loadData
        & $render $txtFilter.Text.Trim()
        $w.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRemoveSel.Enabled = $true; $btnCsv.Enabled = $true; $btnAddCol.Enabled = $true; $btnClose.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Retrait termine.`nReussite : $okc`nEchec : $failc","Bilan",'OK','Information') | Out-Null
    })

    $w.CancelButton = $btnClose
    [void]$w.ShowDialog($script:MainForm)
    if ($script:btnListMembers) { $script:btnListMembers.Enabled = ($null -ne $script:ValidatedGroup -and -not $script:IsProcessing) }
}

# ============================================================================
#  9. Execution (Ajout / Retrait)
# ============================================================================
function Invoke-Operation {
    param([ValidateSet('Add','Remove')][string]$Action)

    if (-not $script:ValidatedGroup -or -not $script:SimulationDone -or -not $script:SimData) { return }

    $gid  = $script:ValidatedGroup.Id
    $name = $script:ValidatedGroup.Name

    if ($Action -eq 'Add') {
        $targets = @($script:SimData | Where-Object Category -eq 'HORS_GROUPE')
    } else {
        $targets = @($script:SimData | Where-Object Category -eq 'MEMBRE')
    }

    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur eligible pour cette operation dans la recherche courante.","Rien a faire",'OK','Information') | Out-Null
        return
    }

    $skipped = @($script:SimData | Where-Object Category -in @('NON_TROUVE','AMBIGU','INVALIDE')).Count
    $other   = $script:SimData.Count - $targets.Count - $skipped

    if ($Action -eq 'Add') {
        $msg = "AJOUT de $($targets.Count) utilisateur(s) dans :`n`n  $name`n  $gid`n`n" +
               "Ignores : deja membres ($other) / non resolus ($skipped)`n`nConfirmer l'ajout ?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg,"Confirmation d'ajout",'YesNo','Warning')
        if ($r -ne 'Yes') { Write-Log "Ajout annule par l'utilisateur." 'WARN'; return }
    } else {
        $prompt = "RETRAIT de $($targets.Count) utilisateur(s) du groupe :`n  $name  ($gid)`n`n" +
                  "Operation potentiellement destructrice.`nPour confirmer, saisissez exactement l'ObjectId du groupe :`n`n$gid"
        if (-not (Confirm-ByTyping -Expected $gid -Prompt $prompt)) {
            Write-Log "Retrait annule (ID non confirme)." 'WARN'; return
        }
    }

    Initialize-LogFile
    Set-ProcessingState $true
    $script:CancelRequested = $false
    Write-Log "===== $($Action.ToUpper()) : $($targets.Count) utilisateur(s) sur '$name' ($gid) =====" 'INFO'
    Write-Log "La fenetre peut afficher (Ne repond pas) pendant le traitement : c'est normal, ne pas forcer la fermeture." 'WARN'

    $ok = 0; $fail = 0
    $script:progress.Minimum = 0; $script:progress.Maximum = $targets.Count; $script:progress.Value = 0

    $i = 0
    foreach ($t in $targets) {
        if ($script:CancelRequested) { Write-Log "Operation interrompue par l'utilisateur." 'WARN'; break }
        $i++; $script:progress.Value = $i
        # Identite utilisee : ObjectId si dispo, sinon e-mail.
        $link = if ($t.ObjectId) { $t.ObjectId } else { $t.Smtp }
        try {
            if ($Action -eq 'Add') {
                Add-UnifiedGroupLinks -Identity $gid -LinkType Members -Links $link -ErrorAction Stop | Out-Null
                Write-Log "OK  AJOUT  $($t.Name) ($link)" 'OK'
            } else {
                Remove-UnifiedGroupLinks -Identity $gid -LinkType Members -Links $link -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Log "OK  RETRAIT  $($t.Name) ($link)" 'OK'
            }
            $ok++
            Update-GridRowStatus $t.Name ($(if($Action -eq 'Add'){'AJOUTE'}else{'RETIRE'})) ([System.Drawing.Color]::FromArgb(200,235,200))
        }
        catch {
            $fail++
            Write-Log "ECHEC  $($t.Name) ($link) : $($_.Exception.Message)" 'ERR'
            Update-GridRowStatus $t.Name 'ECHEC' ([System.Drawing.Color]::FromArgb(250,210,210))
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $script:progress.Value = 0

    Write-Log "----- Bilan $($Action.ToUpper()) : reussite $ok / echec $fail -----" $(if($fail -gt 0){'WARN'}else{'OK'})

    $script:SimulationDone = $false
    Set-ProcessingState $false
    [System.Windows.Forms.MessageBox]::Show("Termine.`nReussite : $ok`nEchec : $fail`n`nRelancez une recherche pour toute autre action.","Bilan",'OK','Information') | Out-Null
}

function Update-GridRowStatus {
    param([string]$Name,[string]$Status,[System.Drawing.Color]$Color)
    foreach ($row in $script:grid.Rows) {
        if ($row.Cells['c1'].Value -eq $Name) {
            $row.Cells['c4'].Value = $Status
            $row.DefaultCellStyle.BackColor = $Color
            break
        }
    }
}

# ============================================================================
#  10. Interface
# ============================================================================
$script:MainForm = New-Object System.Windows.Forms.Form
$script:MainForm.Text = "Microsoft 365 - Ajout / retrait d'utilisateurs dans un groupe (Exchange Online)"
$script:MainForm.Size = New-Object System.Drawing.Size(960,860)
$script:MainForm.MinimumSize = New-Object System.Drawing.Size(880,760)
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI',9)

$lblCtx = New-Object System.Windows.Forms.Label
$lblCtx.SetBounds(15,12,910,20); $lblCtx.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)

# --- Groupe ---
$grpGrp = New-Object System.Windows.Forms.GroupBox
$grpGrp.Text = "1. Groupe Microsoft 365 cible"; $grpGrp.SetBounds(15,40,910,95); $grpGrp.Anchor = 'Top,Left,Right'

$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Text = "Nom, e-mail ou GUID :"; $lbl1.SetBounds(15,28,130,20)
$script:txtGroup = New-Object System.Windows.Forms.TextBox
$script:txtGroup.SetBounds(150,25,375,24); $script:txtGroup.Anchor = 'Top,Left,Right'
$script:btnValidate = New-Object System.Windows.Forms.Button
$script:btnValidate.Text = "Valider le groupe"; $script:btnValidate.SetBounds(535,24,175,26); $script:btnValidate.Anchor = 'Top,Right'
$script:btnListMembers = New-Object System.Windows.Forms.Button
$script:btnListMembers.Text = "Lister les membres"; $script:btnListMembers.SetBounds(715,24,175,26); $script:btnListMembers.Anchor = 'Top,Right'; $script:btnListMembers.Enabled = $false
$script:lblGroupInfo = New-Object System.Windows.Forms.Label
$script:lblGroupInfo.Text = "Groupe non valide."; $script:lblGroupInfo.SetBounds(15,58,875,28)
$script:lblGroupInfo.ForeColor = [System.Drawing.Color]::DimGray; $script:lblGroupInfo.Anchor = 'Top,Left,Right'
$grpGrp.Controls.AddRange(@($lbl1,$script:txtGroup,$script:btnValidate,$script:btnListMembers,$script:lblGroupInfo))

# --- Utilisateurs ---
$grpUsr = New-Object System.Windows.Forms.GroupBox
$grpUsr.Text = "2. Utilisateurs (UPN ou e-mail, un par ligne ou colles ; separateurs espace/virgule/point-virgule)"
$grpUsr.SetBounds(15,145,910,170); $grpUsr.Anchor = 'Top,Left,Right'

$script:txtUsers = New-Object System.Windows.Forms.TextBox
$script:txtUsers.Multiline = $true; $script:txtUsers.ScrollBars = 'Vertical'
$script:txtUsers.SetBounds(15,25,720,130); $script:txtUsers.Font = New-Object System.Drawing.Font('Consolas',9)
$script:txtUsers.Anchor = 'Top,Left,Right,Bottom'

$script:btnLoadTxt = New-Object System.Windows.Forms.Button
$script:btnLoadTxt.Text = "Charger .txt"; $script:btnLoadTxt.SetBounds(750,25,145,30); $script:btnLoadTxt.Anchor = 'Top,Right'
$script:btnClear = New-Object System.Windows.Forms.Button
$script:btnClear.Text = "Vider la liste"; $script:btnClear.SetBounds(750,62,145,30); $script:btnClear.Anchor = 'Top,Right'
$script:lblCount = New-Object System.Windows.Forms.Label
$script:lblCount.Text = "0 entree unique"; $script:lblCount.SetBounds(750,105,145,40); $script:lblCount.Anchor = 'Top,Right'
$grpUsr.Controls.AddRange(@($script:txtUsers,$script:btnLoadTxt,$script:btnClear,$script:lblCount))

# --- Log ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Dossier de log :"; $lblLog.SetBounds(20,332,95,20)
$script:txtLogFolder = New-Object System.Windows.Forms.TextBox
$script:txtLogFolder.SetBounds(120,329,470,24)
$script:txtLogFolder.Text = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'M365_GroupBulk_Logs')
$script:btnBrowseLog = New-Object System.Windows.Forms.Button
$script:btnBrowseLog.Text = "Parcourir"; $script:btnBrowseLog.SetBounds(600,328,90,26)

# --- Actions ---
$script:btnSimulate = New-Object System.Windows.Forms.Button
$script:btnSimulate.Text = "RECHERCHER"; $script:btnSimulate.SetBounds(15,375,220,38); $script:btnSimulate.Enabled = $false
$script:btnSimulate.BackColor = [System.Drawing.Color]::FromArgb(210,225,245)
$script:btnAdd = New-Object System.Windows.Forms.Button
$script:btnAdd.Text = "AJOUTER"; $script:btnAdd.SetBounds(245,375,220,38); $script:btnAdd.Enabled = $false
$script:btnAdd.BackColor = [System.Drawing.Color]::FromArgb(210,240,210)
$script:btnRemove = New-Object System.Windows.Forms.Button
$script:btnRemove.Text = "RETIRER"; $script:btnRemove.SetBounds(475,375,220,38); $script:btnRemove.Enabled = $false
$script:btnRemove.BackColor = [System.Drawing.Color]::FromArgb(245,215,215)
$script:btnStop = New-Object System.Windows.Forms.Button
$script:btnStop.Text = "STOP"; $script:btnStop.SetBounds(705,375,220,38); $script:btnStop.Enabled = $false

$script:progress = New-Object System.Windows.Forms.ProgressBar
$script:progress.SetBounds(15,422,910,18); $script:progress.Anchor = 'Top,Left,Right'

$script:grid = New-Object System.Windows.Forms.DataGridView
$script:grid.SetBounds(15,448,910,190); $script:grid.Anchor = 'Top,Left,Right'
$script:grid.ReadOnly = $true; $script:grid.AllowUserToAddRows = $false
$script:grid.SelectionMode = 'FullRowSelect'; $script:grid.RowHeadersVisible = $false
$script:grid.AutoSizeColumnsMode = 'Fill'
[void]$script:grid.Columns.Add('c1','Entree demandee')
[void]$script:grid.Columns.Add('c2','Categorie')
[void]$script:grid.Columns.Add('c3','ObjectId')
[void]$script:grid.Columns.Add('c5','Nom affiche')
[void]$script:grid.Columns.Add('c4','Detail / Statut')

$lblJ = New-Object System.Windows.Forms.Label
$lblJ.Text = "Journal :"; $lblJ.SetBounds(15,645,80,18); $lblJ.Anchor = 'Bottom,Left'
$script:txtLog = New-Object System.Windows.Forms.TextBox
$script:txtLog.Multiline = $true; $script:txtLog.ScrollBars = 'Vertical'; $script:txtLog.ReadOnly = $true
$script:txtLog.SetBounds(15,665,790,140); $script:txtLog.Font = New-Object System.Drawing.Font('Consolas',8.5)
$script:txtLog.BackColor = [System.Drawing.Color]::FromArgb(30,30,30); $script:txtLog.ForeColor = [System.Drawing.Color]::Gainsboro
$script:txtLog.Anchor = 'Bottom,Left,Right'
$script:btnExport = New-Object System.Windows.Forms.Button
$script:btnExport.Text = "Exporter CSV"; $script:btnExport.SetBounds(815,665,110,32); $script:btnExport.Anchor = 'Bottom,Right'
$script:btnOpenLog = New-Object System.Windows.Forms.Button
$script:btnOpenLog.Text = "Ouvrir le log"; $script:btnOpenLog.SetBounds(815,702,110,32); $script:btnOpenLog.Anchor = 'Bottom,Right'

$script:MainForm.Controls.AddRange(@(
    $lblCtx,$grpGrp,$grpUsr,
    $lblLog,$script:txtLogFolder,$script:btnBrowseLog,
    $script:btnSimulate,$script:btnAdd,$script:btnRemove,$script:btnStop,
    $script:progress,$script:grid,$lblJ,$script:txtLog,$script:btnExport,$script:btnOpenLog
))

# ============================================================================
#  11. Evenements
# ============================================================================
$script:btnValidate.Add_Click({ Validate-Group })
$script:btnListMembers.Add_Click({ Show-GroupMembers })
$script:txtGroup.Add_TextChanged({ Invalidate-Group })

$script:txtUsers.Add_TextChanged({
    $n = (Get-InputUserNames).Count
    $script:lblCount.Text = "$n entree(s) unique(s)"
    Reset-Simulation
})

$script:btnLoadTxt.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Get-Content -LiteralPath $ofd.FileName -ErrorAction Stop
            if ($script:txtUsers.Text.Trim() -ne '') { $script:txtUsers.AppendText("`r`n") }
            $script:txtUsers.AppendText(($lines -join "`r`n"))
            Write-Log "Fichier charge : $($ofd.FileName) ($($lines.Count) ligne(s))" 'INFO'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Lecture impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        }
    }
})

$script:btnClear.Add_Click({ $script:txtUsers.Clear() })

$script:btnBrowseLog.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') { $script:txtLogFolder.Text = $fbd.SelectedPath; $script:LogFile = $null }
})

$script:btnSimulate.Add_Click({ Invoke-Simulation })
$script:btnAdd.Add_Click({ Invoke-Operation -Action 'Add' })
$script:btnRemove.Add_Click({ Invoke-Operation -Action 'Remove' })
$script:btnStop.Add_Click({ $script:CancelRequested = $true; Write-Log "Arret demande..." 'WARN' })

$script:btnExport.Add_Click({
    if (-not $script:SimData -or $script:SimData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucune donnee a exporter (lancez d'abord une recherche).","Export",'OK','Information') | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV (*.csv)|*.csv"; $sfd.FileName = "M365_Recherche_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $script:SimData | Select-Object Name,Category,ObjectId,DisplayName,Smtp,Detail |
                Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            Write-Log "Export CSV : $($sfd.FileName)" 'OK'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        }
    }
})

$script:btnOpenLog.Add_Click({
    if ($script:LogFile -and (Test-Path -LiteralPath $script:LogFile)) { Start-Process notepad.exe $script:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Aucun fichier de log pour l'instant.","Log",'OK','Information') | Out-Null }
})

$script:MainForm.Add_FormClosing({
    if ($script:IsProcessing) {
        $_.Cancel = $true
        $script:CancelRequested = $true
        [System.Windows.Forms.MessageBox]::Show("Un traitement est en cours. Il a ete interrompu ; reessayez de fermer.","Fermeture",'OK','Warning') | Out-Null
    }
})

# ============================================================================
#  12. Demarrage
# ============================================================================
if (-not (Initialize-ExoContext)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Aucun contexte Exchange Online detecte.`n`n" +
        "Lancez ce script depuis une console ou Connect-ExchangeOnline a ete effectue " +
        "(module ExchangeOnlineManagement).",
        "Contexte Exchange manquant",'OK','Error') | Out-Null
    $lblCtx.Text = "Contexte Exchange Online : NON DETECTE - actions bloquees."
    $lblCtx.ForeColor = [System.Drawing.Color]::Firebrick
    $script:btnValidate.Enabled = $false
} else {
    $lblCtx.Text = "Exchange Online detecte - Compte : $($script:ExoUser)   |   Organisation : $($script:ExoOrg)"
    $lblCtx.ForeColor = [System.Drawing.Color]::DarkGreen
}

Update-ButtonsState
[void]$script:MainForm.ShowDialog()
