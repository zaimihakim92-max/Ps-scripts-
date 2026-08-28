#Requires -Version 5.1
<#
.SYNOPSIS
    Audit du parc - Active Directory + SCCM (interface graphique)

.DESCRIPTION
    Vue detaillee du parc machines avec deux (trois) onglets :
      * AD     : ordinateurs des OU choisies (selecteur d'OU en arborescence)
                 -> Description, dernier logon machine, OS, activation,
                    ancienneté (inactivite 90 / 160 jours), dernier utilisateur
                    (rapproche depuis SCCM quand disponible).
      * SCCM   : TOUS les postes (Get-CMDevice) -> dernier utilisateur,
                 derniere activite client, inactivite 90 / 160 jours.
      * Rapprochement AD/SCCM (double audit) : jointure par nom de machine,
                 presence AD seul / SCCM seul / les deux, divergences,
                 machines obsoletes des deux cotes -> base pour la gestion de parc.

    Chaque onglet reprend les memes mecanismes que l'outil "dernier utilisateur" :
      tri par clic d'en-tete, redimensionnement/reordonnancement des colonnes,
      filtre texte temps reel, filtre par statut, coloration des lignes,
      menu contextuel (copier), export CSV avec choix et ordre des colonnes.

    Un export consolide "gestion de parc" reprend la table de rapprochement
    (source de verite par machine) prete a alimenter un referentiel.

.NOTES
    A lancer en STA : powershell.exe -STA -File .\audit_parc_ad_sccm_gui.ps1
    Doit tourner dans une session ou Get-CMDevice est disponible
    (console Configuration Manager / lecteur de site charge) pour la partie SCCM.
    La partie AD ne necessite que le module ActiveDirectory : si SCCM est absent,
    l'outil fonctionne en mode AD seul (onglets SCCM / Rapprochement vides).

    Seuils d'inactivite par defaut : 90 jours (seuil bas) et 160 jours (seuil haut),
    modifiables dans la barre haute avant de lancer l'audit.

    Remarque technique : cote AD, l'inactivite s'appuie sur LastLogonDate
    (lastLogonTimestamp, precision de replication ~9-14 jours) - suffisant pour
    reperer les objets obsoletes, pas pour un horodatage a la minute.
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

# SCCM : verification souple. Absent -> mode AD seul.
$script:SccmAvailable = [bool](Get-Command Get-CMDevice -ErrorAction SilentlyContinue)

# ==================================================================
#  Configuration
# ==================================================================
$script:LowDays  = 90    # seuil bas  (inactivite = a surveiller)
$script:HighDays = 160   # seuil haut (inactivite = obsolete)

$script:SelectedOUs = @()      # DN des OU choisies
$script:Recurse     = $true    # inclure les sous-OU (Subtree) sinon OneLevel
$script:IncludeHardware = $true # enrichir modele / n. serie via l'inventaire materiel SCCM

# --- Specifications des trois vues -------------------------------
# Chaque vue : ordre des colonnes, type, libelles, colonnes filtrables texte.
# L'ordre des cles = l'ordre d'affichage (les infos utiles d'abord).
$script:Spec = @{
    AD = @{
        Kind    = 'AD'
        Columns = [ordered]@{
            name          = [string]
            lastLogonUser = [string]
            model         = [string]
            serialNumber  = [string]
            description   = [string]
            os            = [string]
            lastLogon     = [datetime]
            inactiveDays  = [int]
            enabled       = [bool]
            pwdLastSet    = [datetime]
            dn            = [string]
            status        = [string]
        }
        Headers = @{
            name='Machine'; lastLogonUser='Dernier utilisateur'; model='Modele'
            serialNumber='N. serie'; description='Description'; os='OS'
            lastLogon='Dernier logon (AD)'; inactiveDays='Inactif (j)'
            enabled='Active'; pwdLastSet='MDP change le'; dn='OU (DN)'; status='Statut'
        }
        TextCols     = @('name','lastLogonUser','model','serialNumber','description','os','dn','status')
        DateCols     = @('lastLogon','pwdLastSet')
        StatusItems  = @('Tous','OK','Inactif seuil bas','Inactif seuil haut','Desactive','Inconnu')
        ExportName   = 'Audit_AD_Parc'
    }
    SCCM = @{
        Kind    = 'SCCM'
        Columns = [ordered]@{
            name          = [string]
            lastLogonUser = [string]
            manufacturer  = [string]
            model         = [string]
            serialNumber  = [string]
            os            = [string]
            lastActive    = [datetime]
            inactiveDays  = [int]
            clientStatus  = [string]
            resourceId    = [string]
            status        = [string]
        }
        Headers = @{
            name='Machine'; lastLogonUser='Dernier utilisateur'; manufacturer='Fabricant'
            model='Modele'; serialNumber='N. serie'; os='OS'
            lastActive='Derniere activite'; inactiveDays='Inactif (j)'
            clientStatus='Client'; resourceId='ResourceID'; status='Statut'
        }
        TextCols     = @('name','lastLogonUser','manufacturer','model','serialNumber','clientStatus','os','resourceId','status')
        DateCols     = @('lastActive')
        StatusItems  = @('Tous','OK','Inactif seuil bas','Inactif seuil haut','Inconnu')
        ExportName   = 'Audit_SCCM_Parc'
    }
    RECON = @{
        Kind    = 'RECON'
        Columns = [ordered]@{
            machine          = [string]
            sccmLastUser     = [string]
            model            = [string]
            serialNumber     = [string]
            inAD             = [bool]
            inSCCM           = [bool]
            description      = [string]
            adLastLogon      = [datetime]
            adInactiveDays   = [int]
            sccmLastActive   = [datetime]
            sccmInactiveDays = [int]
            presence         = [string]
            flag             = [string]
        }
        Headers = @{
            machine='Machine'; sccmLastUser='Dernier utilisateur'; model='Modele'; serialNumber='N. serie'
            inAD='Dans AD'; inSCCM='Dans SCCM'; description='Description (AD)'
            adLastLogon='Logon AD'; adInactiveDays='Inactif AD (j)'
            sccmLastActive='Activite SCCM'; sccmInactiveDays='Inactif SCCM (j)'
            presence='Presence'; flag='Diagnostic'
        }
        TextCols     = @('machine','sccmLastUser','model','serialNumber','description','presence','flag')
        DateCols     = @('adLastLogon','sccmLastActive')
        StatusItems  = @('Tous','AD seul','SCCM seul','Les deux','Obsolete (AD+SCCM)','A verifier')
        ExportName   = 'Gestion_Parc_Consolide'
    }
}

$script:Views = @{}   # rempli a la construction de l'UI : Kind -> objet vue

# ==================================================================
#  Fonctions - utilitaires
# ==================================================================
function Get-Prop {
    # Lecture defensive d'une propriete (objets WMI/SCCM aux schemas variables)
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value } else { return $null }
}

function Get-ShortName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return ($Name -split '\.')[0].Trim().ToUpper()
}

function Get-InactiveDays {
    param($LastDate)
    if (-not $LastDate) { return $null }
    try {
        $d = [datetime]$LastDate
        if ($d -eq [datetime]::MinValue) { return $null }
        return [int]([datetime]::Now - $d).TotalDays
    } catch { return $null }
}

function Set-Cell {
    # Affecte une valeur ou DBNull si vide/nulle, selon le type de colonne
    param([System.Data.DataRow]$Row, [string]$Name, $Value)
    if ($null -eq $Value) { $Row[$Name] = [DBNull]::Value; return }
    if ($Value -is [string] -and $Value -eq '') {
        # Chaine vide : on garde '' pour les colonnes texte
        $Row[$Name] = ''
        return
    }
    $Row[$Name] = $Value
}

function New-TableFromSpec {
    param([hashtable]$ViewSpec)
    $dt = New-Object System.Data.DataTable
    foreach ($name in $ViewSpec.Columns.Keys) {
        [void]$dt.Columns.Add($name, $ViewSpec.Columns[$name])
    }
    return ,$dt
}

function Copy-ToClipboard {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    try { [System.Windows.Forms.Clipboard]::SetText($Text); return $true } catch { return $false }
}

# ==================================================================
#  Fonctions - lecture AD / SCCM
# ==================================================================
function Get-AdComputerObjects {
    param([string[]]$OUs, [bool]$Recurse)

    $scope = if ($Recurse) { 'Subtree' } else { 'OneLevel' }
    $props = @('Description','OperatingSystem','LastLogonDate','Enabled',
               'DNSHostName','whenCreated','PasswordLastSet','DistinguishedName')
    $out = @()
    foreach ($ou in $OUs) {
        $out += Get-ADComputer -SearchBase $ou -SearchScope $scope -Filter * `
                    -Properties $props -ErrorAction Stop
    }
    return $out
}

function Get-SccmDeviceObjects {
    # Normalise Get-CMDevice en objets simples et defensifs.
    # LastActiveTime peut etre nul selon la sante client -> repli LastClientCheckTime.
    $devices = Get-CMDevice -ErrorAction Stop
    foreach ($d in $devices) {
        $la = Get-Prop $d 'LastActiveTime'
        if (-not $la) { $la = Get-Prop $d 'LastClientCheckTime' }
        if ($la -and ([datetime]$la) -eq [datetime]::MinValue) { $la = $null }

        $cas = Get-Prop $d 'ClientActiveStatus'
        $client = switch ($cas) {
            1       { 'Actif' }
            0       { 'Inactif' }
            default { if ((Get-Prop $d 'IsClient')) { 'Client' } else { 'Sans client' } }
        }

        [pscustomobject]@{
            Name          = [string](Get-Prop $d 'Name')
            ResourceID    = [string](Get-Prop $d 'ResourceID')
            LastLogonUser = [string](Get-Prop $d 'LastLogonUser')
            LastActive    = if ($la) { [datetime]$la } else { $null }
            ClientStatus  = $client
            OS            = [string](Get-Prop $d 'DeviceOS')
            Manufacturer  = ''
            Model         = ''
            SerialNumber  = ''
        }
    }
}

function Get-SccmHardwareIndex {
    # Enrichissement modele / fabricant / n. serie via l'inventaire materiel SCCM
    # (fournisseur SMS). Indexe par ResourceID. Degrade proprement si indisponible.
    $res = @{ Model = @{}; Serial = @{}; Ok = $false; Message = '' }
    try {
        $drive    = Get-PSDrive -PSProvider CMSite -ErrorAction Stop | Select-Object -First 1
        $siteCode = $drive.Name
        $provider = [string]$drive.Root
        if ([string]::IsNullOrWhiteSpace($provider) -or $provider -match '^[A-Za-z0-9]{3}:') {
            $provider = [string](Get-Prop $drive 'SiteServer')
        }
        if ([string]::IsNullOrWhiteSpace($provider)) { $provider = $env:COMPUTERNAME }
        $ns = "root\sms\site_$siteCode"

        # Fabricant + modele (SMS_G_System_COMPUTER_SYSTEM)
        Get-CimInstance -ComputerName $provider -Namespace $ns `
            -ClassName SMS_G_System_COMPUTER_SYSTEM -ErrorAction Stop |
            ForEach-Object {
                $rid = [string]$_.ResourceID
                if ($rid) {
                    $res.Model[$rid] = [pscustomobject]@{
                        Manufacturer = [string]$_.Manufacturer
                        Model        = [string]$_.Model
                    }
                }
            }

        # Numero de serie / service tag (SMS_G_System_PC_BIOS)
        Get-CimInstance -ComputerName $provider -Namespace $ns `
            -ClassName SMS_G_System_PC_BIOS -ErrorAction Stop |
            ForEach-Object {
                $rid = [string]$_.ResourceID
                $sn  = [string]$_.SerialNumber
                if ($rid -and $sn) { $res.Serial[$rid] = $sn.Trim() }
            }

        $res.Ok = $true
    }
    catch {
        $res.Message = $_.Exception.Message
    }
    return $res
}

# ==================================================================
#  Fonctions - construction des tables
# ==================================================================
function Build-AllTables {
    param(
        [System.Windows.Forms.ToolStripProgressBar]$Progress,
        [System.Windows.Forms.ToolStripStatusLabel]$Status
    )

    $low  = $script:LowDays
    $high = $script:HighDays

    $tAD    = $script:Views.AD.Table
    $tSCCM  = $script:Views.SCCM.Table
    $tRECON = $script:Views.RECON.Table
    $tAD.Rows.Clear(); $tSCCM.Rows.Clear(); $tRECON.Rows.Clear()

    # ---------- AD ----------
    $Status.Text = "Lecture Active Directory..."
    [System.Windows.Forms.Application]::DoEvents()
    $adComps = Get-AdComputerObjects -OUs $script:SelectedOUs -Recurse $script:Recurse

    # ---------- SCCM ----------
    $sccmDevices = @()
    if ($script:SccmAvailable) {
        $Status.Text = "Lecture SCCM (tous les postes)..."
        [System.Windows.Forms.Application]::DoEvents()
        try   { $sccmDevices = @(Get-SccmDeviceObjects) }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Lecture SCCM impossible :`n$($_.Exception.Message)`n`nPoursuite en mode AD seul.",
                "SCCM", 'OK', 'Warning')
            $sccmDevices = @()
        }
    }

    # ---------- Enrichissement materiel (modele / fabricant / n. serie) ----------
    if ($script:IncludeHardware -and $sccmDevices.Count -gt 0) {
        $Status.Text = "Lecture inventaire materiel SCCM (modele / n. serie)..."
        [System.Windows.Forms.Application]::DoEvents()
        $hw = Get-SccmHardwareIndex
        if ($hw.Ok) {
            foreach ($s in $sccmDevices) {
                $rid = [string]$s.ResourceID
                if ($rid -and $hw.Model.ContainsKey($rid)) {
                    $s.Manufacturer = $hw.Model[$rid].Manufacturer
                    $s.Model        = $hw.Model[$rid].Model
                }
                if ($rid -and $hw.Serial.ContainsKey($rid)) {
                    $s.SerialNumber = $hw.Serial[$rid]
                }
            }
        }
        else {
            $script:lblStatus.Text = "Inventaire materiel indisponible (modele/serie vides) : $($hw.Message)"
        }
    }

    # Index SCCM par nom court (pour le rapprochement + enrichissement cote AD)
    $sccmByName = @{}
    foreach ($s in $sccmDevices) {
        $k = Get-ShortName $s.Name
        if ($k -and -not $sccmByName.ContainsKey($k)) { $sccmByName[$k] = $s }
    }

    $Progress.Minimum = 0
    $Progress.Maximum = [Math]::Max(1, ($adComps.Count + $sccmDevices.Count))
    $Progress.Value   = 0
    $step = 0

    # ---------- Remplissage table AD ----------
    $adByName = @{}
    foreach ($c in $adComps) {
        $step++; if ($step % 25 -eq 0) { $Progress.Value = [Math]::Min($step,$Progress.Maximum); [System.Windows.Forms.Application]::DoEvents() }

        $key   = Get-ShortName $c.Name
        $days  = Get-InactiveDays $c.LastLogonDate
        $enab  = [bool]$c.Enabled

        # Enrichissement depuis SCCM si le poste y est connu
        $lastUser = ''; $model = ''; $serial = ''
        if ($sccmByName.ContainsKey($key)) {
            $sx       = $sccmByName[$key]
            $lastUser = [string]$sx.LastLogonUser
            $model    = [string]$sx.Model
            $serial   = [string]$sx.SerialNumber
        }

        $state =
            if (-not $enab)                    { 'Desactive' }
            elseif ($null -eq $days)           { 'Inconnu (jamais vu)' }
            elseif ($days -ge $high)           { "Inactif >= $high j" }
            elseif ($days -ge $low)            { "Inactif >= $low j" }
            else                               { 'OK' }

        $r = $tAD.NewRow()
        Set-Cell $r 'name'          ([string]$c.Name)
        Set-Cell $r 'lastLogonUser' $lastUser
        Set-Cell $r 'model'         $model
        Set-Cell $r 'serialNumber'  $serial
        Set-Cell $r 'description'   ([string]$c.Description)
        Set-Cell $r 'os'            ([string]$c.OperatingSystem)
        Set-Cell $r 'lastLogon'     $c.LastLogonDate
        Set-Cell $r 'inactiveDays'  $days
        Set-Cell $r 'enabled'       $enab
        Set-Cell $r 'pwdLastSet'    $c.PasswordLastSet
        Set-Cell $r 'dn'            ([string]$c.DistinguishedName)
        Set-Cell $r 'status'        $state
        [void]$tAD.Rows.Add($r)

        if ($key -and -not $adByName.ContainsKey($key)) {
            $adByName[$key] = [pscustomobject]@{
                Description = [string]$c.Description
                LastLogon   = $c.LastLogonDate
                Days        = $days
                Enabled     = $enab
            }
        }
    }

    # ---------- Remplissage table SCCM ----------
    foreach ($s in $sccmDevices) {
        $step++; if ($step % 25 -eq 0) { $Progress.Value = [Math]::Min($step,$Progress.Maximum); [System.Windows.Forms.Application]::DoEvents() }

        $days = Get-InactiveDays $s.LastActive
        $state =
            if     ($null -eq $days) { 'Inconnu (activite absente)' }
            elseif ($days -ge $high) { "Inactif >= $high j" }
            elseif ($days -ge $low)  { "Inactif >= $low j" }
            else                     { 'OK' }

        $r = $tSCCM.NewRow()
        Set-Cell $r 'name'          ([string]$s.Name)
        Set-Cell $r 'lastLogonUser' ([string]$s.LastLogonUser)
        Set-Cell $r 'manufacturer'  ([string]$s.Manufacturer)
        Set-Cell $r 'model'         ([string]$s.Model)
        Set-Cell $r 'serialNumber'  ([string]$s.SerialNumber)
        Set-Cell $r 'os'            ([string]$s.OS)
        Set-Cell $r 'lastActive'    $s.LastActive
        Set-Cell $r 'inactiveDays'  $days
        Set-Cell $r 'clientStatus'  ([string]$s.ClientStatus)
        Set-Cell $r 'resourceId'    ([string]$s.ResourceID)
        Set-Cell $r 'status'        $state
        [void]$tSCCM.Rows.Add($r)
    }

    # ---------- Table de rapprochement (double audit) ----------
    $Status.Text = "Rapprochement AD / SCCM..."
    [System.Windows.Forms.Application]::DoEvents()

    $allKeys = @($adByName.Keys) + @($sccmByName.Keys) | Select-Object -Unique
    foreach ($k in $allKeys) {
        $inAd   = $adByName.ContainsKey($k)
        $inScc  = $sccmByName.ContainsKey($k)
        $ad     = if ($inAd)  { $adByName[$k] }  else { $null }
        $sc     = if ($inScc) { $sccmByName[$k] } else { $null }

        $adDays   = if ($ad) { $ad.Days } else { $null }
        $scDays   = if ($sc) { Get-InactiveDays $sc.LastActive } else { $null }

        $presence =
            if ($inAd -and $inScc) { 'AD + SCCM' }
            elseif ($inAd)         { 'AD seul' }
            else                   { 'SCCM seul' }

        $flag =
            if     (-not $inScc) { 'Present AD, absent SCCM' }
            elseif (-not $inAd)  { 'Present SCCM, absent AD' }
            elseif (($adDays -ne $null -and $adDays -ge $high) -and ($scDays -ne $null -and $scDays -ge $high)) { 'Obsolete (AD+SCCM)' }
            elseif (($adDays -ne $null -and $adDays -ge $low)  -or  ($scDays -ne $null -and $scDays -ge $low))  { 'A verifier' }
            else { 'OK / actif' }

        $r = $tRECON.NewRow()
        Set-Cell $r 'machine'          $k
        Set-Cell $r 'sccmLastUser'     $(if ($sc) { [string]$sc.LastLogonUser } else { '' })
        Set-Cell $r 'model'            $(if ($sc) { [string]$sc.Model } else { '' })
        Set-Cell $r 'serialNumber'     $(if ($sc) { [string]$sc.SerialNumber } else { '' })
        Set-Cell $r 'inAD'             $inAd
        Set-Cell $r 'inSCCM'           $inScc
        Set-Cell $r 'description'      $(if ($ad) { [string]$ad.Description } else { '' })
        Set-Cell $r 'adLastLogon'      $(if ($ad) { $ad.LastLogon } else { $null })
        Set-Cell $r 'adInactiveDays'   $adDays
        Set-Cell $r 'sccmLastActive'   $(if ($sc) { $sc.LastActive } else { $null })
        Set-Cell $r 'sccmInactiveDays' $scDays
        Set-Cell $r 'presence'         $presence
        Set-Cell $r 'flag'             $flag
        [void]$tRECON.Rows.Add($r)
    }

    $Progress.Value = $Progress.Maximum
    $Status.Text = "Termine. AD : $($tAD.Rows.Count)  |  SCCM : $($tSCCM.Rows.Count)  |  Rapprochement : $($tRECON.Rows.Count)"
}

# ==================================================================
#  Fonctions - filtre / couleurs / compteurs (generique par vue)
# ==================================================================
function Update-ViewColors {
    param($View)
    $low  = $script:LowDays
    $high = $script:HighDays
    $red   = [System.Drawing.Color]::FromArgb(255,224,224)
    $amber = [System.Drawing.Color]::FromArgb(255,245,204)
    $gray  = [System.Drawing.Color]::FromArgb(230,230,230)
    $blue  = [System.Drawing.Color]::FromArgb(224,238,255)
    $purple= [System.Drawing.Color]::FromArgb(240,226,255)

    foreach ($row in $View.Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $bg = [System.Drawing.Color]::Empty

        switch ($View.Kind) {
            'AD' {
                $st = [string]$row.Cells['status'].Value
                $d  = $row.Cells['inactiveDays'].Value
                if     ($st -eq 'Desactive')            { $bg = $gray }
                elseif ($d -isnot [DBNull] -and $d -ge $high) { $bg = $red }
                elseif ($d -isnot [DBNull] -and $d -ge $low)  { $bg = $amber }
            }
            'SCCM' {
                $d = $row.Cells['inactiveDays'].Value
                if     ($d -isnot [DBNull] -and $d -ge $high) { $bg = $red }
                elseif ($d -isnot [DBNull] -and $d -ge $low)  { $bg = $amber }
            }
            'RECON' {
                $f = [string]$row.Cells['flag'].Value
                if     ($f -eq 'Obsolete (AD+SCCM)')       { $bg = $red }
                elseif ($f -eq 'A verifier')               { $bg = $amber }
                elseif ($f -eq 'Present AD, absent SCCM')  { $bg = $blue }
                elseif ($f -eq 'Present SCCM, absent AD')  { $bg = $purple }
            }
        }
        $row.DefaultCellStyle.BackColor = $bg
    }
}

function Get-StatusClause {
    param($View, [string]$Selected)
    $low  = $script:LowDays
    $high = $script:HighDays
    switch ($View.Kind) {
        'AD' {
            switch ($Selected) {
                'OK'                 { return "status = 'OK'" }
                'Inactif seuil bas'  { return "inactiveDays >= $low" }
                'Inactif seuil haut' { return "inactiveDays >= $high" }
                'Desactive'          { return "enabled = false" }
                'Inconnu'            { return "inactiveDays IS NULL" }
            }
        }
        'SCCM' {
            switch ($Selected) {
                'OK'                 { return "status = 'OK'" }
                'Inactif seuil bas'  { return "inactiveDays >= $low" }
                'Inactif seuil haut' { return "inactiveDays >= $high" }
                'Inconnu'            { return "inactiveDays IS NULL" }
            }
        }
        'RECON' {
            switch ($Selected) {
                'AD seul'            { return "inAD = true AND inSCCM = false" }
                'SCCM seul'          { return "inAD = false AND inSCCM = true" }
                'Les deux'           { return "inAD = true AND inSCCM = true" }
                'Obsolete (AD+SCCM)' { return "flag = 'Obsolete (AD+SCCM)'" }
                'A verifier'         { return "flag = 'A verifier'" }
            }
        }
    }
    return $null
}

function Apply-ViewFilter {
    param($View)
    $clauses = @()

    $q = $View.Filter.Text.Trim()
    $q = $q -replace '([\[\]%*])','[$1]'
    $q = $q.Replace("'","''")
    if ($q) {
        $parts = foreach ($c in $View.Spec.TextCols) { "$c LIKE '%$q%'" }
        $clauses += '(' + ($parts -join ' OR ') + ')'
    }

    $sc = Get-StatusClause -View $View -Selected ([string]$View.Status.SelectedItem)
    if ($sc) { $clauses += $sc }

    $View.Binding.Filter = if ($clauses.Count) { $clauses -join ' AND ' } else { $null }
    Update-ViewColors -View $View
    Update-ViewCounts -View $View
}

function Update-ViewCounts {
    param($View)
    $total = $View.Table.Rows.Count
    if ($total -eq 0) { $View.Count.Text = ''; return }
    $shown = $View.Binding.Count
    $extra = ''
    switch ($View.Kind) {
        'AD' {
            $inLow  = @($View.Table.Select("inactiveDays >= $($script:LowDays)")).Count
            $inHigh = @($View.Table.Select("inactiveDays >= $($script:HighDays)")).Count
            $off    = @($View.Table.Select("enabled = false")).Count
            $extra  = "   |   >=$($script:LowDays)j : $inLow   |   >=$($script:HighDays)j : $inHigh   |   Desactives : $off"
        }
        'SCCM' {
            $inLow  = @($View.Table.Select("inactiveDays >= $($script:LowDays)")).Count
            $inHigh = @($View.Table.Select("inactiveDays >= $($script:HighDays)")).Count
            $extra  = "   |   >=$($script:LowDays)j : $inLow   |   >=$($script:HighDays)j : $inHigh"
        }
        'RECON' {
            $adOnly  = @($View.Table.Select("inAD = true AND inSCCM = false")).Count
            $scOnly  = @($View.Table.Select("inAD = false AND inSCCM = true")).Count
            $obs     = @($View.Table.Select("flag = 'Obsolete (AD+SCCM)'")).Count
            $extra   = "   |   AD seul : $adOnly   |   SCCM seul : $scOnly   |   Obsoletes : $obs"
        }
    }
    $View.Count.Text = "Affiche : $shown / $total$extra"
}

# ==================================================================
#  Fonction - dialogue d'export (colonnes + ordre), generique
# ==================================================================
function Show-ExportDialog {
    param([System.Windows.Forms.DataGridView]$Grid)

    $cols = $Grid.Columns | Sort-Object DisplayIndex

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Colonnes et ordre de l'export"
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
    $clb.Location     = New-Object System.Drawing.Point(12,44)
    $clb.Size         = New-Object System.Drawing.Size(230,338)
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
#  Fonctions - export Excel colore (COM) + repli CSV
# ==================================================================
function ConvertTo-OleColor {
    param([System.Drawing.Color]$Color)
    return [int]$Color.R + ([int]$Color.G -shl 8) + ([int]$Color.B -shl 16)
}

function Export-RecordsToExcel {
    # Records : objets @{ Cells = object[] ; Color = [System.Drawing.Color] ou $null }
    # Cree un .xlsx : entete en-tete stylee, corps en un transfert, lignes colorees,
    # autofiltre, volet fige, colonnes ajustees. Necessite Excel installe (COM).
    param(
        [string[]]$Headers,
        $Records,
        [string]$Path,
        [string]$SheetName = 'Audit'
    )

    $recs     = @($Records)
    $rowCount = $recs.Count
    $colCount = $Headers.Count

    $excel = $null; $wb = $null; $ws = $null
    try   { $excel = New-Object -ComObject Excel.Application }
    catch { return @{ Ok = $false; Message = "Excel introuvable (COM indisponible)." } }

    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        try {
            $clean = ($SheetName -replace '[\\/\?\*\[\]:]',' ')
            if ($clean.Length -gt 31) { $clean = $clean.Substring(0,31) }
            $ws.Name = $clean
        } catch {}

        # En-tetes
        for ($c = 0; $c -lt $colCount; $c++) { $ws.Cells.Item(1, $c+1).Value2 = [string]$Headers[$c] }

        # Corps : tableau 2D transfere en une seule fois (rapide)
        if ($rowCount -gt 0) {
            $arr = New-Object 'object[,]' $rowCount, $colCount
            for ($r = 0; $r -lt $rowCount; $r++) {
                $cells = $recs[$r].Cells
                for ($c = 0; $c -lt $colCount; $c++) {
                    $v = $cells[$c]
                    if     ($null -eq $v -or $v -is [DBNull]) { $v = '' }
                    elseif ($v -is [datetime])                { $v = $v.ToString('yyyy-MM-dd') }
                    elseif ($v -is [bool])                    { $v = if ($v) { 'Oui' } else { 'Non' } }
                    $arr[$r,$c] = $v
                }
            }
            $ws.Range($ws.Cells.Item(2,1), $ws.Cells.Item($rowCount+1, $colCount)).Value2 = $arr
        }

        # Style de l'entete
        $hdr = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$colCount))
        $hdr.Font.Bold          = $true
        $hdr.Font.Color         = 16777215     # blanc
        $hdr.Interior.Color     = 7949855      # bleu fonce (#1F4E79)
        $hdr.HorizontalAlignment = -4108       # centre

        # Coloration des lignes (reprend exactement les couleurs de la grille)
        for ($r = 0; $r -lt $rowCount; $r++) {
            $col = $recs[$r].Color
            if ($col -and -not $col.IsEmpty) {
                $ole = ConvertTo-OleColor $col
                $ws.Range($ws.Cells.Item($r+2,1), $ws.Cells.Item($r+2,$colCount)).Interior.Color = $ole
            }
        }

        # Mise en forme finale
        $used = $ws.UsedRange
        try { [void]$used.AutoFilter() } catch {}
        try {
            $ws.Activate()
            $excel.ActiveWindow.SplitRow    = 1
            $excel.ActiveWindow.FreezePanes = $true
        } catch {}
        try { [void]$used.EntireColumn.AutoFit() } catch {}

        $wb.SaveAs($Path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        return @{ Ok = $true; Message = '' }
    }
    catch {
        return @{ Ok = $false; Message = $_.Exception.Message }
    }
    finally {
        if ($excel) { try { $excel.Quit() } catch {} }
        foreach ($o in @($ws,$wb,$excel)) {
            if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Export-RecordsToCsv {
    param([string[]]$Headers, $Records, [string]$Path)
    $recs = @($Records)
    $rows = foreach ($rec in $recs) {
        $o = [ordered]@{}
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $v = $rec.Cells[$i]
            if ($v -is [DBNull]) { $v = '' }
            elseif ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd') }
            elseif ($v -is [bool])     { $v = if ($v) { 'Oui' } else { 'Non' } }
            $o[[string]$Headers[$i]] = $v
        }
        [pscustomobject]$o
    }
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Get-GridExportRecords {
    # Construit entetes lisibles + enregistrements (valeurs + couleur) depuis la grille affichee
    param($View, [string[]]$Columns)
    $headers = foreach ($cn in $Columns) {
        if ($View.Spec.Headers.ContainsKey($cn)) { $View.Spec.Headers[$cn] } else { $cn }
    }
    $records = foreach ($row in $View.Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $cells = foreach ($cn in $Columns) { $row.Cells[$cn].Value }
        $bg = $row.DefaultCellStyle.BackColor
        [pscustomobject]@{
            Cells = @($cells)
            Color = if ($bg -and -not $bg.IsEmpty) { $bg } else { $null }
        }
    }
    return @{ Headers = @($headers); Records = @($records) }
}

function Get-ReconColorFromFlag {
    param([string]$Flag)
    switch ($Flag) {
        'Obsolete (AD+SCCM)'      { return [System.Drawing.Color]::FromArgb(255,224,224) }
        'A verifier'              { return [System.Drawing.Color]::FromArgb(255,245,204) }
        'Present AD, absent SCCM' { return [System.Drawing.Color]::FromArgb(224,238,255) }
        'Present SCCM, absent AD' { return [System.Drawing.Color]::FromArgb(240,226,255) }
        default                   { return $null }
    }
}

function Invoke-ViewExport {
    param($View, [string]$SuggestedName)
    if ($View.Table.Rows.Count -eq 0) { return }

    $ordered = Show-ExportDialog -Grid $View.Grid
    if (-not $ordered) { return }
    if ($ordered.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Aucune colonne selectionnee.","Export",'OK','Warning'); return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter           = "Excel colore (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv"
    $sfd.Title            = "Enregistrer l'export"
    $sfd.FileName         = "$SuggestedName.xlsx"
    $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $path  = $sfd.FileName
    $data  = Get-GridExportRecords -View $View -Columns $ordered
    $isCsv = ([System.IO.Path]::GetExtension($path) -ieq '.csv')

    if ($isCsv) {
        try {
            Export-RecordsToCsv -Headers $data.Headers -Records $data.Records -Path $path
            $script:lblStatus.Text = "Export CSV : $path"
            if ([System.Windows.Forms.MessageBox]::Show("Export CSV termine.`n`nOuvrir le dossier ?","Export",'YesNo','Information') -eq 'Yes') {
                Start-Process explorer.exe "/select,`"$path`""
            }
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show("Erreur d'export :`n$($_.Exception.Message)","Erreur",'OK','Error')
        }
        return
    }

    $script:lblStatus.Text = "Generation du fichier Excel..."
    $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $res = Export-RecordsToExcel -Headers $data.Headers -Records $data.Records -Path $path -SheetName $View.Kind
    $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($res.Ok) {
        $script:lblStatus.Text = "Export Excel : $path"
        if ([System.Windows.Forms.MessageBox]::Show("Export Excel termine (donnees colorees).`n`nOuvrir le fichier ?","Export Excel",'YesNo','Information') -eq 'Yes') {
            Start-Process $path
        }
    }
    else {
        $csvPath = [System.IO.Path]::ChangeExtension($path, '.csv')
        $ask = [System.Windows.Forms.MessageBox]::Show(
            "Excel n'a pas pu etre utilise :`n$($res.Message)`n`nExporter en CSV (sans couleurs) a la place ?",
            "Excel indisponible", 'YesNo', 'Warning')
        if ($ask -eq 'Yes') {
            try {
                Export-RecordsToCsv -Headers $data.Headers -Records $data.Records -Path $csvPath
                $script:lblStatus.Text = "Export CSV : $csvPath"
                if ([System.Windows.Forms.MessageBox]::Show("Export CSV termine.`n`nOuvrir le dossier ?","Export",'YesNo','Information') -eq 'Yes') {
                    Start-Process explorer.exe "/select,`"$csvPath`""
                }
            } catch {
                [void][System.Windows.Forms.MessageBox]::Show("Erreur d'export CSV :`n$($_.Exception.Message)","Erreur",'OK','Error')
            }
        }
    }
}

# ==================================================================
#  Fonction - selecteur d'OU (arborescence)
# ==================================================================
function Show-OUPicker {
    # Retourne @{ OUs = @(dn...); Recurse = $bool } ou $null
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Choix des unites d'organisation (OU)"
    $dlg.Size            = New-Object System.Drawing.Size(560,600)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.MinimumSize     = New-Object System.Drawing.Size(420,400)

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.CheckBoxes = $true
    $tv.Dock       = 'Fill'
    $tv.Font       = New-Object System.Drawing.Font('Segoe UI',9)

    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Height = 116
    $bottom.Dock   = 'Bottom'

    $chkRecurse = New-Object System.Windows.Forms.CheckBox
    $chkRecurse.Text     = "Inclure les sous-OU (toute l'arborescence sous chaque OU cochee)"
    $chkRecurse.Location = New-Object System.Drawing.Point(12,8)
    $chkRecurse.AutoSize = $true
    $chkRecurse.Checked  = $true

    # Retour visuel : nombre et noms des OU cochees
    $lblSel = New-Object System.Windows.Forms.Label
    $lblSel.Location  = New-Object System.Drawing.Point(12,34)
    $lblSel.Size      = New-Object System.Drawing.Size(520,34)
    $lblSel.ForeColor = [System.Drawing.Color]::FromArgb(0,90,158)
    $lblSel.Anchor    = 'Top,Left,Right'
    $lblSel.Text      = "0 OU cochee(s)"

    $btnExpand = New-Object System.Windows.Forms.Button
    $btnExpand.Text     = "Tout deplier"
    $btnExpand.Location = New-Object System.Drawing.Point(12,74)
    $btnExpand.Size     = New-Object System.Drawing.Size(110,30)

    $btnCollapse = New-Object System.Windows.Forms.Button
    $btnCollapse.Text     = "Tout replier"
    $btnCollapse.Location = New-Object System.Drawing.Point(128,74)
    $btnCollapse.Size     = New-Object System.Drawing.Size(110,30)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Valider"
    $btnOk.Size         = New-Object System.Drawing.Size(120,30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuler"
    $btnCancel.Size         = New-Object System.Drawing.Size(100,30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $bottom.Controls.AddRange(@($chkRecurse,$lblSel,$btnExpand,$btnCollapse,$btnOk,$btnCancel))
    $dlg.Controls.Add($tv)
    $dlg.Controls.Add($bottom)
    $bottom.BringToFront()
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    # Placement fiable des boutons a droite (recalcule a l'affichage et au redimensionnement)
    $placeButtons = {
        $btnCancel.Top  = 74
        $btnOk.Top      = 74
        $btnCancel.Left = $bottom.ClientSize.Width - $btnCancel.Width - 12
        $btnOk.Left     = $btnCancel.Left - $btnOk.Width - 8
    }
    $bottom.Add_Resize($placeButtons)
    $dlg.Add_Shown($placeButtons)

    $btnExpand.Add_Click(  { $tv.ExpandAll() })
    $btnCollapse.Add_Click({ $tv.CollapseAll() })

    # Recalcule le compteur + la liste des noms d'OU cochees
    $updateSel = {
        $names = @()
        $stack = New-Object System.Collections.Stack
        foreach ($x in $tv.Nodes) { $stack.Push($x) }
        while ($stack.Count) {
            $node = $stack.Pop()
            if ($node.Checked -and $node.Tag) { $names += [string]$node.Text }
            foreach ($c in $node.Nodes) { $stack.Push($c) }
        }
        if ($names.Count -eq 0) { $lblSel.Text = "0 OU cochee(s)"; return }
        $show   = $names | Select-Object -First 6
        $suffix = if ($names.Count -gt 6) { " (+$($names.Count - 6))" } else { '' }
        $lblSel.Text = "$($names.Count) OU cochee(s) : " + ($show -join ', ') + $suffix
    }

    # Coche/decoche recursif de TOUTE la descendance quand on coche un noeud
    $tv.Add_AfterCheck({
        param($s,$e)
        if ($e.Action -eq [System.Windows.Forms.TreeViewAction]::Unknown) { return }
        $st = New-Object System.Collections.Stack
        foreach ($c in $e.Node.Nodes) { $st.Push($c) }
        while ($st.Count) {
            $c = $st.Pop()
            $c.Checked = $e.Node.Checked
            foreach ($g in $c.Nodes) { $st.Push($g) }
        }
        & $updateSel
    })

    # --- Construction de l'arbre a partir des DN ---
    $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $rootDN = $domain.DistinguishedName
    } catch { $rootDN = '' }

    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Text = if ($rootDN) { "$($domain.DNSRoot)  ($rootDN)" } else { "Domaine" }
    $rootNode.Tag  = $rootDN
    [void]$tv.Nodes.Add($rootNode)

    $map = @{}
    if ($rootDN) { $map[$rootDN] = $rootNode }

    $ous = Get-ADOrganizationalUnit -Filter * -ErrorAction Stop |
           Sort-Object { ($_.DistinguishedName -split ',').Count }

    foreach ($ou in $ous) {
        $dn     = $ou.DistinguishedName
        $parent = $dn.Substring($dn.IndexOf(',') + 1)
        $node   = New-Object System.Windows.Forms.TreeNode
        $node.Text = $ou.Name
        $node.Tag  = $dn
        if ($map.ContainsKey($parent)) { [void]$map[$parent].Nodes.Add($node) }
        else                           { [void]$rootNode.Nodes.Add($node) }
        $map[$dn] = $node
    }

    $rootNode.Expand()
    $dlg.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $selected = @()
    $stack = New-Object System.Collections.Stack
    foreach ($n in $tv.Nodes) { $stack.Push($n) }
    while ($stack.Count) {
        $n = $stack.Pop()
        if ($n.Checked -and $n.Tag) { $selected += [string]$n.Tag }
        foreach ($c in $n.Nodes) { $stack.Push($c) }
    }
    # Si recurse actif, on dedoublonne les OU imbriquees (on garde les plus hautes)
    $selected = $selected | Select-Object -Unique
    if ($chkRecurse.Checked -and $selected.Count -gt 1) {
        $selected = $selected | Where-Object {
            $me = $_
            -not ($selected | Where-Object { $_ -ne $me -and $me.EndsWith(",$_") })
        }
    }

    return @{ OUs = @($selected); Recurse = [bool]$chkRecurse.Checked }
}

# ==================================================================
#  Fonction - construction d'un onglet (vue)
# ==================================================================
function New-AuditView {
    param([hashtable]$ViewSpec)

    $table   = New-TableFromSpec -ViewSpec $ViewSpec
    $binding = New-Object System.Windows.Forms.BindingSource
    $binding.DataSource = $table

    # --- Barre de filtre (haut de l'onglet) ---
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Height = 40
    $bar.Dock   = 'Top'

    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = "Filtre :"; $lblF.AutoSize = $true
    $lblF.Location = New-Object System.Drawing.Point(8,12)

    $txtF = New-Object System.Windows.Forms.TextBox
    $txtF.Location = New-Object System.Drawing.Point(58,9)
    $txtF.Size     = New-Object System.Drawing.Size(230,24)

    $btnX = New-Object System.Windows.Forms.Button
    $btnX.Text = "X"; $btnX.Size = New-Object System.Drawing.Size(26,24)
    $btnX.Location = New-Object System.Drawing.Point(292,8)

    $lblS = New-Object System.Windows.Forms.Label
    $lblS.Text = "Statut :"; $lblS.AutoSize = $true
    $lblS.Location = New-Object System.Drawing.Point(330,12)

    $cboS = New-Object System.Windows.Forms.ComboBox
    $cboS.DropDownStyle = 'DropDownList'
    $cboS.Location = New-Object System.Drawing.Point(382,9)
    $cboS.Size     = New-Object System.Drawing.Size(180,24)
    [void]$cboS.Items.AddRange($ViewSpec.StatusItems)
    $cboS.SelectedIndex = 0

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.AutoSize  = $true
    $lblCount.Anchor    = 'Top,Right'
    $lblCount.TextAlign = 'MiddleRight'
    $lblCount.Location  = New-Object System.Drawing.Point(580,12)

    $btnExp = New-Object System.Windows.Forms.Button
    $btnExp.Text   = "Export Excel..."
    $btnExp.Size   = New-Object System.Drawing.Size(120,28)
    $btnExp.Anchor = 'Top,Right'
    $btnExp.Enabled = $false

    $bar.Controls.AddRange(@($lblF,$txtF,$btnX,$lblS,$cboS,$lblCount,$btnExp))

    # --- Grille ---
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock                      = 'Fill'
    $grid.DataSource                = $binding
    $grid.AllowUserToAddRows        = $false
    $grid.AllowUserToDeleteRows     = $false
    $grid.ReadOnly                  = $true
    $grid.AllowUserToOrderColumns   = $true
    $grid.AllowUserToResizeColumns  = $true
    $grid.AllowUserToResizeRows     = $false
    $grid.AutoSizeColumnsMode       = 'AllCells'
    $grid.SelectionMode             = 'FullRowSelect'
    $grid.MultiSelect               = $true
    $grid.RowHeadersVisible         = $false
    $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
    $grid.ClipboardCopyMode         = 'EnableWithoutHeaderText'
    $grid.BorderStyle               = 'None'
    $grid.BackgroundColor           = [System.Drawing.Color]::White
    $grid.Font                      = New-Object System.Drawing.Font('Segoe UI',9)
    $grid.RowTemplate.Height        = 24
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(247,249,252)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)

    $headers  = $ViewSpec.Headers
    $dateCols = $ViewSpec.DateCols
    $grid.Add_DataBindingComplete({
        foreach ($c in $grid.Columns) {
            if ($headers.ContainsKey($c.Name)) { $c.HeaderText = $headers[$c.Name] }
            if ($dateCols -contains $c.Name)   { $c.DefaultCellStyle.Format = 'yyyy-MM-dd' }
        }
        if ($grid.Columns.Count -gt 0) { $grid.Columns[0].Frozen = $true }
    }.GetNewClosure())

    # --- Menu contextuel ---
    $ctx = New-Object System.Windows.Forms.ContextMenuStrip
    $miCell = $ctx.Items.Add("Copier la cellule")
    $miRow  = $ctx.Items.Add("Copier la ligne")
    [void]$ctx.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miAll  = $ctx.Items.Add("Tout selectionner")
    $grid.ContextMenuStrip = $ctx

    $grid.Add_CellMouseDown({
        param($s,$e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
            $row = $grid.Rows[$e.RowIndex]
            if (-not $row.Selected) { $grid.ClearSelection(); $row.Selected = $true }
            if ($e.ColumnIndex -ge 0) { $grid.CurrentCell = $row.Cells[$e.ColumnIndex] }
        }
    }.GetNewClosure())

    $miCell.Add_Click({
        if ($grid.CurrentCell) { [void](Copy-ToClipboard ([string]$grid.CurrentCell.Value)) }
    }.GetNewClosure())
    $miRow.Add_Click({
        if (-not $grid.CurrentRow) { return }
        $vals = foreach ($c in ($grid.Columns | Sort-Object DisplayIndex)) {
            if ($c.Visible) { [string]$grid.CurrentRow.Cells[$c.Name].Value }
        }
        [void](Copy-ToClipboard ($vals -join "`t"))
    }.GetNewClosure())
    $miAll.Add_Click({ $grid.SelectAll() }.GetNewClosure())

    # --- Assemblage de l'onglet ---
    $page = New-Object System.Windows.Forms.TabPage
    $page.Padding = New-Object System.Windows.Forms.Padding(2)
    $page.Controls.Add($grid)
    $page.Controls.Add($bar)
    $bar.BringToFront()

    $view = [pscustomobject]@{
        Kind    = $ViewSpec.Kind
        Spec    = $ViewSpec
        Table   = $table
        Binding = $binding
        Grid    = $grid
        Filter  = $txtF
        Status  = $cboS
        Count   = $lblCount
        Export  = $btnExp
        Page    = $page
        Bar     = $bar
    }

    # --- Evenements de la vue ---
    $txtF.Add_TextChanged({ Apply-ViewFilter -View $view }.GetNewClosure())
    $cboS.Add_SelectedIndexChanged({ Apply-ViewFilter -View $view }.GetNewClosure())
    $btnX.Add_Click({ $txtF.Clear(); $cboS.SelectedIndex = 0 }.GetNewClosure())
    $btnExp.Add_Click({ Invoke-ViewExport -View $view -SuggestedName $ViewSpec.ExportName }.GetNewClosure())

    return $view
}

# ==================================================================
#  Interface principale
# ==================================================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text          = "Audit du parc - Active Directory + SCCM"
$Form.Size          = New-Object System.Drawing.Size(1200,760)
$Form.StartPosition = 'CenterScreen'
$Form.MinimumSize   = New-Object System.Drawing.Size(900,560)
$script:MainForm    = $Form

# --- Barre haute ---
$Top = New-Object System.Windows.Forms.Panel
$Top.Height = 84
$Top.Dock   = 'Top'

$btnOU = New-Object System.Windows.Forms.Button
$btnOU.Text     = "Choisir les OU (AD)..."
$btnOU.Location = New-Object System.Drawing.Point(10,10)
$btnOU.Size     = New-Object System.Drawing.Size(160,28)

$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text      = "Aucune OU selectionnee"
$lblOU.Location  = New-Object System.Drawing.Point(180,15)
$lblOU.AutoSize  = $true
$lblOU.ForeColor = [System.Drawing.Color]::FromArgb(0,90,158)
$ttMain = New-Object System.Windows.Forms.ToolTip
$ttMain.AutoPopDelay = 30000
$ttMain.SetToolTip($lblOU, "Aucune OU selectionnee")

$lblSeuilBas = New-Object System.Windows.Forms.Label
$lblSeuilBas.Text     = "Seuil bas (j) :"
$lblSeuilBas.Location = New-Object System.Drawing.Point(10,50)
$lblSeuilBas.AutoSize = $true

$numLow = New-Object System.Windows.Forms.NumericUpDown
$numLow.Location = New-Object System.Drawing.Point(96,47)
$numLow.Size     = New-Object System.Drawing.Size(60,24)
$numLow.Minimum  = 1; $numLow.Maximum = 3650; $numLow.Value = $script:LowDays

$lblSeuilHaut = New-Object System.Windows.Forms.Label
$lblSeuilHaut.Text     = "Seuil haut (j) :"
$lblSeuilHaut.Location = New-Object System.Drawing.Point(170,50)
$lblSeuilHaut.AutoSize = $true

$numHigh = New-Object System.Windows.Forms.NumericUpDown
$numHigh.Location = New-Object System.Drawing.Point(262,47)
$numHigh.Size     = New-Object System.Drawing.Size(60,24)
$numHigh.Minimum  = 1; $numHigh.Maximum = 3650; $numHigh.Value = $script:HighDays

# Case a cocher : enrichissement materiel via l'inventaire SCCM
$chkHardware = New-Object System.Windows.Forms.CheckBox
$chkHardware.Text     = "Modele + n. serie (inventaire SCCM)"
$chkHardware.Location = New-Object System.Drawing.Point(335,49)
$chkHardware.AutoSize = $true
$chkHardware.Checked  = $script:IncludeHardware
$chkHardware.Enabled  = $script:SccmAvailable
$ttMain.SetToolTip($chkHardware, "Interroge SMS_G_System_COMPUTER_SYSTEM et SMS_G_System_PC_BIOS`r`npour remplir Modele / Fabricant / N. serie. Peut rallonger le traitement.")

$lblSccm = New-Object System.Windows.Forms.Label
$lblSccm.Location = New-Object System.Drawing.Point(600,50)
$lblSccm.AutoSize = $true
$lblSccm.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
$lblSccm.Text = if ($script:SccmAvailable) { "SCCM : disponible" } else { "SCCM : indisponible (AD seul)" }
$lblSccm.ForeColor = if ($script:SccmAvailable) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::Firebrick }
$sccmTip = if ($script:SccmAvailable) { "Le parc SCCM complet est interroge au lancement de l'audit.`r`nPositionne-toi sur le lecteur du site (PS <CODE>:\) pour que Get-CMDevice renvoie des postes." } else { "Get-CMDevice introuvable : onglets SCCM et Rapprochement vides." }
$ttMain.SetToolTip($lblSccm, $sccmTip)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text    = "Lancer l'audit"
$btnRun.Size    = New-Object System.Drawing.Size(140,28)
$btnRun.Anchor  = 'Top,Right'
$btnRun.Enabled = $false

$btnAsset = New-Object System.Windows.Forms.Button
$btnAsset.Text    = "Export gestion de parc..."
$btnAsset.Size    = New-Object System.Drawing.Size(180,28)
$btnAsset.Anchor  = 'Top,Right'
$btnAsset.Enabled = $false

$Top.Controls.AddRange(@($btnOU,$lblOU,$lblSeuilBas,$numLow,$lblSeuilHaut,$numHigh,$chkHardware,$lblSccm,$btnRun,$btnAsset))

# --- Onglets ---
$Tabs = New-Object System.Windows.Forms.TabControl
$Tabs.Dock = 'Fill'

$viewAD    = New-AuditView -ViewSpec $script:Spec.AD
$viewSCCM  = New-AuditView -ViewSpec $script:Spec.SCCM
$viewRECON = New-AuditView -ViewSpec $script:Spec.RECON

$viewAD.Page.Text    = "AD"
$viewSCCM.Page.Text  = "SCCM"
$viewRECON.Page.Text = "Rapprochement AD / SCCM"

[void]$Tabs.TabPages.Add($viewAD.Page)
[void]$Tabs.TabPages.Add($viewSCCM.Page)
[void]$Tabs.TabPages.Add($viewRECON.Page)

$script:Views = @{ AD = $viewAD; SCCM = $viewSCCM; RECON = $viewRECON }

# --- Barre d'etat ---
$Status    = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Text      = if ($script:SccmAvailable) { "Pret. Choisissez des OU puis lancez l'audit." } else { "Pret (mode AD seul). Choisissez des OU puis lancez l'audit." }
$script:lblStatus.Spring    = $true
$script:lblStatus.TextAlign = 'MiddleLeft'
$Progress  = New-Object System.Windows.Forms.ToolStripProgressBar
$Progress.Size = New-Object System.Drawing.Size(240,16)
[void]$Status.Items.Add($script:lblStatus)
[void]$Status.Items.Add($Progress)

# Ordre d'ajout : Fill d'abord, puis Top / Bottom au premier plan
$Form.Controls.Add($Tabs)
$Form.Controls.Add($Top);    $Top.BringToFront()
$Form.Controls.Add($Status); $Status.BringToFront()

# ==================================================================
#  Evenements
# ==================================================================
$Form.Add_Shown({
    # Positionne les boutons ancres a droite de la barre haute
    $btnRun.Left   = $Top.ClientSize.Width - $btnRun.Width - 10
    $btnRun.Top    = 10
    $btnAsset.Left = $Top.ClientSize.Width - $btnAsset.Width - 10
    $btnAsset.Top  = 46

    # Positionne export + compteur a droite dans chaque barre d'onglet
    foreach ($v in @($viewAD,$viewSCCM,$viewRECON)) {
        $v.Export.Left = $v.Bar.ClientSize.Width - $v.Export.Width - 8
        $v.Export.Top  = 8
        $v.Count.Left  = $v.Export.Left - $v.Count.Width - 12
        $v.Count.Top   = 12
    }
})

$btnOU.Add_Click({
    $res = Show-OUPicker
    if (-not $res) { return }
    if ($res.OUs.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Aucune OU cochee.","Choix des OU",'OK','Information')
        return
    }
    $script:SelectedOUs = $res.OUs
    $script:Recurse     = $res.Recurse
    $scopeTxt = if ($res.Recurse) { "avec sous-OU" } else { "niveau direct" }

    # Noms courts (composant OU= de tete de chaque DN) pour l'affichage
    $names = foreach ($dn in $res.OUs) {
        $first = ($dn -split ',')[0]
        ($first -replace '^\s*OU=','').Trim()
    }
    $shown  = @($names | Select-Object -First 5)
    $suffix = if ($names.Count -gt 5) { " (+$($names.Count - 5))" } else { '' }
    $lblOU.Text = "$($res.OUs.Count) OU ($scopeTxt) : " + ($shown -join ', ') + $suffix

    # Liste complete des DN dans l'infobulle
    $ttMain.SetToolTip($lblOU, ($res.OUs -join "`r`n"))
    $btnRun.Enabled = $true
})

$btnRun.Add_Click({
    if ($script:SelectedOUs.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Choisissez au moins une OU.","Audit",'OK','Warning'); return
    }
    $script:LowDays  = [int]$numLow.Value
    $script:HighDays = [int]$numHigh.Value
    $script:IncludeHardware = [bool]$chkHardware.Checked
    if ($script:HighDays -lt $script:LowDays) {
        [void][System.Windows.Forms.MessageBox]::Show("Le seuil haut doit etre >= au seuil bas.","Seuils",'OK','Warning'); return
    }

    $btnOU.Enabled = $false; $btnRun.Enabled = $false; $btnAsset.Enabled = $false
    foreach ($v in $script:Views.Values) { $v.Export.Enabled = $false }
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Build-AllTables -Progress $Progress -Status $script:lblStatus
        foreach ($v in @($viewAD,$viewSCCM,$viewRECON)) {
            Apply-ViewFilter -View $v
            $v.Export.Enabled = ($v.Table.Rows.Count -gt 0)
        }
        $btnAsset.Enabled = ($viewRECON.Table.Rows.Count -gt 0)
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show("Erreur pendant l'audit :`n$($_.Exception.Message)","Erreur",'OK','Error')
        $script:lblStatus.Text = "Erreur."
    }
    finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnOU.Enabled = $true; $btnRun.Enabled = $true
    }
})

# Export consolide "gestion de parc" = table de rapprochement complete, en Excel colore
$btnAsset.Add_Click({
    if ($viewRECON.Table.Rows.Count -eq 0) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter           = "Excel colore (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv"
    $sfd.Title            = "Export consolide gestion de parc"
    $sfd.FileName         = "Gestion_Parc_Consolide.xlsx"
    $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $path = $sfd.FileName

    # Toutes les colonnes de la table de rapprochement, en-tetes lisibles, couleur par diagnostic
    $cols    = @($viewRECON.Table.Columns | ForEach-Object { $_.ColumnName })
    $headers = foreach ($c in $cols) { if ($viewRECON.Spec.Headers.ContainsKey($c)) { $viewRECON.Spec.Headers[$c] } else { $c } }
    $records = foreach ($dr in $viewRECON.Table.Rows) {
        $cells = foreach ($c in $cols) { $dr[$c] }
        [pscustomobject]@{
            Cells = @($cells)
            Color = Get-ReconColorFromFlag ([string]$dr['flag'])
        }
    }

    $isCsv = ([System.IO.Path]::GetExtension($path) -ieq '.csv')
    if ($isCsv) {
        try {
            Export-RecordsToCsv -Headers @($headers) -Records $records -Path $path
            $script:lblStatus.Text = "Export gestion de parc (CSV) : $path"
            if ([System.Windows.Forms.MessageBox]::Show("Export termine.`n`nOuvrir le dossier ?","Gestion de parc",'YesNo','Information') -eq 'Yes') {
                Start-Process explorer.exe "/select,`"$path`""
            }
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show("Erreur d'export :`n$($_.Exception.Message)","Erreur",'OK','Error')
        }
        return
    }

    $script:lblStatus.Text = "Generation du fichier Excel consolide..."
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $res = Export-RecordsToExcel -Headers @($headers) -Records $records -Path $path -SheetName 'Gestion parc'
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($res.Ok) {
        $script:lblStatus.Text = "Export gestion de parc (Excel) : $path"
        if ([System.Windows.Forms.MessageBox]::Show("Export consolide termine (donnees colorees).`n`nOuvrir le fichier ?","Gestion de parc",'YesNo','Information') -eq 'Yes') {
            Start-Process $path
        }
    }
    else {
        $csvPath = [System.IO.Path]::ChangeExtension($path, '.csv')
        if ([System.Windows.Forms.MessageBox]::Show("Excel indisponible :`n$($res.Message)`n`nExporter en CSV a la place ?","Excel indisponible",'YesNo','Warning') -eq 'Yes') {
            try {
                Export-RecordsToCsv -Headers @($headers) -Records $records -Path $csvPath
                $script:lblStatus.Text = "Export gestion de parc (CSV) : $csvPath"
                if ([System.Windows.Forms.MessageBox]::Show("Export termine.`n`nOuvrir le dossier ?","Gestion de parc",'YesNo','Information') -eq 'Yes') {
                    Start-Process explorer.exe "/select,`"$csvPath`""
                }
            } catch {
                [void][System.Windows.Forms.MessageBox]::Show("Erreur d'export CSV :`n$($_.Exception.Message)","Erreur",'OK','Error')
            }
        }
    }
})

# ==================================================================
#  Lancement
# ==================================================================
[void]$Form.ShowDialog()
$Form.Dispose()
