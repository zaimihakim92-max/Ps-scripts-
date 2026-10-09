<#
.SYNOPSIS
    Audit des licences du tenant (true-up) - Interface graphique, arborescence depliable.
    Connexion : Az (Connect-AzAccount). Interroge Microsoft Graph via un jeton obtenu depuis Az
    (aucun module Graph requis).

.DESCRIPTION
    Pour chaque licence (SKU) : total / consomme / disponible, affectations DIRECTES, affectations
    HERITEES regroupees par groupe, et DOUBLONS (direct + heritage). Objectif true-up.

    Affichage optimise :
      - compteurs sur chaque licence (direct / groupe / doublons, taux d'utilisation),
      - categorie Payante / Essai-gratuite,
      - case "Masquer essais & non consommees" pour se concentrer sur l'essentiel,
      - couleurs : sur-provisionnement (rouge), tendu >= 90 % (orange), doublons (rouge).

    Exports :
      - Rapport HTML (synthese propre, imprimable) : livrable true-up.
      - Classeur .xlsx NATIF (sans Excel ni module) : donnees completes a plat.
      - CSV : donnees completes a plat.

.PREREQUISITES
    - Module Az (Az.Accounts) + session Connect-AzAccount avec un compte capable de lire l'annuaire.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:GraphToken = $null
$script:Audit      = @()
$script:Flat       = New-Object System.Collections.Generic.List[object]
$script:Account    = ""

# ==================================================================
# Noms conviviaux de SKU
# ==================================================================
$script:SkuNames = @{
    'SPE_E3'='Microsoft 365 E3'; 'SPE_E5'='Microsoft 365 E5'; 'SPE_F1'='Microsoft 365 F3'
    'ENTERPRISEPACK'='Office 365 E3'; 'ENTERPRISEPREMIUM'='Office 365 E5'; 'STANDARDPACK'='Office 365 E1'; 'DESKLESSPACK'='Office 365 F3'
    'EMS'='EMS E3'; 'EMSPREMIUM'='EMS E5'; 'AAD_PREMIUM'='Entra ID P1'; 'AAD_PREMIUM_P2'='Entra ID P2'
    'POWER_BI_PRO'='Power BI Pro'; 'POWER_BI_STANDARD'='Power BI (free)'; 'FLOW_FREE'='Power Automate Free'; 'FLOW_PER_USER'='Power Automate per user'
    'PROJECTPROFESSIONAL'='Project Plan 3'; 'PROJECTPREMIUM'='Project Plan 5'; 'PROJECT_P1'='Project Plan 1'
    'VISIOCLIENT'='Visio Plan 2'; 'VISIO_PLAN1_DEPT'='Visio Plan 1'; 'PBI_PREMIUM_PER_USER'='Power BI Premium per user'
    'EXCHANGESTANDARD'='Exchange Online P1'; 'EXCHANGEENTERPRISE'='Exchange Online P2'
    'MCOSTANDARD'='Skype Entreprise'; 'MCOEV'='Teams Phone'; 'MCOMEETADV'='Teams Audio Conferencing'; 'TEAMS_EXPLORATORY'='Teams Exploratory'
    'PHONESYSTEM_VIRTUALUSER'='Teams Phone Resource Account'; 'MDATP_XPLAT'='Defender for Endpoint'
    'WIN_DEF_ATP'='Defender for Endpoint'; 'ATP_ENTERPRISE'='Defender for Office P1'
    'Microsoft_365_Copilot'='Microsoft 365 Copilot'; 'Microsoft_Teams_Rooms_Pro'='Teams Rooms Pro'
    'O365_BUSINESS_PREMIUM'='M365 Business Standard'; 'O365_BUSINESS_ESSENTIALS'='M365 Business Basic'; 'SPB'='M365 Business Premium'
}
function Get-SkuName { param([string]$part) if ($script:SkuNames.ContainsKey($part)) { return $script:SkuNames[$part] } return $part }

# Essai / gratuit (categorisation, ne masque rien par defaut)
function Test-FreeSku {
    param([string]$part)
    return ($part -match '(?i)(TRIAL|VIRAL|PREVIEW|_FREE$|FLOW_FREE|POWER_BI_STANDARD|_DEV$|_IW$|ADHOC|SPZA|WINDOWS_STORE|MCOPSTNC|SHAREPOINTSTORAGE|STREAM|FORMS_PRO|VIVA_INSIGHTS_COHORT|Win10_VDA|ONBOARDING)')
}

# ==================================================================
# Az / Graph
# ==================================================================
function Test-AzConnected { try { return [bool](Get-AzContext -ErrorAction Stop) } catch { return $false } }

function Get-GraphToken {
    $t = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
    if ($t.Token -is [System.Security.SecureString]) { return (New-Object System.Net.NetworkCredential("", $t.Token)).Password }
    return [string]$t.Token
}

function Invoke-GraphAll {
    param([string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $headers = @{ Authorization = "Bearer $script:GraphToken"; 'ConsistencyLevel' = 'eventual' }
    $next = $Uri
    while ($next) {
        $resp = Invoke-RestMethod -Method GET -Uri $next -Headers $headers -ErrorAction Stop
        if ($null -ne $resp.value) { foreach ($v in $resp.value) { $items.Add($v) | Out-Null } }
        $next = $resp.'@odata.nextLink'
        [System.Windows.Forms.Application]::DoEvents()
    }
    return $items
}

# ==================================================================
# Form
# ==================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Audit des licences du tenant (true-up)"
$form.Size = New-Object System.Drawing.Size(1020, 830)
$form.MinimumSize = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connexion Azure"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Etat : verification..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(800,18); $lblConn.Anchor="Top,Left,Right"

$btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text="Lancer l'audit"; $btnLoad.Location=New-Object System.Drawing.Point(15,50); $btnLoad.Size=New-Object System.Drawing.Size(160,28)
$btnLoad.BackColor=[System.Drawing.Color]::FromArgb(0,120,215); $btnLoad.ForeColor=[System.Drawing.Color]::White; $btnLoad.FlatStyle="Flat"; $form.Controls.Add($btnLoad)
Add-Lbl $form "Filtre :" 185 55 | Out-Null
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilter.Location=New-Object System.Drawing.Point(235,52); $txtFilter.Size=New-Object System.Drawing.Size(210,23); $form.Controls.Add($txtFilter)
$chkHide = New-Object System.Windows.Forms.CheckBox; $chkHide.Text="Masquer essais & non consommees"; $chkHide.Location=New-Object System.Drawing.Point(455,54); $chkHide.AutoSize=$true; $form.Controls.Add($chkHide)
$btnExpand = New-Object System.Windows.Forms.Button; $btnExpand.Text="Deplier"; $btnExpand.Location=New-Object System.Drawing.Point(690,50); $btnExpand.Size=New-Object System.Drawing.Size(80,28); $form.Controls.Add($btnExpand)
$btnCollapse = New-Object System.Windows.Forms.Button; $btnCollapse.Text="Replier"; $btnCollapse.Location=New-Object System.Drawing.Point(775,50); $btnCollapse.Size=New-Object System.Drawing.Size(80,28); $form.Controls.Add($btnCollapse)
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text="Exporter..."; $btnExport.Location=New-Object System.Drawing.Point(865,50); $btnExport.Size=New-Object System.Drawing.Size(130,28); $btnExport.Anchor="Top,Right"; $btnExport.Enabled=$false; $form.Controls.Add($btnExport)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location=New-Object System.Drawing.Point(15,88); $tree.Size=New-Object System.Drawing.Size(985,570); $tree.Anchor="Top,Bottom,Left,Right"
$tree.HideSelection=$false; $tree.Font=New-Object System.Drawing.Font("Segoe UI",9)
$form.Controls.Add($tree)

Add-Lbl $form "Journal :" 15 666 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,686); $txtLog.Size=New-Object System.Drawing.Size(985,100); $txtLog.Anchor="Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# UI helpers
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

function Populate-Tree {
    param([string]$Filter = "")
    $f = $Filter.Trim().ToLower()
    $hide = $chkHide.Checked
    $tree.BeginUpdate(); $tree.Nodes.Clear()
    foreach ($lic in $script:Audit) {
        if ($hide -and ($lic.Free -or $lic.Consumed -eq 0)) { continue }
        $licMatch = (-not $f) -or ($lic.Part.ToLower().Contains($f)) -or ($lic.Friendly.ToLower().Contains($f))

        $root = New-Object System.Windows.Forms.TreeNode
        $root.Text = "{0} ({1})  -  {2}/{3} · dispo {4} · util {5}%  |  direct {6} · groupe {7} · doublons {8}" -f `
            $lic.Friendly, $lic.Part, $lic.Consumed, $lic.Enabled, $lic.Avail, $lic.Util, $lic.DirectCount, $lic.GroupCount, $lic.DupCount
        if ($lic.Avail -lt 0) { $root.ForeColor = [System.Drawing.Color]::Firebrick }
        elseif ($lic.Free) { $root.ForeColor = [System.Drawing.Color]::Gray }
        elseif ($lic.Consumed -eq 0) { $root.ForeColor = [System.Drawing.Color]::DimGray }

        $added = $false
        $direct = @($lic.Direct | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) })
        if ($direct.Count -gt 0) {
            $nd = New-Object System.Windows.Forms.TreeNode ("Direct ({0})" -f $direct.Count)
            foreach ($u in $direct) { $un=New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}{2}" -f $u.Name,$u.Upn,$(if($u.State -and $u.State -ne 'Active'){" [$($u.State)]"}else{""})); if($u.State -and $u.State -ne 'Active'){$un.ForeColor=[System.Drawing.Color]::DarkOrange}; [void]$nd.Nodes.Add($un) }
            [void]$root.Nodes.Add($nd); $added=$true
        }
        foreach ($gid in $lic.Groups.Keys) {
            $gname = $lic.GroupNames[$gid]
            $gusers = @($lic.Groups[$gid] | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) -or ("$gname").ToLower().Contains($f) })
            if ($gusers.Count -gt 0) {
                $ng = New-Object System.Windows.Forms.TreeNode ("Via groupe : {0} ({1})" -f $gname, $gusers.Count); $ng.ForeColor=[System.Drawing.Color]::FromArgb(0,90,160)
                foreach ($u in $gusers) { [void]$ng.Nodes.Add((New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}" -f $u.Name,$u.Upn))) }
                [void]$root.Nodes.Add($ng); $added=$true
            }
        }
        $dups = @($lic.Duplicates | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) })
        if ($dups.Count -gt 0) {
            $ndup = New-Object System.Windows.Forms.TreeNode ("Doublons direct + heritage ({0})" -f $dups.Count); $ndup.ForeColor=[System.Drawing.Color]::Firebrick
            foreach ($u in $dups) { [void]$ndup.Nodes.Add((New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}" -f $u.Name,$u.Upn))) }
            [void]$root.Nodes.Add($ndup); $added=$true
        }
        if ($licMatch -or $added) { [void]$tree.Nodes.Add($root) }
    }
    $tree.EndUpdate()
}

# ==================================================================
# Audit model
# ==================================================================
function Start-Audit {
    $btnLoad.Enabled=$false
    try { $script:GraphToken = Get-GraphToken }
    catch { Write-Log "ERREUR jeton Graph : $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Impossible d'obtenir un jeton Graph depuis Az.`n$($_.Exception.Message)","Audit","OK","Error")|Out-Null; $btnLoad.Enabled=$true; return }
    try { $script:Account = (Get-AzContext).Account.Id } catch { }

    try {
        Write-Log "Lecture des licences (subscribedSkus)..."
        $skus = Invoke-GraphAll "https://graph.microsoft.com/v1.0/subscribedSkus"
        Write-Log "Lecture des groupes avec licences..."
        $groups = Invoke-GraphAll "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,assignedLicenses&`$top=999"
        Write-Log "Lecture des utilisateurs et de leurs affectations..."
        $users = Invoke-GraphAll "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,licenseAssignmentStates&`$top=999"
        Write-Log ("Recu : {0} SKU, {1} groupe(s), {2} utilisateur(s)." -f $skus.Count, $groups.Count, $users.Count)
    } catch {
        Write-Log "ERREUR Graph : $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Appel Graph refuse.`n$($_.Exception.Message)`n`nLe compte connecte doit pouvoir lire l'annuaire.","Audit","OK","Error")|Out-Null
        $btnLoad.Enabled=$true; return
    }

    $gName = @{}
    foreach ($g in $groups) { if ($g.id) { $gName[$g.id] = $g.displayName } }

    $model = @{}
    foreach ($s in $skus) {
        $model[$s.skuId] = [PSCustomObject]@{
            SkuId=$s.skuId; Part=$s.skuPartNumber; Friendly=(Get-SkuName $s.skuPartNumber); Free=(Test-FreeSku $s.skuPartNumber)
            Enabled=[int]$s.prepaidUnits.enabled; Consumed=[int]$s.consumedUnits
            Direct=(New-Object System.Collections.Generic.List[object]); Groups=@{}; GroupNames=@{}
            DirectUpns=(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
            GroupUpns=(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
            Duplicates=(New-Object System.Collections.Generic.List[object])
            Avail=0; Util=0; DirectCount=0; GroupCount=0; DupCount=0
        }
    }

    foreach ($u in $users) {
        if (-not $u.licenseAssignmentStates) { continue }
        foreach ($st in $u.licenseAssignmentStates) {
            $sid = $st.skuId; if (-not $model.ContainsKey($sid)) { continue }
            $lic = $model[$sid]
            $uObj = [PSCustomObject]@{ Name=$u.displayName; Upn=$u.userPrincipalName; State=$st.state }
            if ([string]::IsNullOrEmpty($st.assignedByGroup)) { $lic.Direct.Add($uObj)|Out-Null; [void]$lic.DirectUpns.Add($u.userPrincipalName) }
            else {
                $gid = $st.assignedByGroup
                if (-not $lic.Groups.ContainsKey($gid)) { $lic.Groups[$gid]=New-Object System.Collections.Generic.List[object]; $lic.GroupNames[$gid]=$(if($gName.ContainsKey($gid)){$gName[$gid]}else{"(groupe $gid)"}) }
                $lic.Groups[$gid].Add($uObj)|Out-Null; [void]$lic.GroupUpns.Add($u.userPrincipalName)
            }
        }
    }

    $script:Flat.Clear()
    foreach ($lic in $model.Values) {
        foreach ($upn in $lic.DirectUpns) { if ($lic.GroupUpns.Contains($upn)) { $nm=($lic.Direct | Where-Object { $_.Upn -ieq $upn } | Select-Object -First 1).Name; $lic.Duplicates.Add([PSCustomObject]@{Name=$nm;Upn=$upn})|Out-Null } }
        $lic.Avail = $lic.Enabled - $lic.Consumed
        $lic.Util  = if ($lic.Enabled -gt 0) { [math]::Round(100*$lic.Consumed/$lic.Enabled) } else { 0 }
        $lic.DirectCount = $lic.Direct.Count
        $lic.GroupCount  = (@($lic.Groups.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum; if (-not $lic.GroupCount) { $lic.GroupCount = 0 }
        $lic.DupCount    = $lic.Duplicates.Count
        foreach ($u in $lic.Direct) { $script:Flat.Add([PSCustomObject]@{Licence=$lic.Friendly;Sku=$lic.Part;Type='Direct';Groupe='';Utilisateur=$u.Name;UPN=$u.Upn;Etat=$u.State})|Out-Null }
        foreach ($gid in $lic.Groups.Keys) { foreach ($u in $lic.Groups[$gid]) { $script:Flat.Add([PSCustomObject]@{Licence=$lic.Friendly;Sku=$lic.Part;Type='Groupe';Groupe=$lic.GroupNames[$gid];Utilisateur=$u.Name;UPN=$u.Upn;Etat=$u.State})|Out-Null } }
    }
    $script:Audit = @($model.Values | Sort-Object @{E={$_.Free}}, @{E={$_.Avail}}, @{E={$_.Friendly}})

    Populate-Tree $txtFilter.Text
    $dupTotal = (@($script:Audit | ForEach-Object { $_.DupCount }) | Measure-Object -Sum).Sum
    $over = @($script:Audit | Where-Object { -not $_.Free -and $_.Avail -lt 0 }).Count
    Write-Log ("Audit termine : {0} licence(s), {1} affectation(s), {2} doublon(s), {3} licence(s) payante(s) sur-provisionnee(s)." -f $script:Audit.Count, $script:Flat.Count, $dupTotal, $over)
    $btnExport.Enabled = ($script:Audit.Count -gt 0)
    $btnLoad.Enabled=$true
}

# ==================================================================
# Export .xlsx NATIF (OpenXML, sans Excel ni module - compatible PowerShell 7)
# ==================================================================
function ConvertTo-Xml { param([string]$s) if ($null -eq $s) { return "" }; $s = $s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]',''; return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;') }
function Get-ColLetter { param([int]$n) $s=''; while ($n -gt 0) { $m=($n-1)%26; $s=[char](65+$m)+$s; $n=[int][math]::Floor(($n-1)/26) }; return $s }

function New-XlsxFile {
    param($Rows, [string[]]$Keys, [string]$Path, [string]$SheetName='Donnees')
    $rowsArr = @($Rows); $nCols = $Keys.Count; $lastRow = $rowsArr.Count + 1; $lastCol = Get-ColLetter $nCols
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    function _put([string]$name,[string]$content) { $e=$zip.CreateEntry($name); $w=New-Object System.IO.StreamWriter($e.Open(),$utf8); $w.Write($content); $w.Close() }

    _put '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    _put '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    _put 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    _put 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + (ConvertTo-Xml $SheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>')
    _put 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'

    $e = $zip.CreateEntry('xl/worksheets/sheet1.xml'); $w = New-Object System.IO.StreamWriter($e.Open(), $utf8)
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>')
    # En-tete
    $w.Write('<row r="1">')
    for ($c=0; $c -lt $nCols; $c++) { $ref=(Get-ColLetter ($c+1))+'1'; $w.Write('<c r="'+$ref+'" t="inlineStr" s="1"><is><t xml:space="preserve">'+(ConvertTo-Xml $Keys[$c])+'</t></is></c>') }
    $w.Write('</row>')
    # Donnees
    $rn = 1
    foreach ($item in $rowsArr) {
        $rn++; $w.Write('<row r="'+$rn+'">')
        for ($c=0; $c -lt $nCols; $c++) { $ref=(Get-ColLetter ($c+1))+$rn; $val=ConvertTo-Xml ([string]$item.$($Keys[$c])); $w.Write('<c r="'+$ref+'" t="inlineStr"><is><t xml:space="preserve">'+$val+'</t></is></c>') }
        $w.Write('</row>')
        if (($rn % 5000) -eq 0) { Write-Log "  ecriture $rn lignes..."; [System.Windows.Forms.Application]::DoEvents() }
    }
    $w.Write('</sheetData><autoFilter ref="A1:'+$lastCol+$lastRow+'"/></worksheet>')
    $w.Close()
    $zip.Dispose(); $fs.Close()
}

# ==================================================================
# Rapport HTML (synthese true-up)
# ==================================================================
function Export-HtmlReport {
    param([string]$Path)
    $enc = { param($s) if ($null -eq $s){return ""}; return ([string]$s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;') }
    $total   = $script:Audit.Count
    $paid    = @($script:Audit | Where-Object { -not $_.Free })
    $over     = @($paid | Where-Object { $_.Avail -lt 0 })
    $dupTotal = (@($script:Audit | ForEach-Object { $_.DupCount }) | Measure-Object -Sum).Sum
    $assign   = $script:Flat.Count
    $now = Get-Date -Format "dd/MM/yyyy HH:mm"

    $css = @'
*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:#1b2733;background:#f4f6f9}
header{background:#1f3a5f;color:#fff;padding:22px 32px}
header h1{margin:0;font-size:20px;font-weight:600}
header .sub{opacity:.8;font-size:13px;margin-top:4px}
.wrap{max-width:1180px;margin:0 auto;padding:24px 32px}
.kpis{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:28px}
.kpi{flex:1;min-width:190px;background:#fff;border:1px solid #e2e8f0;border-left:4px solid #1f3a5f;border-radius:8px;padding:16px 18px}
.kpi .n{font-size:26px;font-weight:700}.kpi .l{font-size:12px;color:#5b6b7b;margin-top:2px}
.kpi.red{border-left-color:#c0392b}.kpi.amber{border-left-color:#d08700}
h2{font-size:15px;margin:28px 0 12px;color:#1f3a5f}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;font-size:13px}
th{background:#eef2f7;text-align:left;padding:9px 12px;font-weight:600;color:#33485c;border-bottom:1px solid #e2e8f0}
td{padding:8px 12px;border-bottom:1px solid #f0f3f6}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.num{text-align:right;font-variant-numeric:tabular-nums}
.pill{display:inline-block;padding:2px 9px;border-radius:11px;font-size:11px;font-weight:600}
.ok{background:#e4f5e9;color:#1e7d43}.tight{background:#fdf0d8;color:#9a6700}.over{background:#fbe2df;color:#b4281c}.unused{background:#eceff2;color:#63717f}
.cat{font-size:11px;color:#63717f}
.rowover td{background:#fdf1ef}
.bar{height:7px;background:#eceff2;border-radius:4px;overflow:hidden;min-width:70px}
.bar>span{display:block;height:100%;background:#1f7a3d}
.bar.o>span{background:#d08700}.bar.x>span{background:#c0392b}
footer{max-width:1180px;margin:0 auto;padding:16px 32px 40px;color:#8794a1;font-size:12px}
@media print{body{background:#fff}tr:hover td{background:none}header{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
'@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Audit des licences</title><style>'+$css+'</style></head><body>')
    [void]$sb.Append('<header><h1>Audit des licences du tenant &mdash; true-up</h1><div class="sub">'+(& $enc $script:Account)+' &nbsp;·&nbsp; genere le '+$now+'</div></header>')
    [void]$sb.Append('<div class="wrap">')
    [void]$sb.Append('<div class="kpis">')
    [void]$sb.Append('<div class="kpi"><div class="n">'+$total+'</div><div class="l">Licences suivies</div></div>')
    [void]$sb.Append('<div class="kpi"><div class="n">'+$paid.Count+'</div><div class="l">dont payantes</div></div>')
    [void]$sb.Append('<div class="kpi '+$(if($over.Count){'red'}else{''})+'"><div class="n">'+$over.Count+'</div><div class="l">Payantes sur-provisionnees</div></div>')
    [void]$sb.Append('<div class="kpi '+$(if($dupTotal){'amber'}else{''})+'"><div class="n">'+$dupTotal+'</div><div class="l">Doublons direct + heritage</div></div>')
    [void]$sb.Append('<div class="kpi"><div class="n">'+$assign+'</div><div class="l">Affectations totales</div></div>')
    [void]$sb.Append('</div>')

    # Tableau synthese
    [void]$sb.Append('<h2>Synthese des licences</h2>')
    [void]$sb.Append('<table><thead><tr><th>Licence</th><th>SKU</th><th>Categorie</th><th class="num">Total</th><th class="num">Consomme</th><th class="num">Dispo</th><th>Utilisation</th><th class="num">Direct</th><th class="num">Groupe</th><th class="num">Doublons</th><th>Statut</th></tr></thead><tbody>')
    foreach ($l in $script:Audit) {
        $st = if ($l.Avail -lt 0) {'over'} elseif ($l.Consumed -eq 0) {'unused'} elseif ($l.Util -ge 90) {'tight'} else {'ok'}
        $stTxt = @{over='Sur-provisionne';unused='Non utilisee';tight='Tendu';ok='OK'}[$st]
        $barCls = if ($st -eq 'over') {'bar x'} elseif ($st -eq 'tight') {'bar o'} else {'bar'}
        $barW = [math]::Min(100,$l.Util)
        $rowCls = if ($st -eq 'over') { ' class="rowover"' } else { '' }
        [void]$sb.Append('<tr'+$rowCls+'><td>'+(& $enc $l.Friendly)+'</td><td class="cat">'+(& $enc $l.Part)+'</td><td class="cat">'+$(if($l.Free){'Essai / gratuite'}else{'Payante'})+'</td>')
        [void]$sb.Append('<td class="num">'+$l.Enabled+'</td><td class="num">'+$l.Consumed+'</td><td class="num">'+$l.Avail+'</td>')
        [void]$sb.Append('<td><div class="'+$barCls+'"><span style="width:'+$barW+'%"></span></div> <span class="cat">'+$l.Util+'%</span></td>')
        [void]$sb.Append('<td class="num">'+$l.DirectCount+'</td><td class="num">'+$l.GroupCount+'</td><td class="num">'+$l.DupCount+'</td>')
        [void]$sb.Append('<td><span class="pill '+$st+'">'+$stTxt+'</span></td></tr>')
    }
    [void]$sb.Append('</tbody></table>')

    # Points d'attention
    [void]$sb.Append('<h2>Points d''attention (true-up)</h2>')
    if ($over.Count -gt 0) {
        [void]$sb.Append('<table><thead><tr><th>Licence payante sur-provisionnee</th><th>SKU</th><th class="num">Consomme</th><th class="num">Total</th><th class="num">Depassement</th></tr></thead><tbody>')
        foreach ($l in ($over | Sort-Object Avail)) { [void]$sb.Append('<tr class="rowover"><td>'+(& $enc $l.Friendly)+'</td><td class="cat">'+(& $enc $l.Part)+'</td><td class="num">'+$l.Consumed+'</td><td class="num">'+$l.Enabled+'</td><td class="num">'+([math]::Abs($l.Avail))+'</td></tr>') }
        [void]$sb.Append('</tbody></table>')
    } else { [void]$sb.Append('<p class="cat">Aucune licence payante en depassement.</p>') }

    $dupLics = @($script:Audit | Where-Object { $_.DupCount -gt 0 } | Sort-Object DupCount -Descending)
    if ($dupLics.Count -gt 0) {
        [void]$sb.Append('<h2>Doublons par licence (affectation directe + heritage de groupe)</h2>')
        [void]$sb.Append('<table><thead><tr><th>Licence</th><th>SKU</th><th class="num">Utilisateurs en doublon</th></tr></thead><tbody>')
        foreach ($l in $dupLics) { [void]$sb.Append('<tr><td>'+(& $enc $l.Friendly)+'</td><td class="cat">'+(& $enc $l.Part)+'</td><td class="num">'+$l.DupCount+'</td></tr>') }
        [void]$sb.Append('</tbody></table><p class="cat">Detail nominatif des doublons : voir l''export de donnees completes (.xlsx / .csv).</p>')
    }

    [void]$sb.Append('</div><footer>Rapport genere le '+$now+' &nbsp;·&nbsp; '+$total+' licences &nbsp;·&nbsp; '+$assign+' affectations &nbsp;·&nbsp; source : Microsoft Graph via session Az.</footer></body></html>')

    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
}

# ==================================================================
# Events
# ==================================================================
$btnConnect.Add_Click({
    if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) { [System.Windows.Forms.MessageBox]::Show("Module Az introuvable.","Prerequis","OK","Error")|Out-Null; return }
    $btnConnect.Enabled=$false; Set-Conn $false "Connexion en cours..."
    try { Connect-AzAccount -ErrorAction Stop | Out-Null } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Connexion","OK","Error")|Out-Null }
    $ctx=Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connecte : $($ctx.Account.Id)" } else { Set-Conn $false "Non connecte." }
    $btnConnect.Enabled=$true
})
$btnLoad.Add_Click({ if (-not (Test-AzConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a Azure.","Audit","OK","Warning")|Out-Null; return }; Start-Audit })
$txtFilter.Add_TextChanged({ if ($script:Audit.Count -gt 0) { Populate-Tree $txtFilter.Text } })
$chkHide.Add_CheckedChanged({ if ($script:Audit.Count -gt 0) { Populate-Tree $txtFilter.Text } })
$btnExpand.Add_Click({ $tree.ExpandAll(); if ($tree.Nodes.Count -gt 0) { $tree.Nodes[0].EnsureVisible() } })
$btnCollapse.Add_Click({ $tree.CollapseAll() })

$btnExport.Add_Click({
    if ($script:Audit.Count -eq 0) { return }
    $dlg=New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter="Rapport HTML (synthese) (*.html)|*.html|Classeur Excel - donnees completes (*.xlsx)|*.xlsx|CSV - donnees completes (*.csv)|*.csv"
    $dlg.FileName="Audit_Licences.html"
    if ($dlg.ShowDialog() -ne "OK") { return }
    $file=$dlg.FileName
    try {
        if ($file -match '\.html?$') {
            Export-HtmlReport -Path $file
            Write-Log "Rapport HTML : $file"
            if ([System.Windows.Forms.MessageBox]::Show("Rapport genere :`n$file`n`nL'ouvrir maintenant ?","Export","YesNo","Information") -eq "Yes") { Start-Process $file }
        }
        elseif ($file -match '\.xlsx$') {
            $keys=@('Licence','Sku','Type','Groupe','Utilisateur','UPN','Etat')
            Write-Log "Generation .xlsx ($($script:Flat.Count) lignes)..."
            New-XlsxFile -Rows $script:Flat -Keys $keys -Path $file -SheetName 'Licences'
            Write-Log "Export .xlsx : $file"
            [System.Windows.Forms.MessageBox]::Show("Export termine :`n$file","Export","OK","Information")|Out-Null
        }
        else {
            $keys=@('Licence','Sku','Type','Groupe','Utilisateur','UPN','Etat')
            $script:Flat | Select-Object $keys | Export-Csv -Path $file -Delimiter ";" -NoTypeInformation -Encoding UTF8
            Write-Log "Export CSV : $file"
            [System.Windows.Forms.MessageBox]::Show("Export termine :`n$file","Export","OK","Information")|Out-Null
        }
    } catch { Write-Log "ERREUR export : $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Export","OK","Error")|Out-Null }
})

# ==================================================================
# Adopt active Az session
# ==================================================================
if (Get-Command Get-AzContext -ErrorAction SilentlyContinue) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx) { Set-Conn $true "Connexion Az active reutilisee : $($ctx.Account.Id)"; Write-Log "Connexion Az detectee : $($ctx.Account.Id)" }
    else { Set-Conn $false "Non connecte. Cliquez sur 'Connexion Azure'." }
} else { Set-Conn $false "Module Az requis (Install-Module Az)." }

[void]$form.ShowDialog()
