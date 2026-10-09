<#
.SYNOPSIS
    Audit des licences du tenant (true-up) - Interface graphique, arborescence depliable.
    Connexion : Az (Connect-AzAccount). Interroge Microsoft Graph via un jeton obtenu depuis Az
    (pas de module Graph requis).

.DESCRIPTION
    Pour chaque licence (SKU) du tenant, affiche dans un arbre :
      - le total / consomme / disponible (donnees subscribedSkus),
      - les affectations DIRECTES,
      - les affectations HERITEES, regroupees par groupe (licence basee sur les groupes),
      - les DOUBLONS (utilisateur ayant la meme licence en direct ET par heritage).

    Objectif true-up : voir d'un coup qui consomme quoi et par quel chemin, reperer les doublons
    et les licences en erreur. Filtre (licence ou utilisateur) et export a plat (.xlsx / .csv).

.PREREQUISITES
    - Module Az (Az.Accounts) + session Connect-AzAccount avec un compte administrateur.
    - Lecture annuaire (l'app Azure PowerShell lit l'annuaire au nom de l'utilisateur connecte).
    - Export .xlsx : Microsoft Excel installe (pilote en COM). Sinon export .csv.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:GraphToken = $null
$script:Audit      = @()   # modele : liste d'objets licence
$script:Flat       = New-Object System.Collections.Generic.List[object]

# ==================================================================
# Noms conviviaux de SKU (repli sur le skuPartNumber si absent)
# ==================================================================
$script:SkuNames = @{
    'SPE_E3'='Microsoft 365 E3'; 'SPE_E5'='Microsoft 365 E5'; 'SPE_F1'='Microsoft 365 F3'
    'ENTERPRISEPACK'='Office 365 E3'; 'ENTERPRISEPREMIUM'='Office 365 E5'; 'STANDARDPACK'='Office 365 E1'; 'DESKLESSPACK'='Office 365 F3'
    'EMS'='EMS E3'; 'EMSPREMIUM'='EMS E5'; 'AAD_PREMIUM'='Entra ID P1'; 'AAD_PREMIUM_P2'='Entra ID P2'
    'POWER_BI_PRO'='Power BI Pro'; 'POWER_BI_STANDARD'='Power BI (free)'; 'FLOW_FREE'='Power Automate Free'
    'PROJECTPROFESSIONAL'='Project Plan 3'; 'PROJECTPREMIUM'='Project Plan 5'; 'PROJECT_P1'='Project Plan 1'
    'VISIOCLIENT'='Visio Plan 2'; 'VISIO_PLAN1_DEPT'='Visio Plan 1'
    'EXCHANGESTANDARD'='Exchange Online P1'; 'EXCHANGEENTERPRISE'='Exchange Online P2'
    'MCOSTANDARD'='Skype Entreprise'; 'MCOEV'='Teams Phone'; 'MCOMEETADV'='Teams Audio Conferencing'; 'TEAMS_EXPLORATORY'='Teams Exploratory'
    'WIN_DEF_ATP'='Defender for Endpoint'; 'ATP_ENTERPRISE'='Defender for Office P1'
    'Microsoft_365_Copilot'='Microsoft 365 Copilot'
    'O365_BUSINESS_PREMIUM'='M365 Business Standard'; 'O365_BUSINESS_ESSENTIALS'='M365 Business Basic'; 'SPB'='M365 Business Premium'
}
function Get-SkuName { param([string]$part) if ($script:SkuNames.ContainsKey($part)) { return $script:SkuNames[$part] } return $part }

# ==================================================================
# Fonctions Az / Graph
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
$form.Size = New-Object System.Drawing.Size(1000, 820)
$form.MinimumSize = New-Object System.Drawing.Size(880, 640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Add-Lbl { param($parent,$text,$x,$y) $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $parent.Controls.Add($l); return $l }

$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Text="Connexion Azure"; $btnConnect.Location=New-Object System.Drawing.Point(15,12); $btnConnect.Size=New-Object System.Drawing.Size(160,28); $form.Controls.Add($btnConnect)
$lblConn = Add-Lbl $form "Etat : verification..." 190 18; $lblConn.AutoSize=$false; $lblConn.Size=New-Object System.Drawing.Size(780,18); $lblConn.Anchor="Top,Left,Right"

$btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text="Lancer l'audit"; $btnLoad.Location=New-Object System.Drawing.Point(15,50); $btnLoad.Size=New-Object System.Drawing.Size(190,28)
$btnLoad.BackColor=[System.Drawing.Color]::FromArgb(0,120,215); $btnLoad.ForeColor=[System.Drawing.Color]::White; $btnLoad.FlatStyle="Flat"; $form.Controls.Add($btnLoad)
Add-Lbl $form "Filtre :" 220 55 | Out-Null
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilter.Location=New-Object System.Drawing.Point(270,52); $txtFilter.Size=New-Object System.Drawing.Size(240,23); $form.Controls.Add($txtFilter)
$btnExpand = New-Object System.Windows.Forms.Button; $btnExpand.Text="Deplier"; $btnExpand.Location=New-Object System.Drawing.Point(525,50); $btnExpand.Size=New-Object System.Drawing.Size(90,28); $form.Controls.Add($btnExpand)
$btnCollapse = New-Object System.Windows.Forms.Button; $btnCollapse.Text="Replier"; $btnCollapse.Location=New-Object System.Drawing.Point(620,50); $btnCollapse.Size=New-Object System.Drawing.Size(90,28); $form.Controls.Add($btnCollapse)
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text="Exporter..."; $btnExport.Location=New-Object System.Drawing.Point(820,50); $btnExport.Size=New-Object System.Drawing.Size(160,28); $btnExport.Anchor="Top,Right"; $btnExport.Enabled=$false; $form.Controls.Add($btnExport)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Location=New-Object System.Drawing.Point(15,88); $tree.Size=New-Object System.Drawing.Size(965,560); $tree.Anchor="Top,Bottom,Left,Right"
$tree.HideSelection=$false; $tree.Font=New-Object System.Drawing.Font("Segoe UI",9)
$form.Controls.Add($tree)

Add-Lbl $form "Journal :" 15 656 | Out-Null
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location=New-Object System.Drawing.Point(15,676); $txtLog.Size=New-Object System.Drawing.Size(965,100); $txtLog.Anchor="Bottom,Left,Right"
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"; $txtLog.ReadOnly=$true; $txtLog.Font=New-Object System.Drawing.Font("Consolas",9); $form.Controls.Add($txtLog)

# ==================================================================
# UI helpers
# ==================================================================
function Write-Log { param([string]$m) $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format HH:mm:ss), $m)); [System.Windows.Forms.Application]::DoEvents() }
function Set-Conn { param([bool]$c,[string]$m) $lblConn.ForeColor= if($c){[System.Drawing.Color]::ForestGreen}else{[System.Drawing.Color]::Gray}; $lblConn.Text=$m }

function Populate-Tree {
    param([string]$Filter = "")
    $f = $Filter.Trim().ToLower()
    $tree.BeginUpdate(); $tree.Nodes.Clear()
    foreach ($lic in $script:Audit) {
        $licMatch = (-not $f) -or ($lic.Part.ToLower().Contains($f)) -or ($lic.Friendly.ToLower().Contains($f))

        # Noeud racine licence
        $root = New-Object System.Windows.Forms.TreeNode
        $root.Text = "{0} ({1})  -  consomme {2} / {3}  (dispo {4})" -f $lic.Friendly, $lic.Part, $lic.Consumed, $lic.Enabled, ($lic.Enabled - $lic.Consumed)
        $root.ForeColor = if (($lic.Enabled - $lic.Consumed) -lt 0) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black }

        $added = $false

        # Direct
        $direct = @($lic.Direct | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) })
        if ($direct.Count -gt 0) {
            $nd = New-Object System.Windows.Forms.TreeNode ("Direct ({0})" -f $direct.Count)
            foreach ($u in $direct) { $un=New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}{2}" -f $u.Name,$u.Upn, $(if($u.State -and $u.State -ne 'Active'){" [$($u.State)]"}else{""})); if($u.State -and $u.State -ne 'Active'){$un.ForeColor=[System.Drawing.Color]::DarkOrange}; [void]$nd.Nodes.Add($un) }
            [void]$root.Nodes.Add($nd); $added=$true
        }

        # Par groupe
        foreach ($gid in $lic.Groups.Keys) {
            $gname = $lic.GroupNames[$gid]
            $gusers = @($lic.Groups[$gid] | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) -or ("$gname").ToLower().Contains($f) })
            if ($gusers.Count -gt 0) {
                $ng = New-Object System.Windows.Forms.TreeNode ("Via groupe : {0} ({1})" -f $gname, $gusers.Count)
                $ng.ForeColor = [System.Drawing.Color]::FromArgb(0,90,160)
                foreach ($u in $gusers) { [void]$ng.Nodes.Add((New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}" -f $u.Name,$u.Upn))) }
                [void]$root.Nodes.Add($ng); $added=$true
            }
        }

        # Doublons direct + heritage
        $dups = @($lic.Duplicates | Where-Object { (-not $f) -or $licMatch -or $_.Upn.ToLower().Contains($f) -or $_.Name.ToLower().Contains($f) })
        if ($dups.Count -gt 0) {
            $ndup = New-Object System.Windows.Forms.TreeNode ("Doublons direct + heritage ({0})" -f $dups.Count)
            $ndup.ForeColor = [System.Drawing.Color]::Firebrick
            foreach ($u in $dups) { [void]$ndup.Nodes.Add((New-Object System.Windows.Forms.TreeNode ("{0}  -  {1}" -f $u.Name,$u.Upn))) }
            [void]$root.Nodes.Add($ndup); $added=$true
        }

        if ($licMatch -or $added) { [void]$tree.Nodes.Add($root) }
    }
    $tree.EndUpdate()
}

# ==================================================================
# Build audit model
# ==================================================================
function Start-Audit {
    $btnLoad.Enabled=$false
    try { $script:GraphToken = Get-GraphToken }
    catch { Write-Log "ERREUR jeton Graph : $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Impossible d'obtenir un jeton Graph depuis Az.`n$($_.Exception.Message)","Audit","OK","Error")|Out-Null; $btnLoad.Enabled=$true; return }

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
        [System.Windows.Forms.MessageBox]::Show("Appel Graph refuse.`n$($_.Exception.Message)`n`nLe compte connecte doit pouvoir lire l'annuaire (admin).","Audit","OK","Error")|Out-Null
        $btnLoad.Enabled=$true; return
    }

    # Table des groupes (id -> nom)
    $gName = @{}
    foreach ($g in $groups) { if ($g.id) { $gName[$g.id] = $g.displayName } }

    # Index SKU
    $model = @{}
    foreach ($s in $skus) {
        $model[$s.skuId] = [PSCustomObject]@{
            SkuId = $s.skuId; Part = $s.skuPartNumber; Friendly = (Get-SkuName $s.skuPartNumber)
            Enabled = [int]$s.prepaidUnits.enabled; Consumed = [int]$s.consumedUnits
            Direct = (New-Object System.Collections.Generic.List[object])
            Groups = @{}; GroupNames = @{}
            DirectUpns = (New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
            GroupUpns  = (New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
            Duplicates = (New-Object System.Collections.Generic.List[object])
        }
    }

    # Parcours des utilisateurs
    foreach ($u in $users) {
        if (-not $u.licenseAssignmentStates) { continue }
        foreach ($st in $u.licenseAssignmentStates) {
            $sid = $st.skuId
            if (-not $model.ContainsKey($sid)) { continue }
            $lic = $model[$sid]
            $uObj = [PSCustomObject]@{ Name=$u.displayName; Upn=$u.userPrincipalName; State=$st.state }
            if ([string]::IsNullOrEmpty($st.assignedByGroup)) {
                $lic.Direct.Add($uObj) | Out-Null; [void]$lic.DirectUpns.Add($u.userPrincipalName)
            } else {
                $gid = $st.assignedByGroup
                if (-not $lic.Groups.ContainsKey($gid)) { $lic.Groups[$gid] = New-Object System.Collections.Generic.List[object]; $lic.GroupNames[$gid] = $(if($gName.ContainsKey($gid)){$gName[$gid]}else{"(groupe $gid)"}) }
                $lic.Groups[$gid].Add($uObj) | Out-Null; [void]$lic.GroupUpns.Add($u.userPrincipalName)
            }
        }
    }

    # Doublons + modele final trie + liste a plat
    $script:Flat.Clear()
    foreach ($lic in $model.Values) {
        foreach ($upn in $lic.DirectUpns) {
            if ($lic.GroupUpns.Contains($upn)) {
                $nm = ($lic.Direct | Where-Object { $_.Upn -ieq $upn } | Select-Object -First 1).Name
                $lic.Duplicates.Add([PSCustomObject]@{ Name=$nm; Upn=$upn }) | Out-Null
            }
        }
        # a plat pour export
        foreach ($u in $lic.Direct) { $script:Flat.Add([PSCustomObject]@{ Licence=$lic.Friendly; Sku=$lic.Part; Type='Direct'; Groupe=''; Utilisateur=$u.Name; UPN=$u.Upn; Etat=$u.State }) | Out-Null }
        foreach ($gid in $lic.Groups.Keys) { foreach ($u in $lic.Groups[$gid]) { $script:Flat.Add([PSCustomObject]@{ Licence=$lic.Friendly; Sku=$lic.Part; Type='Groupe'; Groupe=$lic.GroupNames[$gid]; Utilisateur=$u.Name; UPN=$u.Upn; Etat=$u.State }) | Out-Null } }
    }
    $script:Audit = @($model.Values | Sort-Object Friendly)

    Populate-Tree
    $dupTotal = @($script:Audit | ForEach-Object { $_.Duplicates.Count } | Measure-Object -Sum).Sum
    Write-Log ("Audit termine : {0} licence(s), {1} affectation(s) a plat, {2} doublon(s) direct+heritage." -f $script:Audit.Count, $script:Flat.Count, $dupTotal)
    $btnExport.Enabled = ($script:Flat.Count -gt 0)
    $btnLoad.Enabled=$true
}

# ==================================================================
# Export (.xlsx via Excel COM, ou .csv)
# ==================================================================
function Export-ToXlsxCom {
    param($Rows, [string[]]$Keys, [string]$Path)
    $rowsArr = @($Rows); $nCols = $Keys.Count
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("lic_audit_{0}.txt" -f ([guid]::NewGuid().ToString('N'))))
    $sw = New-Object System.IO.StreamWriter($tmp, $false, (New-Object System.Text.UTF8Encoding($true)))
    $sw.WriteLine(($Keys -join "`t"))
    foreach ($item in $rowsArr) { $vals = foreach ($k in $Keys) { ("" + $item.$k) -replace "[`t`r`n]"," " }; $sw.WriteLine(($vals -join "`t")) }
    $sw.Close()
    $xl = $null
    try { $xl = New-Object -ComObject Excel.Application }
    catch { [System.Windows.Forms.MessageBox]::Show("Excel introuvable (requis pour .xlsx). Utilisez .csv.","Export","OK","Warning")|Out-Null; Remove-Item $tmp -Force -ErrorAction SilentlyContinue; return }
    $xl.Visible=$false; $xl.DisplayAlerts=$false
    try {
        $xl.Workbooks.OpenText($tmp,65001,1,1,-4142,$false,$true,$false,$false,$false,$false) | Out-Null
        $wb=$xl.ActiveWorkbook; $ws=$wb.Worksheets.Item(1); $ws.Name="Licences"
        $hdr=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$nCols)); $hdr.Font.Bold=$true
        $xl.ActiveWindow.SplitRow=1; $xl.ActiveWindow.FreezePanes=$true
        $hdr.AutoFilter() | Out-Null; $ws.Columns.AutoFit() | Out-Null
        $wb.SaveAs($Path,51); $wb.Close($false)
    } finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
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

$btnLoad.Add_Click({
    if (-not (Test-AzConnected)) { [System.Windows.Forms.MessageBox]::Show("Non connecte a Azure.","Audit","OK","Warning")|Out-Null; return }
    Start-Audit
})

$txtFilter.Add_TextChanged({ if ($script:Audit.Count -gt 0) { Populate-Tree $txtFilter.Text } })
$btnExpand.Add_Click({ $tree.ExpandAll(); if ($tree.Nodes.Count -gt 0) { $tree.Nodes[0].EnsureVisible() } })
$btnCollapse.Add_Click({ $tree.CollapseAll() })

$btnExport.Add_Click({
    if ($script:Flat.Count -eq 0) { return }
    $keys = @('Licence','Sku','Type','Groupe','Utilisateur','UPN','Etat')
    $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="Classeur Excel (*.xlsx)|*.xlsx|Fichier CSV (*.csv)|*.csv"; $dlg.FileName="Audit_Licences.xlsx"
    if ($dlg.ShowDialog() -ne "OK") { return }
    $file=$dlg.FileName
    try {
        if ($file -match '\.xlsx$') { Export-ToXlsxCom -Rows $script:Flat -Keys $keys -Path $file }
        else { $script:Flat | Select-Object $keys | Export-Csv -Path $file -Delimiter ";" -NoTypeInformation -Encoding UTF8 }
        Write-Log "Export : $file"
        [System.Windows.Forms.MessageBox]::Show("Export termine :`n$file","Export","OK","Information")|Out-Null
    } catch { [System.Windows.Forms.MessageBox]::Show("Echec :`n$($_.Exception.Message)","Export","OK","Error")|Out-Null }
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
