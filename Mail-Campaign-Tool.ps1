<#
.SYNOPSIS
    Mail-Campaign-Tool - Outil de campagne mailing (WinForms + Exchange / SMTP)

.DESCRIPTION
    Interface graphique complète pour générer et envoyer des mails en masse :
      1. Destinataires  : import bulk TXT/CSV (une adresse par ligne), saisie manuelle,
                          nettoyage, dédoublonnage, validation, export.
      2. Serveur SMTP   : Exchange / relais SMTP, avec ou sans authentification,
                          SSL/TLS, test de connexion, mail de test.
      3. Éditeur        : éditeur HTML WYSIWYG complet (gras, italique, couleurs, polices,
                          titres, listes, alignement, liens, tableaux, images par chemin
                          ou URL, blocs Info/Attention/Succès/Code, variables, modèles
                          "Procédure", "Information", "Maintenance", "Sécurité"),
                          vue source HTML, pièces jointes, brouillons.
      4. Signature      : éditeur de signature avec images, sauvegardée dans %APPDATA%.
      5. Envoi          : mode "1 mail par destinataire (personnalisé)" ou "Cci par lots",
                          délai entre envois, simulation (dry-run), journal coloré,
                          export CSV des résultats.

    Les images locales (chemin) sont automatiquement embarquées dans le mail (Content-ID),
    donc visibles dans Outlook sans pièce jointe apparente.

    Variables utilisables dans l'objet et le corps :
        {{EMAIL}} {{PRENOM}} {{NOM}} {{NOMCOMPLET}} {{DATE}} {{HEURE}}
    (PRENOM / NOM sont déduits de la partie locale de l'adresse : prenom.nom@domaine.fr)

    Raccourcis clavier dans l'éditeur : Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+Z / Ctrl+Y,
    Ctrl+C / Ctrl+V (le collage depuis Word/Outlook conserve la mise en forme).

.NOTES
    Auteur   : Hakim
    Version  : 1.0
    Prérequis: Windows PowerShell 5.1, .NET Framework 4.x, Windows (WebBrowser/IE engine)
    Lancement: powershell.exe -STA -ExecutionPolicy Bypass -File .\Mail-Campaign-Tool.ps1
               (le script se relance automatiquement en STA si nécessaire)
    Fichiers : %APPDATA%\Mail-Campaign-Tool\config.json     (config, sans mot de passe)
               %APPDATA%\Mail-Campaign-Tool\signature.html  (signature)
               %APPDATA%\Mail-Campaign-Tool\logs\*.log      (journaux)
#>

#region ================= INITIALISATION =================
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11

$script:AppName       = 'Mail-Campaign-Tool'
$script:AppVersion    = '1.0'
$script:AppDir        = Join-Path $env:APPDATA $script:AppName
$script:LogDir        = Join-Path $script:AppDir 'logs'
$script:TempDir       = Join-Path $env:TEMP $script:AppName
foreach ($d in @($script:AppDir, $script:LogDir, $script:TempDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:ConfigFile    = Join-Path $script:AppDir 'config.json'
$script:SignatureFile = Join-Path $script:AppDir 'signature.html'
$script:LogFile       = Join-Path $script:LogDir ("campaign_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
$script:EditorFile    = Join-Path $script:TempDir 'editor.html'
$script:SigEditorFile = Join-Path $script:TempDir 'signature_editor.html'
$script:PreviewFile   = Join-Path $script:TempDir 'preview.html'
$script:CancelSend    = $false
$script:SourceMode    = $false
$script:Results       = New-Object System.Collections.ArrayList
$script:EditorsReady  = $false

# Force le moteur IE11 pour le contrôle WebBrowser (sinon mode IE7 : CSS moderne cassé)
function Set-BrowserEmulation {
    try {
        $key = 'HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION'
        if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
        $exe = [System.IO.Path]::GetFileName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
        $cur = (Get-ItemProperty -Path $key -Name $exe -ErrorAction SilentlyContinue).$exe
        if ($cur -ne 11001) { New-ItemProperty -Path $key -Name $exe -Value 11001 -PropertyType DWord -Force | Out-Null }
    } catch { }
}
Set-BrowserEmulation

# Polices
$FontUI     = New-Object System.Drawing.Font('Segoe UI', 9)
$FontBold   = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$FontItalic = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
$FontUnder  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
$FontStrike = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Strikeout)
$FontMono   = New-Object System.Drawing.Font('Consolas', 10)
$FontTitle  = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$FontBig    = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
#endregion

#region ================= HELPERS UI =================
function New-Ctl {
    param([string]$Type, [hashtable]$Props = @{}, $Parent)
    $c = New-Object ("System.Windows.Forms.$Type")
    # Minimum / Maximum d'abord (NumericUpDown) pour éviter un Value hors bornes
    $keys = @($Props.Keys | Sort-Object { switch ($_) { 'Minimum' { 0 } 'Maximum' { 1 } default { 2 } } })
    foreach ($k in $keys) {
        $v = $Props[$k]
        switch ($k) {
            'Location'    { $c.Location    = New-Object System.Drawing.Point($v[0], $v[1]) }
            'Size'        { $c.Size        = New-Object System.Drawing.Size($v[0], $v[1]) }
            'MinimumSize' { $c.MinimumSize = New-Object System.Drawing.Size($v[0], $v[1]) }
            default       { $c.$k = $v }
        }
    }
    if ($Parent) { [void]$Parent.Controls.Add($c) }
    $c
}

function Add-Field {
    param($Parent, [string]$Label, [int]$X, [int]$Y, [int]$Width = 260, [string]$Default = '', [switch]$Password, [int]$LabelWidth = 150)
    New-Ctl 'Label' @{ Text = $Label; Location = @($X, $Y + 3); AutoSize = $true } $Parent | Out-Null
    $t = New-Ctl 'TextBox' @{ Location = @($X + $LabelWidth, $Y); Size = @($Width, 23); Text = $Default } $Parent
    if ($Password) { $t.UseSystemPasswordChar = $true }
    $t
}

function Show-Msg {
    param([string]$Text, [string]$Title = $script:AppName, [string]$Icon = 'Information')
    [System.Windows.Forms.MessageBox]::Show($Text, $Title, 'OK', $Icon) | Out-Null
}
function Confirm-Msg {
    param([string]$Text, [string]$Title = $script:AppName)
    return ([System.Windows.Forms.MessageBox]::Show($Text, $Title, 'YesNo', 'Question') -eq 'Yes')
}
function Read-Input {
    param([string]$Prompt, [string]$Title = $script:AppName, [string]$Default = '')
    return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','OK','WARN','ERROR','DEBUG')][string]$Level = 'INFO')
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    $color = switch ($Level) {
        'OK'    { [System.Drawing.Color]::ForestGreen }
        'WARN'  { [System.Drawing.Color]::DarkOrange }
        'ERROR' { [System.Drawing.Color]::Firebrick }
        'DEBUG' { [System.Drawing.Color]::Gray }
        default { [System.Drawing.Color]::Black }
    }
    if ($rtbLog) {
        $rtbLog.SelectionStart  = $rtbLog.TextLength
        $rtbLog.SelectionLength = 0
        $rtbLog.SelectionColor  = $color
        $rtbLog.AppendText($line + "`r`n")
        $rtbLog.SelectionColor  = $rtbLog.ForeColor
        $rtbLog.ScrollToCaret()
    }
    if ($lblStatus) { $lblStatus.Text = $Message }
    try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch { }
    [System.Windows.Forms.Application]::DoEvents()
}
#endregion

#region ================= HELPERS METIER =================
function Test-EmailAddress {
    param([string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    return ($Address -match '^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]{2,}$')
}

function Get-RecipientList {
    param([switch]$ValidOnly)
    $raw = $txtRecipients.Text -split '[\r\n;,\s]+' |
           ForEach-Object { $_.Trim().Trim('<', '>', '"', "'") } |
           Where-Object { $_ }
    $seen = @{}
    $list = foreach ($r in $raw) {
        $k = $r.ToLowerInvariant()
        if (-not $seen.ContainsKey($k)) { $seen[$k] = 1; $r }
    }
    if ($ValidOnly) { $list = $list | Where-Object { Test-EmailAddress $_ } }
    return ,@($list | Where-Object { $_ })
}

function Get-VarsForEmail {
    param([string]$Email)
    $prenom = ''; $nom = ''
    if ($Email) {
        $local = ($Email -split '@')[0]
        $parts = @($local -split '[._\-]+' | Where-Object { $_ } | ForEach-Object {
            if ($_.Length -gt 1) { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() } else { $_.ToUpper() }
        })
        if ($parts.Count -gt 0) { $prenom = $parts[0] }
        if ($parts.Count -gt 1) { $nom = ($parts[1..($parts.Count-1)] -join ' ') }
    }
    return @{
        EMAIL      = $Email
        PRENOM     = $prenom
        NOM        = $nom
        NOMCOMPLET = ("$prenom $nom").Trim()
        DATE       = (Get-Date -Format 'dd/MM/yyyy')
        HEURE      = (Get-Date -Format 'HH:mm')
    }
}

function Expand-Variables {
    param([string]$Text, [hashtable]$Vars)
    if (-not $Text) { return '' }
    foreach ($k in $Vars.Keys) {
        $val = ([string]$Vars[$k]) -replace '\$', '$$$$'
        $Text = [regex]::Replace($Text, "\{\{\s*$k\s*\}\}", $val, 'IgnoreCase')
    }
    return $Text
}

function ConvertTo-PlainText {
    param([string]$Html)
    if (-not $Html) { return '' }
    $t = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', ''
    $t = $t -replace '(?i)<br\s*/?>', "`r`n"
    $t = $t -replace '(?i)</(p|div|li|tr|h[1-6]|table)>', "`r`n"
    $t = $t -replace '<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace '[ \t]+', ' '
    $t = $t -replace '(\r?\n[ \t]*){3,}', "`r`n`r`n"
    return $t.Trim()
}

function Get-MimeType {
    param([string]$Path)
    switch ([System.IO.Path]::GetExtension($Path).ToLower()) {
        '.png'  { 'image/png' }
        '.jpg'  { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.gif'  { 'image/gif' }
        '.bmp'  { 'image/bmp' }
        '.svg'  { 'image/svg+xml' }
        '.webp' { 'image/webp' }
        default { 'application/octet-stream' }
    }
}

function ConvertTo-HtmlColor {
    param([System.Drawing.Color]$Color)
    return ('#{0:X2}{1:X2}{2:X2}' -f $Color.R, $Color.G, $Color.B)
}

function ConvertTo-FileUri {
    param([string]$Path)
    return 'file:///' + (($Path -replace '\\', '/') -replace ' ', '%20')
}

# Remplace les <img src="file:///..."> ou src="C:\..." par des cid: embarqués
function ConvertTo-InlineImages {
    param([string]$Html, [System.Collections.ArrayList]$Resources)
    $rx = [regex]'(?is)(<img[^>]*?\ssrc\s*=\s*["''])([^"'']+)(["''])'
    $mc = $rx.Matches($Html)
    if ($mc.Count -eq 0) { return $Html }
    $sb   = New-Object System.Text.StringBuilder
    $last = 0
    $known = @{}
    foreach ($m in $mc) {
        [void]$sb.Append($Html.Substring($last, $m.Index - $last))
        $src  = $m.Groups[2].Value
        $path = $null
        if ($src -match '^file:///?') {
            $path = [System.Uri]::UnescapeDataString(($src -replace '^file:///?', '')) -replace '/', '\'
        } elseif ($src -match '^([a-zA-Z]:\\|\\\\)') {
            $path = $src
        }
        $replacement = $m.Value
        if ($path -and (Test-Path -LiteralPath $path)) {
            $key = $path.ToLowerInvariant()
            if (-not $known.ContainsKey($key)) {
                $lr = New-Object System.Net.Mail.LinkedResource($path, (Get-MimeType $path))
                $lr.ContentId        = [guid]::NewGuid().ToString('N')
                $lr.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
                [void]$Resources.Add($lr)
                $known[$key] = $lr.ContentId
            }
            $replacement = $m.Groups[1].Value + 'cid:' + $known[$key] + $m.Groups[3].Value
        }
        [void]$sb.Append($replacement)
        $last = $m.Index + $m.Length
    }
    [void]$sb.Append($Html.Substring($last))
    return $sb.ToString()
}

function Get-EditorHtml {
    param($Browser)
    try {
        if ($Browser -eq $wbEditor -and $script:SourceMode) { return $txtSource.Text }
        $h = $Browser.Document.Body.InnerHtml
        if ($null -eq $h) { return '' }
        return $h
    } catch { return '' }
}
function Set-EditorHtml {
    param($Browser, [string]$Html)
    try { $Browser.Document.Body.InnerHtml = $Html } catch { Write-Log "Impossible de charger le contenu : $($_.Exception.Message)" 'ERROR' }
}

function Get-SignatureHtml {
    if ($script:EditorsReady -and $wbSig -and $wbSig.Document) {
        $h = Get-EditorHtml $wbSig
        if ($h -and ($h -replace '<[^>]+>|&nbsp;|\s', '') -ne '') { return $h }
        return ''
    }
    if (Test-Path $script:SignatureFile) { return (Get-Content -Path $script:SignatureFile -Raw -Encoding UTF8) }
    return ''
}

function Get-FinalHtml {
    param([switch]$IncludeSignature)
    $body = Get-EditorHtml $wbEditor
    $sig  = ''
    if ($IncludeSignature) {
        $s = Get-SignatureHtml
        if ($s) { $sig = "<br><br>$s" }
    }
    $title = [System.Net.WebUtility]::HtmlEncode($txtSubject.Text)
    return @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>$title</title></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222222;">
$body
$sig
</body></html>
"@
}
#endregion

#region ================= MODELES & BLOCS HTML =================
$script:Templates = [ordered]@{}

$script:Templates['Procédure (pas à pas)'] = @'
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;">
<div style="background:#1F4E79;color:#ffffff;padding:14px 18px;font-size:16pt;font-weight:bold;">PROCÉDURE &ndash; [Titre de la procédure]</div>
<table style="width:100%;border-collapse:collapse;margin:12px 0;font-size:10pt;">
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#f3f6fa;width:22%;"><b>Référence</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">PROC-IT-001</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#f3f6fa;"><b>Date</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">{{DATE}}</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#f3f6fa;"><b>Public concerné</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">Tous les utilisateurs</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#f3f6fa;"><b>Durée estimée</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">5 minutes</td></tr>
</table>
<p>Bonjour {{PRENOM}},</p>
<p>Veuillez trouver ci-dessous la procédure à suivre pour <b>[objectif de la procédure]</b>.</p>
<h3 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:3px;margin-top:20px;">1. Contexte</h3>
<p>[Expliquer pourquoi cette procédure est mise en place.]</p>
<h3 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:3px;margin-top:20px;">2. Prérequis</h3>
<ul><li>[Prérequis 1]</li><li>[Prérequis 2]</li></ul>
<h3 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:3px;margin-top:20px;">3. Étapes à suivre</h3>
<ol>
<li style="margin-bottom:10px;"><b>Étape 1 &ndash; [Titre]</b><br>[Description détaillée]<br><i style="color:#888;">[Insérer une capture d'écran : Insertion &gt; Image]</i></li>
<li style="margin-bottom:10px;"><b>Étape 2 &ndash; [Titre]</b><br>[Description détaillée]</li>
<li style="margin-bottom:10px;"><b>Étape 3 &ndash; [Titre]</b><br>[Description détaillée]</li>
</ol>
<div style="border-left:5px solid #E67E22;background:#FDF2E9;padding:10px 14px;margin:14px 0;"><b>&#9888; Important :</b> [Point de vigilance, action irréversible, etc.]</div>
<h3 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:3px;margin-top:20px;">4. Vérification</h3>
<p>[Comment vérifier que l'opération a réussi.]</p>
<h3 style="color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:3px;margin-top:20px;">5. En cas de problème</h3>
<p>Contactez le support informatique : <a href="mailto:support@domaine.fr">support@domaine.fr</a> &ndash; Tél. 01 23 45 67 89</p>
<p>Cordialement,</p>
</div>
'@

$script:Templates['Information / Communication IT'] = @'
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;">
<div style="background:#2E75B6;color:#ffffff;padding:12px 18px;font-size:15pt;font-weight:bold;">&#8505; Information &ndash; [Sujet]</div>
<p style="margin-top:14px;">Bonjour {{PRENOM}},</p>
<p>Le service informatique vous informe que <b>[nouveauté / changement]</b> à compter du <b>[date]</b>.</p>
<h3 style="color:#2E75B6;">Ce qui change pour vous</h3>
<ul><li>[Changement 1]</li><li>[Changement 2]</li></ul>
<h3 style="color:#2E75B6;">Ce que vous devez faire</h3>
<ul><li>[Action attendue]</li></ul>
<div style="border-left:5px solid #2E75B6;background:#EAF2FB;padding:10px 14px;margin:14px 0;"><b>&#8505; Bon à savoir :</b> [Précision utile]</div>
<p>Nous restons à votre disposition pour toute question.</p>
<p>Cordialement,</p>
</div>
'@

$script:Templates['Maintenance planifiée'] = @'
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;">
<div style="background:#C00000;color:#ffffff;padding:12px 18px;font-size:15pt;font-weight:bold;">&#9881; Maintenance planifiée &ndash; [Service concerné]</div>
<table style="width:100%;border-collapse:collapse;margin:12px 0;font-size:10pt;">
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#fbeaea;width:22%;"><b>Début</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">[JJ/MM/AAAA à HHhMM]</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#fbeaea;"><b>Fin prévue</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">[JJ/MM/AAAA à HHhMM]</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#fbeaea;"><b>Services impactés</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">[Messagerie, SharePoint, VPN...]</td></tr>
<tr><td style="padding:5px 8px;border:1px solid #d0d0d0;background:#fbeaea;"><b>Impact</b></td><td style="padding:5px 8px;border:1px solid #d0d0d0;">[Interruption totale / dégradation]</td></tr>
</table>
<p>Bonjour {{PRENOM}},</p>
<p>Une opération de maintenance est planifiée sur <b>[service]</b>. Pendant cette période, <b>[description de l'impact]</b>.</p>
<div style="border-left:5px solid #C00000;background:#FBEAEA;padding:10px 14px;margin:14px 0;"><b>&#9888; Recommandation :</b> enregistrez vos travaux et fermez vos sessions avant le début de l'intervention.</div>
<p>Nous vous prions de nous excuser pour la gêne occasionnée.</p>
<p>Cordialement,</p>
</div>
'@

$script:Templates['Sensibilisation sécurité'] = @'
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222;">
<div style="background:#385723;color:#ffffff;padding:12px 18px;font-size:15pt;font-weight:bold;">&#128274; Sécurité &ndash; [Sujet : phishing, mots de passe, MFA...]</div>
<p style="margin-top:14px;">Bonjour {{PRENOM}},</p>
<p>[Contexte : campagne de phishing en cours, rappel des bonnes pratiques...]</p>
<h3 style="color:#385723;">Les bons réflexes</h3>
<ol>
<li>Vérifiez toujours l'adresse de l'expéditeur.</li>
<li>Ne cliquez jamais sur un lien suspect ; survolez-le pour voir la destination réelle.</li>
<li>Ne communiquez jamais votre mot de passe, même au service informatique.</li>
<li>En cas de doute, transférez le mail à <a href="mailto:securite@domaine.fr">securite@domaine.fr</a>.</li>
</ol>
<div style="border-left:5px solid #385723;background:#EEF5E9;padding:10px 14px;margin:14px 0;"><b>&#10004; À retenir :</b> le service informatique ne vous demandera jamais vos identifiants par mail.</div>
<p>Merci pour votre vigilance.</p>
<p>Cordialement,</p>
</div>
'@

$script:Templates['Vide (page blanche)'] = '<p>Bonjour {{PRENOM}},</p><p><br></p><p>Cordialement,</p>'

$script:Blocks = @{
    Info  = '<div style="border-left:5px solid #2E75B6;background:#EAF2FB;padding:10px 14px;margin:12px 0;"><b>&#8505; Information :</b> [texte]</div><p><br></p>'
    Warn  = '<div style="border-left:5px solid #E67E22;background:#FDF2E9;padding:10px 14px;margin:12px 0;"><b>&#9888; Attention :</b> [texte]</div><p><br></p>'
    Ok    = '<div style="border-left:5px solid #2E8B57;background:#EAF7EE;padding:10px 14px;margin:12px 0;"><b>&#10004; Succès :</b> [texte]</div><p><br></p>'
    Err   = '<div style="border-left:5px solid #C00000;background:#FBEAEA;padding:10px 14px;margin:12px 0;"><b>&#10006; Erreur :</b> [texte]</div><p><br></p>'
    Code  = '<pre style="background:#f4f4f4;border:1px solid #ddd;border-radius:4px;padding:10px;font-family:Consolas,''Courier New'',monospace;font-size:10pt;">[commande ou code]</pre><p><br></p>'
    Step  = '<div style="margin:10px 0;padding:8px 12px;border:1px solid #d0d0d0;border-radius:4px;background:#fafafa;"><span style="display:inline-block;background:#1F4E79;color:#fff;border-radius:50%;width:24px;height:24px;line-height:24px;text-align:center;font-weight:bold;margin-right:8px;">N</span><b>[Titre de l''étape]</b><br>[Description]</div><p><br></p>'
    Button= '<p><a href="https://" style="display:inline-block;background:#1F4E79;color:#ffffff;padding:10px 22px;text-decoration:none;border-radius:4px;font-weight:bold;">[Texte du bouton]</a></p>'
}
#endregion

#region ================= EDITEUR : COMMANDES =================
function Invoke-EditorCommand {
    param($Browser, [string]$Command, $Value = $null)
    try {
        $Browser.Document.Focus()
        $Browser.Document.ExecCommand($Command, $false, $Value)
    } catch { Write-Log "Commande '$Command' impossible : $($_.Exception.Message)" 'WARN' }
}

function Insert-EditorHtml {
    param($Browser, [string]$Html)
    try {
        $Browser.Document.Focus()
        $doc   = $Browser.Document.DomDocument
        $sel   = $doc.selection
        $range = $sel.createRange()
        $range.pasteHTML($Html)
    } catch {
        try { $Browser.Document.Body.InnerHtml = $Browser.Document.Body.InnerHtml + $Html } catch { }
    }
}

function Get-EditorSelectionText {
    param($Browser)
    try { return [string]$Browser.Document.DomDocument.selection.createRange().text } catch { return '' }
}

function Invoke-EditorAction {
    param($Browser, [string]$Action)
    switch ($Action) {
        'ForeColor' {
            $dlg = New-Object System.Windows.Forms.ColorDialog; $dlg.FullOpen = $true
            if ($dlg.ShowDialog() -eq 'OK') { Invoke-EditorCommand $Browser 'ForeColor' (ConvertTo-HtmlColor $dlg.Color) }
        }
        'BackColor' {
            $dlg = New-Object System.Windows.Forms.ColorDialog; $dlg.FullOpen = $true
            if ($dlg.ShowDialog() -eq 'OK') { Invoke-EditorCommand $Browser 'BackColor' (ConvertTo-HtmlColor $dlg.Color) }
        }
        'Link' {
            $url = Read-Input "Adresse du lien (http://, https://, mailto:) :" 'Insérer un lien' 'https://'
            if (-not $url) { return }
            $selText = Get-EditorSelectionText $Browser
            if ($selText.Trim()) {
                Invoke-EditorCommand $Browser 'CreateLink' $url
            } else {
                $txt = Read-Input "Texte affiché pour le lien :" 'Insérer un lien' $url
                if (-not $txt) { $txt = $url }
                Insert-EditorHtml $Browser ("<a href=`"{0}`">{1}</a>" -f $url, [System.Net.WebUtility]::HtmlEncode($txt))
            }
        }
        'Image' {
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title  = "Sélectionner une image (elle sera embarquée dans le mail)"
            $dlg.Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.bmp|Tous les fichiers|*.*"
            if ($dlg.ShowDialog() -ne 'OK') { return }
            Insert-ImageFromPath $Browser $dlg.FileName
        }
        'ImagePath' {
            $p = Read-Input "Chemin complet de l'image (local ou UNC) :" 'Insérer une image par chemin' ''
            if (-not $p) { return }
            $p = $p.Trim('"')
            if (-not (Test-Path -LiteralPath $p)) { Show-Msg "Fichier introuvable :`n$p" 'Image' 'Warning'; return }
            Insert-ImageFromPath $Browser $p
        }
        'ImageUrl' {
            $u = Read-Input "URL de l'image (https://...) :`nAttention : l'image ne sera pas embarquée, elle doit rester accessible en ligne." 'Insérer une image distante' 'https://'
            if (-not $u -or $u -eq 'https://') { return }
            Insert-EditorHtml $Browser ("<img src=`"{0}`" style=`"max-width:100%;`" alt=`"`">" -f $u)
        }
        'Table' {
            $r = Read-Input "Nombre de lignes :" 'Insérer un tableau' '3'
            if (-not $r) { return }
            $c = Read-Input "Nombre de colonnes :" 'Insérer un tableau' '3'
            if (-not $c) { return }
            if ($r -notmatch '^\d+$' -or $c -notmatch '^\d+$') { Show-Msg 'Valeurs numériques attendues.' 'Tableau' 'Warning'; return }
            $rows = [int]$r; $cols = [int]$c
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.Append('<table style="border-collapse:collapse;width:100%;margin:10px 0;">')
            for ($i = 0; $i -lt $rows; $i++) {
                [void]$sb.Append('<tr>')
                for ($j = 0; $j -lt $cols; $j++) {
                    if ($i -eq 0) { [void]$sb.Append('<th style="border:1px solid #bfbfbf;background:#1F4E79;color:#fff;padding:6px 8px;text-align:left;">Titre</th>') }
                    else          { [void]$sb.Append('<td style="border:1px solid #bfbfbf;padding:6px 8px;">&nbsp;</td>') }
                }
                [void]$sb.Append('</tr>')
            }
            [void]$sb.Append('</table><p><br></p>')
            Insert-EditorHtml $Browser $sb.ToString()
        }
        'Hr'        { Invoke-EditorCommand $Browser 'InsertHorizontalRule' }
        'BlockInfo' { Insert-EditorHtml $Browser $script:Blocks.Info }
        'BlockWarn' { Insert-EditorHtml $Browser $script:Blocks.Warn }
        'BlockOk'   { Insert-EditorHtml $Browser $script:Blocks.Ok }
        'BlockErr'  { Insert-EditorHtml $Browser $script:Blocks.Err }
        'BlockCode' { Insert-EditorHtml $Browser $script:Blocks.Code }
        'BlockStep' { Insert-EditorHtml $Browser $script:Blocks.Step }
        'BlockBtn'  { Insert-EditorHtml $Browser $script:Blocks.Button }
        'Date'      { Insert-EditorHtml $Browser (Get-Date -Format 'dd/MM/yyyy') }
        'Signature' {
            $s = Get-SignatureHtml
            if (-not $s) { Show-Msg "Aucune signature définie. Créez-la dans l'onglet 'Signature'." 'Signature' 'Warning'; return }
            Insert-EditorHtml $Browser ("<br><br>" + $s)
        }
        'RemoveFormat' { Invoke-EditorCommand $Browser 'RemoveFormat' }
        'SelectAll'    { Invoke-EditorCommand $Browser 'SelectAll' }
    }
}

function Insert-ImageFromPath {
    param($Browser, [string]$Path)
    $w = Read-Input "Largeur en pixels (vide = taille réelle) :" 'Taille de l''image' ''
    $style = 'max-width:100%;'
    if ($w -match '^\d+$') { $style = "width:${w}px;max-width:100%;" }
    $html = "<img src=`"{0}`" style=`"{1}`" alt=`"{2}`">" -f (ConvertTo-FileUri $Path), $style, [System.IO.Path]::GetFileName($Path)
    Insert-EditorHtml $Browser $html
    Write-Log "Image insérée : $Path" 'INFO'
}

# Handlers génériques (Tag = @{ Browser; Cmd; Value } ou @{ Browser; Action })
$script:OnCmd    = { $t = $this.Tag; Invoke-EditorCommand -Browser $t.Browser -Command $t.Cmd -Value $t.Value }
$script:OnAction = { $t = $this.Tag; Invoke-EditorAction  -Browser $t.Browser -Action  $t.Action }
$script:OnInsert = { $t = $this.Tag; Insert-EditorHtml    -Browser $t.Browser -Html    $t.Html }

function Add-TsButton {
    param($Strip, [string]$Text, [string]$Tip, [hashtable]$Tag, [scriptblock]$Handler, $Font)
    $b = New-Object System.Windows.Forms.ToolStripButton
    $b.Text = $Text; $b.ToolTipText = $Tip; $b.DisplayStyle = 'Text'; $b.Tag = $Tag
    if ($Font) { $b.Font = $Font }
    $b.Add_Click($Handler)
    [void]$Strip.Items.Add($b)
}
function Add-TsSep { param($Strip) [void]$Strip.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) }

function Build-FormatToolbar {
    param($Strip, $Browser)
    # Police
    $cmbFont = New-Object System.Windows.Forms.ToolStripComboBox
    $cmbFont.Items.AddRange(@('Calibri','Arial','Segoe UI','Verdana','Tahoma','Trebuchet MS','Georgia','Times New Roman','Courier New','Consolas'))
    $cmbFont.DropDownStyle = 'DropDownList'; $cmbFont.Width = 130; $cmbFont.ToolTipText = 'Police'; $cmbFont.Tag = $Browser
    $cmbFont.Add_SelectedIndexChanged({ if ($this.SelectedItem) { Invoke-EditorCommand -Browser $this.Tag -Command 'FontName' -Value $this.SelectedItem } })
    [void]$Strip.Items.Add($cmbFont)
    # Taille
    $cmbSize = New-Object System.Windows.Forms.ToolStripComboBox
    $cmbSize.Items.AddRange(@('8','10','12','14','18','24','36'))
    $cmbSize.DropDownStyle = 'DropDownList'; $cmbSize.Width = 55; $cmbSize.ToolTipText = 'Taille'; $cmbSize.Tag = $Browser
    $cmbSize.Add_SelectedIndexChanged({ if ($this.SelectedIndex -ge 0) { Invoke-EditorCommand -Browser $this.Tag -Command 'FontSize' -Value ($this.SelectedIndex + 1) } })
    [void]$Strip.Items.Add($cmbSize)
    # Style de bloc
    $cmbBlock = New-Object System.Windows.Forms.ToolStripComboBox
    $cmbBlock.Items.AddRange(@('Paragraphe','Titre 1','Titre 2','Titre 3','Titre 4'))
    $cmbBlock.DropDownStyle = 'DropDownList'; $cmbBlock.Width = 100; $cmbBlock.ToolTipText = 'Style de paragraphe'; $cmbBlock.Tag = $Browser
    $cmbBlock.Add_SelectedIndexChanged({
        $map = @('<P>','<H1>','<H2>','<H3>','<H4>')
        if ($this.SelectedIndex -ge 0) { Invoke-EditorCommand -Browser $this.Tag -Command 'FormatBlock' -Value $map[$this.SelectedIndex] }
    })
    [void]$Strip.Items.Add($cmbBlock)
    Add-TsSep $Strip
    Add-TsButton $Strip 'G'  'Gras (Ctrl+B)'      @{ Browser=$Browser; Cmd='Bold' }          $script:OnCmd $FontBold
    Add-TsButton $Strip 'I'  'Italique (Ctrl+I)'  @{ Browser=$Browser; Cmd='Italic' }        $script:OnCmd $FontItalic
    Add-TsButton $Strip 'S'  'Souligné (Ctrl+U)'  @{ Browser=$Browser; Cmd='Underline' }     $script:OnCmd $FontUnder
    Add-TsButton $Strip 'ab' 'Barré'              @{ Browser=$Browser; Cmd='StrikeThrough' } $script:OnCmd $FontStrike
    Add-TsButton $Strip 'x²' 'Exposant'           @{ Browser=$Browser; Cmd='Superscript' }   $script:OnCmd
    Add-TsSep $Strip
    Add-TsButton $Strip 'A Couleur'   'Couleur du texte'     @{ Browser=$Browser; Action='ForeColor' } $script:OnAction
    Add-TsButton $Strip 'Surlignage'  'Couleur de fond'      @{ Browser=$Browser; Action='BackColor' } $script:OnAction
    Add-TsSep $Strip
    Add-TsButton $Strip '≡ G' 'Aligner à gauche' @{ Browser=$Browser; Cmd='JustifyLeft' }   $script:OnCmd
    Add-TsButton $Strip '≡ C' 'Centrer'          @{ Browser=$Browser; Cmd='JustifyCenter' } $script:OnCmd
    Add-TsButton $Strip '≡ D' 'Aligner à droite' @{ Browser=$Browser; Cmd='JustifyRight' }  $script:OnCmd
    Add-TsButton $Strip '≡ J' 'Justifier'        @{ Browser=$Browser; Cmd='JustifyFull' }   $script:OnCmd
    Add-TsSep $Strip
    Add-TsButton $Strip '• Liste'  'Liste à puces'    @{ Browser=$Browser; Cmd='InsertUnorderedList' } $script:OnCmd
    Add-TsButton $Strip '1. Liste' 'Liste numérotée'  @{ Browser=$Browser; Cmd='InsertOrderedList' }   $script:OnCmd
    Add-TsButton $Strip '→ Retrait' 'Augmenter le retrait' @{ Browser=$Browser; Cmd='Indent' }  $script:OnCmd
    Add-TsButton $Strip '← Retrait' 'Diminuer le retrait'  @{ Browser=$Browser; Cmd='Outdent' } $script:OnCmd
    Add-TsSep $Strip
    Add-TsButton $Strip '↶ Annuler'  'Annuler (Ctrl+Z)'  @{ Browser=$Browser; Cmd='Undo' } $script:OnCmd
    Add-TsButton $Strip '↷ Rétablir' 'Rétablir (Ctrl+Y)' @{ Browser=$Browser; Cmd='Redo' } $script:OnCmd
    Add-TsButton $Strip 'Effacer format' 'Supprimer la mise en forme de la sélection' @{ Browser=$Browser; Action='RemoveFormat' } $script:OnAction
    Add-TsButton $Strip 'Tout sélect.' 'Tout sélectionner' @{ Browser=$Browser; Action='SelectAll' } $script:OnAction
}

function Build-InsertToolbar {
    param($Strip, $Browser, [switch]$WithTemplates)
    $lbl = New-Object System.Windows.Forms.ToolStripLabel; $lbl.Text = 'Insérer :'; $lbl.Font = $FontBold; [void]$Strip.Items.Add($lbl)
    Add-TsButton $Strip 'Lien'            'Insérer un lien hypertexte'                @{ Browser=$Browser; Action='Link' }      $script:OnAction
    Add-TsButton $Strip 'Image (fichier)' 'Parcourir et insérer une image (embarquée)' @{ Browser=$Browser; Action='Image' }     $script:OnAction
    Add-TsButton $Strip 'Image (chemin)'  'Saisir le chemin d''une image'             @{ Browser=$Browser; Action='ImagePath' } $script:OnAction
    Add-TsButton $Strip 'Image (URL)'     'Image distante (non embarquée)'            @{ Browser=$Browser; Action='ImageUrl' }  $script:OnAction
    Add-TsButton $Strip 'Tableau'         'Insérer un tableau'                        @{ Browser=$Browser; Action='Table' }     $script:OnAction
    Add-TsButton $Strip 'Ligne'           'Ligne horizontale'                         @{ Browser=$Browser; Action='Hr' }        $script:OnAction
    Add-TsButton $Strip 'Date'            'Insérer la date du jour'                   @{ Browser=$Browser; Action='Date' }      $script:OnAction
    Add-TsSep $Strip
    # Blocs
    $ddBlocks = New-Object System.Windows.Forms.ToolStripDropDownButton
    $ddBlocks.Text = 'Blocs'; $ddBlocks.ToolTipText = 'Encadrés prêts à l''emploi'
    $blockItems = [ordered]@{
        'Bloc Information (bleu)' = 'BlockInfo'; 'Bloc Attention (orange)' = 'BlockWarn'
        'Bloc Succès (vert)' = 'BlockOk'; 'Bloc Erreur (rouge)' = 'BlockErr'
        'Bloc Code / Commande' = 'BlockCode'; 'Étape numérotée' = 'BlockStep'; 'Bouton (lien)' = 'BlockBtn'
    }
    foreach ($k in $blockItems.Keys) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem; $mi.Text = $k
        $mi.Tag = @{ Browser=$Browser; Action=$blockItems[$k] }
        $mi.Add_Click($script:OnAction)
        [void]$ddBlocks.DropDownItems.Add($mi)
    }
    [void]$Strip.Items.Add($ddBlocks)
    # Variables
    $ddVars = New-Object System.Windows.Forms.ToolStripDropDownButton
    $ddVars.Text = 'Variables'; $ddVars.ToolTipText = 'Remplacées à l''envoi pour chaque destinataire'
    foreach ($v in @('{{PRENOM}}','{{NOM}}','{{NOMCOMPLET}}','{{EMAIL}}','{{DATE}}','{{HEURE}}')) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem; $mi.Text = $v
        $mi.Tag = @{ Browser=$Browser; Html=$v }
        $mi.Add_Click($script:OnInsert)
        [void]$ddVars.DropDownItems.Add($mi)
    }
    [void]$Strip.Items.Add($ddVars)
    Add-TsSep $Strip
    if ($WithTemplates) {
        $ddTpl = New-Object System.Windows.Forms.ToolStripDropDownButton
        $ddTpl.Text = 'Modèles'; $ddTpl.ToolTipText = 'Charger un modèle de mail (remplace le contenu)'; $ddTpl.Font = $FontBold
        foreach ($k in $script:Templates.Keys) {
            $mi = New-Object System.Windows.Forms.ToolStripMenuItem; $mi.Text = $k
            $mi.Tag = @{ Browser=$Browser; Key=$k }
            $mi.Add_Click({
                $t = $this.Tag
                $cur = Get-EditorHtml $t.Browser
                if (($cur -replace '<[^>]+>|&nbsp;|\s', '') -ne '' -and -not (Confirm-Msg "Le contenu actuel de l'éditeur sera remplacé par le modèle '$($t.Key)'.`nContinuer ?")) { return }
                Set-EditorHtml $t.Browser $script:Templates[$t.Key]
                Write-Log "Modèle chargé : $($t.Key)" 'INFO'
            })
            [void]$ddTpl.DropDownItems.Add($mi)
        }
        [void]$Strip.Items.Add($ddTpl)
        Add-TsButton $Strip 'Signature' 'Insérer la signature à la position du curseur' @{ Browser=$Browser; Action='Signature' } $script:OnAction $FontBold
    }
}

function New-EditorDocument {
    param([string]$Path)
    $html = @'
<!DOCTYPE html>
<html><head><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta charset="utf-8">
<style>
 body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color:#222; margin: 14px; background:#fff; }
 table { border-collapse: collapse; }
 a { color:#1F4E79; }
</style></head><body></body></html>
'@
    [System.IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($true)))
}

function Initialize-Editor {
    param($Browser, [string]$File, [string]$InitialHtml = '')
    New-EditorDocument $File
    $Browser.Navigate($File)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($Browser.ReadyState -ne 'Complete' -and $sw.ElapsedMilliseconds -lt 8000) {
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 40
    }
    try {
        $Browser.Document.Body.SetAttribute('contentEditable', 'true')
        if ($InitialHtml) { $Browser.Document.Body.InnerHtml = $InitialHtml }
    } catch { Write-Log "Initialisation éditeur : $($_.Exception.Message)" 'ERROR' }
}
#endregion

#region ================= SMTP / ENVOI =================
function New-SmtpClient {
    $c = New-Object System.Net.Mail.SmtpClient($txtServer.Text.Trim(), [int]$numPort.Value)
    $c.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $c.EnableSsl      = $chkSsl.Checked
    $c.Timeout        = [int]$numTimeout.Value * 1000
    $c.UseDefaultCredentials = $false
    if ($chkAuth.Checked) {
        $c.Credentials = New-Object System.Net.NetworkCredential($txtUser.Text, $txtPass.Text)
    } else {
        $c.Credentials = $null   # relais anonyme (Exchange receive connector)
    }
    if ($chkIgnoreCert.Checked) {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    } else {
        [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }
    return $c
}

function Test-SmtpSettings {
    if ([string]::IsNullOrWhiteSpace($txtServer.Text)) { Show-Msg "Renseignez le serveur SMTP (onglet Serveur)." 'Configuration' 'Warning'; return $false }
    if (-not (Test-EmailAddress $txtFrom.Text.Trim())) { Show-Msg "L'adresse expéditeur est invalide (onglet Serveur)." 'Configuration' 'Warning'; return $false }
    if ($chkAuth.Checked -and [string]::IsNullOrWhiteSpace($txtUser.Text)) { Show-Msg "Authentification activée : renseignez l'utilisateur." 'Configuration' 'Warning'; return $false }
    if ($txtReplyTo.Text.Trim() -and -not (Test-EmailAddress $txtReplyTo.Text.Trim())) { Show-Msg "L'adresse 'Répondre à' est invalide." 'Configuration' 'Warning'; return $false }
    return $true
}

function Send-OneMail {
    param($Client, [string[]]$To, [string[]]$Bcc = @(), [string]$Subject, [string]$Html, [hashtable]$Vars, [switch]$DryRun)
    $msg = New-Object System.Net.Mail.MailMessage
    try {
        $msg.From = New-Object System.Net.Mail.MailAddress($txtFrom.Text.Trim(), $txtFromName.Text.Trim(), [System.Text.Encoding]::UTF8)
        foreach ($t in $To)  { if ($t) { $msg.To.Add($t) } }
        foreach ($b in $Bcc) { if ($b) { $msg.Bcc.Add($b) } }
        foreach ($cc in ($txtCc.Text -split '[;,\s]+' | Where-Object { Test-EmailAddress $_ })) { $msg.CC.Add($cc) }
        if ($txtReplyTo.Text.Trim()) { $msg.ReplyToList.Add($txtReplyTo.Text.Trim()) }
        $msg.Subject         = Expand-Variables $Subject $Vars
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        $msg.HeadersEncoding = [System.Text.Encoding]::UTF8
        $msg.Priority        = switch ($cmbPriority.SelectedIndex) { 1 { 'High' } 2 { 'Low' } default { 'Normal' } }
        $msg.Headers.Add('X-Mailer', "$($script:AppName)/$($script:AppVersion)")

        $body = Expand-Variables $Html $Vars
        $res  = New-Object System.Collections.ArrayList
        $body = ConvertTo-InlineImages -Html $body -Resources $res

        $plainView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString((ConvertTo-PlainText $body), [System.Text.Encoding]::UTF8, 'text/plain')
        $htmlView  = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($body, [System.Text.Encoding]::UTF8, 'text/html')
        foreach ($lr in $res) { $htmlView.LinkedResources.Add($lr) }
        $msg.AlternateViews.Add($plainView)
        $msg.AlternateViews.Add($htmlView)
        $msg.IsBodyHtml = $true

        foreach ($item in $lstAttach.Items) {
            if (Test-Path -LiteralPath $item) { $msg.Attachments.Add((New-Object System.Net.Mail.Attachment([string]$item))) }
        }
        if ($DryRun) {
            Write-Log ("SIMULATION -> To={0} Bcc={1} Objet='{2}' Images={3} PJ={4}" -f ($To -join ','), $Bcc.Count, $msg.Subject, $res.Count, $msg.Attachments.Count) 'DEBUG'
            return
        }
        $Client.Send($msg)
    } finally {
        $msg.Dispose()
    }
}

function Add-Result {
    param([string]$Email, [string]$Status, [string]$Detail)
    [void]$script:Results.Add([pscustomobject]@{
        Horodatage = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Destinataire = $Email; Statut = $Status; Detail = $Detail
    })
}

function Test-SmtpConnection {
    $srv = $txtServer.Text.Trim(); $port = [int]$numPort.Value
    if (-not $srv) { Show-Msg "Renseignez le serveur SMTP." 'Test' 'Warning'; return }
    Write-Log "Test de connexion vers ${srv}:$port ..."
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($srv, $port, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne(5000)) { throw "Délai de connexion dépassé (5 s)" }
        $tcp.EndConnect($ar)
        $banner = ''
        if (-not ($chkSsl.Checked -and $port -eq 465)) {
            try {
                $stream = $tcp.GetStream(); $stream.ReadTimeout = 3000
                $buf = New-Object byte[] 1024
                $n = $stream.Read($buf, 0, $buf.Length)
                if ($n -gt 0) { $banner = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n).Trim() }
            } catch { }
        }
        $tcp.Close()
        Write-Log "Connexion TCP réussie vers ${srv}:$port. $banner" 'OK'
        Show-Msg "Connexion réussie vers ${srv}:$port`n`nBannière : $banner" 'Test de connexion'
    } catch {
        Write-Log "Échec de connexion : $($_.Exception.Message)" 'ERROR'
        Show-Msg "Échec de connexion :`n$($_.Exception.Message)" 'Test de connexion' 'Error'
    }
}

function Send-TestMail {
    if (-not (Test-SmtpSettings)) { return }
    $to = Read-Input "Adresse de destination du mail de test :" 'Mail de test' $txtFrom.Text
    if (-not (Test-EmailAddress $to)) { return }
    $client = New-SmtpClient
    try {
        $subject = if ($txtSubject.Text.Trim()) { "[TEST] " + $txtSubject.Text } else { "[TEST] $($script:AppName)" }
        Send-OneMail -Client $client -To @($to) -Subject $subject -Html (Get-FinalHtml -IncludeSignature:($chkAutoSig.Checked)) -Vars (Get-VarsForEmail $to)
        Write-Log "Mail de test envoyé à $to" 'OK'
        Show-Msg "Mail de test envoyé à $to" 'Mail de test'
    } catch {
        Write-Log "Échec du mail de test : $($_.Exception.Message)" 'ERROR'
        Show-Msg "Échec :`n$($_.Exception.Message)" 'Mail de test' 'Error'
    } finally { $client.Dispose() }
}

function Start-Campaign {
    $recips = Get-RecipientList -ValidOnly
    $all    = Get-RecipientList
    if ($recips.Count -eq 0) { Show-Msg "Aucun destinataire valide (onglet Destinataires)." 'Envoi' 'Warning'; return }
    if (-not (Test-SmtpSettings)) { return }
    if ([string]::IsNullOrWhiteSpace($txtSubject.Text)) { Show-Msg "L'objet du mail est vide (onglet Éditeur)." 'Envoi' 'Warning'; return }
    $body = Get-EditorHtml $wbEditor
    if (($body -replace '<[^>]+>|&nbsp;|\s', '') -eq '' -and -not (Confirm-Msg "Le corps du mail est vide. Envoyer quand même ?")) { return }

    $final   = Get-FinalHtml -IncludeSignature:($chkAutoSig.Checked)
    $perUser = $rbPerRecipient.Checked
    $dry     = $chkDryRun.Checked
    $invalid = $all.Count - $recips.Count
    if (-not $perUser -and ($final -match '\{\{\s*(PRENOM|NOM|NOMCOMPLET|EMAIL)\s*\}\}')) {
        if (-not (Confirm-Msg "Le mail contient des variables personnalisées ({{PRENOM}}, {{NOM}}...) mais le mode 'Cci par lots' est sélectionné : elles seront remplacées par du vide.`n`nContinuer ?")) { return }
    }
    $modeTxt = if ($perUser) { "1 mail par destinataire (personnalisé)" } else { "Cci par lots de $([int]$numBatch.Value)" }
    $confirm = "Destinataires valides : $($recips.Count)" + $(if ($invalid) { "  (ignorés : $invalid)" }) + "`n" +
               "Mode : $modeTxt`n" +
               "Serveur : $($txtServer.Text):$($numPort.Value)  (auth : $(if ($chkAuth.Checked) {'oui'} else {'non'}), SSL : $(if ($chkSsl.Checked) {'oui'} else {'non'}))`n" +
               "Expéditeur : $($txtFrom.Text)`n" +
               "Objet : $($txtSubject.Text)`n" +
               "Pièces jointes : $($lstAttach.Items.Count)`n" +
               $(if ($dry) { "`n*** MODE SIMULATION : aucun mail ne sera envoyé ***`n" }) +
               "`nLancer la campagne ?"
    if (-not (Confirm-Msg $confirm "Confirmation d'envoi")) { return }

    $script:CancelSend = $false
    $script:Results.Clear()
    $btnSend.Enabled = $false; $btnCancel.Enabled = $true; $btnTest.Enabled = $false
    $pbSend.Value = 0; $pbSend.Maximum = $recips.Count
    $ok = 0; $ko = 0; $n = $recips.Count
    $delay = [int]$numDelay.Value
    Write-Log ("===== DÉBUT CAMPAGNE : {0} destinataire(s), mode '{1}'{2} =====" -f $n, $modeTxt, $(if ($dry) { ' [SIMULATION]' })) 'INFO'
    $client = New-SmtpClient
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($perUser) {
            $i = 0
            foreach ($r in $recips) {
                if ($script:CancelSend) { Write-Log "Envoi annulé par l'utilisateur." 'WARN'; break }
                $i++
                try {
                    Send-OneMail -Client $client -To @($r) -Subject $txtSubject.Text -Html $final -Vars (Get-VarsForEmail $r) -DryRun:$dry
                    $ok++; Write-Log "($i/$n) OK -> $r" 'OK'; Add-Result $r 'OK' ''
                } catch {
                    $ko++; Write-Log "($i/$n) ÉCHEC -> $r : $($_.Exception.Message)" 'ERROR'; Add-Result $r 'ECHEC' $_.Exception.Message
                }
                $pbSend.Value = $i
                [System.Windows.Forms.Application]::DoEvents()
                if ($delay -gt 0 -and $i -lt $n -and -not $dry) { Start-Sleep -Milliseconds $delay }
            }
        } else {
            $size = [int]$numBatch.Value; $b = 0; $lot = 0
            while ($b -lt $n) {
                if ($script:CancelSend) { Write-Log "Envoi annulé par l'utilisateur." 'WARN'; break }
                $lot++
                $chunk = @($recips[$b..([Math]::Min($b + $size, $n) - 1)])
                try {
                    Send-OneMail -Client $client -To @($txtFrom.Text.Trim()) -Bcc $chunk -Subject $txtSubject.Text -Html $final -Vars (Get-VarsForEmail '') -DryRun:$dry
                    $ok += $chunk.Count
                    Write-Log "Lot $lot OK -> $($chunk.Count) destinataire(s) en Cci" 'OK'
                    foreach ($c in $chunk) { Add-Result $c 'OK' "Lot $lot" }
                } catch {
                    $ko += $chunk.Count
                    Write-Log "Lot $lot ÉCHEC ($($chunk.Count) destinataires) : $($_.Exception.Message)" 'ERROR'
                    foreach ($c in $chunk) { Add-Result $c 'ECHEC' "Lot $lot : $($_.Exception.Message)" }
                }
                $b += $size
                $pbSend.Value = [Math]::Min($b, $n)
                [System.Windows.Forms.Application]::DoEvents()
                if ($delay -gt 0 -and $b -lt $n -and -not $dry) { Start-Sleep -Milliseconds $delay }
            }
        }
    } finally {
        $client.Dispose()
        $btnSend.Enabled = $true; $btnCancel.Enabled = $false; $btnTest.Enabled = $true
    }
    $sw.Stop()
    Write-Log ("===== FIN CAMPAGNE : {0} OK, {1} échec(s), durée {2:mm\:ss} =====" -f $ok, $ko, $sw.Elapsed) $(if ($ko) { 'WARN' } else { 'OK' })
    Save-Config
    Show-Msg "Campagne terminée.`n`nEnvoyés : $ok`nÉchecs : $ko`nDurée : $($sw.Elapsed.ToString('mm\:ss'))" 'Envoi' $(if ($ko) { 'Warning' } else { 'Information' })
}
#endregion

#region ================= CONFIG =================
function Save-Config {
    try {
        $cfg = [ordered]@{
            Server = $txtServer.Text; Port = [int]$numPort.Value; Ssl = $chkSsl.Checked; Auth = $chkAuth.Checked
            User = $txtUser.Text; IgnoreCert = $chkIgnoreCert.Checked
            From = $txtFrom.Text; FromName = $txtFromName.Text; ReplyTo = $txtReplyTo.Text; Cc = $txtCc.Text
            DelayMs = [int]$numDelay.Value; TimeoutSec = [int]$numTimeout.Value; PerRecipient = $rbPerRecipient.Checked
            BatchSize = [int]$numBatch.Value; AutoSignature = $chkAutoSig.Checked; Priority = $cmbPriority.SelectedIndex
            Subject = $txtSubject.Text
        }
        $cfg | ConvertTo-Json | Set-Content -Path $script:ConfigFile -Encoding UTF8
    } catch { Write-Log "Sauvegarde config : $($_.Exception.Message)" 'WARN' }
}
function Import-Config {
    if (-not (Test-Path $script:ConfigFile)) { return }
    try {
        $c = Get-Content -Path $script:ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($c.Server)   { $txtServer.Text = $c.Server }
        if ($c.Port)     { $numPort.Value = [int]$c.Port }
        $chkSsl.Checked        = [bool]$c.Ssl
        $chkAuth.Checked       = [bool]$c.Auth
        $chkIgnoreCert.Checked = [bool]$c.IgnoreCert
        if ($c.User)     { $txtUser.Text = $c.User }
        if ($c.From)     { $txtFrom.Text = $c.From }
        if ($c.FromName) { $txtFromName.Text = $c.FromName }
        if ($c.ReplyTo)  { $txtReplyTo.Text = $c.ReplyTo }
        if ($c.Cc)       { $txtCc.Text = $c.Cc }
        if ($null -ne $c.DelayMs)    { $numDelay.Value = [int]$c.DelayMs }
        if ($c.TimeoutSec)           { $numTimeout.Value = [int]$c.TimeoutSec }
        if ($null -ne $c.PerRecipient) { $rbPerRecipient.Checked = [bool]$c.PerRecipient; $rbBcc.Checked = -not [bool]$c.PerRecipient }
        if ($c.BatchSize)            { $numBatch.Value = [int]$c.BatchSize }
        if ($null -ne $c.AutoSignature) { $chkAutoSig.Checked = [bool]$c.AutoSignature }
        if ($null -ne $c.Priority)   { $cmbPriority.SelectedIndex = [int]$c.Priority }
        if ($c.Subject)  { $txtSubject.Text = $c.Subject }
        Write-Log "Configuration chargée depuis $($script:ConfigFile)" 'DEBUG'
    } catch { Write-Log "Lecture config : $($_.Exception.Message)" 'WARN' }
}
#endregion

#region ================= FORMULAIRE PRINCIPAL =================
$form = New-Ctl 'Form' @{
    Text = "$($script:AppName) v$($script:AppVersion) - Campagne mailing Exchange / SMTP"
    Size = @(1280, 840); MinimumSize = @(1050, 700); StartPosition = 'CenterScreen'; Font = $FontUI
}
$tabs        = New-Ctl 'TabControl' @{ Dock = 'Fill'; Font = $FontBold } $form
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$lblStatus   = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Prêt.'; [void]$statusStrip.Items.Add($lblStatus)
[void]$form.Controls.Add($statusStrip)

$tabRecip  = New-Ctl 'TabPage' @{ Text = '  1. Destinataires  '; Font = $FontUI; Padding = (New-Object System.Windows.Forms.Padding(8)) }
$tabSmtp   = New-Ctl 'TabPage' @{ Text = '  2. Serveur SMTP  ';  Font = $FontUI; Padding = (New-Object System.Windows.Forms.Padding(8)) }
$tabEditor = New-Ctl 'TabPage' @{ Text = '  3. Éditeur de mail  '; Font = $FontUI }
$tabSig    = New-Ctl 'TabPage' @{ Text = '  4. Signature  ';     Font = $FontUI }
$tabSend   = New-Ctl 'TabPage' @{ Text = '  5. Envoi & Journal  '; Font = $FontUI; Padding = (New-Object System.Windows.Forms.Padding(8)) }
$tabs.TabPages.AddRange(@($tabRecip, $tabSmtp, $tabEditor, $tabSig, $tabSend))
#endregion

#region ================= ONGLET 1 : DESTINATAIRES =================
$txtRecipients = New-Ctl 'TextBox' @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Both'; AcceptsReturn = $true; WordWrap = $false; Font = $FontMono } $tabRecip
$lblRecipHead  = New-Ctl 'Label' @{ Dock = 'Top'; Height = 42; Text = "Saisissez une adresse par ligne (les séparateurs ; , et espace sont aussi acceptés), ou importez un fichier TXT / CSV.`nLes doublons et adresses invalides sont détectés automatiquement." } $tabRecip
$pnlRecipRight = New-Ctl 'Panel' @{ Dock = 'Right'; Width = 300; Padding = (New-Object System.Windows.Forms.Padding(10, 0, 0, 0)) } $tabRecip

$y = 0
$btnImport = New-Ctl 'Button' @{ Text = 'Importer un fichier TXT / CSV...'; Location = @(10, $y); Size = @(280, 34); Font = $FontBold } $pnlRecipRight; $y += 40
$btnPaste  = New-Ctl 'Button' @{ Text = 'Ajouter depuis le presse-papier';   Location = @(10, $y); Size = @(280, 30) } $pnlRecipRight; $y += 36
$btnClean  = New-Ctl 'Button' @{ Text = 'Nettoyer et dédoublonner';          Location = @(10, $y); Size = @(280, 30) } $pnlRecipRight; $y += 36
$btnCheck  = New-Ctl 'Button' @{ Text = 'Valider les adresses';              Location = @(10, $y); Size = @(280, 30) } $pnlRecipRight; $y += 36
$btnExport = New-Ctl 'Button' @{ Text = 'Exporter la liste (TXT)';           Location = @(10, $y); Size = @(280, 30) } $pnlRecipRight; $y += 36
$btnClear  = New-Ctl 'Button' @{ Text = 'Vider la liste';                    Location = @(10, $y); Size = @(280, 30) } $pnlRecipRight; $y += 44
$lblCount  = New-Ctl 'Label' @{ Text = '0 adresse'; Location = @(10, $y); AutoSize = $true; Font = $FontTitle; ForeColor = [System.Drawing.Color]::DarkSlateBlue } $pnlRecipRight; $y += 40

$grpMode = New-Ctl 'GroupBox' @{ Text = "Mode d'envoi"; Location = @(10, $y); Size = @(280, 150) } $pnlRecipRight
$rbPerRecipient = New-Ctl 'RadioButton' @{ Text = '1 mail par destinataire (personnalisé)'; Location = @(12, 24); AutoSize = $true; Checked = $true } $grpMode
New-Ctl 'Label' @{ Text = 'Variables {{PRENOM}}, {{NOM}}... remplacées'; Location = @(30, 44); AutoSize = $true; ForeColor = [System.Drawing.Color]::Gray } $grpMode | Out-Null
$rbBcc = New-Ctl 'RadioButton' @{ Text = 'Un seul mail, destinataires en Cci'; Location = @(12, 70); AutoSize = $true } $grpMode
New-Ctl 'Label' @{ Text = 'Taille des lots Cci :'; Location = @(30, 98); AutoSize = $true } $grpMode | Out-Null
$numBatch = New-Ctl 'NumericUpDown' @{ Location = @(160, 95); Size = @(70, 23); Minimum = 1; Maximum = 500; Value = 50 } $grpMode
New-Ctl 'Label' @{ Text = '(limite Exchange par mail)'; Location = @(30, 122); AutoSize = $true; ForeColor = [System.Drawing.Color]::Gray } $grpMode | Out-Null
$y += 160
New-Ctl 'Label' @{ Text = 'Copie (Cc) sur chaque mail :'; Location = @(10, $y); AutoSize = $true } $pnlRecipRight | Out-Null; $y += 20
$txtCc = New-Ctl 'TextBox' @{ Location = @(10, $y); Size = @(280, 23) } $pnlRecipRight

function Update-RecipientCount {
    $all = Get-RecipientList; $valid = Get-RecipientList -ValidOnly
    $lblCount.Text = "$($valid.Count) adresse(s) valide(s)" + $(if ($all.Count -ne $valid.Count) { "  /  $($all.Count - $valid.Count) invalide(s)" })
}
$txtRecipients.Add_TextChanged({ Update-RecipientCount })

$btnImport.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Importer une liste de destinataires'; $dlg.Filter = 'Fichiers texte (*.txt;*.csv)|*.txt;*.csv|Tous les fichiers|*.*'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $lines = Get-Content -Path $dlg.FileName -Encoding UTF8 | Where-Object { $_.Trim() }
    $before = (Get-RecipientList).Count
    if ($txtRecipients.Text.Trim()) { $txtRecipients.AppendText("`r`n") }
    $txtRecipients.AppendText(($lines -join "`r`n"))
    $after = (Get-RecipientList).Count
    Write-Log "Import : $($lines.Count) ligne(s) lues depuis $($dlg.FileName), $($after - $before) nouvelle(s) adresse(s) unique(s)." 'OK'
})
$btnPaste.Add_Click({
    $clip = [System.Windows.Forms.Clipboard]::GetText()
    if (-not $clip) { Show-Msg 'Le presse-papier est vide.' 'Presse-papier' 'Warning'; return }
    if ($txtRecipients.Text.Trim()) { $txtRecipients.AppendText("`r`n") }
    $txtRecipients.AppendText($clip)
})
$btnClean.Add_Click({
    $all = Get-RecipientList; $valid = Get-RecipientList -ValidOnly
    $txtRecipients.Text = ($valid -join "`r`n")
    $removed = $all.Count - $valid.Count
    Write-Log "Nettoyage : $($valid.Count) adresse(s) conservée(s), $removed invalide(s) retirée(s), doublons supprimés." 'OK'
    Show-Msg "Liste nettoyée.`n`nConservées : $($valid.Count)`nInvalides retirées : $removed" 'Nettoyage'
})
$btnCheck.Add_Click({
    $all = Get-RecipientList
    $bad = @($all | Where-Object { -not (Test-EmailAddress $_) })
    if ($bad.Count -eq 0) { Show-Msg "Toutes les adresses sont valides ($($all.Count))." 'Validation' }
    else { Show-Msg ("{0} adresse(s) invalide(s) :`n`n{1}" -f $bad.Count, (($bad | Select-Object -First 40) -join "`n")) 'Validation' 'Warning' }
})
$btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Fichier texte|*.txt'; $dlg.FileName = "destinataires_$(Get-Date -Format yyyyMMdd).txt"
    if ($dlg.ShowDialog() -eq 'OK') { (Get-RecipientList -ValidOnly) | Set-Content -Path $dlg.FileName -Encoding UTF8; Write-Log "Liste exportée : $($dlg.FileName)" 'OK' }
})
$btnClear.Add_Click({ if (Confirm-Msg 'Vider la liste des destinataires ?') { $txtRecipients.Clear() } })
#endregion

#region ================= ONGLET 2 : SERVEUR SMTP =================
$grpServer = New-Ctl 'GroupBox' @{ Text = 'Serveur Exchange / SMTP'; Location = @(10, 10); Size = @(560, 250) } $tabSmtp
$txtServer = Add-Field $grpServer 'Serveur (FQDN / IP) :' 15 30 380 'mail.domaine.local'
New-Ctl 'Label' @{ Text = 'Port :'; Location = @(15, 63); AutoSize = $true } $grpServer | Out-Null
$numPort = New-Ctl 'NumericUpDown' @{ Location = @(165, 60); Size = @(80, 23); Minimum = 1; Maximum = 65535; Value = 25 } $grpServer
New-Ctl 'Label' @{ Text = '(25 = relais interne, 587 = STARTTLS, 465 = SSL implicite)'; Location = @(255, 63); AutoSize = $true; ForeColor = [System.Drawing.Color]::Gray } $grpServer | Out-Null
$chkSsl = New-Ctl 'CheckBox' @{ Text = 'Activer SSL / TLS (STARTTLS)'; Location = @(165, 92); AutoSize = $true } $grpServer
$chkIgnoreCert = New-Ctl 'CheckBox' @{ Text = 'Ignorer les erreurs de certificat (certificat interne auto-signé)'; Location = @(165, 116); AutoSize = $true } $grpServer
$chkAuth = New-Ctl 'CheckBox' @{ Text = 'Authentification requise (décoché = relais anonyme)'; Location = @(165, 148); AutoSize = $true; Font = $FontBold } $grpServer
$txtUser = Add-Field $grpServer 'Utilisateur :'  15 178 380 '' 
$txtPass = Add-Field $grpServer 'Mot de passe :' 15 208 380 '' -Password
$txtUser.Enabled = $false; $txtPass.Enabled = $false
$chkAuth.Add_CheckedChanged({ $txtUser.Enabled = $chkAuth.Checked; $txtPass.Enabled = $chkAuth.Checked })
New-Ctl 'Label' @{ Text = '(le mot de passe n''est jamais enregistré)'; Location = @(400, 211); AutoSize = $true; ForeColor = [System.Drawing.Color]::Gray } $grpServer | Out-Null

$grpSender = New-Ctl 'GroupBox' @{ Text = 'Expéditeur'; Location = @(590, 10); Size = @(560, 160) } $tabSmtp
$txtFrom     = Add-Field $grpSender 'Adresse expéditeur :' 15 30 380 'support-it@domaine.fr'
$txtFromName = Add-Field $grpSender 'Nom affiché :'        15 60 380 'Support Informatique'
$txtReplyTo  = Add-Field $grpSender 'Répondre à (option) :' 15 90 380 ''
New-Ctl 'Label' @{ Text = 'Priorité :'; Location = @(15, 123); AutoSize = $true } $grpSender | Out-Null
$cmbPriority = New-Ctl 'ComboBox' @{ Location = @(165, 120); Size = @(150, 23); DropDownStyle = 'DropDownList' } $grpSender
$cmbPriority.Items.AddRange(@('Normale', 'Haute', 'Basse')); $cmbPriority.SelectedIndex = 0

$grpOpts = New-Ctl 'GroupBox' @{ Text = "Options d'envoi"; Location = @(590, 180); Size = @(560, 80) } $tabSmtp
New-Ctl 'Label' @{ Text = 'Délai entre envois (ms) :'; Location = @(15, 33); AutoSize = $true } $grpOpts | Out-Null
$numDelay = New-Ctl 'NumericUpDown' @{ Location = @(165, 30); Size = @(90, 23); Minimum = 0; Maximum = 60000; Increment = 100; Value = 200 } $grpOpts
New-Ctl 'Label' @{ Text = 'Timeout (s) :'; Location = @(290, 33); AutoSize = $true } $grpOpts | Out-Null
$numTimeout = New-Ctl 'NumericUpDown' @{ Location = @(380, 30); Size = @(70, 23); Minimum = 5; Maximum = 600; Value = 60 } $grpOpts

$btnTestConn = New-Ctl 'Button' @{ Text = 'Tester la connexion (TCP + bannière)'; Location = @(10, 275); Size = @(270, 36); Font = $FontBold } $tabSmtp
$btnTestMail = New-Ctl 'Button' @{ Text = 'Envoyer un mail de test...';          Location = @(290, 275); Size = @(270, 36) } $tabSmtp
$btnSaveCfg  = New-Ctl 'Button' @{ Text = 'Enregistrer la configuration';        Location = @(590, 275); Size = @(270, 36) } $tabSmtp
New-Ctl 'Label' @{
    Text = "Relais anonyme Exchange : laissez 'Authentification' décoché. L'adresse expéditeur doit être autorisée sur le connecteur de réception (ms-Exch-SMTP-Accept-Any-Sender / Accept-Authoritative-Domain-Sender).`nO365 / SMTP AUTH : port 587 + SSL + authentification (mot de passe d'application si MFA)."
    Location = @(10, 330); Size = @(1140, 60); ForeColor = [System.Drawing.Color]::DimGray
} $tabSmtp | Out-Null
$btnTestConn.Add_Click({ Test-SmtpConnection })
$btnTestMail.Add_Click({ Send-TestMail })
$btnSaveCfg.Add_Click({ Save-Config; Write-Log "Configuration enregistrée dans $($script:ConfigFile)" 'OK'; Show-Msg 'Configuration enregistrée.' })
#endregion

#region ================= ONGLET 3 : EDITEUR =================
$tlEditor = New-Ctl 'TableLayoutPanel' @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 5 } $tabEditor
[void]$tlEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
[void]$tlEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$tlEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$tlEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$tlEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 95)))

# Ligne 0 : objet + fichiers
$pnlSubject = New-Ctl 'Panel' @{ Dock = 'Fill' }
New-Ctl 'Label' @{ Text = 'Objet :'; Location = @(6, 11); AutoSize = $true; Font = $FontBold } $pnlSubject | Out-Null
$txtSubject  = New-Ctl 'TextBox' @{ Location = @(60, 8); Size = @(540, 23); Anchor = 'Top,Left,Right' } $pnlSubject
$btnNew      = New-Ctl 'Button' @{ Text = 'Nouveau';       Size = @(80, 27); Anchor = 'Top,Right' } $pnlSubject
$btnOpen     = New-Ctl 'Button' @{ Text = 'Ouvrir HTML';   Size = @(95, 27); Anchor = 'Top,Right' } $pnlSubject
$btnSave     = New-Ctl 'Button' @{ Text = 'Enregistrer';   Size = @(95, 27); Anchor = 'Top,Right' } $pnlSubject
$btnSource   = New-Ctl 'Button' @{ Text = 'Source HTML';   Size = @(100, 27); Anchor = 'Top,Right' } $pnlSubject
$btnPreview  = New-Ctl 'Button' @{ Text = 'Aperçu';        Size = @(90, 27); Anchor = 'Top,Right'; Font = $FontBold } $pnlSubject
$btnHelp     = New-Ctl 'Button' @{ Text = '?';             Size = @(30, 27); Anchor = 'Top,Right' } $pnlSubject
$pnlSubject.Add_Resize({
    $x = $pnlSubject.Width - 8
    foreach ($b in @($btnHelp, $btnPreview, $btnSource, $btnSave, $btnOpen, $btnNew)) { $x -= $b.Width + 4; $b.Location = New-Object System.Drawing.Point($x, 6) }
    $txtSubject.Width = [Math]::Max(150, $x - 68)
})
$tlEditor.Controls.Add($pnlSubject, 0, 0)

# Lignes 1-2 : barres d'outils
$tsFormat = New-Object System.Windows.Forms.ToolStrip; $tsFormat.Dock = 'Fill'; $tsFormat.GripStyle = 'Hidden'; $tsFormat.AutoSize = $true
$tsInsert = New-Object System.Windows.Forms.ToolStrip; $tsInsert.Dock = 'Fill'; $tsInsert.GripStyle = 'Hidden'; $tsInsert.AutoSize = $true
$tlEditor.Controls.Add($tsFormat, 0, 1)
$tlEditor.Controls.Add($tsInsert, 0, 2)

# Ligne 3 : éditeur WYSIWYG + source
$pnlEdit   = New-Ctl 'Panel' @{ Dock = 'Fill'; BorderStyle = 'FixedSingle' }
$wbEditor  = New-Ctl 'WebBrowser' @{ Dock = 'Fill'; ScriptErrorsSuppressed = $true; IsWebBrowserContextMenuEnabled = $true; WebBrowserShortcutsEnabled = $true; AllowWebBrowserDrop = $false } $pnlEdit
$txtSource = New-Ctl 'TextBox' @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Both'; AcceptsReturn = $true; AcceptsTab = $true; WordWrap = $false; Font = $FontMono; Visible = $false } $pnlEdit
$tlEditor.Controls.Add($pnlEdit, 0, 3)

# Ligne 4 : pièces jointes
$grpAttach = New-Ctl 'GroupBox' @{ Text = 'Pièces jointes (identiques pour tous les destinataires)'; Dock = 'Fill' }
$lstAttach = New-Ctl 'ListBox' @{ Location = @(10, 20); Size = @(700, 60); Anchor = 'Top,Left,Right,Bottom'; HorizontalScrollbar = $true } $grpAttach
$btnAttAdd = New-Ctl 'Button' @{ Text = 'Ajouter...'; Size = @(110, 27); Anchor = 'Top,Right' } $grpAttach
$btnAttDel = New-Ctl 'Button' @{ Text = 'Retirer';    Size = @(110, 27); Anchor = 'Top,Right' } $grpAttach
$grpAttach.Add_Resize({
    $btnAttAdd.Location = New-Object System.Drawing.Point(($grpAttach.Width - 125), 22)
    $btnAttDel.Location = New-Object System.Drawing.Point(($grpAttach.Width - 125), 52)
    $lstAttach.Width = $grpAttach.Width - 145
})
$tlEditor.Controls.Add($grpAttach, 0, 4)

Build-FormatToolbar $tsFormat $wbEditor
Build-InsertToolbar $tsInsert $wbEditor -WithTemplates

$btnNew.Add_Click({
    if (Confirm-Msg "Effacer le contenu de l'éditeur et l'objet ?") {
        if ($script:SourceMode) { $txtSource.Text = '' }
        Set-EditorHtml $wbEditor '<p><br></p>'; $txtSubject.Text = ''
        Write-Log 'Nouveau mail.' 'INFO'
    }
})
$btnOpen.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Fichiers HTML|*.html;*.htm|Tous les fichiers|*.*'; $dlg.Title = 'Ouvrir un brouillon HTML'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $raw = Get-Content -Path $dlg.FileName -Raw -Encoding UTF8
    $inner = if ($raw -match '(?is)<body[^>]*>(.*?)</body>') { $Matches[1] } else { $raw }
    if ($raw -match '(?is)<title>(.*?)</title>') { $t = [System.Net.WebUtility]::HtmlDecode($Matches[1]).Trim(); if ($t -and -not $txtSubject.Text) { $txtSubject.Text = $t } }
    if ($script:SourceMode) { $txtSource.Text = $inner } else { Set-EditorHtml $wbEditor $inner }
    Write-Log "Brouillon chargé : $($dlg.FileName)" 'OK'
})
$btnSave.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Fichier HTML|*.html'; $dlg.FileName = "mail_$(Get-Date -Format yyyyMMdd_HHmm).html"; $dlg.Title = 'Enregistrer le brouillon'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    [System.IO.File]::WriteAllText($dlg.FileName, (Get-FinalHtml), (New-Object System.Text.UTF8Encoding($false)))
    Write-Log "Brouillon enregistré : $($dlg.FileName)" 'OK'
})
$btnSource.Add_Click({
    if (-not $script:SourceMode) {
        $txtSource.Text = $wbEditor.Document.Body.InnerHtml
        $wbEditor.Visible = $false; $txtSource.Visible = $true
        $script:SourceMode = $true; $btnSource.Text = 'Mode visuel'; $tsFormat.Enabled = $false; $tsInsert.Enabled = $false
    } else {
        Set-EditorHtml $wbEditor $txtSource.Text
        $txtSource.Visible = $false; $wbEditor.Visible = $true
        $script:SourceMode = $false; $btnSource.Text = 'Source HTML'; $tsFormat.Enabled = $true; $tsInsert.Enabled = $true
    }
})
$btnPreview.Add_Click({
    $sample = Get-RecipientList -ValidOnly
    $email  = if ($sample.Count -gt 0) { $sample[0] } else { $txtFrom.Text }
    $html   = Expand-Variables (Get-FinalHtml -IncludeSignature:($chkAutoSig.Checked)) (Get-VarsForEmail $email)
    [System.IO.File]::WriteAllText($script:PreviewFile, $html, (New-Object System.Text.UTF8Encoding($true)))
    $f = New-Ctl 'Form' @{ Text = "Aperçu - $(Expand-Variables $txtSubject.Text (Get-VarsForEmail $email))  [variables remplacées pour $email]"; Size = @(950, 750); StartPosition = 'CenterParent' }
    $pw = New-Ctl 'WebBrowser' @{ Dock = 'Fill'; ScriptErrorsSuppressed = $true } $f
    $pw.Navigate($script:PreviewFile)
    [void]$f.ShowDialog($form)
})
$btnHelp.Add_Click({
    Show-Msg @"
ÉDITEUR DE MAIL - AIDE RAPIDE

• Mise en forme : sélectionnez du texte puis utilisez la barre d'outils.
  Raccourcis : Ctrl+B gras, Ctrl+I italique, Ctrl+U souligné, Ctrl+Z / Ctrl+Y.
• Collage : Ctrl+V depuis Word / Outlook conserve la mise en forme.
• Images : 'Image (fichier)' ou 'Image (chemin)' insère une image locale qui sera
  EMBARQUÉE dans le mail (Content-ID) -> visible sans pièce jointe.
  Le fichier doit toujours exister au moment de l'envoi.
• Modèles : 'Procédure', 'Information', 'Maintenance', 'Sécurité' (remplace le contenu).
• Blocs : encadrés Info / Attention / Succès / Erreur / Code / Étape / Bouton.
• Variables : {{PRENOM}} {{NOM}} {{NOMCOMPLET}} {{EMAIL}} {{DATE}} {{HEURE}}
  (déduites de prenom.nom@domaine - actives en mode '1 mail par destinataire').
• Source HTML : bascule vers le code pour les retouches fines.
• Aperçu : rendu final avec signature et variables remplacées.
• Enregistrer / Ouvrir : brouillons HTML réutilisables.
"@ 'Aide'
})
$btnAttAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Multiselect = $true; $dlg.Title = 'Ajouter des pièces jointes'
    if ($dlg.ShowDialog() -eq 'OK') { foreach ($f in $dlg.FileNames) { if (-not $lstAttach.Items.Contains($f)) { [void]$lstAttach.Items.Add($f) } } }
})
$btnAttDel.Add_Click({ if ($lstAttach.SelectedIndex -ge 0) { $lstAttach.Items.RemoveAt($lstAttach.SelectedIndex) } })
#endregion

#region ================= ONGLET 4 : SIGNATURE =================
$tlSig = New-Ctl 'TableLayoutPanel' @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 4 } $tabSig
[void]$tlSig.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
[void]$tlSig.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$tlSig.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$tlSig.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))

$pnlSigTop = New-Ctl 'Panel' @{ Dock = 'Fill' }
New-Ctl 'Label' @{ Text = "Signature (texte + logo / bannière). Enregistrée dans $($script:SignatureFile)"; Location = @(6, 12); AutoSize = $true } $pnlSigTop | Out-Null
$btnSigSave   = New-Ctl 'Button' @{ Text = 'Enregistrer la signature'; Size = @(170, 27); Anchor = 'Top,Right'; Font = $FontBold } $pnlSigTop
$btnSigReload = New-Ctl 'Button' @{ Text = 'Recharger';               Size = @(90, 27);  Anchor = 'Top,Right' } $pnlSigTop
$btnSigImport = New-Ctl 'Button' @{ Text = 'Importer HTML';           Size = @(110, 27); Anchor = 'Top,Right' } $pnlSigTop
$btnSigModel  = New-Ctl 'Button' @{ Text = 'Modèle signature';        Size = @(130, 27); Anchor = 'Top,Right' } $pnlSigTop
$pnlSigTop.Add_Resize({
    $x = $pnlSigTop.Width - 8
    foreach ($b in @($btnSigSave, $btnSigReload, $btnSigImport, $btnSigModel)) { $x -= $b.Width + 4; $b.Location = New-Object System.Drawing.Point($x, 8) }
})
$tlSig.Controls.Add($pnlSigTop, 0, 0)
$tsSigFormat = New-Object System.Windows.Forms.ToolStrip; $tsSigFormat.Dock = 'Fill'; $tsSigFormat.GripStyle = 'Hidden'; $tsSigFormat.AutoSize = $true
$tsSigInsert = New-Object System.Windows.Forms.ToolStrip; $tsSigInsert.Dock = 'Fill'; $tsSigInsert.GripStyle = 'Hidden'; $tsSigInsert.AutoSize = $true
$tlSig.Controls.Add($tsSigFormat, 0, 1)
$tlSig.Controls.Add($tsSigInsert, 0, 2)
$pnlSigEdit = New-Ctl 'Panel' @{ Dock = 'Fill'; BorderStyle = 'FixedSingle' }
$wbSig = New-Ctl 'WebBrowser' @{ Dock = 'Fill'; ScriptErrorsSuppressed = $true; IsWebBrowserContextMenuEnabled = $true; AllowWebBrowserDrop = $false } $pnlSigEdit
$tlSig.Controls.Add($pnlSigEdit, 0, 3)
Build-FormatToolbar $tsSigFormat $wbSig
Build-InsertToolbar $tsSigInsert $wbSig

$script:SignatureModel = @'
<table style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#333;border-collapse:collapse;">
<tr>
<td style="padding-right:14px;border-right:3px solid #1F4E79;vertical-align:top;"><i style="color:#888;">[Logo : Insérer &gt; Image (fichier)]</i></td>
<td style="padding-left:14px;vertical-align:top;">
<b style="font-size:12pt;color:#1F4E79;">Prénom NOM</b><br>
Technicien support informatique N2<br>
<span style="color:#666;">Société &ndash; Service Informatique</span><br>
&#9742; 01 23 45 67 89 &nbsp;|&nbsp; &#9993; <a href="mailto:support@domaine.fr" style="color:#1F4E79;">support@domaine.fr</a><br>
<a href="https://intranet.domaine.fr" style="color:#1F4E79;">Portail intranet</a>
</td></tr></table>
'@

$btnSigSave.Add_Click({
    $h = Get-EditorHtml $wbSig
    [System.IO.File]::WriteAllText($script:SignatureFile, $h, (New-Object System.Text.UTF8Encoding($false)))
    Write-Log "Signature enregistrée : $($script:SignatureFile)" 'OK'
    Show-Msg 'Signature enregistrée.' 'Signature'
})
$btnSigReload.Add_Click({
    if (Test-Path $script:SignatureFile) { Set-EditorHtml $wbSig (Get-Content -Path $script:SignatureFile -Raw -Encoding UTF8); Write-Log 'Signature rechargée.' 'INFO' }
    else { Show-Msg 'Aucune signature enregistrée.' 'Signature' 'Warning' }
})
$btnSigImport.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter = 'Fichiers HTML|*.html;*.htm'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $raw = Get-Content -Path $dlg.FileName -Raw -Encoding UTF8
    $inner = if ($raw -match '(?is)<body[^>]*>(.*?)</body>') { $Matches[1] } else { $raw }
    Set-EditorHtml $wbSig $inner
})
$btnSigModel.Add_Click({
    if ((Get-EditorHtml $wbSig) -and -not (Confirm-Msg 'Remplacer la signature actuelle par le modèle ?')) { return }
    Set-EditorHtml $wbSig $script:SignatureModel
})
#endregion

#region ================= ONGLET 5 : ENVOI & JOURNAL =================
$pnlSendTop = New-Ctl 'Panel' @{ Dock = 'Top'; Height = 150 } $tabSend
$lblSummary = New-Ctl 'Label' @{ Location = @(6, 6); Size = @(620, 95); Font = $FontUI; BorderStyle = 'FixedSingle'; Padding = (New-Object System.Windows.Forms.Padding(6)); BackColor = [System.Drawing.Color]::WhiteSmoke } $pnlSendTop
$chkAutoSig = New-Ctl 'CheckBox' @{ Text = 'Ajouter automatiquement la signature en fin de mail'; Location = @(640, 8); AutoSize = $true; Checked = $true } $pnlSendTop
$chkDryRun  = New-Ctl 'CheckBox' @{ Text = 'SIMULATION (dry-run) : tout préparer sans envoyer'; Location = @(640, 32); AutoSize = $true; ForeColor = [System.Drawing.Color]::DarkOrange; Font = $FontBold } $pnlSendTop
$btnTest    = New-Ctl 'Button' @{ Text = 'Mail de test...';      Location = @(640, 62); Size = @(150, 34) } $pnlSendTop
$btnSend    = New-Ctl 'Button' @{ Text = 'LANCER LA CAMPAGNE';   Location = @(800, 62); Size = @(220, 34); Font = $FontBig; BackColor = [System.Drawing.Color]::FromArgb(31, 78, 121); ForeColor = [System.Drawing.Color]::White; FlatStyle = 'Flat' } $pnlSendTop
$btnCancel  = New-Ctl 'Button' @{ Text = 'Annuler';              Location = @(1030, 62); Size = @(100, 34); Enabled = $false } $pnlSendTop
$pbSend     = New-Ctl 'ProgressBar' @{ Location = @(6, 110); Size = @(1124, 22); Anchor = 'Top,Left,Right' } $pnlSendTop
$btnExportLog = New-Ctl 'Button' @{ Text = 'Exporter les résultats (CSV)'; Location = @(6, 0); Size = @(200, 27) }
$btnClearLog  = New-Ctl 'Button' @{ Text = 'Effacer le journal';           Location = @(212, 0); Size = @(150, 27) }
$btnOpenLogs  = New-Ctl 'Button' @{ Text = 'Ouvrir le dossier des logs';   Location = @(368, 0); Size = @(190, 27) }
$pnlLogBtns = New-Ctl 'Panel' @{ Dock = 'Bottom'; Height = 32 }
$pnlLogBtns.Controls.AddRange(@($btnExportLog, $btnClearLog, $btnOpenLogs))
$rtbLog = New-Ctl 'RichTextBox' @{ Dock = 'Fill'; ReadOnly = $true; Font = $FontMono; BackColor = [System.Drawing.Color]::White; DetectUrls = $false }
$tabSend.Controls.Add($rtbLog)
$tabSend.Controls.Add($pnlLogBtns)
$rtbLog.BringToFront()

function Update-Summary {
    $valid = Get-RecipientList -ValidOnly
    $mode  = if ($rbPerRecipient.Checked) { "1 mail par destinataire (personnalisé)" } else { "Cci par lots de $([int]$numBatch.Value)" }
    $lblSummary.Text = "Destinataires valides : $($valid.Count)`n" +
                       "Serveur : $($txtServer.Text):$([int]$numPort.Value)  |  Auth : $(if ($chkAuth.Checked) {'oui'} else {'non (anonyme)'})  |  SSL : $(if ($chkSsl.Checked) {'oui'} else {'non'})`n" +
                       "Expéditeur : $($txtFromName.Text) <$($txtFrom.Text)>`n" +
                       "Objet : $($txtSubject.Text)`n" +
                       "Mode : $mode  |  Délai : $([int]$numDelay.Value) ms  |  PJ : $($lstAttach.Items.Count)"
}
$tabs.Add_SelectedIndexChanged({ if ($tabs.SelectedTab -eq $tabSend) { Update-Summary } })
$btnSend.Add_Click({ Start-Campaign })
$btnCancel.Add_Click({ $script:CancelSend = $true; Write-Log 'Annulation demandée...' 'WARN' })
$btnTest.Add_Click({ Send-TestMail })
$btnClearLog.Add_Click({ $rtbLog.Clear() })
$btnOpenLogs.Add_Click({ Start-Process explorer.exe $script:LogDir })
$btnExportLog.Add_Click({
    if ($script:Results.Count -eq 0) { Show-Msg 'Aucun résultat à exporter (lancez une campagne).' 'Export' 'Warning'; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = "resultats_campagne_$(Get-Date -Format yyyyMMdd_HHmm).csv"
    if ($dlg.ShowDialog() -eq 'OK') {
        $script:Results | Export-Csv -Path $dlg.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Log "Résultats exportés : $($dlg.FileName)" 'OK'
    }
})
#endregion

#region ================= DEMARRAGE / FERMETURE =================
$form.Add_Shown({
    Write-Log "$($script:AppName) v$($script:AppVersion) démarré. Journal : $($script:LogFile)" 'INFO'
    Import-Config
    # Les WebBrowser doivent être visibles pour créer leur handle : on parcourt les onglets une fois
    $tabs.SelectedTab = $tabEditor
    [System.Windows.Forms.Application]::DoEvents()
    Initialize-Editor $wbEditor $script:EditorFile '<p><br></p>'
    $tabs.SelectedTab = $tabSig
    [System.Windows.Forms.Application]::DoEvents()
    $sig = if (Test-Path $script:SignatureFile) { Get-Content -Path $script:SignatureFile -Raw -Encoding UTF8 } else { '' }
    Initialize-Editor $wbSig $script:SigEditorFile $sig
    $script:EditorsReady = $true
    $tabs.SelectedTab = $tabRecip
    Update-RecipientCount
    $pnlSubject.PerformLayout(); $pnlSubject.Invalidate()
    $lblStatus.Text = 'Prêt.'
})
$form.Add_FormClosing({
    Save-Config
    try {
        if ($script:EditorsReady) {
            $h = Get-EditorHtml $wbSig
            if ($h) { [System.IO.File]::WriteAllText($script:SignatureFile, $h, (New-Object System.Text.UTF8Encoding($false))) }
        }
    } catch { }
})

[void]$form.ShowDialog()
#endregion
