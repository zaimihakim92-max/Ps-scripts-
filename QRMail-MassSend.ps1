<#
================================================================================
                 ENVOI DE MASSE - QR CODES TOKENS PAR EMAIL
                  (avec prévisualisation avant chaque envoi)
================================================================================

  OBJECTIF
  --------
  À partir du CSV enrichi produit par le générateur de QR codes
  (colonnes : Nom ; Email ; Token ; CheminQR), ce script :
      - identifie l'adresse email de chaque ligne
      - récupère l'image QR correspondante (colonne CheminQR)
      - l'injecte dans le corps du mail (image INLINE, pas une simple PJ)
      - affiche une PRÉVISUALISATION fidèle de chaque mail
      - n'envoie qu'après validation (mail par mail, ou mode auto)

  SCHÉMA DE FONCTIONNEMENT
  ------------------------

      ┌──────────────────────────┐
      │  CSV enrichi  (;)        │   Nom ; Email ; ... ; CheminQR
      └────────────┬─────────────┘
                   │  [GUI] chargement + mapping colonnes Email / Nom / CheminQR
                   ▼
      ┌──────────────────────────┐
      │  Paramètres SMTP         │   serveur / port / SSL / expéditeur
      │  (laissés VIERGES,       │   + identifiants optionnels
      │   à renseigner)          │
      └────────────┬─────────────┘
                   ▼
      ┌──────────────────────────┐
      │  File d'attente          │   1 destinataire = 1 mail
      └────────────┬─────────────┘
                   ▼
      ┌──────────────────────────┐     [Envoyer]      ┌──────────────────┐
      │  PRÉVISUALISATION        │───────────────────▶│  SMTP (inline    │
      │  destinataire / objet /  │     [Passer]       │  cid: QR code)   │
      │  corps HTML + QR affiché │──────┐             └──────────────────┘
      └────────────┬─────────────┘      │
                   │ [Tout envoyer]     ▼
                   ▼              destinataire suivant
      ┌──────────────────────────┐
      │  Journal + bilan final   │   OK / échec / passés
      └──────────────────────────┘

  MODES D'ENVOI
  -------------
  - "Envoyer"            : envoie le mail affiché, passe au suivant
  - "Passer"             : ignore le destinataire affiché, passe au suivant
  - "Tout envoyer"       : envoie tous les mails RESTANTS sans nouvelle
                           confirmation (à utiliser après avoir validé
                           quelques prévisualisations)

  POINTS DE SÉCURITÉ
  ------------------
  - Le QR est injecté en image inline (Content-ID) : le secret ne transite
    que vers le destinataire concerné, une seule image par mail.
  - Le mot de passe SMTP éventuel est saisi en champ masqué et converti en
    PSCredential ; il n'est jamais écrit sur disque ni dans le journal.
  - Anti-erreur d'aiguillage : contrôle de cohérence avant chaque envoi
    (email non vide + fichier QR existant), sinon mise en échec et passage
    au suivant.
  - Pause configurable entre deux envois (mode "Tout envoyer") pour ne pas
    déclencher les seuils anti-spam / throttling du serveur.

  BOUTON "ENVOYER UN TEST"
  ------------------------
  Envoie le mail du destinataire sélectionné (le 1er par défaut) vers une
  adresse de votre choix (la vôtre), objet préfixé [TEST] et bandeau optionnel.
  Ne compte ni dans le bilan ni dans la preuve d'envoi.

  ENCODAGE DU FICHIER
  -------------------
  Enregistrer ce script en "UTF-8 avec BOM". Sans BOM, Windows PowerShell 5.1
  le lit en ANSI : certains caractères UTF-8 deviennent des guillemets, le
  script ne se compile plus et PowerShell affiche en cascade des erreurs
  "Unable to find type [System.Windows.Forms...]".
  Les symboles de l'interface sont désormais générés par code ([char]0x...)
  pour limiter ce risque.

  PRÉREQUIS
  ---------
  - Windows PowerShell 5.1 ou PowerShell 7+ (Windows)
  - Un serveur SMTP joignable et un compte autorisé à émettre
    (relais interne, connecteur dédié, etc.)
  - Le CSV enrichi ET les images QR encore présents sur le disque
    (ne pas purger avant la fin de la campagne !)
================================================================================
#>

# ==============================================================================
# CONFIGURATION - MODÈLE DU MAIL (personnalisable)
# ==============================================================================

# Chemin de l'IMAGE DE SIGNATURE (logo, bannière de service, signature scannée...)
# - Renseignez un chemin complet, ex : "C:\Outils\signature.png"
# - Laissez vide ("") pour ne pas inclure de signature.
# L'image est injectée en INLINE (cid:) comme le QR : elle s'affiche dans le
# corps du mail sans apparaître comme pièce jointe à ouvrir.
$CheminSignature = ""

# Largeur d'affichage de la signature dans le mail (en pixels)
$LargeurSignaturePx = 300

# Objet du mail. {NOM} sera remplacé par le nom du destinataire.
$ModeleObjet = "Votre code d'activation personnel"

# Corps HTML. Variables disponibles :
#   {NOM}        -> nom du destinataire
#   {QR}         -> emplacement où l'image QR est injectée (NE PAS SUPPRIMER)
#   {SIGNATURE}  -> emplacement de l'image de signature (retirée automatiquement
#                   si $CheminSignature est vide ou si le fichier est introuvable)
$ModeleCorps = @"
<html>
<body style="font-family: Segoe UI, Arial, sans-serif; color:#222; max-width:600px;">
  <p>Bonjour {NOM},</p>
  <p>Veuillez trouver ci-dessous votre code QR personnel d'activation.</p>
  <p>Scannez-le avec l'application prévue à cet effet :</p>
  <p style="text-align:center; margin:25px 0;">{QR}</p>
  <p>Ce code est strictement personnel : ne le transférez à personne.</p>
  <p>En cas de difficulté, contactez le support informatique.</p>
  <p>Cordialement,<br/>Le support informatique</p>
  <p style="margin-top:20px;">{SIGNATURE}</p>
</body>
</html>
"@

# Pause entre deux envois en mode "Tout envoyer" (en millisecondes)
$PauseEntreEnvoisMs = 1500

# ==============================================================================
# CHARGEMENT DES ASSEMBLIES GUI
# ==============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Symboles générés par code : aucun caractère spécial littéral dans les chaînes,
# le script reste compilable même si le fichier perd son BOM UTF-8.
$SymOK   = [string][char]0x2713   # coche
$SymKO   = [string][char]0x2717   # croix
$SymWarn = [string][char]0x26A0   # avertissement

# ==============================================================================
# 1. FONCTIONS
# ==============================================================================

function Detect-CsvDelimiter {
    <#
        Détecte le délimiteur d'un fichier CSV automatiquement
        Teste en ordre : ; (France) → , (US) → TAB → espace
        Retourne le délimiteur détecté
    #>
    param([string]$FilePath)
    
    try {
        $firstLine = Get-Content $FilePath -TotalCount 1 -Encoding UTF8
        
        # Tester les délimiteurs courants (ordre de priorité)
        if ($firstLine -like "*;*") { return ";" }        # Point-virgule (France)
        if ($firstLine -like "*,*") { return "," }        # Virgule (US)
        if ($firstLine -like "*`t*") { return "`t" }      # TAB
        if ($firstLine -like "* *") { return " " }        # Espace
        
        return ";"  # Par défaut
    }
    catch {
        Write-Log "$SymWarn Erreur détection délimiteur : $($_.Exception.Message) -> par défaut ';'"
        return ";"
    }
}

function Send-QRMail {
    <#
        Envoie UN mail avec le QR code en image inline (Content-ID).
        Utilise System.Net.Mail.SmtpClient (et non Send-MailMessage) car
        c'est le seul moyen propre d'embarquer une image inline via cid:.

        Retourne $true si l'envoi a réussi, sinon lève une exception
        (interceptée par l'appelant pour journalisation).
    #>
    param(
        [string]$Serveur,
        [int]$Port,
        [bool]$UtiliserSSL,
        [pscredential]$Credential,   # $null = authentification Windows intégrée (relais AD)
        [string]$Expediteur,
        [string]$Destinataire,
        [string]$Objet,
        [string]$CorpsHtml,          # doit contenir le marqueur {QR}
        [string]$CheminQR
    )

    # --- Construction du corps : {QR} devient <img cid:...>, {SIGNATURE} idem ---
    # IMPORTANT : Content-ID UNIQUE par mail (GUID). Avec un CID fixe, certains
    # clients (Outlook en vue conversation) mettent l'image en cache et
    # réaffichent le PREMIER QR reçu dans tous les mails suivants.
    $cidQR  = "qr-"  + [guid]::NewGuid().ToString("N")
    $cidSig = "sig-" + [guid]::NewGuid().ToString("N")
    $html   = $CorpsHtml -replace '\{QR\}', "<img src=""cid:$cidQR"" alt=""QR code"" style=""width:280px;height:280px;"" />"

    # Signature : injectée seulement si le fichier existe, sinon le marqueur est retiré
    $sigOk  = (-not [string]::IsNullOrWhiteSpace($script:CheminSignature)) -and (Test-Path $script:CheminSignature)
    if ($sigOk) {
        $html = $html -replace '\{SIGNATURE\}', "<img src=""cid:$cidSig"" alt=""Signature"" style=""width:$($script:LargeurSignaturePx)px;"" />"
    } else {
        $html = $html -replace '\{SIGNATURE\}', ""
    }

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From       = New-Object System.Net.Mail.MailAddress($Expediteur)
    $mail.Subject    = $Objet
    $mail.IsBodyHtml = $true
    $mail.To.Add($Destinataire) | Out-Null

    # --- Vue HTML + ressources liées (QR + signature, référencées par Content-ID) ---
    $vue = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString(
        $html, $null, [System.Net.Mime.MediaTypeNames+Text]::Html)

    $ressource = New-Object System.Net.Mail.LinkedResource($CheminQR, "image/png")
    $ressource.ContentId = $cidQR
    $ressource.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
    $vue.LinkedResources.Add($ressource)

    $ressourceSig = $null
    if ($sigOk) {
        # Type MIME déduit de l'extension (png/jpg/gif pris en charge)
        $mime = switch ([IO.Path]::GetExtension($script:CheminSignature).ToLower()) {
            '.jpg'  { 'image/jpeg' }
            '.jpeg' { 'image/jpeg' }
            '.gif'  { 'image/gif'  }
            default { 'image/png'  }
        }
        $ressourceSig = New-Object System.Net.Mail.LinkedResource($script:CheminSignature, $mime)
        $ressourceSig.ContentId = $cidSig
        $ressourceSig.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
        $vue.LinkedResources.Add($ressourceSig)
    }

    $mail.AlternateViews.Add($vue)

    # --- Client SMTP ---
    $smtp = New-Object System.Net.Mail.SmtpClient($Serveur, $Port)
    $smtp.EnableSsl = $UtiliserSSL
    if ($Credential) {
        # Compte explicite (ex: compte de service)
        $smtp.Credentials = $Credential.GetNetworkCredential()
    } else {
        # Authentification Windows intégrée : utilise le compte AD de la session
        # (cas typique d'un relais SMTP interne autorisant le compte machine/utilisateur)
        $smtp.UseDefaultCredentials = $true
    }

    try {
        $smtp.Send($mail)
        return $true
    }
    finally {
        # Libération systématique des ressources (fichiers verrouillés sinon)
        $ressource.Dispose()
        if ($ressourceSig) { $ressourceSig.Dispose() }
        $vue.Dispose()
        $mail.Dispose()
        $smtp.Dispose()
    }
}

function Find-Column {
    <# Auto-détection d'une colonne par liste de motifs regex (ordre = priorité). #>
    param([string[]]$Colonnes, [string[]]$MotsCles)
    foreach ($mc in $MotsCles) {
        $match = $Colonnes | Where-Object { $_ -match $mc } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

# ==============================================================================
# 2. CONSTRUCTION DE L'INTERFACE
# ==============================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Envoi de masse - QR codes par email"
$form.Size            = New-Object System.Drawing.Size(900, 720)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ------------------------------------------------------------------ CSV enrichi
$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = "1. CSV enrichi (avec colonne CheminQR) :"
$lblCsv.Location = '15,12'; $lblCsv.AutoSize = $true
$form.Controls.Add($lblCsv)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = '15,33'; $txtCsv.Size = '650,25'; $txtCsv.ReadOnly = $true
$form.Controls.Add($txtCsv)

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "Parcourir..."
$btnCsv.Location = '675,31'; $btnCsv.Size = '95,27'
$form.Controls.Add($btnCsv)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Aucun fichier chargé."
$lblCount.Location = '780,36'; $lblCount.AutoSize = $true
$lblCount.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblCount)

# ------------------------------------------------------------- Mapping colonnes
$grpMap = New-Object System.Windows.Forms.GroupBox
$grpMap.Text = "2. Mapping des colonnes"
$grpMap.Location = '15,65'; $grpMap.Size = '420,85'
$form.Controls.Add($grpMap)

$combos = @{}
$mapDefs = @(
    @{ Cle = 'Email';    Libelle = 'Email :';      X = 10  },
    @{ Cle = 'Nom';      Libelle = 'Nom :';        X = 145 },
    @{ Cle = 'CheminQR'; Libelle = 'Chemin QR :';  X = 280 }
)
foreach ($d in $mapDefs) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $d.Libelle; $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($d.X, 22)
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'; $cmb.Size = '125,25'
    $cmb.Location = New-Object System.Drawing.Point($d.X, 44)
    $grpMap.Controls.Add($lbl); $grpMap.Controls.Add($cmb)
    $combos[$d.Cle] = $cmb
}

# -------------------------------------------------------------- Paramètres SMTP
$grpSmtp = New-Object System.Windows.Forms.GroupBox
$grpSmtp.Text = "3. Paramètres SMTP (à renseigner)"
$grpSmtp.Location = '445,65'; $grpSmtp.Size = '435,150'
$form.Controls.Add($grpSmtp)

# Serveur SMTP -- LAISSÉ VIERGE volontairement
$lblSrv = New-Object System.Windows.Forms.Label
$lblSrv.Text = "Serveur :"; $lblSrv.Location = '10,25'; $lblSrv.AutoSize = $true
$grpSmtp.Controls.Add($lblSrv)
$txtSrv = New-Object System.Windows.Forms.TextBox
$txtSrv.Location = '70,22'; $txtSrv.Size = '200,25'
$txtSrv.Text = ""                                   # <-- serveur SMTP : VIERGE
$grpSmtp.Controls.Add($txtSrv)

$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Text = "Port :"; $lblPort.Location = '280,25'; $lblPort.AutoSize = $true
$grpSmtp.Controls.Add($lblPort)
$txtPort = New-Object System.Windows.Forms.TextBox
$txtPort.Location = '320,22'; $txtPort.Size = '50,25'
$txtPort.Text = "25"                                # 25 = relais interne classique (587 si soumission authentifiée)
$grpSmtp.Controls.Add($txtPort)

$chkSsl = New-Object System.Windows.Forms.CheckBox
$chkSsl.Text = "SSL/TLS"; $chkSsl.Location = '375,23'; $chkSsl.AutoSize = $true
$grpSmtp.Controls.Add($chkSsl)

# Expéditeur -- LAISSÉ VIERGE volontairement
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "Expéditeur :"; $lblFrom.Location = '10,55'; $lblFrom.AutoSize = $true
$grpSmtp.Controls.Add($lblFrom)
$txtFrom = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = '85,52'; $txtFrom.Size = '285,25'
$txtFrom.Text = ""                                  # <-- adresse expéditeur : VIERGE
$grpSmtp.Controls.Add($txtFrom)

# Authentification : Windows intégrée (défaut, adaptée à un relais AD interne)
# ou compte explicite (utilisateur + mot de passe masqué)
$chkAuth = New-Object System.Windows.Forms.CheckBox
$chkAuth.Text = "Compte explicite (sinon : authentification Windows intégrée / AD)"
$chkAuth.Location = '10,84'; $chkAuth.AutoSize = $true
$grpSmtp.Controls.Add($chkAuth)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = '10,108'; $txtUser.Size = '175,25'; $txtUser.Enabled = $false
$grpSmtp.Controls.Add($txtUser)
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = '195,108'; $txtPass.Size = '175,25'; $txtPass.Enabled = $false
$txtPass.UseSystemPasswordChar = $true              # saisie masquée
$grpSmtp.Controls.Add($txtPass)
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "(utilisateur / mot de passe)"; $lblUser.Location = '10,136'; $lblUser.AutoSize = $true
$lblUser.ForeColor = [System.Drawing.Color]::Gray
$grpSmtp.Controls.Add($lblUser)

$chkAuth.Add_CheckedChanged({
    $txtUser.Enabled = $chkAuth.Checked
    $txtPass.Enabled = $chkAuth.Checked
})

# ------------------------------------------------------------- Prévisualisation (refactorisée)
$grpPrev = New-Object System.Windows.Forms.GroupBox
$grpPrev.Text = "4. Prévisualisation de tous les mails - sélectionne un destinataire dans la liste"
$grpPrev.Location = '15,222'; $grpPrev.Size = '865,330'
$form.Controls.Add($grpPrev)

# Bouton d'édition du modèle de mail (en haut à droite du groupe)
$btnModele = New-Object System.Windows.Forms.Button
$btnModele.Text = "Modifier le modèle..."
$btnModele.Location = '760,0'; $btnModele.Size = '100,22'
$btnModele.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpPrev.Controls.Add($btnModele)

# --- CheckedListBox des destinataires (gauche) - REDIMENSIONNABLE ---
$listbox = New-Object System.Windows.Forms.CheckedListBox
$listbox.Location = '10,22'; $listbox.Size = '250,295'
$listbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$listbox.IntegralHeight = $false
$listbox.CheckOnClick = $true   # Cocher en cliquant directement sur la case
$grpPrev.Controls.Add($listbox)

$listbox.Add_ItemCheck({
    param($s, $e)
    # ItemCheck est levé AVANT le changement d'état : on corrige avec NewValue
    $nbCoches = $listbox.CheckedItems.Count
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked -and
        $e.CurrentValue -ne [System.Windows.Forms.CheckState]::Checked) { $nbCoches++ }
    elseif ($e.NewValue -ne [System.Windows.Forms.CheckState]::Checked -and
            $e.CurrentValue -eq [System.Windows.Forms.CheckState]::Checked) { $nbCoches-- }
    if ($nbCoches -eq 0) {
        $btnSendAll.Text = "Envoyer tous les mails"
    } else {
        $btnSendAll.Text = "Envoyer $nbCoches mail(s) cochés"
    }
})

$script:suspendSelection = $false
$listbox.Add_SelectedIndexChanged({
    if ($script:suspendSelection) { return }
    if ($listbox.SelectedIndex -ge 0) {
        $script:index = $listbox.SelectedIndex
        [System.Windows.Forms.Application]::DoEvents()
        Show-Apercu
    }
})

$script:splitterX = 260   # bord droit de la liste (position du séparateur)
$script:splitterIsDragging = $false
$script:splitterStartX = 0

# --- TabControl pour afficher Données+QR ou Aperçu Mail ---
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = '270,22'; $tabControl.Size = '585,295'
$grpPrev.Controls.Add($tabControl)

# TAB 1 : Données & QR
$tabDonnees = New-Object System.Windows.Forms.TabPage
$tabDonnees.Text = "Données & QR"
$tabControl.TabPages.Add($tabDonnees)

# DataGridView pour les données
$dgvDonnees = New-Object System.Windows.Forms.DataGridView
$dgvDonnees.Location = '0,0'; $dgvDonnees.Size = '585,130'
$dgvDonnees.AllowUserToAddRows = $false
$dgvDonnees.AllowUserToDeleteRows = $false
$dgvDonnees.ReadOnly = $true
$dgvDonnees.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dgvDonnees.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dgvDonnees.BackgroundColor = [System.Drawing.Color]::White
$dgvDonnees.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$dgvDonnees.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgvDonnees.RowHeadersVisible = $false
$dgvDonnees.Anchor = 'Top,Left,Right'
$tabDonnees.Controls.Add($dgvDonnees)

# Panel pour QR code (bas)
$pnlQR = New-Object System.Windows.Forms.Panel
$pnlQR.Location = '0,130'; $pnlQR.Size = '585,155'
$pnlQR.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlQR.BackColor = [System.Drawing.Color]::White
$pnlQR.AutoScroll = $true
$pnlQR.Anchor = 'Top,Left,Right,Bottom'
$tabDonnees.Controls.Add($pnlQR)

# PictureBox pour afficher le QR code
$picQR = New-Object System.Windows.Forms.PictureBox
$picQR.Location = '5,5'; $picQR.Size = '150,150'
$picQR.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$picQR.BackColor = [System.Drawing.Color]::White
$picQR.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlQR.Controls.Add($picQR)

# Double-clic = ouvre le fichier QR (handler posé UNE seule fois, chemin lu dans Tag)
$picQR.Add_DoubleClick({
    if ($picQR.Tag -and (Test-Path $picQR.Tag)) {
        try { Start-Process -FilePath $picQR.Tag }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Impossible d'ouvrir : $($_.Exception.Message)", "Erreur", 'OK', 'Error') | Out-Null
        }
    }
})

# Label pour infos QR
$lblQRInfo = New-Object System.Windows.Forms.Label
$lblQRInfo.Location = '160,5'; $lblQRInfo.Size = '415,150'
$lblQRInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblQRInfo.AutoSize = $false
$lblQRInfo.Text = ""
$pnlQR.Controls.Add($lblQRInfo)

# TAB 2 : Aperçu Mail
$tabMail = New-Object System.Windows.Forms.TabPage
$tabMail.Text = "Aperçu Mail"
$tabControl.TabPages.Add($tabMail)

# Préviz HTML du mail sélectionné
$web = New-Object System.Windows.Forms.WebBrowser
$web.Dock = 'Fill'
# AllowNavigation doit rester à $true : avec $false, seul le PREMIER DocumentText
# s'affiche. Les navigations externes sont bloquées via l'événement Navigating.
$web.AllowNavigation = $true
$web.Add_Navigating({
    param($s, $e)
    if ($e.Url -and $e.Url.ToString() -ne 'about:blank') { $e.Cancel = $true }
})
$web.AllowWebBrowserDrop = $false
$web.IsWebBrowserContextMenuEnabled = $false
$web.WebBrowserShortcutsEnabled = $false
$web.ScriptErrorsSuppressed = $true
$tabMail.Controls.Add($web)

# Redimensionnement par drag-and-drop sur la limite entre ListBox et TabControl
$grpPrev.Add_MouseDown({
    param($s, $e)
    if ($e.X -ge $listbox.Right -and $e.X -le $tabControl.Left -and $e.Y -ge 22 -and $e.Y -le 317) {
        $script:splitterIsDragging = $true
        $script:splitterStartX = $e.X
    }
})

$grpPrev.Add_MouseMove({
    param($s, $e)
    if ($script:splitterIsDragging -or ($e.X -ge $listbox.Right -and $e.X -le $tabControl.Left -and $e.Y -ge 22 -and $e.Y -le 317)) {
        $grpPrev.Cursor = [System.Windows.Forms.Cursors]::VSplit
    } else {
        $grpPrev.Cursor = [System.Windows.Forms.Cursors]::Default
    }

    if ($script:splitterIsDragging) {
        $deltaX = $e.X - $script:splitterStartX
        $newX = [Math]::Max(120, [Math]::Min(600, $script:splitterX + $deltaX))
        $listbox.Width     = $newX - 10
        $tabControl.Left   = $newX + 10
        $tabControl.Width  = 855 - $tabControl.Left
        $grpPrev.Refresh()
    }
})

$grpPrev.Add_MouseUp({
    if ($script:splitterIsDragging) {
        $script:splitterX = $listbox.Right
        $script:splitterIsDragging = $false
    }
})

# -------------------------------------------------------------- Boutons d'action (refactorisés)
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Charger les prévisualisations"
$btnStart.Location = '15,560'; $btnStart.Size = '200,36'
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 160)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'; $btnStart.Enabled = $false
$form.Controls.Add($btnStart)

$btnSendAll = New-Object System.Windows.Forms.Button
$btnSendAll.Text = "Envoyer tous les mails"
$btnSendAll.Location = '385,560'; $btnSendAll.Size = '190,36'
$btnSendAll.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
$btnSendAll.ForeColor = [System.Drawing.Color]::White
$btnSendAll.FlatStyle = 'Flat'; $btnSendAll.Enabled = $false
$form.Controls.Add($btnSendAll)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Envoyer un test..."
$btnTest.Location = '225,560'; $btnTest.Size = '150,36'
$btnTest.BackColor = [System.Drawing.Color]::FromArgb(200, 130, 0)
$btnTest.ForeColor = [System.Drawing.Color]::White
$btnTest.FlatStyle = 'Flat'; $btnTest.Enabled = $false
$form.Controls.Add($btnTest)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Annuler"
$btnCancel.Location = '585,560'; $btnCancel.Size = '90,36'
$form.Controls.Add($btnCancel)
$btnCancel.Add_Click({ $form.Close() })

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '685,567'; $progress.Size = '195,22'
$form.Controls.Add($progress)

# ----------------------------------------------------------------------- Journal
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = '15,604'; $txtLog.Size = '865,70'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtLog)

function Write-Log {
    param([string]$Message)
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Message))
}

# ==============================================================================
# 3. ÉTAT DE LA CAMPAGNE
# ==============================================================================
$script:donnees         = $null    # lignes du CSV
$script:index           = -1       # index de la ligne en cours de prévisualisation
$script:stats           = @{ OK = 0; KO = 0; Passes = 0 }
$script:envoiEffectues  = @()      # liste des envois pour la preuve
$script:derniereAdresseTest = ""   # mémorise l'adresse de test pour la session

# ==============================================================================
# 4. FONCTIONS DE PILOTAGE DE LA CAMPAGNE
# ==============================================================================

function Get-LigneCourante { return $script:donnees[$script:index] }

function Get-ChampsLigne {
    <# Extrait (email, nom, cheminQR) de la ligne courante selon le mapping GUI. #>
    $ligne = Get-LigneCourante
    return @{
        Email    = ([string]$ligne.($combos['Email'].SelectedItem)).Trim()
        Nom      = if ($combos['Nom'].SelectedItem) { ([string]$ligne.($combos['Nom'].SelectedItem)).Trim() } else { "" }
        CheminQR = ([string]$ligne.($combos['CheminQR'].SelectedItem)).Trim()
    }
}

function Show-Apercu {
    <#
        Affiche la prévisualisation du mail pour le destinataire à l'index $script:index.
        N'AFFECTE PAS le contenu du ListBox (pour éviter la récursion infinie via
        l'événement SelectedIndexChanged).
    #>
    if ($script:index -lt 0 -or $script:index -ge $script:donnees.Count) { return }

    $c = Get-ChampsLigne
    $num   = $script:index + 1

    # Vérification du QR
    $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)

    # Préviz : {QR} -> image encodée en BASE64 (data URI)
    $imgPrev = if ($qrOk) {
        try {
            $bytesQR = [IO.File]::ReadAllBytes($c.CheminQR)
            $hashQRPref = ([System.Security.Cryptography.SHA256]::Create().ComputeHash($bytesQR) | Select-Object -First 10) -join ""
            Write-Log "Aperçu QR [$num] : hash=$hashQRPref"
            
            $b64 = [Convert]::ToBase64String($bytesQR)
            "<img src=""data:image/png;base64,$b64"" style=""width:280px;height:280px;"" />"
        }
        catch {
            "<div style='color:red;border:2px dashed red;padding:30px;text-align:center;'>ERREUR LECTURE QR :<br/>$($_.Exception.Message)</div>"
        }
    } else {
        "<div style='color:red;border:2px dashed red;padding:30px;text-align:center;'>IMAGE QR INTROUVABLE :<br/>$($c.CheminQR)</div>"
    }

    # Préviz signature
    $sigOk = (-not [string]::IsNullOrWhiteSpace($CheminSignature)) -and (Test-Path $CheminSignature)
    $sigPrev = if ($sigOk) {
        try {
            $mimePrev = switch ([IO.Path]::GetExtension($CheminSignature).ToLower()) {
                '.jpg'  { 'image/jpeg' } '.jpeg' { 'image/jpeg' } '.gif' { 'image/gif' } default { 'image/png' }
            }
            $bytesSig = [IO.File]::ReadAllBytes($CheminSignature)
            $b64s = [Convert]::ToBase64String($bytesSig)
            "<img src=""data:$mimePrev;base64,$b64s"" style=""width:$($LargeurSignaturePx)px;"" />"
        }
        catch {
            "<div style='color:#b06000;'>ERREUR LECTURE SIGNATURE : $($_.Exception.Message)</div>"
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($CheminSignature)) {
        "<div style='color:#b06000;border:1px dashed #b06000;padding:8px;'>Signature configurée mais introuvable : $CheminSignature</div>"
    } else {
        ""
    }

    $html = (($ModeleCorps -replace '\{NOM\}', $c.Nom) -replace '\{QR\}', $imgPrev) -replace '\{SIGNATURE\}', $sigPrev
    $web.DocumentText = $html

    # === AFFICHER LES DONNÉES ET LE QR CODE DANS LE TAB "Données & QR" ===
    
    # Remplir le DataGridView avec les données du destinataire
    $dgvDonnees.DataSource = $null
    $dgvDonnees.Rows.Clear()
    $dgvDonnees.Columns.Clear()
    
    $colChamp  = $dgvDonnees.Columns.Add("Champ", "Champ")
    $colValeur = $dgvDonnees.Columns.Add("Valeur", "Valeur")
    $dgvDonnees.Columns[$colChamp].FillWeight  = 30
    $dgvDonnees.Columns[$colValeur].FillWeight = 70

    # Toutes les colonnes de la ligne CSV (et non les propriétés du hashtable $c)
    $ligne = Get-LigneCourante
    foreach ($prop in $ligne.PSObject.Properties) {
        $dgvDonnees.Rows.Add($prop.Name, [string]$prop.Value) | Out-Null
    }
    $colToken = $ligne.PSObject.Properties.Name | Where-Object { $_ -match 'token' } | Select-Object -First 1
    $token    = if ($colToken) { [string]$ligne.$colToken } else { "(pas de colonne Token)" }

    # Libère l'image précédente
    if ($picQR.Image) { $picQR.Image.Dispose(); $picQR.Image = $null }
    $picQR.Tag = $null

    # Afficher l'image QR (copie en mémoire : FromFile verrouillerait le fichier)
    if ($qrOk) {
        try {
            $ms  = New-Object System.IO.MemoryStream(,[IO.File]::ReadAllBytes($c.CheminQR))
            $tmp = [System.Drawing.Image]::FromStream($ms)
            $picQR.Image = New-Object System.Drawing.Bitmap($tmp)
            $tmp.Dispose(); $ms.Dispose()
            $picQR.Tag = $c.CheminQR
            $lblQRInfo.Text = "$SymOK QR Code : $([System.IO.Path]::GetFileName($c.CheminQR))`n`nToken : $token`n`n(Double-clic pour ouvrir)"
        }
        catch {
            $lblQRInfo.Text = "$SymKO Erreur lecture QR`n`n$($_.Exception.Message)"
        }
    } else {
        $lblQRInfo.Text = "$SymWarn QR Code INTROUVABLE`n`nChemin attendu :`n$($c.CheminQR)"
    }

    $progress.Maximum = $script:donnees.Count
    $progress.Value   = $num
}

function Update-ListboxItemStatus {
    <#
        Met à jour UNIQUEMENT le texte de l'item du ListBox à l'index donné,
        SANS déclencher l'événement SelectedIndexChanged (sinon boucle infinie).
    #>
    param([int]$Index, [string]$ItemText)
    
    $script:suspendSelection = $true
    try {
        if ($listbox.Items.Count -gt $Index) {
            $etaitCoche = $listbox.GetItemChecked($Index)
            $listbox.Items[$Index] = $ItemText
            $listbox.SetItemChecked($Index, $etaitCoche)
        }
    }
    finally { $script:suspendSelection = $false }
}

function Invoke-EnvoiCourant {
    <#
        Envoie le mail de la ligne courante.
        Retourne $true (succès) / $false (échec ou ligne invalide).
        Enregistre chaque tentative dans $script:envoiEffectues pour la preuve d'envoi.
        Toutes les erreurs sont journalisées, jamais bloquantes pour la campagne.
    #>
    $c = Get-ChampsLigne
    $num = $script:index + 1
    $timestampEnvoi = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $statut = "ERREUR"
    $message = ""

    # Garde-fous : ligne incomplète = échec contrôlé, pas d'envoi partiel
    if ([string]::IsNullOrWhiteSpace($c.Email)) {
        Write-Log "[$num] ÉCHEC : email vide."
        $message = "Email vide"
        $script:stats.KO++
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = "N/A"
            Statut       = "ÉCHEC"
            Message      = $message
        }
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($c.CheminQR) -or -not (Test-Path $c.CheminQR)) {
        Write-Log "[$num] ÉCHEC : fichier QR introuvable pour $($c.Email)."
        $message = "Fichier QR introuvable : $($c.CheminQR)"
        $script:stats.KO++
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
            Statut       = "ÉCHEC"
            Message      = $message
        }
        return $false
    }

    # Identifiants : compte explicite si coché, sinon authentification Windows (AD)
    $cred = $null
    if ($chkAuth.Checked) {
        if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
            Write-Log "[$num] ÉCHEC : compte explicite coché mais utilisateur vide."
            $message = "Compte explicite coché mais utilisateur vide"
            $script:stats.KO++
            $script:envoiEffectues += [PSCustomObject]@{
                Timestamp    = $timestampEnvoi
                Destinataire = $c.Email
                Objet        = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
                Statut       = "ÉCHEC"
                Message      = $message
            }
            return $false
        }
        $secure = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $cred   = New-Object System.Management.Automation.PSCredential($txtUser.Text, $secure)
    }

    $objet = ($ModeleObjet -replace '\{NOM\}', $c.Nom)
    try {
        Send-QRMail -Serveur     $txtSrv.Text.Trim() `
                    -Port        ([int]$txtPort.Text) `
                    -UtiliserSSL $chkSsl.Checked `
                    -Credential  $cred `
                    -Expediteur  $txtFrom.Text.Trim() `
                    -Destinataire $c.Email `
                    -Objet       $objet `
                    -CorpsHtml   ($ModeleCorps -replace '\{NOM\}', $c.Nom) `
                    -CheminQR    $c.CheminQR | Out-Null
        Write-Log "[$num] ENVOYÉ : $($c.Email)"
        $statut = "OK"
        $message = "Envoi réussi"
        $script:stats.OK++
        $succes = $true
    }
    catch {
        Write-Log "[$num] ÉCHEC SMTP pour $($c.Email) : $($_.Exception.Message)"
        $statut = "ERREUR SMTP"
        $message = $_.Exception.Message
        $script:stats.KO++
        $succes = $false
    }
    finally {
        # Enregistrer la preuve d'envoi pour CHAQUE tentative
        $script:envoiEffectues += [PSCustomObject]@{
            Timestamp    = $timestampEnvoi
            Destinataire = $c.Email
            Objet        = $objet
            Statut       = $statut
            Message      = $message
        }
    }

    return $succes
}

function Test-ParametresSmtp {
    <# Vérifie que serveur + expéditeur sont renseignés avant tout envoi. #>
    if ([string]::IsNullOrWhiteSpace($txtSrv.Text) -or [string]::IsNullOrWhiteSpace($txtFrom.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Renseignez le serveur SMTP et l'adresse expéditeur avant d'envoyer.",
            "Paramètres SMTP incomplets", 'OK', 'Warning') | Out-Null
        return $false
    }
    if (-not ($txtPort.Text -match '^\d+$')) {
        [System.Windows.Forms.MessageBox]::Show("Le port SMTP doit être numérique.", "Port invalide", 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

function Show-DialogueTest {
    <#
        Petite fenêtre : adresse de réception du test (pré-remplie) + option bandeau.
        Retourne @{ Adresse = ...; Bandeau = $true/$false } ou $null si annulé.
    #>
    param([string]$Defaut, [string]$Contexte)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Envoyer un mail de test"
    $dlg.Size = New-Object System.Drawing.Size(480, 235)
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Adresse qui recevra le test :"; $l1.Location = '15,15'; $l1.AutoSize = $true
    $dlg.Controls.Add($l1)

    $tAdr = New-Object System.Windows.Forms.TextBox
    $tAdr.Location = '15,37'; $tAdr.Size = '435,25'; $tAdr.Text = $Defaut
    $dlg.Controls.Add($tAdr)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = "Contenu utilisé : $Contexte"; $l2.Location = '15,70'; $l2.Size = '435,36'
    $l2.ForeColor = [System.Drawing.Color]::Gray
    $dlg.Controls.Add($l2)

    $chkBandeau = New-Object System.Windows.Forms.CheckBox
    $chkBandeau.Text = "Ajouter un bandeau TEST en haut du mail"
    $chkBandeau.Location = '15,108'; $chkBandeau.AutoSize = $true; $chkBandeau.Checked = $true
    $dlg.Controls.Add($chkBandeau)

    $bOk = New-Object System.Windows.Forms.Button
    $bOk.Text = "Envoyer le test"; $bOk.Location = '225,145'; $bOk.Size = '120,32'
    $dlg.Controls.Add($bOk); $dlg.AcceptButton = $bOk

    $bAnn = New-Object System.Windows.Forms.Button
    $bAnn.Text = "Annuler"; $bAnn.Location = '355,145'; $bAnn.Size = '95,32'
    $bAnn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($bAnn); $dlg.CancelButton = $bAnn

    $bOk.Add_Click({
        try   { $null = New-Object System.Net.Mail.MailAddress($tAdr.Text.Trim()) }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Adresse email invalide.", "Test", 'OK', 'Warning') | Out-Null
            return
        }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $resultat = $null
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $resultat = @{ Adresse = $tAdr.Text.Trim(); Bandeau = $chkBandeau.Checked }
    }
    $dlg.Dispose()
    return $resultat
}

function Invoke-EnvoiTest {
    <#
        Envoie le mail de la ligne $Index vers $AdresseTest au lieu du vrai destinataire.
        Même QR, même corps, même signature que l'envoi réel ; objet préfixé [TEST].
        N'alimente NI les statistiques NI la preuve d'envoi.
    #>
    param([int]$Index, [string]$AdresseTest, [bool]$Bandeau)

    $script:index = $Index
    $c   = Get-ChampsLigne
    $num = $Index + 1

    if ([string]::IsNullOrWhiteSpace($c.CheminQR) -or -not (Test-Path $c.CheminQR)) {
        Write-Log "[TEST] Ligne $num : fichier QR introuvable ($($c.CheminQR))."
        [System.Windows.Forms.MessageBox]::Show("Fichier QR introuvable pour la ligne $num :`n$($c.CheminQR)",
            "Test impossible", 'OK', 'Warning') | Out-Null
        return
    }

    $cred = $null
    if ($chkAuth.Checked) {
        if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Compte explicite coché mais utilisateur vide.",
                "Test impossible", 'OK', 'Warning') | Out-Null
            return
        }
        $secure = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $cred   = New-Object System.Management.Automation.PSCredential($txtUser.Text, $secure)
    }

    $objet = "[TEST] " + ($script:ModeleObjet -replace '\{NOM\}', $c.Nom)
    $corps = $script:ModeleCorps -replace '\{NOM\}', $c.Nom
    if ($Bandeau) {
        $emailHtml = [System.Net.WebUtility]::HtmlEncode($c.Email)
        $bandeauHtml = "<div style='background:#fff3cd;border:1px solid #e0b000;padding:8px;margin-bottom:14px;" +
                       "font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#5a4500;'>" +
                       "MAIL DE TEST - contenu prévu pour <b>$emailHtml</b> (ligne $num du CSV)</div>"
        $m = [regex]::Match($corps, '(?i)<body[^>]*>')
        $corps = if ($m.Success) { $corps.Insert($m.Index + $m.Length, $bandeauHtml) } else { $bandeauHtml + $corps }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Send-QRMail -Serveur      $txtSrv.Text.Trim() `
                    -Port         ([int]$txtPort.Text) `
                    -UtiliserSSL  $chkSsl.Checked `
                    -Credential   $cred `
                    -Expediteur   $txtFrom.Text.Trim() `
                    -Destinataire $AdresseTest `
                    -Objet        $objet `
                    -CorpsHtml    $corps `
                    -CheminQR     $c.CheminQR | Out-Null
        Write-Log "[TEST] Envoyé à $AdresseTest (contenu de la ligne $num : $($c.Email))"
        [System.Windows.Forms.MessageBox]::Show("Mail de test envoyé à $AdresseTest.`n`nVérifiez la réception, l'affichage du QR et de la signature.",
            "Test envoyé", 'OK', 'Information') | Out-Null
    }
    catch {
        Write-Log "[TEST] ÉCHEC SMTP vers $AdresseTest : $($_.Exception.Message)"
        $detail = $_.Exception.Message
        if ($_.Exception.InnerException) { $detail += "`n`n" + $_.Exception.InnerException.Message }
        [System.Windows.Forms.MessageBox]::Show("Échec de l'envoi de test :`n`n$detail", "Test en échec", 'OK', 'Error') | Out-Null
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ==============================================================================
# 5. ÉDITEUR DE MODÈLE (objet, corps HTML, signature)
# ==============================================================================
function Show-EditeurModele {
    <#
        Ouvre une fenêtre modale d'édition du modèle de mail :
        - Objet (avec variable {NOM})
        - Corps HTML (variables {NOM}, {QR} obligatoire, {SIGNATURE} optionnelle)
        - Chemin de l'image de signature (+ bouton Parcourir) et largeur en px
        Les modifications sont appliquées en mémoire ($script:...) :
        - immédiatement visibles dans la prévisualisation
        - valables pour toute la session ; le script lui-même n'est pas modifié.
    #>
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Édition du modèle de mail"
    $dlg.Size            = New-Object System.Drawing.Size(720, 560)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox     = $false
    $dlg.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

    # --- Objet -----------------------------------------------------------------
    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Objet du mail  ({NOM} = nom du destinataire) :"
    $l1.Location = '15,12'; $l1.AutoSize = $true
    $dlg.Controls.Add($l1)

    $tObjet = New-Object System.Windows.Forms.TextBox
    $tObjet.Location = '15,33'; $tObjet.Size = '675,25'
    $tObjet.Text = $script:ModeleObjet
    $dlg.Controls.Add($tObjet)

    # --- Corps HTML --------------------------------------------------------------
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = "Corps HTML  -  variables : {NOM}   {QR} (OBLIGATOIRE)   {SIGNATURE} (optionnelle) :"
    $l2.Location = '15,68'; $l2.AutoSize = $true
    $dlg.Controls.Add($l2)

    $tCorps = New-Object System.Windows.Forms.TextBox
    $tCorps.Location = '15,89'; $tCorps.Size = '675,290'
    $tCorps.Multiline = $true; $tCorps.ScrollBars = 'Vertical'
    $tCorps.AcceptsReturn = $true; $tCorps.WordWrap = $false
    $tCorps.Font = New-Object System.Drawing.Font("Consolas", 9)
    $tCorps.Text = $script:ModeleCorps
    $dlg.Controls.Add($tCorps)

    # --- Signature ------------------------------------------------------------------
    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = "Image de signature (vide = pas de signature) :"
    $l3.Location = '15,392'; $l3.AutoSize = $true
    $dlg.Controls.Add($l3)

    $tSig = New-Object System.Windows.Forms.TextBox
    $tSig.Location = '15,413'; $tSig.Size = '480,25'
    $tSig.Text = $script:CheminSignature
    $dlg.Controls.Add($tSig)

    $bSig = New-Object System.Windows.Forms.Button
    $bSig.Text = "Parcourir..."
    $bSig.Location = '505,411'; $bSig.Size = '90,27'
    $bSig.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Images (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif|Tous les fichiers (*.*)|*.*"
        $ofd.Title  = "Sélectionner l'image de signature"
        if ($ofd.ShowDialog() -eq 'OK') { $tSig.Text = $ofd.FileName }
    })
    $dlg.Controls.Add($bSig)

    $lLarg = New-Object System.Windows.Forms.Label
    $lLarg.Text = "Largeur (px) :"
    $lLarg.Location = '605,392'; $lLarg.AutoSize = $true
    $dlg.Controls.Add($lLarg)

    $numLarg = New-Object System.Windows.Forms.NumericUpDown
    $numLarg.Location = '605,413'; $numLarg.Size = '85,25'
    $numLarg.Minimum = 50; $numLarg.Maximum = 800
    $numLarg.Value = [Math]::Max(50, [Math]::Min(800, $script:LargeurSignaturePx))
    $dlg.Controls.Add($numLarg)

    # --- Boutons Enregistrer / Annuler ------------------------------------------------
    $bOk = New-Object System.Windows.Forms.Button
    $bOk.Text = "Enregistrer"
    $bOk.Location = '460,470'; $bOk.Size = '110,32'
    $bOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
    $bOk.ForeColor = [System.Drawing.Color]::White
    $bOk.FlatStyle = 'Flat'
    $dlg.Controls.Add($bOk)

    $bCancel = New-Object System.Windows.Forms.Button
    $bCancel.Text = "Annuler"
    $bCancel.Location = '580,470'; $bCancel.Size = '110,32'
    $bCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($bCancel)
    $dlg.CancelButton = $bCancel

    $bOk.Add_Click({
        # Garde-fou : le marqueur {QR} est indispensable, sinon le token n'est jamais envoyé
        if ($tCorps.Text -notmatch '\{QR\}') {
            [System.Windows.Forms.MessageBox]::Show(
                "Le corps du mail doit contenir le marqueur {QR} : c'est lui qui reçoit le code QR du destinataire.",
                "Marqueur {QR} manquant", 'OK', 'Warning') | Out-Null
            return
        }
        # Avertissement non bloquant si la signature est renseignée mais introuvable
        if (-not [string]::IsNullOrWhiteSpace($tSig.Text) -and -not (Test-Path $tSig.Text.Trim())) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                "L'image de signature est introuvable :`n$($tSig.Text)`n`nEnregistrer quand même ? (les mails partiront sans signature)",
                "Signature introuvable", 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }

        # Application en mémoire (portée script -> utilisée par préviz ET envoi)
        $script:ModeleObjet        = $tObjet.Text
        $script:ModeleCorps        = $tCorps.Text
        $script:CheminSignature    = $tSig.Text.Trim()
        $script:LargeurSignaturePx = [int]$numLarg.Value
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    if ($dlg.ShowDialog($form) -eq 'OK') {
        Write-Log "Modèle de mail mis à jour (objet/corps/signature)."
        # Rafraîchit la prévisualisation en cours pour refléter le nouveau modèle
        if ($script:index -ge 0 -and $script:index -lt $script:donnees.Count) {
            Show-Apercu
        }
    }
    $dlg.Dispose()
}

$btnModele.Add_Click({ Show-EditeurModele })

# ==============================================================================
# 6. ÉVÉNEMENTS
# ==============================================================================

# --- Chargement du CSV enrichi -------------------------------------------------
$btnCsv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $ofd.Title  = "Sélectionner le CSV enrichi (avec colonne CheminQR)"
    if ($ofd.ShowDialog() -ne 'OK') { return }

    $txtCsv.Text = $ofd.FileName

    # Détection AUTOMATIQUE du délimiteur (;, ,, TAB, espace)
    $delimiter = Detect-CsvDelimiter -FilePath $ofd.FileName
    Write-Log "$SymOK Délimiteur détecté : $(if ($delimiter -eq ';') { 'Point-virgule' } elseif ($delimiter -eq ',') { 'Virgule' } elseif ($delimiter -eq "`t") { 'TAB' } else { 'Espace' })"
    
    # Import en UTF-8 avec délimiteur détecté, repli sur l'encodage par défaut (ANSI) si échec
    try   { $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter $delimiter -Encoding UTF8) }
    catch { $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter $delimiter) }

    if (-not $script:donnees -or $script:donnees.Count -eq 0) {
        Write-Log "ERREUR : CSV vide ou illisible."
        return
    }

    $colonnes = @($script:donnees[0].PSObject.Properties.Name)
    foreach ($k in $combos.Keys) {
        $combos[$k].Items.Clear()
        $combos[$k].Items.AddRange($colonnes)
    }

    # Auto-détection des colonnes (mêmes conventions que le générateur)
    $autoMail = Find-Column $colonnes @('mail', 'courriel')
    $autoNom  = Find-Column $colonnes @('^nom$', 'nom', 'name', 'utilisateur', 'user')
    $autoQR   = Find-Column $colonnes @('cheminqr', 'qr', 'chemin', 'path')
    if ($autoMail) { $combos['Email'].SelectedItem    = $autoMail }
    if ($autoNom)  { $combos['Nom'].SelectedItem      = $autoNom }
    if ($autoQR)   { $combos['CheminQR'].SelectedItem = $autoQR }

    $lblCount.Text = "$($script:donnees.Count) ligne(s)"
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(0,120,90)
    Write-Log "CSV chargé : $($script:donnees.Count) destinataire(s) potentiel(s)."
    $btnStart.Enabled = $true
})

# --- Charger les prévisualisations -----------------------------------------------
$btnStart.Add_Click({
    if (-not $combos['Email'].SelectedItem -or -not $combos['CheminQR'].SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show(
            "Les colonnes Email et Chemin QR doivent être mappées.",
            "Mapping incomplet", 'OK', 'Warning') | Out-Null
        return
    }

    $listbox.Items.Clear()
    $script:index              = -1
    $script:stats              = @{ OK = 0; KO = 0; Passes = 0 }
    $script:envoiEffectues     = @()    # réinitialise la liste des envois
    
    # Peuple la CheckedListBox avec tous les destinataires (sans déclencher SelectedIndexChanged)
    for ($i = 0; $i -lt $script:donnees.Count; $i++) {
        $script:index = $i
        $c = Get-ChampsLigne
        $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)
        $etat = if ($qrOk) { $SymOK } else { $SymKO }
        $itemText = "[$($i+1)] $($c.Email) - $etat"
        [void]$listbox.Items.Add($itemText, $true)   # $true = cocher par défaut
    }

    Write-Log "--- Prévisualisations chargées : $($script:donnees.Count) destinataire(s) (tous cochés) ---"
    $btnStart.Enabled = $false
    $btnSendAll.Enabled = $true
    $btnTest.Enabled    = $true
    $btnSendAll.Text = "Envoyer $($script:donnees.Count) mail(s) cochés"
    
    # Affiche la première préviz (SelectedIndexChanged appelle Show-Apercu)
    $script:index = 0
    [System.Windows.Forms.Application]::DoEvents()
    $listbox.SelectedIndex = 0
})

# --- Envoyer un test -----------------------------------------------------------
$btnTest.Add_Click({
    if (-not (Test-ParametresSmtp)) { return }
    if (-not $script:donnees -or $script:donnees.Count -eq 0) { return }

    # Ligne modèle : destinataire sélectionné, sinon le premier
    $idx = if ($listbox.SelectedIndex -ge 0) { $listbox.SelectedIndex } else { 0 }
    $script:index = $idx
    $c = Get-ChampsLigne

    $defaut = if ($script:derniereAdresseTest) { $script:derniereAdresseTest } else { $txtFrom.Text.Trim() }
    $choix  = Show-DialogueTest -Defaut $defaut -Contexte "ligne $($idx + 1) - $($c.Email)"
    if (-not $choix) { return }

    $script:derniereAdresseTest = $choix.Adresse
    Invoke-EnvoiTest -Index $idx -AdresseTest $choix.Adresse -Bandeau $choix.Bandeau
})

# --- Envoyer tous les mails -------------------------------------------------------
$btnSendAll.Add_Click({
    if (-not (Test-ParametresSmtp)) { return }

    $nbCoches = $listbox.CheckedItems.Count
    if ($nbCoches -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun destinataire coché.", "Envoi", 'OK', 'Information') | Out-Null
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Envoyer $nbCoches mail(s) ?",
        "Confirmation envoi de masse", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }

    $btnSendAll.Enabled = $false
    $btnStart.Enabled = $false
    $btnTest.Enabled = $false
    $listbox.Enabled = $false
    $script:stats = @{ OK = 0; KO = 0; Passes = 0 }
    $script:envoiEffectues = @()    # réinitialise pour cette campagne d'envoi

    for ($i = 0; $i -lt $script:donnees.Count; $i++) {
        # IMPORTANT : envoyer SEULEMENT si l'utilisateur est coché
        if (-not $listbox.GetItemChecked($i)) {
            Write-Log "[$($i+1)] PASSÉ (non coché)"
            continue
        }

        $script:index = $i
        $c = Get-ChampsLigne
        $num = $i + 1

        $progress.Maximum = $script:donnees.Count
        $progress.Value   = $num
        [System.Windows.Forms.Application]::DoEvents()

        $envoiOk = Invoke-EnvoiCourant

        # Met à jour l'item de la liste avec le statut RÉEL de l'envoi
        $qrOk = (-not [string]::IsNullOrWhiteSpace($c.CheminQR)) -and (Test-Path $c.CheminQR)
        $etat = if ($qrOk) { $SymOK } else { $SymKO }
        $statusText = if ($envoiOk) { "ENVOYÉ" } else { "ÉCHEC" }
        $itemText = "[$num] $($c.Email) - $etat - $statusText"
        Update-ListboxItemStatus -Index $i -ItemText $itemText
        
        Start-Sleep -Milliseconds $PauseEntreEnvoisMs
    }

    $btnSendAll.Enabled = $true
    $btnStart.Enabled = $true
    $btnTest.Enabled = $true
    $listbox.Enabled = $true
    
    # --- Génération du fichier de preuve d'envoi (à côté du CSV source) ---
    # ($txtOut n'existait pas dans ce script : Join-Path échouait)
    $preuveFilename = "preuve_envoi_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $preuveFilepath = Join-Path (Split-Path -Parent $txtCsv.Text) $preuveFilename
    try {
        $script:envoiEffectues | Export-Csv -Path $preuveFilepath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Write-Log "Preuve d'envoi générée : $preuveFilepath"
    }
    catch {
        Write-Log "ERREUR écriture preuve d'envoi : $($_.Exception.Message)"
    }
    
    $bilan = "Bilan final : $($script:stats.OK) envoyé(s), $($script:stats.KO) échec(s).`n`nPreuve d'envoi : $preuveFilename"
    Write-Log "--- $bilan ---"
    [System.Windows.Forms.MessageBox]::Show($bilan, "Campagne terminée", 'OK', 'Information') | Out-Null
})

# ==============================================================================
# 7. LANCEMENT
# ==============================================================================
Write-Log "Prêt. Chargez le CSV enrichi, renseignez le SMTP, puis démarrez la campagne."
[void]$form.ShowDialog()
