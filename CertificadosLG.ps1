# ============================================================
# Disenado por Eduardo Rubio
# Certificados LG - Interfaz Grafica
# Basado en: Manual Generacion Certificado Firmado CA (Pentatel)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$PassRAR  = 'Larraskitu48!'
$PassPFX  = 'DIGITALSIGNAGE'
$Root     = (Get-Location).Path
$OpenSSL  = Join-Path $Root "openssl.exe"
if (!(Test-Path $OpenSSL)) { $OpenSSL = Join-Path $Root "bin\openssl.exe" }

$ColorFondo       = [System.Drawing.Color]::FromArgb(22, 24, 30)
$ColorPanel        = [System.Drawing.Color]::FromArgb(32, 35, 44)
$ColorPanelHover   = [System.Drawing.Color]::FromArgb(42, 46, 58)
$ColorAcento       = [System.Drawing.Color]::FromArgb(60, 140, 230)
$ColorAcentoHover  = [System.Drawing.Color]::FromArgb(80, 160, 245)
$ColorVerde        = [System.Drawing.Color]::FromArgb(80, 200, 130)
$ColorRojo         = [System.Drawing.Color]::FromArgb(230, 90, 90)
$ColorAmarillo     = [System.Drawing.Color]::FromArgb(220, 190, 70)
$ColorTexto        = [System.Drawing.Color]::FromArgb(200, 205, 215)
$ColorTextoSub     = [System.Drawing.Color]::FromArgb(130, 138, 158)
$ColorBorde        = [System.Drawing.Color]::FromArgb(58, 62, 78)
$ColorLogFondo     = [System.Drawing.Color]::FromArgb(18, 20, 26)
$ColorTextoClaro   = [System.Drawing.Color]::FromArgb(170, 178, 195)

$FuenteTitulo  = New-Object System.Drawing.Font("Segoe UI Semibold", 15)
$FuenteNormal  = New-Object System.Drawing.Font("Segoe UI", 9.5)
$FuentePequena = New-Object System.Drawing.Font("Segoe UI", 8.5)
$FuenteBoton   = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5)
$FuenteMono    = New-Object System.Drawing.Font("Consolas", 9.5)
$FuenteEstado  = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$FuenteSeccion = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Certificados LG  -  Eduardo Rubio"
$Form.Size = New-Object System.Drawing.Size(780, 700)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox = $false
$Form.BackColor = $ColorFondo
$Form.Font = $FuenteNormal
$Form.ForeColor = $ColorTexto

$PanelHeader = New-Object System.Windows.Forms.Panel
$PanelHeader.Size = New-Object System.Drawing.Size(780, 64)
$PanelHeader.Location = New-Object System.Drawing.Point(0, 0)
$PanelHeader.BackColor = $ColorPanel
$Form.Controls.Add($PanelHeader)

$LblTitulo = New-Object System.Windows.Forms.Label
$LblTitulo.Text = "[+]  Certificados LG"
$LblTitulo.Font = $FuenteTitulo
$LblTitulo.ForeColor = $ColorTexto
$LblTitulo.AutoSize = $true
$LblTitulo.Location = New-Object System.Drawing.Point(18, 10)
$PanelHeader.Controls.Add($LblTitulo)

$LblSubtitulo = New-Object System.Windows.Forms.Label
$LblSubtitulo.Text = "Generador automatico  |  Disenado por Eduardo Rubio"
$LblSubtitulo.Font = $FuentePequena
$LblSubtitulo.ForeColor = $ColorTextoSub
$LblSubtitulo.AutoSize = $true
$LblSubtitulo.Location = New-Object System.Drawing.Point(20, 40)
$PanelHeader.Controls.Add($LblSubtitulo)

$LineaAcento = New-Object System.Windows.Forms.Panel
$LineaAcento.Size = New-Object System.Drawing.Size(780, 2)
$LineaAcento.Location = New-Object System.Drawing.Point(0, 64)
$LineaAcento.BackColor = $ColorAcento
$Form.Controls.Add($LineaAcento)

$LblRarTitulo = New-Object System.Windows.Forms.Label
$LblRarTitulo.Text = "CARPETA DE ARCHIVOS RAR"
$LblRarTitulo.Font = $FuenteSeccion
$LblRarTitulo.ForeColor = $ColorAcento
$LblRarTitulo.AutoSize = $true
$LblRarTitulo.Location = New-Object System.Drawing.Point(25, 82)
$Form.Controls.Add($LblRarTitulo)

$PanelRuta = New-Object System.Windows.Forms.Panel
$PanelRuta.Size = New-Object System.Drawing.Size(720, 42)
$PanelRuta.Location = New-Object System.Drawing.Point(25, 104)
$PanelRuta.BackColor = $ColorPanel
$Form.Controls.Add($PanelRuta)

$PanelRuta.Add_Paint({
    $g = $_.Graphics
    $pen = New-Object System.Drawing.Pen($ColorBorde, 1)
    $g.DrawRectangle($pen, 0, 0, $PanelRuta.Width - 1, $PanelRuta.Height - 1)
    $pen.Dispose()
})

$TxtRuta = New-Object System.Windows.Forms.TextBox
$TxtRuta.Size = New-Object System.Drawing.Size(590, 28)
$TxtRuta.Location = New-Object System.Drawing.Point(12, 10)
$TxtRuta.Font = $FuenteNormal
$TxtRuta.BackColor = $ColorPanel
$TxtRuta.ForeColor = $ColorTextoSub
$TxtRuta.BorderStyle = "None"
$TxtRuta.Text = "Seleccione una carpeta..."
$TxtRuta.ReadOnly = $true
$PanelRuta.Controls.Add($TxtRuta)

$BtnExaminar = New-Object System.Windows.Forms.Button
$BtnExaminar.Text = "Examinar..."
$BtnExaminar.Size = New-Object System.Drawing.Size(105, 30)
$BtnExaminar.Location = New-Object System.Drawing.Point(608, 6)
$BtnExaminar.FlatStyle = "Flat"
$BtnExaminar.FlatAppearance.BorderSize = 0
$BtnExaminar.BackColor = $ColorAcento
$BtnExaminar.ForeColor = [System.Drawing.Color]::White
$BtnExaminar.Font = $FuentePequena
$BtnExaminar.Cursor = [System.Windows.Forms.Cursors]::Hand
$PanelRuta.Controls.Add($BtnExaminar)

$PanelInfo = New-Object System.Windows.Forms.Panel
$PanelInfo.Size = New-Object System.Drawing.Size(720, 50)
$PanelInfo.Location = New-Object System.Drawing.Point(25, 156)
$PanelInfo.BackColor = $ColorPanel
$Form.Controls.Add($PanelInfo)

$PanelInfo.Add_Paint({
    $g = $_.Graphics
    $pen = New-Object System.Drawing.Pen($ColorBorde, 1)
    $g.DrawRectangle($pen, 0, 0, $PanelInfo.Width - 1, $PanelInfo.Height - 1)
    $pen.Dispose()
})

$sslOk = Test-Path $OpenSSL
$zipOk = Test-Path $SevenZip

$LblSSL = New-Object System.Windows.Forms.Label
$LblSSL.Text = if ($sslOk) { "[OK] OpenSSL encontrado" } else { "[!!] OpenSSL no encontrado" }
$LblSSL.Font = $FuentePequena
$LblSSL.ForeColor = if ($sslOk) { $ColorVerde } else { $ColorRojo }
$LblSSL.AutoSize = $true
$LblSSL.Location = New-Object System.Drawing.Point(15, 8)
$PanelInfo.Controls.Add($LblSSL)

$LblSSLPath = New-Object System.Windows.Forms.Label
$LblSSLPath.Text = $OpenSSL
$LblSSLPath.Font = $FuentePequena
$LblSSLPath.ForeColor = $ColorTextoSub
$LblSSLPath.AutoSize = $true
$LblSSLPath.Location = New-Object System.Drawing.Point(15, 28)
$PanelInfo.Controls.Add($LblSSLPath)

$Lbl7z = New-Object System.Windows.Forms.Label
$Lbl7z.Text = if ($zipOk) { "[OK] 7-Zip encontrado" } else { "[!!] 7-Zip no encontrado" }
$Lbl7z.Font = $FuentePequena
$Lbl7z.ForeColor = if ($zipOk) { $ColorVerde } else { $ColorRojo }
$Lbl7z.AutoSize = $true
$Lbl7z.Location = New-Object System.Drawing.Point(400, 8)
$PanelInfo.Controls.Add($Lbl7z)

$Lbl7zPath = New-Object System.Windows.Forms.Label
$Lbl7zPath.Text = $SevenZip
$Lbl7zPath.Font = $FuentePequena
$Lbl7zPath.ForeColor = $ColorTextoSub
$Lbl7zPath.AutoSize = $true
$Lbl7zPath.Location = New-Object System.Drawing.Point(400, 28)
$PanelInfo.Controls.Add($Lbl7zPath)

$LblRutasTitulo = New-Object System.Windows.Forms.Label
$LblRutasTitulo.Text = "RUTAS DE SALIDA"
$LblRutasTitulo.Font = $FuenteSeccion
$LblRutasTitulo.ForeColor = $ColorAcento
$LblRutasTitulo.AutoSize = $true
$LblRutasTitulo.Location = New-Object System.Drawing.Point(25, 218)
$Form.Controls.Add($LblRutasTitulo)

$LblRutaCerts = New-Object System.Windows.Forms.Label
$LblRutaCerts.Text = "Certificados:      .\Certificados\<hostname>\   (extraccion + .pfx + .pem intermedios)"
$LblRutaCerts.Font = $FuentePequena
$LblRutaCerts.ForeColor = $ColorTextoClaro
$LblRutaCerts.AutoSize = $true
$LblRutaCerts.Location = New-Object System.Drawing.Point(25, 238)
$Form.Controls.Add($LblRutaCerts)

$LblRutaLG = New-Object System.Windows.Forms.Label
$LblRutaLG.Text = "LG Certificates:   .\LG Certificates\<hostname>\  (ca_certificate + client_certificate + client_key)"
$LblRutaLG.Font = $FuentePequena
$LblRutaLG.ForeColor = $ColorTextoClaro
$LblRutaLG.AutoSize = $true
$LblRutaLG.Location = New-Object System.Drawing.Point(25, 256)
$Form.Controls.Add($LblRutaLG)

$LblLogTitulo = New-Object System.Windows.Forms.Label
$LblLogTitulo.Text = "REGISTRO DE ACTIVIDAD"
$LblLogTitulo.Font = $FuenteSeccion
$LblLogTitulo.ForeColor = $ColorAcento
$LblLogTitulo.AutoSize = $true
$LblLogTitulo.Location = New-Object System.Drawing.Point(25, 282)
$Form.Controls.Add($LblLogTitulo)

$TxtLog = New-Object System.Windows.Forms.RichTextBox
$TxtLog.Size = New-Object System.Drawing.Size(720, 230)
$TxtLog.Location = New-Object System.Drawing.Point(25, 302)
$TxtLog.BackColor = $ColorLogFondo
$TxtLog.ForeColor = $ColorTexto
$TxtLog.Font = $FuenteMono
$TxtLog.ReadOnly = $true
$TxtLog.BorderStyle = "None"
$TxtLog.ScrollBars = "Vertical"
$Form.Controls.Add($TxtLog)

$ProgressBack = New-Object System.Windows.Forms.Panel
$ProgressBack.Size = New-Object System.Drawing.Size(720, 5)
$ProgressBack.Location = New-Object System.Drawing.Point(25, 540)
$ProgressBack.BackColor = $ColorPanel
$Form.Controls.Add($ProgressBack)

$ProgressBar = New-Object System.Windows.Forms.Panel
$ProgressBar.Size = New-Object System.Drawing.Size(0, 5)
$ProgressBar.Location = New-Object System.Drawing.Point(0, 0)
$ProgressBar.BackColor = $ColorAcento
$ProgressBack.Controls.Add($ProgressBar)

$LblEstado = New-Object System.Windows.Forms.Label
$LblEstado.Text = "Esperando..."
$LblEstado.Font = $FuenteEstado
$LblEstado.ForeColor = $ColorTextoSub
$LblEstado.Size = New-Object System.Drawing.Size(500, 22)
$LblEstado.Location = New-Object System.Drawing.Point(25, 553)
$Form.Controls.Add($LblEstado)

$LblContador = New-Object System.Windows.Forms.Label
$LblContador.Text = ""
$LblContador.Font = $FuenteEstado
$LblContador.ForeColor = $ColorTextoSub
$LblContador.AutoSize = $true
$LblContador.Location = New-Object System.Drawing.Point(640, 553)
$Form.Controls.Add($LblContador)

$BtnProcesar = New-Object System.Windows.Forms.Button
$BtnProcesar.Text = ">  PROCESAR"
$BtnProcesar.Size = New-Object System.Drawing.Size(190, 46)
$BtnProcesar.Location = New-Object System.Drawing.Point(25, 585)
$BtnProcesar.FlatStyle = "Flat"
$BtnProcesar.FlatAppearance.BorderSize = 0
$BtnProcesar.BackColor = $ColorAcento
$BtnProcesar.ForeColor = [System.Drawing.Color]::White
$BtnProcesar.Font = $FuenteBoton
$BtnProcesar.Cursor = [System.Windows.Forms.Cursors]::Hand
$BtnProcesar.Enabled = $false
$Form.Controls.Add($BtnProcesar)

$BtnAbrir = New-Object System.Windows.Forms.Button
$BtnAbrir.Text = "Abrir LG Certificates"
$BtnAbrir.Size = New-Object System.Drawing.Size(155, 46)
$BtnAbrir.Location = New-Object System.Drawing.Point(230, 585)
$BtnAbrir.FlatStyle = "Flat"
$BtnAbrir.FlatAppearance.BorderSize = 1
$BtnAbrir.FlatAppearance.BorderColor = $ColorBorde
$BtnAbrir.BackColor = $ColorPanel
$BtnAbrir.ForeColor = $ColorTexto
$BtnAbrir.Font = $FuentePequena
$BtnAbrir.Cursor = [System.Windows.Forms.Cursors]::Hand
$BtnAbrir.Enabled = $false
$Form.Controls.Add($BtnAbrir)

$BtnBorrar = New-Object System.Windows.Forms.Button
$BtnBorrar.Text = "Borrar RAR"
$BtnBorrar.Size = New-Object System.Drawing.Size(120, 46)
$BtnBorrar.Location = New-Object System.Drawing.Point(400, 585)
$BtnBorrar.FlatStyle = "Flat"
$BtnBorrar.FlatAppearance.BorderSize = 1
$BtnBorrar.FlatAppearance.BorderColor = $ColorBorde
$BtnBorrar.BackColor = $ColorPanel
$BtnBorrar.ForeColor = $ColorRojo
$BtnBorrar.Font = $FuentePequena
$BtnBorrar.Cursor = [System.Windows.Forms.Cursors]::Hand
$BtnBorrar.Enabled = $false
$Form.Controls.Add($BtnBorrar)

$BtnSalir = New-Object System.Windows.Forms.Button
$BtnSalir.Text = "Salir"
$BtnSalir.Size = New-Object System.Drawing.Size(85, 46)
$BtnSalir.Location = New-Object System.Drawing.Point(660, 585)
$BtnSalir.FlatStyle = "Flat"
$BtnSalir.FlatAppearance.BorderSize = 1
$BtnSalir.FlatAppearance.BorderColor = $ColorBorde
$BtnSalir.BackColor = $ColorPanel
$BtnSalir.ForeColor = $ColorTextoSub
$BtnSalir.Font = $FuentePequena
$BtnSalir.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($BtnSalir)

$script:RarFolder = ""
$script:RarFiles  = @()
$script:LGDest    = ""

function Log($texto, $color) {
    $TxtLog.SelectionStart = $TxtLog.TextLength
    $TxtLog.SelectionColor = $color
    $TxtLog.AppendText("$texto`n")
    $TxtLog.ScrollToCaret()
    $Form.Refresh()
}

function SetProgress($pct) {
    $w = [math]::Round(720 * ($pct / 100))
    $ProgressBar.Size = New-Object System.Drawing.Size($w, 5)
    if ($pct -ge 100) { $ProgressBar.BackColor = $ColorVerde }
    else { $ProgressBar.BackColor = $ColorAcento }
    $Form.Refresh()
}

function SetEstado($texto, $color) {
    $LblEstado.Text = $texto
    $LblEstado.ForeColor = $color
    $Form.Refresh()
}

$BtnExaminar.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Seleccione la carpeta con los archivos SCTASK*.rar"
    $fb.ShowNewFolderButton = $false
    if ($fb.ShowDialog() -eq 'OK') {
        $script:RarFolder = $fb.SelectedPath
        $script:RarFiles = Get-ChildItem -Path $script:RarFolder -Filter "SCTASK*.rar" -File
        $TxtRuta.Text = $script:RarFolder
        $TxtRuta.ForeColor = $ColorTexto
        if ($script:RarFiles.Count -eq 0) {
            SetEstado "Sin archivos SCTASK*.rar en esa carpeta" $ColorAmarillo
            $BtnProcesar.Enabled = $false
            $TxtLog.Clear()
            Log "[!!] No se encontraron archivos SCTASK*.rar en:" $ColorAmarillo
            Log "     $($script:RarFolder)" $ColorTextoSub
        } else {
            $canRun = $sslOk -and $zipOk
            $BtnProcesar.Enabled = $canRun
            if (!$canRun) { SetEstado "Faltan dependencias (ver panel superior)" $ColorRojo }
            else { SetEstado "Listo para procesar" $ColorVerde }
            $TxtLog.Clear()
            Log "Carpeta seleccionada: $($script:RarFolder)" $ColorTexto
            Log "Archivos encontrados: $($script:RarFiles.Count)" $ColorAcento
            Log "" $ColorTexto
            foreach ($r in $script:RarFiles) {
                Log "   - $($r.Name)  ($([math]::Round($r.Length/1KB, 1)) KB)" $ColorTextoSub
            }
            $LblContador.Text = "0 / $($script:RarFiles.Count)"
        }
    }
})

$BtnExaminar.Add_MouseEnter({ $BtnExaminar.BackColor = $ColorAcentoHover })
$BtnExaminar.Add_MouseLeave({ $BtnExaminar.BackColor = $ColorAcento })
$BtnProcesar.Add_MouseEnter({ if ($BtnProcesar.Enabled) { $BtnProcesar.BackColor = $ColorAcentoHover } })
$BtnProcesar.Add_MouseLeave({ $BtnProcesar.BackColor = $ColorAcento })
$BtnAbrir.Add_MouseEnter({ $BtnAbrir.BackColor = $ColorPanelHover })
$BtnAbrir.Add_MouseLeave({ $BtnAbrir.BackColor = $ColorPanel })
$BtnBorrar.Add_MouseEnter({ $BtnBorrar.BackColor = $ColorPanelHover })
$BtnBorrar.Add_MouseLeave({ $BtnBorrar.BackColor = $ColorPanel })
$BtnSalir.Add_MouseEnter({ $BtnSalir.BackColor = $ColorPanelHover })
$BtnSalir.Add_MouseLeave({ $BtnSalir.BackColor = $ColorPanel })

$BtnProcesar.Add_Click({
    $BtnProcesar.Enabled = $false
    $BtnExaminar.Enabled = $false
    $BtnAbrir.Enabled = $false
    $BtnBorrar.Enabled = $false
    $TxtLog.Clear()
    SetProgress 0

    $CertBase = Join-Path $Root "Certificados"
    $script:LGDest = Join-Path $Root "LG Certificates"
    if (!(Test-Path $CertBase))      { New-Item -ItemType Directory -Path $CertBase | Out-Null }
    if (!(Test-Path $script:LGDest)) { New-Item -ItemType Directory -Path $script:LGDest | Out-Null }

    $total = $script:RarFiles.Count
    $procesados = 0
    $errores = 0
    $i = 0

    Log "Iniciando procesamiento de $total archivo(s)..." $ColorAcento
    Log "========================================================" $ColorBorde
    Log "" $ColorTexto

    foreach ($Rar in $script:RarFiles) {
        $i++
        $pctBase = (($i - 1) / $total) * 100
        $LblContador.Text = "$i / $total"
        SetEstado "Procesando: $($Rar.Name)" $ColorAcento
        Log "[$i/$total] $($Rar.Name)" $ColorAcento

        $TmpDir = Join-Path $script:RarFolder "temp_$([System.IO.Path]::GetRandomFileName())"
        New-Item -ItemType Directory -Path $TmpDir | Out-Null
        & $SevenZip x $Rar.FullName "-o$TmpDir" "-p$PassRAR" -y 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Log "   [ERROR] No se pudo extraer el RAR" $ColorRojo
            Log "" $ColorTexto
            $errores++
            Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
            SetProgress (($i / $total) * 100)
            continue
        }
        Log "   [OK] Extraido" $ColorVerde

        $CerFile = Get-ChildItem -Path $TmpDir -Filter "*.cer" -Recurse | Select-Object -First 1
        if (!$CerFile) {
            Log "   [ERROR] No se encontro archivo .cer dentro del RAR" $ColorRojo
            Log "" $ColorTexto
            $errores++
            Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
            SetProgress (($i / $total) * 100)
            continue
        }

        $CerName = $CerFile.BaseName
        Log "   Hostname: $CerName" $ColorTextoClaro

        $CertDir = Join-Path $CertBase $CerName
        if (!(Test-Path $CertDir)) { New-Item -ItemType Directory -Path $CertDir | Out-Null }

        Get-ChildItem -Path $TmpDir -Recurse -File |
            Where-Object { $_.Extension -notin '.rsp','.req' } |
            ForEach-Object { Move-Item $_.FullName -Destination $CertDir -Force }
        Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue

        SetProgress ($pctBase + (1/$total)*15)

        $CerPath    = Join-Path $CertDir "$CerName.cer"
        $KeyPath    = Join-Path $CertDir "$CerName.key"
        $IssuingCrt = Join-Path $CertDir "Iberdrola Issuing CA v3.crt"
        $RootCrt    = Join-Path $CertDir "Iberdrola Root CA v3.crt"

        $ok = $true
        if (!(Test-Path $CerPath))    { Log "   [FALTA] $CerName.cer" $ColorRojo; $ok = $false }
        if (!(Test-Path $IssuingCrt)) { Log "   [FALTA] Iberdrola Issuing CA v3.crt" $ColorRojo; $ok = $false }
        if (!(Test-Path $RootCrt))    { Log "   [FALTA] Iberdrola Root CA v3.crt" $ColorRojo; $ok = $false }
        if (!$ok) { Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue }

        $PfxPath = Join-Path $CertDir "$CerName.pfx"
        if (Test-Path $KeyPath) {
            Log "   Generando .pfx ..." $ColorTextoSub
            & $OpenSSL pkcs12 -export -out $PfxPath -inkey $KeyPath -in $CerPath -passout "pass:$PassPFX" 2>&1 | Out-Null
            if (!(Test-Path $PfxPath)) {
                Log "   [ERROR] Fallo al generar .pfx" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
            }
            Log "   [OK] .pfx generado" $ColorVerde
        } else {
            Log "   [AVISO] Sin .key - buscando .pfx existente..." $ColorAmarillo
            if (!(Test-Path $PfxPath)) {
                Log "   [ERROR] No hay .key ni .pfx disponible" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
            }
            Log "   [OK] .pfx existente encontrado" $ColorVerde
        }

        SetProgress ($pctBase + (1/$total)*30)

        $IssuingPem = Join-Path $CertDir "Iberdrola Issuing CA v3.pem"
        $RootPem    = Join-Path $CertDir "Iberdrola Root CA v3.pem"
        Log "   Convirtiendo .crt a .pem ..." $ColorTextoSub
        & $OpenSSL x509 -in $IssuingCrt -outform PEM -out $IssuingPem 2>&1 | Out-Null
        & $OpenSSL x509 -in $RootCrt    -outform PEM -out $RootPem    2>&1 | Out-Null
        if (!(Test-Path $IssuingPem) -or !(Test-Path $RootPem)) {
            Log "   [ERROR] Conversion .crt a .pem fallida" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   [OK] .pem intermedios generados" $ColorVerde

        SetProgress ($pctBase + (1/$total)*50)

        $LGDir = Join-Path $script:LGDest $CerName
        if (!(Test-Path $LGDir)) { New-Item -ItemType Directory -Path $LGDir | Out-Null }

        Log "   Generando ca_certificate.pem ..." $ColorTextoSub
        $IssContent  = (Get-Content $IssuingPem -Raw).TrimEnd("`r`n")
        $RootContent = (Get-Content $RootPem    -Raw).TrimEnd("`r`n")
        [System.IO.File]::WriteAllText((Join-Path $LGDir "ca_certificate.pem"), "$IssContent`r`n$RootContent`r`n", [System.Text.UTF8Encoding]::new($false))
        Log "   [OK] ca_certificate.pem" $ColorVerde

        SetProgress ($pctBase + (1/$total)*70)

        Log "   Generando client_certificate.pem ..." $ColorTextoSub
        $ClientCert = Join-Path $LGDir "client_certificate.pem"
        & $OpenSSL pkcs12 -in $PfxPath -clcerts -nokeys -out $ClientCert -passin "pass:$PassPFX" 2>&1 | Out-Null
        if (!(Test-Path $ClientCert)) {
            Log "   [ERROR] client_certificate.pem" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   [OK] client_certificate.pem" $ColorVerde

        SetProgress ($pctBase + (1/$total)*85)

        Log "   Generando client_key.pem ..." $ColorTextoSub
        $ClientKey = Join-Path $LGDir "client_key.pem"
        & $OpenSSL pkcs12 -in $PfxPath -nocerts -nodes -out $ClientKey -passin "pass:$PassPFX" 2>&1 | Out-Null
        if (!(Test-Path $ClientKey)) {
            Log "   [ERROR] client_key.pem" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   [OK] client_key.pem" $ColorVerde

        $LGOk = @("ca_certificate.pem","client_certificate.pem","client_key.pem") | Where-Object { Test-Path (Join-Path $LGDir $_) }
        if ($LGOk.Count -eq 3) {
            Log "   -----------------------------------------------" $ColorVerde
            Log "   [COMPLETO] $CerName  (3/3 archivos LG)" $ColorVerde
            Log "   -----------------------------------------------" $ColorVerde
            $procesados++
        } else {
            Log "   [AVISO] Carpeta LG incompleta ($($LGOk.Count)/3)" $ColorAmarillo
            $errores++
        }
        Log "" $ColorTexto
        SetProgress (($i / $total) * 100)
    }

    Log "========================================================" $ColorBorde
    Log "" $ColorTexto
    Log "RESUMEN: $procesados / $total procesados correctamente" $ColorAcento
    if ($errores -gt 0) { Log "Errores: $errores" $ColorRojo }
    Log "" $ColorTexto
    Log "Certificados:    $CertBase" $ColorTextoSub
    Log "LG Certificates: $($script:LGDest)" $ColorTextoSub

    SetProgress 100
    $LblContador.Text = "$procesados / $total"
    if ($errores -eq 0) { SetEstado "[OK] Completado sin errores" $ColorVerde }
    else { SetEstado "[!!] Completado con $errores error(es)" $ColorAmarillo }

    $BtnAbrir.Enabled = $true
    $BtnBorrar.Enabled = $true
    $BtnExaminar.Enabled = $true
})

$BtnAbrir.Add_Click({
    if ($script:LGDest -and (Test-Path $script:LGDest)) { Start-Process explorer.exe $script:LGDest }
})

$BtnBorrar.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show("Eliminar $($script:RarFiles.Count) archivo(s) RAR originales?`n`nEsta accion no se puede deshacer.", "Confirmar eliminacion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq 'Yes') {
        $script:RarFiles | Remove-Item -Force
        Log "" $ColorTexto
        Log "[X] Archivos RAR eliminados." $ColorRojo
        $BtnBorrar.Enabled = $false
    }
})

$BtnSalir.Add_Click({ $Form.Close() })
$Form.ShowDialog() | Out-Null
