# ============================================================
# Diseñado por Eduardo Rubio
# Certificados LG - Interfaz Grafica
# Basado en: Manual Generacion Certificado Firmado CA (Pentatel)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─── Configuracion ───
$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$PassRAR  = 'Larraskitu48!'
$PassPFX  = 'DIGITALSIGNAGE'
$Root     = (Get-Location).Path
$OpenSSL  = Join-Path $Root "openssl.exe"
if (!(Test-Path $OpenSSL)) { $OpenSSL = Join-Path $Root "bin\openssl.exe" }

# ─── Colores y Fuentes ───
$ColorFondo      = [System.Drawing.Color]::FromArgb(18, 18, 24)
$ColorPanel       = [System.Drawing.Color]::FromArgb(28, 28, 38)
$ColorPanelHover  = [System.Drawing.Color]::FromArgb(38, 38, 50)
$ColorAcento      = [System.Drawing.Color]::FromArgb(0, 150, 255)
$ColorAcentoHover = [System.Drawing.Color]::FromArgb(30, 170, 255)
$ColorVerde       = [System.Drawing.Color]::FromArgb(0, 200, 120)
$ColorRojo        = [System.Drawing.Color]::FromArgb(255, 80, 80)
$ColorAmarillo    = [System.Drawing.Color]::FromArgb(255, 200, 60)
$ColorTexto       = [System.Drawing.Color]::FromArgb(230, 230, 240)
$ColorTextoSub    = [System.Drawing.Color]::FromArgb(140, 140, 165)
$ColorBorde       = [System.Drawing.Color]::FromArgb(55, 55, 75)

$FuenteTitulo  = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$FuenteNormal  = New-Object System.Drawing.Font("Segoe UI", 10)
$FuentePequena = New-Object System.Drawing.Font("Segoe UI", 8.5)
$FuenteBoton   = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$FuenteMono    = New-Object System.Drawing.Font("Cascadia Code, Consolas", 9)
$FuenteEstado  = New-Object System.Drawing.Font("Segoe UI Semibold", 11)

# ═══════════════════════════════════════════════════════════
# FORMULARIO PRINCIPAL
# ═══════════════════════════════════════════════════════════
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Certificados LG"
$Form.Size = New-Object System.Drawing.Size(780, 680)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox = $false
$Form.BackColor = $ColorFondo
$Form.Font = $FuenteNormal
$Form.ForeColor = $ColorTexto

# ─── Header ───
$PanelHeader = New-Object System.Windows.Forms.Panel
$PanelHeader.Size = New-Object System.Drawing.Size(780, 70)
$PanelHeader.Location = New-Object System.Drawing.Point(0, 0)
$PanelHeader.BackColor = $ColorPanel
$Form.Controls.Add($PanelHeader)

$LblTitulo = New-Object System.Windows.Forms.Label
$LblTitulo.Text = "🔒  Certificados LG"
$LblTitulo.Font = $FuenteTitulo
$LblTitulo.ForeColor = $ColorTexto
$LblTitulo.AutoSize = $true
$LblTitulo.Location = New-Object System.Drawing.Point(20, 12)
$PanelHeader.Controls.Add($LblTitulo)

$LblSubtitulo = New-Object System.Windows.Forms.Label
$LblSubtitulo.Text = "Generador automatico  ·  Diseñado por Eduardo Rubio"
$LblSubtitulo.Font = $FuentePequena
$LblSubtitulo.ForeColor = $ColorTextoSub
$LblSubtitulo.AutoSize = $true
$LblSubtitulo.Location = New-Object System.Drawing.Point(22, 44)
$PanelHeader.Controls.Add($LblSubtitulo)

# Linea acento bajo header
$LineaAcento = New-Object System.Windows.Forms.Panel
$LineaAcento.Size = New-Object System.Drawing.Size(780, 2)
$LineaAcento.Location = New-Object System.Drawing.Point(0, 70)
$LineaAcento.BackColor = $ColorAcento
$Form.Controls.Add($LineaAcento)

# ─── Seccion: Carpeta RAR ───
$LblRarTitulo = New-Object System.Windows.Forms.Label
$LblRarTitulo.Text = "CARPETA DE ARCHIVOS RAR"
$LblRarTitulo.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$LblRarTitulo.ForeColor = $ColorAcento
$LblRarTitulo.AutoSize = $true
$LblRarTitulo.Location = New-Object System.Drawing.Point(25, 88)
$Form.Controls.Add($LblRarTitulo)

$PanelRuta = New-Object System.Windows.Forms.Panel
$PanelRuta.Size = New-Object System.Drawing.Size(720, 44)
$PanelRuta.Location = New-Object System.Drawing.Point(25, 110)
$PanelRuta.BackColor = $ColorPanel
$Form.Controls.Add($PanelRuta)

# Borde redondeado simulado
$PanelRuta.Add_Paint({
    $g = $_.Graphics
    $pen = New-Object System.Drawing.Pen($ColorBorde, 1)
    $g.DrawRectangle($pen, 0, 0, $PanelRuta.Width - 1, $PanelRuta.Height - 1)
    $pen.Dispose()
})

$TxtRuta = New-Object System.Windows.Forms.TextBox
$TxtRuta.Size = New-Object System.Drawing.Size(590, 30)
$TxtRuta.Location = New-Object System.Drawing.Point(12, 10)
$TxtRuta.Font = $FuenteNormal
$TxtRuta.BackColor = $ColorPanel
$TxtRuta.ForeColor = $ColorTextoSub
$TxtRuta.BorderStyle = "None"
$TxtRuta.Text = "Seleccione una carpeta..."
$TxtRuta.ReadOnly = $true
$PanelRuta.Controls.Add($TxtRuta)

$BtnExaminar = New-Object System.Windows.Forms.Button
$BtnExaminar.Text = "📁 Examinar"
$BtnExaminar.Size = New-Object System.Drawing.Size(105, 32)
$BtnExaminar.Location = New-Object System.Drawing.Point(608, 6)
$BtnExaminar.FlatStyle = "Flat"
$BtnExaminar.FlatAppearance.BorderSize = 0
$BtnExaminar.BackColor = $ColorAcento
$BtnExaminar.ForeColor = [System.Drawing.Color]::White
$BtnExaminar.Font = $FuentePequena
$BtnExaminar.Cursor = [System.Windows.Forms.Cursors]::Hand
$PanelRuta.Controls.Add($BtnExaminar)

# ─── Info Panel (OpenSSL + 7zip status) ───
$PanelInfo = New-Object System.Windows.Forms.Panel
$PanelInfo.Size = New-Object System.Drawing.Size(720, 50)
$PanelInfo.Location = New-Object System.Drawing.Point(25, 164)
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
$LblSSL.Text = if ($sslOk) { "✅  OpenSSL encontrado" } else { "❌  OpenSSL no encontrado" }
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
$Lbl7z.Text = if ($zipOk) { "✅  7-Zip encontrado" } else { "❌  7-Zip no encontrado" }
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

# ─── Log de actividad ───
$LblLogTitulo = New-Object System.Windows.Forms.Label
$LblLogTitulo.Text = "REGISTRO DE ACTIVIDAD"
$LblLogTitulo.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$LblLogTitulo.ForeColor = $ColorAcento
$LblLogTitulo.AutoSize = $true
$LblLogTitulo.Location = New-Object System.Drawing.Point(25, 228)
$Form.Controls.Add($LblLogTitulo)

$TxtLog = New-Object System.Windows.Forms.RichTextBox
$TxtLog.Size = New-Object System.Drawing.Size(720, 250)
$TxtLog.Location = New-Object System.Drawing.Point(25, 250)
$TxtLog.BackColor = [System.Drawing.Color]::FromArgb(14, 14, 20)
$TxtLog.ForeColor = $ColorTexto
$TxtLog.Font = $FuenteMono
$TxtLog.ReadOnly = $true
$TxtLog.BorderStyle = "None"
$TxtLog.ScrollBars = "Vertical"
$Form.Controls.Add($TxtLog)

# ─── Barra de progreso ───
$ProgressBack = New-Object System.Windows.Forms.Panel
$ProgressBack.Size = New-Object System.Drawing.Size(720, 6)
$ProgressBack.Location = New-Object System.Drawing.Point(25, 508)
$ProgressBack.BackColor = $ColorPanel
$Form.Controls.Add($ProgressBack)

$ProgressBar = New-Object System.Windows.Forms.Panel
$ProgressBar.Size = New-Object System.Drawing.Size(0, 6)
$ProgressBar.Location = New-Object System.Drawing.Point(0, 0)
$ProgressBar.BackColor = $ColorAcento
$ProgressBack.Controls.Add($ProgressBar)

$LblEstado = New-Object System.Windows.Forms.Label
$LblEstado.Text = "Esperando..."
$LblEstado.Font = $FuenteEstado
$LblEstado.ForeColor = $ColorTextoSub
$LblEstado.Size = New-Object System.Drawing.Size(400, 25)
$LblEstado.Location = New-Object System.Drawing.Point(25, 522)
$Form.Controls.Add($LblEstado)

$LblContador = New-Object System.Windows.Forms.Label
$LblContador.Text = ""
$LblContador.Font = $FuenteEstado
$LblContador.ForeColor = $ColorTextoSub
$LblContador.AutoSize = $true
$LblContador.Location = New-Object System.Drawing.Point(630, 522)
$Form.Controls.Add($LblContador)

# ─── Botones inferiores ───
$BtnProcesar = New-Object System.Windows.Forms.Button
$BtnProcesar.Text = "▶  PROCESAR"
$BtnProcesar.Size = New-Object System.Drawing.Size(200, 48)
$BtnProcesar.Location = New-Object System.Drawing.Point(25, 558)
$BtnProcesar.FlatStyle = "Flat"
$BtnProcesar.FlatAppearance.BorderSize = 0
$BtnProcesar.BackColor = $ColorAcento
$BtnProcesar.ForeColor = [System.Drawing.Color]::White
$BtnProcesar.Font = $FuenteBoton
$BtnProcesar.Cursor = [System.Windows.Forms.Cursors]::Hand
$BtnProcesar.Enabled = $false
$Form.Controls.Add($BtnProcesar)

$BtnAbrir = New-Object System.Windows.Forms.Button
$BtnAbrir.Text = "📂  Abrir Carpeta LG"
$BtnAbrir.Size = New-Object System.Drawing.Size(165, 48)
$BtnAbrir.Location = New-Object System.Drawing.Point(240, 558)
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
$BtnBorrar.Text = "🗑  Borrar RAR"
$BtnBorrar.Size = New-Object System.Drawing.Size(140, 48)
$BtnBorrar.Location = New-Object System.Drawing.Point(420, 558)
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
$BtnSalir.Size = New-Object System.Drawing.Size(90, 48)
$BtnSalir.Location = New-Object System.Drawing.Point(655, 558)
$BtnSalir.FlatStyle = "Flat"
$BtnSalir.FlatAppearance.BorderSize = 1
$BtnSalir.FlatAppearance.BorderColor = $ColorBorde
$BtnSalir.BackColor = $ColorPanel
$BtnSalir.ForeColor = $ColorTextoSub
$BtnSalir.Font = $FuentePequena
$BtnSalir.Cursor = [System.Windows.Forms.Cursors]::Hand
$Form.Controls.Add($BtnSalir)

# ═══════════════════════════════════════════════════════════
# FUNCIONES
# ═══════════════════════════════════════════════════════════

# Variables compartidas
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
    $ProgressBar.Size = New-Object System.Drawing.Size($w, 6)
    if ($pct -ge 100) { $ProgressBar.BackColor = $ColorVerde }
    else { $ProgressBar.BackColor = $ColorAcento }
    $Form.Refresh()
}

function SetEstado($texto, $color) {
    $LblEstado.Text = $texto
    $LblEstado.ForeColor = $color
    $Form.Refresh()
}

# ─── Examinar ───
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
            Log "⚠  No se encontraron archivos SCTASK*.rar en:" $ColorAmarillo
            Log "   $($script:RarFolder)" $ColorTextoSub
        } else {
            $canRun = $sslOk -and $zipOk
            $BtnProcesar.Enabled = $canRun
            if (!$canRun) {
                SetEstado "Faltan dependencias (ver arriba)" $ColorRojo
            } else {
                SetEstado "Listo para procesar" $ColorVerde
            }
            $TxtLog.Clear()
            Log "📁  Carpeta seleccionada: $($script:RarFolder)" $ColorTexto
            Log "   Archivos encontrados: $($script:RarFiles.Count)" $ColorAcento
            Log "" $ColorTexto
            foreach ($r in $script:RarFiles) {
                Log "   • $($r.Name)  ($([math]::Round($r.Length/1KB, 1)) KB)" $ColorTextoSub
            }
            $LblContador.Text = "0 / $($script:RarFiles.Count)"
        }
    }
})

# ─── Hover effects ───
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

# ─── PROCESAR ───
$BtnProcesar.Add_Click({
    $BtnProcesar.Enabled = $false
    $BtnExaminar.Enabled = $false
    $BtnAbrir.Enabled = $false
    $BtnBorrar.Enabled = $false
    $TxtLog.Clear()
    SetProgress 0

    $Dest   = Join-Path $script:RarFolder "Certificados"
    $script:LGDest = Join-Path $Dest "LG Certificates"
    if (!(Test-Path $Dest))          { New-Item -ItemType Directory -Path $Dest | Out-Null }
    if (!(Test-Path $script:LGDest)) { New-Item -ItemType Directory -Path $script:LGDest | Out-Null }

    $total      = $script:RarFiles.Count
    $procesados = 0
    $errores    = 0
    $i          = 0

    Log "🚀  Iniciando procesamiento de $total archivo(s)..." $ColorAcento
    Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" $ColorBorde
    Log "" $ColorTexto

    foreach ($Rar in $script:RarFiles) {
        $i++
        $pctBase = (($i - 1) / $total) * 100
        $LblContador.Text = "$i / $total"
        SetEstado "Procesando: $($Rar.Name)" $ColorAcento

        Log "[$i/$total] 📦  $($Rar.Name)" $ColorAcento

        # Extraer
        $TmpDir = Join-Path $script:RarFolder "temp_$([System.IO.Path]::GetRandomFileName())"
        New-Item -ItemType Directory -Path $TmpDir | Out-Null

        & $SevenZip x $Rar.FullName "-o$TmpDir" "-p$PassRAR" -y 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Log "   ❌  Error al extraer RAR" $ColorRojo
            Log "" $ColorTexto
            $errores++
            Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
            SetProgress (($i / $total) * 100)
            continue
        }
        Log "   ✅  Extraido correctamente" $ColorVerde

        $CerFile = Get-ChildItem -Path $TmpDir -Filter "*.cer" -Recurse | Select-Object -First 1
        if (!$CerFile) {
            Log "   ❌  No se encontro .cer" $ColorRojo
            Log "" $ColorTexto
            $errores++
            Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
            SetProgress (($i / $total) * 100)
            continue
        }

        $CerName = $CerFile.BaseName
        Log "   📄  Certificado: $CerName" $ColorTexto
        $CertDir = Join-Path $Dest $CerName
        if (!(Test-Path $CertDir)) { New-Item -ItemType Directory -Path $CertDir | Out-Null }

        Get-ChildItem -Path $TmpDir -Recurse -File |
            Where-Object { $_.Extension -notin '.rsp','.req' } |
            ForEach-Object { Move-Item $_.FullName -Destination $CertDir -Force }
        Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue

        SetProgress ($pctBase + (1/$total)*20)

        # Verificar archivos
        $CerPath    = Join-Path $CertDir "$CerName.cer"
        $KeyPath    = Join-Path $CertDir "$CerName.key"
        $IssuingCrt = Join-Path $CertDir "Iberdrola Issuing CA v3.crt"
        $RootCrt    = Join-Path $CertDir "Iberdrola Root CA v3.crt"

        $ok = $true
        if (!(Test-Path $CerPath))    { Log "   ❌  Falta .cer" $ColorRojo; $ok = $false }
        if (!(Test-Path $IssuingCrt)) { Log "   ❌  Falta Issuing CA v3.crt" $ColorRojo; $ok = $false }
        if (!(Test-Path $RootCrt))    { Log "   ❌  Falta Root CA v3.crt" $ColorRojo; $ok = $false }
        if (!$ok) { Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue }

        # Generar .pfx
        $PfxPath = Join-Path $CertDir "$CerName.pfx"
        if (Test-Path $KeyPath) {
            Log "   ⚙  Generando .pfx ..." $ColorTextoSub
            & $OpenSSL pkcs12 -export -out $PfxPath -inkey $KeyPath -in $CerPath -passout "pass:$PassPFX" 2>&1 | Out-Null
            if (!(Test-Path $PfxPath)) {
                Log "   ❌  Error al generar .pfx" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
            }
            Log "   ✅  .pfx generado" $ColorVerde
        } else {
            Log "   ⚠  Sin .key - buscando .pfx existente..." $ColorAmarillo
            if (!(Test-Path $PfxPath)) {
                Log "   ❌  No hay .key ni .pfx" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
            }
        }

        SetProgress ($pctBase + (1/$total)*40)

        # Convertir .crt a .pem
        $IssuingPem = Join-Path $CertDir "Iberdrola Issuing CA v3.pem"
        $RootPem    = Join-Path $CertDir "Iberdrola Root CA v3.pem"
        Log "   ⚙  Convirtiendo .crt → .pem ..." $ColorTextoSub
        & $OpenSSL x509 -in $IssuingCrt -outform PEM -out $IssuingPem 2>&1 | Out-Null
        & $OpenSSL x509 -in $RootCrt    -outform PEM -out $RootPem    2>&1 | Out-Null
        if (!(Test-Path $IssuingPem) -or !(Test-Path $RootPem)) {
            Log "   ❌  Error en conversion .crt → .pem" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   ✅  .pem generados" $ColorVerde

        SetProgress ($pctBase + (1/$total)*60)

        # Carpeta LG
        $LGDir = Join-Path $script:LGDest $CerName
        if (!(Test-Path $LGDir)) { New-Item -ItemType Directory -Path $LGDir | Out-Null }

        # ca_certificate.pem
        Log "   ⚙  Generando ca_certificate.pem ..." $ColorTextoSub
        $IssContent  = (Get-Content $IssuingPem -Raw).TrimEnd("`r`n")
        $RootContent = (Get-Content $RootPem    -Raw).TrimEnd("`r`n")
        [System.IO.File]::WriteAllText((Join-Path $LGDir "ca_certificate.pem"),
            "$IssContent`r`n$RootContent`r`n", [System.Text.UTF8Encoding]::new($false))
        Log "   ✅  ca_certificate.pem" $ColorVerde

        SetProgress ($pctBase + (1/$total)*75)

        # client_certificate.pem
        Log "   ⚙  Generando client_certificate.pem ..." $ColorTextoSub
        $ClientCert = Join-Path $LGDir "client_certificate.pem"
        & $OpenSSL pkcs12 -in $PfxPath -clcerts -nokeys -out $ClientCert -passin "pass:$PassPFX" 2>&1 | Out-Null
        if (!(Test-Path $ClientCert)) {
            Log "   ❌  Error client_certificate.pem" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   ✅  client_certificate.pem" $ColorVerde

        SetProgress ($pctBase + (1/$total)*88)

        # client_key.pem
        Log "   ⚙  Generando client_key.pem ..." $ColorTextoSub
        $ClientKey = Join-Path $LGDir "client_key.pem"
        & $OpenSSL pkcs12 -in $PfxPath -nocerts -nodes -out $ClientKey -passin "pass:$PassPFX" 2>&1 | Out-Null
        if (!(Test-Path $ClientKey)) {
            Log "   ❌  Error client_key.pem" $ColorRojo; Log "" $ColorTexto; $errores++; SetProgress (($i / $total) * 100); continue
        }
        Log "   ✅  client_key.pem" $ColorVerde

        # Verificar completo
        $LGOk = @("ca_certificate.pem","client_certificate.pem","client_key.pem") |
            Where-Object { Test-Path (Join-Path $LGDir $_) }
        if ($LGOk.Count -eq 3) {
            Log "   ══════════════════════════════════════════" $ColorVerde
            Log "   ✅  CARPETA LG COMPLETA: $CerName" $ColorVerde
            Log "   ══════════════════════════════════════════" $ColorVerde
            $procesados++
        } else {
            Log "   ⚠  Carpeta LG incompleta" $ColorAmarillo
            $errores++
        }
        Log "" $ColorTexto
        SetProgress (($i / $total) * 100)
    }

    # Resumen final
    Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" $ColorBorde
    Log "" $ColorTexto
    Log "📊  RESUMEN: $procesados / $total procesados correctamente" $ColorAcento
    if ($errores -gt 0) { Log "⚠   Errores: $errores" $ColorRojo }
    Log "📂  LG Certificates: $($script:LGDest)" $ColorTextoSub

    SetProgress 100
    $LblContador.Text = "$procesados / $total ✓"

    if ($errores -eq 0) {
        SetEstado "✅  Completado sin errores" $ColorVerde
    } else {
        SetEstado "⚠  Completado con $errores error(es)" $ColorAmarillo
    }

    $BtnAbrir.Enabled = $true
    $BtnBorrar.Enabled = $true
    $BtnExaminar.Enabled = $true
})

# ─── Abrir carpeta ───
$BtnAbrir.Add_Click({
    if ($script:LGDest -and (Test-Path $script:LGDest)) {
        Start-Process explorer.exe $script:LGDest
    }
})

# ─── Borrar RAR ───
$BtnBorrar.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "¿Eliminar $($script:RarFiles.Count) archivo(s) RAR originales?`n`nEsta accion no se puede deshacer.",
        "Confirmar eliminacion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -eq 'Yes') {
        $script:RarFiles | Remove-Item -Force
        Log "" $ColorTexto
        Log "🗑  Archivos RAR eliminados." $ColorRojo
        $BtnBorrar.Enabled = $false
    }
})

# ─── Salir ───
$BtnSalir.Add_Click({ $Form.Close() })

# ─── Mostrar ───
$Form.ShowDialog() | Out-Null
