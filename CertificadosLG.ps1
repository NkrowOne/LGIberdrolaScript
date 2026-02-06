& {
    Add-Type -AssemblyName System.Windows.Forms
    $7z = "C:\Program Files\7-Zip\7z.exe"
    $pr = 'Larraskitu48!'
    $pp = 'DIGITALSIGNAGE'
    $rt = (Get-Location).Path
    $ssl = Join-Path $rt "openssl.exe"
    if (!(Test-Path $ssl)) { $ssl = Join-Path $rt "bin\openssl.exe" }

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  Certificados LG  -  Pentatel S.L." -ForegroundColor White
    Write-Host "  Disenado por Eduardo Rubio" -ForegroundColor DarkGray
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""

    if (!(Test-Path $ssl)) { Write-Host "[!!] OpenSSL no encontrado en: $ssl" -ForegroundColor Red; return }
    else { Write-Host "[OK] OpenSSL: $ssl" -ForegroundColor Green }
    if (!(Test-Path $7z)) { Write-Host "[!!] 7-Zip no encontrado en: $7z" -ForegroundColor Red; return }
    else { Write-Host "[OK] 7-Zip:   $7z" -ForegroundColor Green }

    Write-Host ""
    Write-Host "ANTES DE CONTINUAR, tenga en cuenta:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. Todos los archivos SCTASK*.rar deben estar reunidos en" -ForegroundColor White
    Write-Host "     una misma carpeta antes de seleccionarla." -ForegroundColor White
    Write-Host ""
    Write-Host "  2. La carpeta 'Certificados' debe estar en el mismo directorio" -ForegroundColor White
    Write-Host "     desde el que se generaron las solicitudes de IT NOW, ya que" -ForegroundColor White
    Write-Host "     el script necesita leer la clave privada (.key) de cada" -ForegroundColor White
    Write-Host "     certificado desde sus subcarpetas." -ForegroundColor White
    Write-Host ""
    Write-Host "  3. Los archivos finales (.pem) se generaran dentro de" -ForegroundColor White
    Write-Host "     .\Certificados\<hostname>\LG Cert\" -ForegroundColor Gray
    Write-Host "     junto al .pfx en la carpeta padre." -ForegroundColor White
    Write-Host ""

    Read-Host "Pulsa ENTER para seleccionar la carpeta con los archivos SCTASK*.rar"

    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Carpeta con archivos SCTASK*.rar"
    $fb.ShowNewFolderButton = $false
    if ($fb.ShowDialog() -ne 'OK') { Write-Host "Cancelado." -ForegroundColor Red; return }

    $rars = Get-ChildItem -Path $fb.SelectedPath -Filter "SCTASK*.rar" -File
    if ($rars.Count -eq 0) { Write-Host "[!!] No hay SCTASK*.rar en esa carpeta." -ForegroundColor Red; return }

    Write-Host ""
    Write-Host "[OK] Carpeta: $($fb.SelectedPath)" -ForegroundColor Green
    Write-Host "[OK] Archivos encontrados: $($rars.Count)" -ForegroundColor Green
    Write-Host ""

    $cb = Join-Path $rt "Certificados"
    if (!(Test-Path $cb)) { New-Item -ItemType Directory -Path $cb | Out-Null }

    $ok = 0; $er = 0; $i = 0
    $mapa = @()

    foreach ($r in $rars) {
        $i++
        Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "[$i/$($rars.Count)] $($r.Name)" -ForegroundColor Cyan

        $tmp = Join-Path $fb.SelectedPath "tmp_$([IO.Path]::GetRandomFileName())"
        New-Item -ItemType Directory -Path $tmp | Out-Null
        & $7z x $r.FullName "-o$tmp" "-p$pr" -y 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Host "   [ERROR] Extraccion fallida" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname="(error extraccion)";Estado="ERROR"}; Remove-Item $tmp -Recurse -Force -EA 0; continue }
        Write-Host "   [OK] Extraido" -ForegroundColor Green

        $cer = Get-ChildItem -Path $tmp -Filter "*.cer" -Recurse | Select-Object -First 1
        if (!$cer) { Write-Host "   [ERROR] No hay .cer" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname="(sin .cer)";Estado="ERROR"}; Remove-Item $tmp -Recurse -Force -EA 0; continue }

        $nm = $cer.BaseName
        Write-Host "   Hostname: $nm" -ForegroundColor Gray
        $cd = Join-Path $cb $nm
        if (!(Test-Path $cd)) { New-Item -ItemType Directory -Path $cd | Out-Null }

        Get-ChildItem -Path $tmp -Recurse -File | Where-Object { $_.Extension -notin '.rsp','.req' } | ForEach-Object { Move-Item $_.FullName -Destination $cd -Force }
        Remove-Item $tmp -Recurse -Force -EA 0

        $cp = Join-Path $cd "$nm.cer"
        $kp = Join-Path $cd "$nm.key"
        $ic = Join-Path $cd "Iberdrola Issuing CA v3.crt"
        $rc = Join-Path $cd "Iberdrola Root CA v3.crt"

        $miss = $false
        if (!(Test-Path $cp)) { Write-Host "   [FALTA] $nm.cer" -ForegroundColor Red; $miss = $true }
        if (!(Test-Path $ic)) { Write-Host "   [FALTA] Iberdrola Issuing CA v3.crt" -ForegroundColor Red; $miss = $true }
        if (!(Test-Path $rc)) { Write-Host "   [FALTA] Iberdrola Root CA v3.crt" -ForegroundColor Red; $miss = $true }
        if ($miss) { $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }

        $pfx = Join-Path $cd "$nm.pfx"
        if (Test-Path $kp) {
            & $ssl pkcs12 -export -out $pfx -inkey $kp -in $cp -passout "pass:$pp" 2>&1 | Out-Null
            if (!(Test-Path $pfx)) { Write-Host "   [ERROR] .pfx" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }
            Write-Host "   [OK] .pfx generado" -ForegroundColor Green
        } else {
            Write-Host "   [AVISO] Sin .key, buscando .pfx existente..." -ForegroundColor Yellow
            if (!(Test-Path $pfx)) { Write-Host "   [ERROR] No hay .key ni .pfx" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }
        }

        $ip = Join-Path $cd "Iberdrola Issuing CA v3.pem"
        $rp = Join-Path $cd "Iberdrola Root CA v3.pem"
        & $ssl x509 -in $ic -outform PEM -out $ip 2>&1 | Out-Null
        & $ssl x509 -in $rc -outform PEM -out $rp 2>&1 | Out-Null
        if (!(Test-Path $ip) -or !(Test-Path $rp)) { Write-Host "   [ERROR] .crt a .pem" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }
        Write-Host "   [OK] .pem intermedios" -ForegroundColor Green

        $ld = Join-Path $cd "LG Cert"
        if (!(Test-Path $ld)) { New-Item -ItemType Directory -Path $ld | Out-Null }

        $ic2 = (Get-Content $ip -Raw).TrimEnd("`r`n")
        $rc2 = (Get-Content $rp -Raw).TrimEnd("`r`n")
        [IO.File]::WriteAllText((Join-Path $ld "ca_certificate.pem"), "$ic2`r`n$rc2`r`n", [Text.UTF8Encoding]::new($false))
        Write-Host "   [OK] ca_certificate.pem" -ForegroundColor Green

        & $ssl pkcs12 -in $pfx -clcerts -nokeys -out (Join-Path $ld "client_certificate.pem") -passin "pass:$pp" 2>&1 | Out-Null
        if (!(Test-Path (Join-Path $ld "client_certificate.pem"))) { Write-Host "   [ERROR] client_certificate.pem" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }
        Write-Host "   [OK] client_certificate.pem" -ForegroundColor Green

        & $ssl pkcs12 -in $pfx -nocerts -nodes -out (Join-Path $ld "client_key.pem") -passin "pass:$pp" 2>&1 | Out-Null
        if (!(Test-Path (Join-Path $ld "client_key.pem"))) { Write-Host "   [ERROR] client_key.pem" -ForegroundColor Red; $er++; $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="ERROR"}; continue }
        Write-Host "   [OK] client_key.pem" -ForegroundColor Green

        Write-Host "   [COMPLETO] $nm (3/3)" -ForegroundColor Green
        $ok++
        $mapa += [PSCustomObject]@{SCTASK=$r.BaseName;Hostname=$nm;Estado="OK"}
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  RESUMEN: $ok / $($rars.Count) procesados" -ForegroundColor Cyan
    if ($er -gt 0) { Write-Host "  Errores: $er" -ForegroundColor Red }
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  SCTASK                   HOSTNAME                                    ESTADO" -ForegroundColor White
    Write-Host "  ------------------------ ------------------------------------------- ------" -ForegroundColor DarkGray
    foreach ($m in $mapa) {
        $sc = $m.SCTASK.PadRight(24)
        $hn = $m.Hostname.PadRight(43)
        if ($m.Estado -eq "OK") { Write-Host "  $sc $hn" -ForegroundColor Gray -NoNewline; Write-Host " OK" -ForegroundColor Green }
        else { Write-Host "  $sc $hn" -ForegroundColor Gray -NoNewline; Write-Host " ERROR" -ForegroundColor Red }
    }
    Write-Host ""
    Write-Host "  Certificados: $cb" -ForegroundColor DarkGray
    Write-Host ""

    $op = Read-Host "Abrir carpeta Certificados? (s/n)"
    if ($op -eq 's') { Start-Process explorer.exe $cb }

    $op2 = Read-Host "Borrar archivos RAR originales? (s/n)"
    if ($op2 -eq 's') { $rars | Remove-Item -Force; Write-Host "RAR eliminados." -ForegroundColor Red }
}
