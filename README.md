# Certificados LG - Generador automatico

**Pentatel S.L.** | Disenado por Eduardo Rubio

---

## Requisitos

- **OpenSSL** instalado (carpeta con `openssl.exe` o subcarpeta `bin\openssl.exe`)
- **7-Zip** instalado en `C:\Program Files\7-Zip\7z.exe`
- **PowerShell** (Windows Terminal recomendado)

---

## Instrucciones de uso

### 1. Abre la terminal en la carpeta de OpenSSL

Navega a la carpeta donde se encuentra `openssl.exe`. Puedes hacer clic derecho dentro de la carpeta y seleccionar **"Abrir en Terminal"**, o abrir PowerShell y hacer `cd` hasta la ruta correspondiente.

> Es importante que la terminal se abra desde esta ubicacion, ya que el script genera la carpeta `LG Certificates` en el directorio desde el que se ejecuta.

### 2. Asegurate de conservar la carpeta `Certificados`

La carpeta `Certificados` que se genero durante la creacion de las solicitudes de IT NOW **debe estar en el mismo directorio** donde vas a ejecutar el script. Dentro de ella estan las subcarpetas de cada hostname con sus claves privadas (`.key`), y el script las necesita para generar los archivos finales.

La estructura esperada seria algo asi:

```
Tu carpeta de trabajo/
    openssl.exe
    Certificados/
        HOSTNAME_1/
            HOSTNAME_1.key
            HOSTNAME_1.csr
        HOSTNAME_2/
            HOSTNAME_2.key
            HOSTNAME_2.csr
        ...
```

### 3. Ten los archivos RAR a mano

Los archivos `SCTASK*.rar` que recibiste por correo de Iberdrola pueden estar en cualquier carpeta (por ejemplo, en `Descargas`). **No es necesario que esten en la carpeta de OpenSSL**. Lo unico importante es que todos los `.rar` que quieras procesar esten juntos en una misma carpeta.

### 4. Copia y pega el script

Abre el archivo `Certificados_LG.ps1`, selecciona **todo el contenido** (Ctrl+A), copialo (Ctrl+C) y pegalo directamente en la terminal de PowerShell (Ctrl+V). El script se ejecutara automaticamente.

### 5. Sigue las indicaciones en pantalla

1. El script verificara que OpenSSL y 7-Zip estan disponibles.
2. Te mostrara las instrucciones y esperara a que pulses **Enter**.
3. Se abrira un selector de carpeta: **navega hasta donde tengas los archivos SCTASK*.rar** y seleccionala. El script detectara automaticamente todos los archivos validos, aunque haya otros archivos distintos en esa misma carpeta.
4. Procesara cada archivo RAR y generara los certificados.
5. Al finalizar, mostrara un resumen con la relacion entre cada SCTASK y su hostname.
6. Te preguntara si quieres abrir la carpeta de resultados y si deseas borrar los RAR originales.

### 6. Resultado

Se creara una carpeta `LG Certificates` en el directorio de trabajo con una subcarpeta por cada hostname, conteniendo los tres archivos necesarios para los monitores LG:

```
LG Certificates/
    HOSTNAME_1/
        ca_certificate.pem
        client_certificate.pem
        client_key.pem
    HOSTNAME_2/
        ca_certificate.pem
        client_certificate.pem
        client_key.pem
    ...
```

Estas carpetas son las que se suben a la plataforma para que los tecnicos de instalacion importen los certificados en los monitores.

---

## Notas

- Si usas la **consola clasica de PowerShell** (ventana azul), es posible que al pegar el script necesites pulsar Enter manualmente. Se recomienda usar **Windows Terminal**.
- Si un hostname da error por falta de `.key`, revisa que la carpeta `Certificados` contenga la subcarpeta correspondiente con el archivo `.key` generado durante la solicitud.
