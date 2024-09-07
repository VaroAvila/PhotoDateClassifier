# Necesario para trabajar con la clase Image y leer los metadatos EXIF
Add-Type -AssemblyName System.Drawing

# Script de PowerShell para clasificar archivos en carpetas por fecha de captura con carpetas enumeradas y mensajes de depuración

Write-Host "Iniciando el script..." -ForegroundColor Yellow

# Seleccionar la carpeta
$folder = Read-Host "Introduce la ruta de la carpeta que deseas organizar"
Write-Host "Ruta seleccionada: $folder" -ForegroundColor Yellow

# Comprobar si la carpeta existe
if (!(Test-Path -Path $folder)) {
    Write-Host "La carpeta no existe. Por favor, introduce una ruta válida." -ForegroundColor Red
    exit
}

Write-Host "Carpeta encontrada. Comprobando y creando carpetas de meses enumeradas si no existen..." -ForegroundColor Yellow

# Crear carpetas de meses enumeradas si no existen
$months = @("1_Enero", "2_Febrero", "3_Marzo", "4_Abril", "5_Mayo", "6_Junio", "7_Julio", "8_Agosto", "9_Septiembre", "10_Octubre", "11_Noviembre", "12_Diciembre")
foreach ($month in $months) {
    $monthFolder = Join-Path -Path $folder -ChildPath $month
    if (!(Test-Path -Path $monthFolder)) {
        New-Item -Path $monthFolder -ItemType Directory | Out-Null
        Write-Host "Carpeta creada: $monthFolder" -ForegroundColor Green
    } else {
        Write-Host "Carpeta ya existe: $monthFolder" -ForegroundColor Yellow
    }
}

Write-Host "Obteniendo archivos en la carpeta especificada..." -ForegroundColor Yellow

# Obtener todos los archivos de la carpeta
$files = Get-ChildItem -Path $folder -File
$totalFiles = $files.Count
Write-Host "$totalFiles archivos encontrados en la carpeta." -ForegroundColor Yellow

if ($totalFiles -eq 0) {
    Write-Host "No hay archivos para organizar en la carpeta especificada." -ForegroundColor Yellow
    exit
}

# Función para obtener la fecha de captura de imágenes utilizando EXIF
function Get-ImageCaptureDate {
    param (
        [string]$imagePath
    )

    try {
        $image = [System.Drawing.Image]::FromFile($imagePath)
        $property = $image.GetPropertyItem(36867) # 36867 es el ID del tag de fecha de captura (DateTaken) en EXIF
        $dateTaken = [System.Text.Encoding]::ASCII.GetString($property.Value)
        $dateTaken = $dateTaken.TrimEnd([char]0) # Eliminar caracteres nulos
        $image.Dispose() # Liberar recurso
        return [datetime]::ParseExact($dateTaken, "yyyy:MM:dd HH:mm:ss", $null)
    } catch {
        return $null
    }
}

# Función para obtener la fecha "Medio Creado" de videos
function Get-VideoCreationDate {
    param (
        [string]$videoPath
    )

    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Get-Item $videoPath).DirectoryName)
        $file = $folder.ParseName((Get-Item $videoPath).Name)
        # Índice 277 corresponde a "Medio Creado" en Windows
        $mediaCreated = $folder.GetDetailsOf($file, 277)
        if ($mediaCreated) {
            return [datetime]::Parse($mediaCreated)
        } else {
            return $null
        }
    } catch {
        return $null
    }
}

# Inicializar contador de progreso
$counter = 0

Write-Host "Moviendo archivos a las carpetas correspondientes..." -ForegroundColor Yellow

# Mover archivos a las carpetas correspondientes
foreach ($file in $files) {
    $captureDate = $null

    # Intentar obtener la fecha de captura para imágenes
    if ($file.Extension -match "\.(jpg|jpeg|png|tiff|heic)$") {
        $captureDate = Get-ImageCaptureDate -imagePath $file.FullName
    }
    # Intentar obtener la fecha "Medio Creado" para videos
    elseif ($file.Extension -match "\.(mp4|mov|avi|mkv|wmv)$") {
        $captureDate = Get-VideoCreationDate -videoPath $file.FullName
    }
    
    # Si no se pudo obtener la fecha de captura, usar la fecha de modificación como último recurso
    if (-not $captureDate) {
        $captureDate = $file.LastWriteTime
    }

    Write-Host "Archivo: $($file.Name) - Fecha de captura: $captureDate" -ForegroundColor Cyan

    # Extraer el número del mes de la fecha de captura
    $captureMonthNumber = $captureDate.Month
    $captureMonth = $months[$captureMonthNumber - 1] # Restar 1 porque los arrays empiezan en 0

    Write-Host "Archivo: $($file.Name) - Mes de captura: $captureMonth" -ForegroundColor Cyan

    $destinationFolder = Join-Path -Path $folder -ChildPath $captureMonth

    # Verificar si el archivo no está en uso antes de moverlo
    try {
        Move-Item -Path $file.FullName -Destination $destinationFolder -ErrorAction Stop
        $counter++
        $progress = [math]::Round(($counter / $totalFiles) * 100)
        Write-Host "[$progress%] Archivo movido: $($file.Name) -> $captureMonth" -ForegroundColor Cyan
    } catch {
        Write-Host "Error moviendo el archivo: $($file.Name). Error: $_" -ForegroundColor Red
    }
}

Write-Host "Archivos organizados en carpetas por mes de captura." -ForegroundColor Green
