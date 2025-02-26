# Definir el directorio de inicio
param (
    [string]$rutaInicial
)

# Comprobar si la ruta es válida
if (-Not (Test-Path $rutaInicial)) {
    Write-Host "La ruta especificada no existe: $rutaInicial" -ForegroundColor Red
    exit 1
}

# Obtener archivos .xls recursivamente
$archivosXLS = Get-ChildItem -Path $rutaInicial -Recurse -Filter "*.xls" | Where-Object { $_.Extension -eq ".xls" }

# Comprobar si hay archivos para convertir
if ($archivosXLS.Count -eq 0) {
    Write-Host "No se encontraron archivos .xls en el directorio especificado." -ForegroundColor Yellow
    exit 0
}

# Iniciar Excel
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
} catch {
    Write-Host "Error al iniciar Excel. Asegúrate de que está instalado." -ForegroundColor Red
    exit 1
}

foreach ($archivo in $archivosXLS) {
    $rutaXLS = $archivo.FullName
    $rutaXLSX = $rutaXLS -replace "\.xls$", ".xlsx"
    
    Write-Host "Procesando: $rutaXLS" -ForegroundColor Cyan
    
    try {
        # Abrir el archivo en Excel
        $libro = $excel.Workbooks.Open($rutaXLS)
        
        # Guardar como .xlsx
        $libro.SaveAs($rutaXLSX, 51)  # 51 es el formato para xlsx
        $libro.Close($false)
        
        # Copiar atributos y permisos
        $acl = Get-Acl $rutaXLS
        Set-Acl -Path $rutaXLSX -AclObject $acl
        
        # Copiar fechas de modificación y creación
        (Get-Item $rutaXLSX).CreationTime = (Get-Item $rutaXLS).CreationTime
        (Get-Item $rutaXLSX).LastWriteTime = (Get-Item $rutaXLS).LastWriteTime
        
        # Eliminar el archivo original
        Remove-Item $rutaXLS -Force
        
        Write-Host "Convertido y eliminado: $rutaXLS" -ForegroundColor Green
    } catch {
        Write-Host "Error al procesar: $rutaXLS - $_" -ForegroundColor Red
    }
}

# Cerrar Excel
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Proceso completado." -ForegroundColor Green
