# Exporta la presentacion con PowerPoint (COM, solo rutas): un PNG por diapositiva y, opcionalmente, el PDF.
# Uso:  powershell -NoProfile -File scripts\exportar_presentacion.ps1 -Pptx <ruta.pptx> -CarpetaPng <carpeta> [-Pdf <ruta.pdf>]
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$CarpetaPng,
    [string]$Pdf = ""
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $CarpetaPng | Out-Null
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($Pptx, -1, 0, 0)   # solo lectura, sin ventana
    $pres.Export($CarpetaPng, "PNG", 1600, 900)
    if ($Pdf -ne "") { $pres.SaveAs($Pdf, 32) }         # 32 = ppSaveAsPDF
    $pres.Close()
} finally {
    $app.Quit()
}
Write-Output "listo"
