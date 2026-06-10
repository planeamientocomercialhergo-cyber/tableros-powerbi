# ============================================================
#  Actualizar el tablero PowerBI
#  Sube a GitHub los cambios (ej. ediciones de LINKS POWER BI.xlsx).
#  Vercel detecta el push y republica la web sola en unos segundos.
#  Uso: clic derecho > "Ejecutar con PowerShell"  (o correrlo desde la terminal)
# ============================================================
Set-Location -LiteralPath $PSScriptRoot

$cambios = git status --porcelain
if (-not $cambios) {
    Write-Host "No hay cambios para subir. Ya esta todo actualizado." -ForegroundColor Yellow
    Read-Host "Enter para cerrar"
    return
}

Write-Host "Cambios detectados:" -ForegroundColor Cyan
git status -s

git add -A
$fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "Actualizo tableros ($fecha)"
git push

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "OK. Subido a GitHub. Vercel publicara la version nueva en unos segundos." -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Hubo un problema en el push. Revisa el mensaje de arriba." -ForegroundColor Red
}
Read-Host "Enter para cerrar"
