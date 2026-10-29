# Genera web\index.html (versión instalable) a partir de App.html (fuente única).
# Uso: .\build.ps1   y luego publique la carpeta "web" en un hosting HTTPS (GitHub Pages, Netlify, Firebase...).
$src = Join-Path $PSScriptRoot 'App.html'
$dst = Join-Path $PSScriptRoot 'web\index.html'
$html = [IO.File]::ReadAllText($src, [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($dst, $html, (New-Object Text.UTF8Encoding $false))
Write-Host "Listo: $dst"
