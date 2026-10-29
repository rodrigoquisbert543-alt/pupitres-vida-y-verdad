# Versión instalable (PWA)

1. Edite `App.html` (única fuente) y ejecute `.\build.ps1` → actualiza `web\index.html`.
2. Publique la carpeta `web` en un hosting **HTTPS** (GitHub Pages, Netlify, Firebase Hosting…).
3. En `App.html` ponga esa URL en `PWA_URL` (así el botón “Instalar” también funciona dentro de Apps Script) y use esa misma URL en Ajustes → “URL de la app” para los QR.
4. Abra la URL en Chrome/Edge (PC, Android) o Safari (iPhone) y pulse **Instalar aplicación** (en el login, la cabecera o Ajustes).

`clasp push` ignora `web/` (ver `.claspignore`).
