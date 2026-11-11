$ErrorActionPreference = "Stop"
$Manifest = Join-Path $PSScriptRoot "..\manifest.xml"
Write-Host "Starting Decken for Microsoft PowerPoint desktop debugging session..."
npx --yes office-addin-debugging start $Manifest desktop
