# Build stats_plots.oxt from sources.
#
#   - Copies the static extension layout (extension/) into a staging dir.
#   - Generates the Basic module (stats_plots/Module1.xba) from the canonical
#     src/Module1.bas, XML-escaping the source so it is the single source of truth.
#   - Zips the staging dir into stats_plots.oxt.
$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$staging = Join-Path $root 'build'
$oxt     = Join-Path $root 'stats_plots.oxt'

if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging | Out-Null

# Static extension layout (description.xml, META-INF/, stats_plots/script.xlb).
Copy-Item (Join-Path $root 'extension\*') $staging -Recurse -Force

# Generate Module1.xba from the canonical .bas source.
$src = Get-Content (Join-Path $root 'src\Module1.bas') -Raw
$escaped = $src -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
$header = @'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
'@
$xba = "$header`r`n$escaped`r`n</script:module>`r`n"
Set-Content -Path (Join-Path $staging 'stats_plots\Module1.xba') -Value $xba -Encoding UTF8

# Package as .oxt (a plain zip).
if (Test-Path $oxt) { Remove-Item $oxt -Force }
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath "$oxt.zip" -Force
Move-Item "$oxt.zip" $oxt -Force

Write-Host "Built $oxt"
