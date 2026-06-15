<#
.SYNOPSIS
    Lance Excel et charge l'add-in ExcelDna packe via automation COM (RegisterXLL).

.DESCRIPTION
    Passer un .xll en argument a excel.exe ne le charge que pour la session ET
    declenche une boite de securite (qui, si non validee, empeche le chargement).
    On utilise donc l'API COM RegisterXLL : chargement propre, sans prompt, verifiable.

    - Cible le .xll packe autonome 64 bits : ExcelDnaPoc-AddIn64-packed.xll
    - Debloque le fichier (MOTW) au cas ou.
    - Ouvre Excel (visible), cree un classeur, appelle RegisterXLL et verifie le retour.
    - L'onglet "POC ExcelDna" apparait dans le ruban.

.PARAMETER Configuration
    Debug (defaut) ou Release.

.EXAMPLE
    .\Launch-AddIn.ps1
    .\Launch-AddIn.ps1 -Configuration Release
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$xll  = Join-Path $root "bin\x64\$Configuration\net8.0-windows\publish\ExcelDnaPoc-AddIn64-packed.xll"

if (-not (Test-Path $xll)) {
    Write-Error @"
Add-in introuvable :
  $xll

Compile d'abord le projet :
  dotnet build -c $Configuration
"@
    exit 1
}

try { Unblock-File -Path $xll } catch { }

Write-Host "Chargement de l'add-in dans Excel (COM RegisterXLL) :" -ForegroundColor Cyan
Write-Host "  $xll" -ForegroundColor Gray

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
[void]$excel.Workbooks.Add()

$ok = $excel.RegisterXLL($xll)
if ($ok) {
    Write-Host "OK - add-in charge. Cherche l'onglet 'POC ExcelDna' dans le ruban." -ForegroundColor Green
} else {
    Write-Warning "RegisterXLL a retourne False : l'add-in n'a pas pu etre charge."
}

# On relache la reference COM ; Excel reste ouvert (Visible = $true).
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
