<#
.SYNOPSIS
    Regenere ChuckMacro.xlam a partir de ChuckMacro.bas (Excel cache + injection VBA).

.DESCRIPTION
    ChuckMacro.bas est la SOURCE (texte, versionnee) ; ChuckMacro.xlam est un ARTEFACT
    DERIVE (non versionne, cf .gitignore). Ce script le (re)genere :
      - ouvre une instance Excel CACHEE (distincte de celle de l'utilisateur),
      - importe le module .bas dans un nouveau classeur (VBProject),
      - l'enregistre au format complement .xlam (FileFormat 55 = xlOpenXMLAddIn),
      - via %TEMP% puis Copy-Item (SaveAs direct vers OneDrive echoue).

    Appele AUTOMATIQUEMENT a chaque build (cible MSBuild GenerateChuckXlam) et par
    Launch-AddIn.ps1.

    Necessite Excel installe + "Acces approuve au modele objet du projet VBA"
    (Centre de gestion de la confidentialite). En cas d'echec (pas d'Excel, droit
    absent...), AVERTIT sans casser le build.
#>
[CmdletBinding()]
param()

$root = $PSScriptRoot
$bas  = Join-Path $root 'ChuckMacro.bas'
$dest = Join-Path $root 'ChuckMacro.xlam'
$tmp  = Join-Path $env:TEMP 'ChuckMacro.xlam'

if (-not (Test-Path $bas)) {
    Write-Warning "Generate-Xlam : ChuckMacro.bas introuvable -> .xlam non genere."
    return
}

$xl = $null
try {
    if (Test-Path $tmp) { Remove-Item $tmp -Force }

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import($bas)  # requiert AccessVBOM
    $wb.SaveAs($tmp, 55)                           # 55 = xlOpenXMLAddIn (.xlam)
    $wb.Close($false)

    Copy-Item $tmp $dest -Force
    Write-Host "Generate-Xlam : ChuckMacro.xlam regenere depuis ChuckMacro.bas." -ForegroundColor Gray
}
catch {
    # Ne casse pas le build : pas d'Excel, AccessVBOM desactive, COM occupe, etc.
    Write-Warning ("Generate-Xlam : generation ignoree (" + $_.Exception.Message + ")")
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
}
