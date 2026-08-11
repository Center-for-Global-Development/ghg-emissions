<#
.SYNOPSIS
figA.ps1
Generate Figure_3.xlsx -- static 1850-2024 cumulative dumbbell.

Inputs:
  data/outputs/charts/fig3_dumbbell.csv, dumbbell_raw_*.csv, dumbbell_gmst.csv
  cgd-general/chart-templates/Chart-template-Dumbbell.xlsx

Output:
  data/outputs/charts/Figure_3.xlsx

Pre-req: scripts 01, 07, 09 must have run.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir     = "$QADir\data\outputs\charts"
$outDir     = "$QADir\data\outputs\charts"
$TestPrefix = ""

$MEASURES = @("co2_excLUC","co2_incLUC","ghg_excLUC","ghg_incLUC")
$MLABEL   = @{
    "co2_excLUC" = "CO2 excl LULUCF"
    "co2_incLUC" = "CO2 incl LULUCF"
    "ghg_excLUC" = "GHG excl LULUCF"
    "ghg_incLUC" = "GHG incl LULUCF"
}
$CHART_TITLES = @{
    "fig3" = "Emissions shares and contribution to increase in GMST (1850$([char]0x2013)2024)"
}

. "$PSScriptRoot\_chart_helpers.ps1"

$cgdTeal      = 5983243
$cgdAmber     = 2930175
$cgdTealMid   = 10390042
$cgdAmberDark = 36556
$cgdTealLight = 13679965
$cgdTealMuted = 9866331
$cgdConnLine  = 14407080
$cgdNearBlack = 3355443

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

try {

. "$PSScriptRoot\_chart_dumbbell.ps1"

Write-Host "`nFig 3: Dumbbell 1850-2024 ..."
Build-Dumbbell "fig3_dumbbell.csv" "Figure_3.xlsx" $CHART_TITLES["fig3"]

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
