<#
.SYNOPSIS
figB.ps1
Generate Figure_4.xlsx -- cumulative shares dumbbell, 1990–2024.

Input:
  data/outputs/charts/fig4_dumbbell.csv
  cgd-general/chart-templates/Chart-template-Dumbbell.xlsx

Output:
  data/outputs/charts/Figure_4.xlsx

Pre-req: scripts 01–09 must have run to produce fig4_dumbbell.csv.
#>

param(
    [string]$QADir   = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$TmplDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "templates")
)

$csvDir = "$QADir\data\outputs\charts"
$outDir = "$QADir\data\outputs\charts"

$cgdTeal      = 5983243
$cgdAmber     = 2930175
$cgdTealMid   = 10390042
$cgdAmberDark = 36556
$cgdTealLight = 13679965
$cgdTealMuted = 9866331
$cgdOlive     = 4489605
$cgdConnLine  = 14407080
$cgdNearBlack = 3355443

. "$PSScriptRoot\_chart_helpers.ps1"
. "$PSScriptRoot\_chart_dumbbell.ps1"

Write-Host "Starting Excel ..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible          = $false
$xl.DisplayAlerts    = $false
$xl.AskToUpdateLinks = $false
try { $xl.AutomationSecurity = 3 } catch {}

$em = [char]0x2013

try {

Write-Host "`nFigure 4: Dumbbell 1990-2024 ..."
Build-Dumbbell "fig4_dumbbell.csv" "Figure_4.xlsx" `
    "Emissions shares and contribution to increase in GMST (1990${em}2024)"

} catch {
    Write-Host "ERROR: $_"
    Write-Host $_.ScriptStackTrace
} finally {
    Write-Host "`nClosing Excel ..."
    try { $xl.Quit() } catch {}
    Release-COM $xl
}

Write-Host "Done."
