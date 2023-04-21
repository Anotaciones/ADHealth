Param(
  [Parameter(Mandatory=$true)][string]$ReportsLocation
)

import-module .\HelperFunctions\HelperFunctions.psd1 -Force -DisableNameChecking -WarningAction SilentlyContinue
$Reports = Get-ChildItem -Path $ReportsLocation -Filter ad_hc_*.xml -Recurse

Write-Host "[*] Collecting rules details."
$rulesData = Import-Clixml -Path .\Resources\RulesDescription.xml

Write-Host "[*] Collecting risks details."
$RiskRules = $Reports | %{
  $hc = $_
  $hc_report = $hc.FullName
  $fname = Split-Path $hc_report -leaf
  $Domain = $fname.Replace('ad_hc_','').replace('.xml', '')
  Write-Host ("`t[*] Parsing '{0}' Report:" -f $hc.Name.Replace("ad_hc_", "").Replace(".xml", "") )
  $reportData = ParseSingleReport -File $hc_report
  [PSCustomObject]@{
    Domain = $Domain
    Report = $reportData
  }
}

Write-Host "[*] Generating Excel template."
$Template = Generate-ExcelTemplate -ReportsCount $Reports.Count

Write-Host "[*] Filling rules into Excel template."
$foundRules = RulesTo-ExcelTemplate -Sheet $Template.Rules -Rules $rulesData -RulesIDs ($RiskRules.Report.RuleID | Select-Object -Unique)

Write-Host "[-] Filling Risks into Excel."
RisksTo-ExcelTemplate -Sheet $Template.Risks -Risks $RiskRules -Rules $foundRules

Write-Host "[*] Updating Dashboard Scores."
ScoresByModelTo-ExcelTemplate -Sheet $Template.Dashboard -Risks $RiskRules
ScoresByCategoryTo-ExcelTemplate -Sheet $Template.Dashboard -Risks $RiskRules
ScoresByDomainTo-ExcelTemplate -Sheet $Template.Dashboard -Risks $RiskRules

Write-Host "[*] Making thinks a little bit pretty."
Update-DashboardLinks -Dashboard $Template.Dashboard -Risks $Template.Risks
HighlightTotal -Sheet $Template.Dashboard
PrintLegend -Sheet $Template.Dashboard -ReportsCount $Reports.Count -Book $Template.Book
FreezeColumns -Excel $Template.Excel -Dashboard $Template.DashBoard -Risks $Template.Risks -Rules $Template.Rules

$Template.Excel.Visible = $True

Remove-Module HelperFunctions

Write-Host "[*] All Yours."
