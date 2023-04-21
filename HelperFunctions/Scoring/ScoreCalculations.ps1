Function CalculateTotalScore {
    Param(
      [Parameter(Mandatory=$true)][Object[]]$Risks
    )
    $tot = (($Risks | Group-Object -Property Domain).Group.Score | Measure-Object -Sum).Sum
    $dom = ($Risks | Group-Object -Property Domain).Name
    [PSCustomObject]@{
      Domain = $dom
      Score = $tot
    }
}