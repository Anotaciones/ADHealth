Function GetRuleRow {
    Param(
       [Parameter(Mandatory=$true)][Object[]]$Rules,
       [Parameter(Mandatory=$true)][string]$ruleID
    )
    $Rules.IndexOf($ruleID) + 2
}

Function RulesTo-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][__ComObject]$Sheet,
      [Parameter(Mandatory=$true)][Object[]]$Rules,
      [Parameter(Mandatory=$true)][Object[]]$RulesIDs
    )
    $foundRules = @()
    $row = 2
    $Rules | %{
      $rule = $_
      If ($RulesIDs -contains $rule.RuleID) {
        $CategoryPrettyPrint = $CategoryHeaders | ?{ $_.Replace(' ', '').ToLower() -eq $rule.Category.ToLower() }
        $ModelPrettyPrint = $ModelHeaders | ?{ $_.Replace(' ', '').ToLower() -eq $rule.Model}
        $Sheet.cells.item($row,1) = $CategoryPrettyPrint
        $Sheet.cells.item($row,2) = $ModelPrettyPrint
        $Sheet.cells.item($row,3) = $rule.RuleID
        $Sheet.cells.item($row,4) = $rule.Title
        $Sheet.cells.item($row,5) = $rule.Description
        $Sheet.cells.item($row,6) = $rule.Explanation
        $Sheet.cells.item($row,7) = $rule.Solution
        $foundRules += $rule.RuleID
        $row++  
      }
    }
    $null = $Sheet.Cells.Item(1,1).AutoFilter()
    $foundRules
  }
  
  Function RisksTo-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][__ComObject]$Sheet,
      [Parameter(Mandatory=$true)][Object[]]$Risks,
      [Parameter(Mandatory=$true)][Object[]]$Rules
    )
    $row = $Sheet.UsedRange.Columns.Item(1).Rows.Count + 1

    $Risks | %{
      $risk = $_
      $risk.Report | %{
        $report = $_
        $ruleRow = GetRuleRow -Rules $Rules -ruleID $report.RuleID
        $Sheet.cells.item($row,1) = $report.Category
        $Sheet.cells.item($row,2) = $report.Model
        $Sheet.cells.item($row,3) = $report.RuleID
        $Sheet.cells.item($row,4) = $report.Risk
        $Sheet.cells.item($row,5) = $risk.Domain
        $Sheet.cells.item($row,6) = $report.Score
        $Sheet.Hyperlinks.Add($Sheet.Cells.Item($row,3),"", ("Rules!A{0}" -f $ruleRow)) | Out-Null
        ScoreColorFormating -Cell $Sheet.cells.item($row,6) -Score $report.Score
        $row ++
      }
    }
    $null = $Sheet.Columns.AutoFit()
    $null = $Sheet.Cells.Item(1,1).AutoFilter()    
  }


  Function ScoresByModelTo-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][__ComObject]$Sheet,
      [Parameter(Mandatory=$true)][Object[]]$Risks      
    )
    $Offset = 0
    $Risks | %{
      $risk = $_
      $Domain = $risk.Domain
      $risk.Report | Group-Object -Property Model | %{
        $risk_group = $_
        $i = $ModelHeaders.IndexOf(($ModelHeaders | ?{ $_ -eq $risk_group.Name}))
        $sum = ($risk_group.Group.Score | Measure-Object -Sum).Sum
        $Sheet.cells.item(5 + $Offset, 3+$i) = $sum  
      }
      $Sheet.cells.item(5 + $Offset, 2) = $Domain = $risk.Domain
      $Offset++
    }
  }

  Function ScoresByCategoryTo-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][__ComObject]$Sheet,
      [Parameter(Mandatory=$true)][Object[]]$Risks,
      [Parameter(Mandatory=$false)][int]$Offset = 0
      
    )

    $Offset = $Risks.Count
    $Risks | %{
      $risk = $_
      $Domain = $risk.Domain
      $tot = ($risk.Report.Score | Measure-Object -Sum).Sum 
      $risk.Report | Group-Object -Property Category | %{
        $risk_group = $_
        $sum = ($risk_group.Group.Score | Measure-Object -Sum).Sum
        $i = $CategoryHeaders.IndexOf(($CategoryHeaders | ?{ $_ -eq $risk_group.Name}))
        $Item1 = $Sheet.Cells.Item(7 + $Offset, $CategoryOffsets[$i])
        $Item2 = $Sheet.Cells.Item(7 + $Offset, $CategoryOffsets[$i] + $CategorySizes[$i] - 2)
        $Range = $Sheet.Range($Item1, $Item2)
        try {
          $Range.Value2 = ($sum / $tot)
          $Sheet.Cells.Item(7 + $Offset, $CategoryOffsets[$i] + $CategorySizes[$i] - 1) = $sum
        } catch {
          $Range.Value2 = "0.00"
          $Sheet.Cells.Item(7 + $Offset, $CategoryOffsets[$i] + $CategorySizes[$i] - 1) = 0  
        }
      }
      $Sheet.cells.item(7 + $Offset, 2) = $Domain
      $Offset++
    }
  }

  Function ScoresByDomainTo-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][__ComObject]$Sheet,
      [Parameter(Mandatory=$true)][Object[]]$Risks
    )
    $Offset = $Risks.Count * 2
    $Totals = @()
    $Risks | %{
      $risk = $_
      $Domain = $risk.Domain
      $total = (($risk.Report.Score | Measure-Object -Sum).Sum)
      $Totals += $total
      $Sheet.Cells.Item(9 + $Offset, 36) = $total
      $Sheet.Cells.Item(9 + $Offset, 2) = $Domain
      $Offset++
    }
    $max = ($Totals | Measure-Object -Maximum).Maximum
    $threshold = $max / 33
    $Offset = $Risks.Count * 2
    $Risks | %{
      $risk = $_
      $Item1 = $Sheet.Cells.Item(9 + $Offset, 3)
      $Item2 = $Sheet.Cells.Item(9 + $Offset, 35)
      $Range = $Sheet.Range($Item1, $Item2)
      $Range.Font.Color = $BACKGROUND_GREY2
      $ifCond = $Range.FormatConditions.Add(1, 1, "=0", ("=`$AJ`${0}" -f (9 + $Offset)))
      $ifCond.Interior.Color = (RGB -R 217 -G 150 -B 148)
      $ifCond.Font.Color = (RGB -R 217 -G 150 -B 148)
      Set-CellFontFormat -Cell $Sheet.Cells.Item(9 + $Offset, 36) -FontSize 8 -FontBold $True -FontColor $FONT_COLOR_CAT -InteriorColor $BACKGROUND_GREY2
      0..32 | %{
        $idx2 = $_
        $Sheet.Cells.Item(9 + $Offset, 3 + $idx2) = $idx2 * $threshold
      }
      $Offset++
    }
  }
 
  Function Update-DashboardLinks {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Dashboard,
       [Parameter(Mandatory=$true)][__ComObject]$Risks
       )
    For ($i = 0; $i -lt $ModelHeaders.Length; $i++) {
        $model = $ModelHeaders[$i]
        $findRow = $Risks.Cells.Find($model) # $null if not found
        If ($findRow -ne $null) {
            $found = $findRow.Rows[0].Row + 1
            $Dashboard.Hyperlinks.Add($Dashboard.Cells.Item(3,3+$i),"", ("Risks!A{0}" -f $found)) | Out-Null
        }
    }
}