<# Constantes de formato (Color) #>
$BORDER_COLOR = [int]14408667
$BORDER_COLOR2 = [int]12434877
$BORDER_COLOR3 = [int]16777215
$FONT_COLOR = [long]16777215
$FONT_COLOR_HDR = [long]0
$FONT_COLOR_CAT = [long]8421504
$BACKGROUND_HDR = [long]3342489
$BACKGROUND_GREY = [long]15132390
$BACKGROUND_GREY2 = [long]15921906

#$MODELS_COLORS = @([long]13285804, [long]11573124, [long]5193523, [long]855309)
$MODELS_COLORS = @(
    [long]10642560, 
    [long]9592887, 
    [long]5880731, 
    [long]10257713
)

#$DATABAR_COLORS = @([int]10086143, [int]15189684, [int]11854022, [int]11389944)
$DATABAR_COLORS = @(
    [int]14336460, 
    [int]13995605, 
    [int]10213059, 
    [int]14536083
)

<# Constantes de valor (Cabeceras y posiciones) #>
$ruleHeaders = @(@('Category', 15.0), @('Model', 24.0), @('RuleID', 30.50), @('Title', 30.50), @('Description', 58.0), @('Explanation', 68.50), @('Solution', 80.0))
$riskHeaders = @('Category', 'Model', 'RuleID', 'Risk', 'Domain', 'Score')
$CategoryHeaders = @( 'Stale objects', 'Privileged Accounts', 'Trusts', 'Anomalies')
$CategorySizes = @(8, 8, 6, 11)
$CategoryOffsets = @(3, 11, 19, 25)
$ModelHeaders = @('Inactive User Or Computer', 'Network topography', 'Object Config', 'Obsolete OS', 'Old authentication protocols', 'Provisioning', 'Replication', 'Vulnerability management', 'Account take over', 'ACL Check', 'Admin control', 'Control Path', 'Delegation Check', 'Irreversible change', 'Privilege control', 'RODC', 'Old trust protocol', 'SID Filtering', 'SIDHistory', 'Trust impermeability', 'Trust inactive', 'Trust Azure', 'Audit', 'Backup', 'Certificate take over', 'Golden ticket', 'Local group vulnerability', 'Network sniffing', 'Pass the credential', 'Password retrieval', 'Reconnaissance', 'Temporary Admins', 'Weak password')

<# Generación de documento Excel base con sus correspondientes pestañas #>
Function Global-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel
    )
    $Excel.ActiveSheet
    
    $Workbook = $Excel.workbooks.add()
    $sRules = $Workbook.worksheets.item(1)
    $sRisks = $Workbook.worksheets.Add()
    $sDashboard = $Workbook.worksheets.Add()
    $sDashboard.name = "DashBoard"
    $sRisks.name = "Risks"
    $sRules.name = "Rules"
  
    $Excel.Worksheets | % {
        $_.Activate()
        $Excel.ActiveSheet.PageSetup.PrintGridlines = $False
        $Excel.ActiveWindow.DisplayGridlines = $False
    }
    @{
        Excel     = $Excel
        Book      = $Workbook
        Dashboard = $sDashboard
        Risks     = $sRisks
        Rules     = $sRules
    }   
}

Function Save-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel,
        [Parameter(Mandatory = $true)][string]$Folder
    )
    $d = Get-Date -Format "dd-MM-yyyy HH-MM-ss"
    $f = "{0} Summary Report.xlsx" -f $d.ToString()
    $fname = Join-Path -Path $Folder -ChildPath $f
    $Excel.SaveAs($fname, 51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)
}
<# Formato de pestaña de reglas #>
Function Rules-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel,
        [Parameter(Mandatory = $true)][__ComObject]$Sheet
    )
    Set-CellFontFormat -Cell $Sheet.Cells
    Set-CellsSize -Cells $Sheet.Cells -Height 12.2
    0..($ruleHeaders.Length - 1) | % {
        $Sheet.Cells.Item(1, $_ + 1) = $ruleHeaders[$_][0]
        Set-CellsSize -Cells $Sheet.Columns($_ + 1) -Width $ruleHeaders[$_][1]
        Set-CellAlignment -Cell $Sheet.Cells.Item(1, $_ + 1) -Horizontal 3
        Set-CellFontFormat -Cell $Sheet.Cells.Item(1, $_ + 1) -InteriorColor $BACKGROUND_HDR -FontColor $FONT_COLOR -FontBold $True
    }
    Set-CellAlignment -Cell $Sheet.Rows() -Vertical -4160 -Horizontal -4131
}

<# Formato de pestaña de riesgos #>
Function Risks-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel,
        [Parameter(Mandatory = $true)][__ComObject]$Sheet
    )
    Set-CellsSize -Cells $Sheet.Cells -Height 12.2
    Set-CellFontFormat -Cell $Sheet.Cells
    0..($riskHeaders.Length - 1) | % {
        $Sheet.Cells.Item(1, $_ + 1) = $riskHeaders[$_]
        Set-CellAlignment -Cell $Sheet.Cells.Item(1, $_ + 1) -Horizontal 3
        Set-CellFontFormat -Cell $Sheet.Cells.Item(1, $_ + 1) -InteriorColor $BACKGROUND_HDR -FontColor $FONT_COLOR -FontBold $True  
    }
}

<# Formato base del panel de control #>
Function DashboardGlobal-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel,
        [Parameter(Mandatory = $true)][__ComObject]$Sheet
    )
    Set-CellsSize -Cells $Sheet.Cells -Width 4.01 -Height 14.4
    Set-CellsSize -Cells $Sheet.Columns(1) -Width 1.50
    Set-CellsSize -Cells $Sheet.Columns(2) -Width 22.50
    Set-CellsSize -Cells $Sheet.Columns(36) -Width 10.80
    Set-CellsSize -Cells $Sheet.Columns(38) -Width 48.9
    Set-CellsSize -Cells $Sheet.Rows(3) -Height 130.0
    Set-CellFontFormat -Cell $Sheet.Columns(2) -FontColor $FONT_COLOR_CAT
    Set-CellAlignment -Cell $Sheet.Columns(2) -Horizontal 4 -Vertical -4107
    $Item1 = $Sheet.Cells.Item(1, 3)
    $Item2 = $Sheet.Cells.Item(2, 35)
    $Range = $Sheet.Range($Item1, $Item2)
    $Range.Merge()
    $Range.Value2 = "Active Directory Health Check"
    Set-CellFontFormat -Cell $Range -FontBold $True -FontColor 0 -FontSize 20
    Set-CellAlignment -Cell $Range -Vertical -4108 -Horizontal -4108
}

<# Formato de celdas para cabeceras de panel de control respecto a modelos y categorias #>
<# Para Scores por modelo #>
Function DashboardByModel-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet
    )
    0..($ModelHeaders.Length - 1) | % {
        $Offset = $_
        $Sheet.Cells.Item(3, 3 + $Offset) = $ModelHeaders[$Offset]
        Set-CellAlignment -Cell $Sheet.Cells.Item(3, 3 + $Offset) -Orientation -4171 -Vertical -4107 -Horizontal -4108
        Set-CellFontFormat -Cell $Sheet.Cells.Item(3, 3 + $Offset) -InteriorColor $BACKGROUND_GREY
        $Sheet.Cells.Item(3, 3 + $Offset).Borders.Color = [int]$BORDER_COLOR
        Set-CellBorders -Cell $Sheet.Cells.Item(3, 3 + $Offset) -L $BORDER_COLOR -R $BORDER_COLOR -T $BORDER_COLOR -B $BORDER_COLOR
    }
    $Offset = 3
    0..($CategoryHeaders.Length - 1) | % {
        $idx = $_
        $1 = $Sheet.Cells.Item(4, $Offset)
        $2 = $Sheet.Cells.Item(4, $Offset + $CategorySizes[$idx] - 1)
        $3 = $Sheet.Range($1, $2)  
        $Sheet.Cells.Item(4, $Offset) = $CategoryHeaders[$idx]
        $3.Merge()
        Set-CellFontFormat -Cell $Sheet.Cells.Item(4, $Offset) -InteriorColor $MODELS_COLORS[$idx] -FontColor $FONT_COLOR -FontBold $True
        Set-CellAlignment -Cell $Sheet.Cells.Item(4, $Offset) -Horizontal 3
        $Offset += $CategorySizes[$idx]
    }
}

<# Formato de celdas para cabeceras de panel de control respecto a categorias #>
<# Para Scores por categoría #>
Function DashboardByCategory-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    $Offset = 3
    0..($CategoryHeaders.Length - 1) | % {
        # Cond Format
        $idx = $_
        $1 = $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset)
        $2 = $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset + $CategorySizes[$idx] - 2)
        $3 = $Sheet.Range($1, $2)
        $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset) = $CategoryHeaders[$idx]
        $3.Merge()
        Set-CellFontFormat -Cell $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset) -InteriorColor $MODELS_COLORS[$idx] -FontColor $FONT_COLOR -FontBold $True
        Set-CellAlignment -Cell $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset) -Horizontal -4108
        $Offset += $CategorySizes[$idx]
        # Score
        Set-CellFontFormat -Cell $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset - 1) -InteriorColor $MODELS_COLORS[$idx] -FontColor $FONT_COLOR -FontBold $True -FontSize 8
        $Sheet.Cells.Item(5 + $ReportsCount + 1, $Offset - 1) = "Score"
    }
}

<# Formato condicional de celdas para Scores por modelo (IconSet's) #>
Function DashboardIconSets-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Book,
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    $Item1 = $Sheet.Cells.Item(5, 3)
    $Item2 = $Sheet.Cells.Item(5 + $ReportsCount - 1, $ModelHeaders.Length + 2)
    $Range = $Sheet.Range($Item1, $Item2) 
    Set-CellFontFormat -Cell $Range -InteriorColor $BACKGROUND_GREY2
    Set-CellAlignment -Cell $Range -Horizontal  -4108 -Vertical -4108
    $iset = $Range.FormatConditions.AddIconSetCondition()
    $iset.IconSet = $Book.iconsets(13)
    $iset.ShowIconOnly = $True
    $iset.IconCriteria(1).Icon = 10
    $iset.IconCriteria(2).Type = 0
    $iset.IconCriteria(2).Operator = 5
    $iset.IconCriteria(2).Value = 0
    $iset.IconCriteria(2).Icon = 11
    $iset.IconCriteria(3).Type = 0
    $iset.IconCriteria(3).Operator = 7
    $iset.IconCriteria(3).Value = 11
    $iset.IconCriteria(3).Icon = 30
    $iset.IconCriteria(4).Type = 0
    $iset.IconCriteria(4).Operator = 7
    $iset.IconCriteria(4).Value = 31
    $iset.IconCriteria(4).Icon = 29
}

<# Formato condicional de celdas para Scores por categoría (DataBar's) #>
Function DashboardDataBars-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    0..($ReportsCount - 1) | % {
        $idx = $_
        $Offset = 3
        0..($CategoryHeaders.Length - 1) | % {
            $idx2 = $_
            $mod = $CategoryHeaders[$idx2]
            $Item1 = $Sheet.Cells.Item(7 + $ReportsCount + $idx, $Offset)
            $Item2 = $Sheet.Cells.Item(7 + $ReportsCount + $idx, $Offset + $CategorySizes[$idx2] - 2)
            $Range = $Sheet.Range($Item1, $Item2) 
            $null = $Range.Merge() 
            Set-CellFontFormat -Cell $Range -FontSize 8 -FontColor 0 -InteriorColor $BACKGROUND_GREY2
            $Range.NumberFormat = "0.00%"
            $dbar = $Sheet.Cells.Item(7 + $ReportsCount + $idx, $Offset).FormatConditions.AddDatabar()
            $dbar.MinPoint.Modify(0, 0)
            $dbar.MaxPoint.Modify(0, 1)
            $dbar.BarBorder.Type = 0
            $dbar.BarFillType = 0
            $dbar.Direction = -5004
            $dbar.BarColor.Color = $DATABAR_COLORS[$idx2]
            $cScore = $Sheet.Cells.Item(7 + $ReportsCount + $idx, $Offset + $CategorySizes[$idx2] - 1)
            Set-CellFontFormat -Cell $cScore -FontSize 8 -FontBold $True -FontColor $FONT_COLOR_CAT -InteriorColor $BACKGROUND_GREY2
            $Offset += $CategorySizes[$idx2]
        } 
    }
}

Function DashboardWaffles-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    $Item1 = $Sheet.Cells.Item(8 + ($ReportsCount * 2), 3)
    $Item2 = $Sheet.Cells.Item(8 + ($ReportsCount * 2), 36)
    $Range = $Sheet.Range($Item1, $Item2)
    $Range.Merge()
    $Range.Value2 = "Total Scores"
    Set-CellFontFormat -Cell $Range -InteriorColor (RGB -R 149 -G 55 -B 53) -FontBold $True -FontColor 16777215
    Set-CellAlignment -Cell $Range -Horizontal 4
    0..($ReportsCount - 1) | % {
        $idx1 = $_
        0..($ModelHeaders.Length) | % {
            $idx2 = $_
            $cell = $Sheet.Cells.Item(9 + ($ReportsCount * 2) + $idx1, 3 + $idx2)
            Set-CellFontFormat -Cell $cell -InteriorColor $BACKGROUND_GREY2
            Set-CellBorders -Cell $cell -B $BORDER_COLOR
            If ($idx2 -eq 0) {
                Set-CellBorders -Cell $cell -L $BORDER_COLOR
            }
            If ($idx2 -ge $ModelHeaders.Length-1) {
                Set-CellBorders -Cell $cell -R $BORDER_COLOR
            }
        }
    }
}

<# Formato de bordes en panel de control #>
Function DashboardIconSetsBorders-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    0..($ReportsCount - 1) | % {
        $idx = $_
        $Offset = 3
        0..($ModelHeaders.Length - 1) | % {
            $idx2 = $_
            Set-CellBorders -Cell $Sheet.Cells.Item(5 + $idx, $Offset + $idx2) -L $BORDER_COLOR -R $BORDER_COLOR
            If ($idx2 -eq 0) {
                Set-CellBorders -Cell $Sheet.Cells.Item(5 + $idx, $Offset + $idx2) -L $BORDER_COLOR2
            }        
            If ($idx2 -eq ($ModelHeaders.Length - 1)) {
                Set-CellBorders -Cell $Sheet.Cells.Item(5 + $idx, $Offset + $idx2) -R $BORDER_COLOR2
            }      
            If ($idx -eq ($ReportsCount - 1)) {
                Set-CellBorders -Cell $Sheet.Cells.Item(5 + $idx, $Offset + $idx2) -B $BORDER_COLOR2
            }
        }
    }
    0..($ReportsCount - 1) | % {
        $idx = $_  
        $Offset = 3
        0..($CategoryHeaders.Length - 1) | % {
            $mod = $CategoryHeaders[$_]
            Set-CellBorders -Cell $Sheet.Cells.Item(5 + $idx, $Offset) -L $BORDER_COLOR2
            $Offset += $mod[1]
        }
    } 
}
Function DashboardDataBarsBorders-ExcelTemplate {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Sheet,
        [Parameter(Mandatory = $true)][Int32]$ReportsCount
    )
    0..($ReportsCount) | % {
        $idx = $_
        $Offset = 3
        0..($CategoryHeaders.Length-1) | % {
            $idx2 = $_
            $mod = $CategoryHeaders[$idx2]
            $Item1 = $Sheet.Cells.Item(6 + $ReportsCount + $idx, $Offset)
            $Item2 = $Sheet.Cells.Item(6 + $ReportsCount + $idx, $Offset + $mod[1] - 2)
            $Range = $Sheet.Range($Item1, $Item2)
            If ($idx -eq 0) {
                Set-CellBorders -Cell $Range -L $BORDER_COLOR3 -R $BORDER_COLOR3
            } Else {
                Set-CellBorders -Cell $Range -L $BORDER_COLOR -R $BORDER_COLOR
            }
            $Offset += $mod[1]
        }
        Set-CellBorders -Cell $Sheet.Cells.Item(6 + $ReportsCount + $idx, $Offset) -L $BORDER_COLOR
    }
    $Offset = 3
    0..($ModelHeaders.Length-1) | % {
        $idx = $_
        Set-CellBorders -Cell $Sheet.Cells.Item(5 + ($ReportsCount * 2) + 1, $Offset) -B $BORDER_COLOR
        $Offset += 1
    }
}

Function HighlightTotal {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Sheet
    )
    $tot = ($Sheet.Cells.Range("AJ5:AJ65535").Value2 | ?{ ($_ -ne $null) -and ($_.GetType() -ne [String]) }).Count
    $max = ($Sheet.Cells.Range("AJ5:AJ65535").Value2 | ?{ ($_ -ne $null) -and ($_.GetType() -ne [String]) } | measure -Maximum).Maximum
    $row = $Sheet.Cells.Range("AJ5:AJ65535").Find($max).Row
    $Sheet.Cells.Range("B{0}:AI{0}" -f ($row - $tot - 2)).Interior.Color = (RGB -R 206 -G 206 -B 206)
    $Sheet.Cells.Range("B{0}:AI{0}" -f ($row - ($tot*2) - 4)).Interior.Color = (RGB -R 206 -G 206 -B 206)
    $Sheet.Cells.Range("B{0}:AJ{0}" -f ($row)).Interior.Color = (RGB -R 206 -G 206 -B 206)

}

Function PrintLegend {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Book,
       [Parameter(Mandatory=$true)][__ComObject]$Sheet,
       [Parameter(Mandatory=$true)][Int32]$ReportsCount
    )
   
    $Item1 = $Sheet.Cells.Item(5, 37)
    $Item2 = $Sheet.Cells.Item(8, 37)
    $Range = $Sheet.Range($Item1, $Item2) 
    $iset = $Range.FormatConditions.AddIconSetCondition()
    $iset.IconSet = $Book.iconsets(13)
    $iset.ShowIconOnly = $True
    $iset.IconCriteria(1).Icon = 10
    $iset.IconCriteria(2).Type = 0
    $iset.IconCriteria(2).Operator = 5
    $iset.IconCriteria(2).Value = 0
    $iset.IconCriteria(2).Icon = 11
    $iset.IconCriteria(3).Type = 0
    $iset.IconCriteria(3).Operator = 7
    $iset.IconCriteria(3).Value = 11
    $iset.IconCriteria(3).Icon = 30
    $iset.IconCriteria(4).Type = 0
    $iset.IconCriteria(4).Operator = 7
    $iset.IconCriteria(4).Value = 31
    $iset.IconCriteria(4).Icon = 29
    $Sheet.Cells.Item(5, 37) = 0
    $Sheet.Cells.Item(5, 38) = "Score is 0 - no risk identified but some improvements detected"
    $Sheet.Cells.Item(6, 37) = 10
    $Sheet.Cells.Item(6, 38) = "Score between 1 and 10 - a few actions have been identified"
    $Sheet.Cells.Item(7, 37) = 22
    $Sheet.Cells.Item(7, 38) = "Score between 10 and 30 - rules should be looked with attention"
    $Sheet.Cells.Item(8, 37) = 33
    $Sheet.Cells.Item(8, 38) = "Score higher than 30 - major risks identified"
    $Sheet.Cells.Item(9, 38) = "Blank cells: no matched rules"

    Set-CellFontFormat -Cell $Sheet.Cells.Item(5, 37) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(5, 38) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(6, 37) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(6, 38) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(7, 37) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(7, 38) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(8, 37) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(8, 38) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(9, 37) -InteriorColor (RGB -R 242 -G 242 -B 242)
    Set-CellFontFormat -Cell $Sheet.Cells.Item(9, 38) -InteriorColor (RGB -R 242 -G 242 -B 242)

}

Function FreezeColumns {
    Param(
        [Parameter(Mandatory = $true)][__ComObject]$Excel,
        [Parameter(Mandatory = $true)][__ComObject]$Dashboard,
        [Parameter(Mandatory = $true)][__ComObject]$Risks,
        [Parameter(Mandatory = $true)][__ComObject]$Rules
    )
    $Excel.ActiveWindow.WindowState = -4137
    $Rules.Activate()
    $Excel.ActiveWindow.Zoom = 90
    $null = $Rules.Cells.Item(2, 4).Select()
    $Excel.ActiveWindow.FreezePanes = $True

    $Risks.Activate()
    $Excel.ActiveWindow.Zoom = 90
    $null = $Risks.Cells.Item(2, 4).Select()
    $Excel.ActiveWindow.FreezePanes = $True

    $Dashboard.Activate()
    $Excel.ActiveWindow.Zoom = 80
    $null = $Dashboard.Rows(3).Select()
    $Excel.ActiveWindow.FreezePanes = $True
    $null = $Dashboard.Cells.Item(1,1).Select()
}

<# Funcion genérica para crear el excel base de trabajo #>
Function Generate-ExcelTemplate {
    Param(
      [Parameter(Mandatory=$true)][Int32]$ReportsCount
    )
    $Excel = New-Object -ComObject Excel.Application
    $Template = Global-ExcelTemplate -Excel $Excel
    Rules-ExcelTemplate -Sheet $Template.Rules -Excel $Template.Excel
    Risks-ExcelTemplate -Sheet $Template.Risks -Excel $Template.Excel
    DashboardGlobal-ExcelTemplate -Sheet $Template.Dashboard -Excel $Template.Excel
    DashboardByModel-ExcelTemplate -Sheet $Template.Dashboard
    DashboardByCategory-ExcelTemplate -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    DashboardIconSets-ExcelTemplate -Book $Template.Book -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    DashboardDataBars-ExcelTemplate -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    DashboardWaffles-ExcelTemplate -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    DashboardIconSetsBorders-ExcelTemplate -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    DashboardDataBarsBorders-ExcelTemplate -Sheet $Template.Dashboard -ReportsCount $ReportsCount
    $Template
  }
