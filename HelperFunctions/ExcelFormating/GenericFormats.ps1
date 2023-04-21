Function RGB {
    Param(
        [Parameter(Mandatory = $true)][int]$R,
        [Parameter(Mandatory = $true)][int]$G,
        [Parameter(Mandatory = $true)][int]$B
    )
    [long] ($R + ($G * 256) + ($B * 256 * 256))
}

Function Set-CellFontFormat {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Cell,
       [Parameter(Mandatory=$false)][long]$InteriorColor = [long]16777215,
       [Parameter(Mandatory=$false)][long]$FontColor = [long]0,
       [Parameter(Mandatory=$false)][String]$FontName = 'Calibri Light',
       [Parameter(Mandatory=$false)][int]$FontSize = 10,
       [Parameter(Mandatory=$false)][Boolean]$FontBold = $False
    )
    $Cell.Interior.Color = $InteriorColor
    $Cell.Font.Color = $FontColor
    $Cell.Font.Name = $FontName
    $Cell.Font.Size = $FontSize
    $Cell.Font.Bold = $FontBold
}

Function Set-CellsSize {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Cells,
       [Parameter(Mandatory=$false)][double]$Width = -1.0,
       [Parameter(Mandatory=$false)][double]$Height = -1.0
    )
    If ($Width -ne -1.0) {$Cells.ColumnWidth = $Width}
    If ($Height -ne -1.0) {$Cells.RowHeight = $Height}
}

Function Set-CellAlignment {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Cell,
       [Parameter(Mandatory=$false)][Int32]$Horizontal = -1,
       [Parameter(Mandatory=$false)][Int32]$Vertical = -1,
       [Parameter(Mandatory=$false)][Int32]$Orientation = -1
    )
    If ($Horizontal -ne -1)  {$Cell.HorizontalAlignment = $Horizontal}
    If ($Vertical -ne -1)    {$Cell.VerticalAlignment = $Vertical}
    If ($Orientation -ne -1) {$Cell.Orientation = $Orientation}
}

Function Set-CellBorders {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Cell,
       [Parameter(Mandatory=$false)][int]$L = -1,
       [Parameter(Mandatory=$false)][int]$R = -1,
       [Parameter(Mandatory=$false)][int]$T = -1,
       [Parameter(Mandatory=$false)][int]$B = -1
    )
    If ($L -ne -1) {$Cell.Borders(1).Color = $L}
    If ($R -ne -1) {$Cell.Borders(2).Color = $R}
    If ($T -ne -1) {$Cell.Borders(3).Color = $T}
    If ($B -ne -1) {$Cell.Borders(4).Color = $B}
}


Function ScoreColorFormating {
    Param(
       [Parameter(Mandatory=$true)][__ComObject]$Cell,
       [Parameter(Mandatory=$true)][int]$Score
    )
    If ($Score -ge 31) {
        Set-CellFontFormat -Cell $Cell -FontColor 2162853 -InteriorColor 5263615
    } ElseIf ($Score -ge 11) {
        Set-CellFontFormat -Cell $Cell -FontColor 801923 -InteriorColor 14083324
    } ElseIf ($Score -ge 1) {
        Set-CellFontFormat -Cell $Cell -FontColor 24704 -InteriorColor 13431551
    } Else {
        Set-CellFontFormat -Cell $Cell -FontColor 6567712 -InteriorColor 15123099
    }
}