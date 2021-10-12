$TestPath  = 'C:\Users\royco\Documents\Test'
$BlankXlsx = 'blank.xlsx'
'https://webapps.nyuhs.org/popcare/'

$LocalIcons  = @()
$GithubIcons = @(
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/CareTime_1x1.png'
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/Radiation_Oncology_1x1.png'
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/Cardiology_1x1.png'
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/MDN_Glucose_1x1.png'
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/MyDining_1x1.png'
    'https://raw.githubusercontent.com/roycoutts/Load-Helper/main/UHSCareCheck_1x1.png'
)

$msoShapeFlowchartConnector = 73
$msoReflectionType1         = 1
$msoTrue                    = -1
$msoFalse                   = 0

function New-Xlsx {
    [CmdletBinding()]
    param(
        [STRING]$Path,
        [STRING]$Filename
    )
    $filepath = (Join-Path -Path $Path -ChildPath $Filename)
    $b64 = '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'
    $bytes = [Convert]::FromBase64String($b64)
    [System.IO.File]::WriteAllBytes($filepath, $bytes)
    Return $filepath
}

function Get-GithubIcons {
    $LocalDriveIcons = @()
    $GithubIcons | ForEach({ 
        $url = $PSItem
        $bytes = (Invoke-WebRequest -Uri $url).Content
        $filename = (Join-Path -Path $TestPath -ChildPath (($url -split '/')[($url -split '/').Count-1]))
        if (Test-Path -Path $filename) {Remove-Item -Path $filename -Force}
        [System.IO.File]::WriteAllBytes($filename, $bytes)
        $LocalDriveIcons += $filename
    })
    Return $LocalDriveIcons
}

function Get-LeftByColumn ([string]$c) {
    $Column = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T')
    [int]$Points = $Column.IndexOf($c.ToUpper()) * 72
    Return $Points
}

function Get-TopByRow ([int]$r) {
    [int]$Points = ($r - 1) * 14.25
    Return $Points
}

function New-ExcelObject {
    # Create Excel.Application COM object
    $excel = New-Object -ComObject excel.application
    $excel.visible = $False
    Return $excel
}

function Open-Workbook {
    [CmdletBinding()]
    param(
        $InputObject,
        [STRING]$Path
    )
    $wb = $InputObject.Workbooks.Open($Path)
    Return $wb
}

function Set-CellSizes {
    [CmdletBinding()]
    param(
        $WorkBookObject,
        [int]$WorkSheetNumber,
        [double]$ColumnWidth,
        [double]$RowHeight
    )
    $Workbook = $WorkBookObject
    $Workbook.Worksheets[$WorkSheetNumber].Cells.ColumnWidth = $ColumnWidth
    $Workbook.Worksheets[$WorkSheetNumber].Cells.RowHeight   = $RowHeight
    Return $Workbook
}

function Create-Icon {
    [CmdletBinding()]
    param(
        $WorkBookObject,
        $WorkSheetNumber,
        $ShapeNumber,
        $PlaceInColumn,
        $PlaceInRow,
        $Width,
        $Height,
        $IconPath
    )
        Write-Host "[int]$WorkSheetNumber"
        Write-Host "[int]$ShapeNumber"
        Write-Host "[string]$PlaceInColumn"
        Write-Host "[int]$PlaceInRow"
        Write-Host "[int]$Width"
        Write-Host "[int]$Height"
        Write-Host "[string]$IconPath"
        Write-Host "$msoShapeFlowchartConnector"
        $Left = Get-LeftByColumn $PlaceInColumn
        $Top = Get-TopByRow $PlaceInRow
        $ExecLine = @()
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Left = $Left"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Top  = $Top"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Width  = $Width"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Height = $Height"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.Visible = $msoTrue"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.UserPicture(""$IconPath"")"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.TextureTile = $msoFalse"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.RotateWithObject = $msoTrue"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Line.Visible = $msoTrue"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Line.Weight = 3"
    $ExecLine += "`$Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Reflection.Type = $msoReflectionType1"
<#
    $Workbook.Worksheets[$WorkSheetNumber].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Left = $Left
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Top  = $Top
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Width  = $Width
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Height = $Height
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.Visible = $msoTrue
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.UserPicture($IconPath)
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.TextureTile = $msoFalse
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Fill.RotateWithObject = $msoTrue
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Line.Visible = $msoTrue
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Line.Weight = 3
    $Workbook.Worksheets[$WorkSheetNumber].Shapes[$ShapeNumber].Reflection.Type = $msoReflectionType1
    #>
    Return $ExecLine
}

# Create Blank XLSX File
$NewXlsxFilePath = New-Xlsx -Path $TestPath -Filename $BlankXlsx

# Create New Excel Object
$Excel = New-ExcelObject

# Open Workbook for Excel
$Workbook = Open-Workbook -InputObject $Excel -Path $NewXlsxFilePath

# Set Cell Sizes
$Workbook = (Set-CellSizes -WorkBookObject $Workbook -WorkSheetNumber 1 -ColumnWidth 13 -RowHeight 14.5)


# Download Icons from Github / Return Local Path to Icons
$Icons = Get-GithubIcons

# Create Icons in Workbook
$RunLines = (Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 1 -PlaceInColumn B -PlaceInRow 2 -Width 72 -Height 72 -IconPath $Icons[0])
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })
$RunLines = Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 2 -PlaceInColumn 'B' -PlaceInRow 10 -Width 72 -Height 72 -IconPath $Icons[1]
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })
$RunLines = Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 3 -PlaceInColumn 'B' -PlaceInRow 18 -Width 72 -Height 72 -IconPath $Icons[2]
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })
$RunLines = Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 4 -PlaceInColumn 'B' -PlaceInRow 26 -Width 72 -Height 72 -IconPath $Icons[3]
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })
$RunLines = Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 5 -PlaceInColumn 'B' -PlaceInRow 34 -Width 72 -Height 72 -IconPath $Icons[4]
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })
$RunLines = Create-Icon -WorkBookObject $Workbook -WorkSheetNumber 1 -ShapeNumber 6 -PlaceInColumn 'B' -PlaceInRow 42 -Width 72 -Height 72 -IconPath $Icons[5]
$RunLines | ForEach ({ Invoke-Expression -Command $PSItem })


$workbook.Worksheets[1].Columns("A:A").ColumnWidth = 2

$workbook.Save()

Stop-Process -Name EXCEL

<#


    Range("C2").Select
    ActiveCell.FormulaR1C1 = "UHS CareTime"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "UHS Radiation Oncology"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "UHS Cardiology"
    Range("C26").Select
    ActiveCell.FormulaR1C1 = "MDN Glucose"
    Range("C34").Select
    ActiveCell.FormulaR1C1 = "My Dining"
    Range("C35").Select

            Write-Host "[int]$WorkSheetNumber"
        Write-Host "[int]$ShapeNumber"
        Write-Host "[string]$PlaceInColumn"
        Write-Host "[int]$PlaceInRow"
        Write-Host "[int]$Width"
        Write-Host "[int]$Height"
        Write-Host "[string]$IconPath"
        Write-Host "$msoShapeFlowchartConnector"

        

$workbook.Worksheets[1].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
$workbook.Worksheets[1].Shapes[1].Left = (Get-LeftByColumn "B")
$workbook.Worksheets[1].Shapes[1].Top  = (Get-TopByRow 2)
$workbook.Worksheets[1].Shapes[1].Width  = 72
$workbook.Worksheets[1].Shapes[1].Height = 72
$workbook.Worksheets[1].Shapes[1].Fill.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[1].Fill.UserPicture($LocalIcons[0])
$workbook.Worksheets[1].Shapes[1].Fill.TextureTile = $msoFalse
$workbook.Worksheets[1].Shapes[1].Fill.RotateWithObject = $msoTrue
$workbook.Worksheets[1].Shapes[1].Line.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[1].Line.Weight = 3
$workbook.Worksheets[1].Shapes[1].Reflection.Type = $msoReflectionType1

$workbook.Worksheets[1].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
$workbook.Worksheets[1].Shapes[2].Left = (Get-LeftByColumn "B")
$workbook.Worksheets[1].Shapes[2].Top  = (Get-TopByRow 10)
$workbook.Worksheets[1].Shapes[2].Width  = 72
$workbook.Worksheets[1].Shapes[2].Height = 72
$workbook.Worksheets[1].Shapes[2].Fill.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[2].Fill.UserPicture($LocalIcons[1])
$workbook.Worksheets[1].Shapes[2].Fill.TextureTile = $msoFalse
$workbook.Worksheets[1].Shapes[2].Fill.RotateWithObject = $msoTrue
$workbook.Worksheets[1].Shapes[2].Line.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[2].Line.Weight = 3
$workbook.Worksheets[1].Shapes[2].Reflection.Type = $msoReflectionType1

$workbook.Worksheets[1].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
$workbook.Worksheets[1].Shapes[3].Left = (Get-LeftByColumn "B")
$workbook.Worksheets[1].Shapes[3].Top  = (Get-TopByRow 18)
$workbook.Worksheets[1].Shapes[3].Width  = 72
$workbook.Worksheets[1].Shapes[3].Height = 72
$workbook.Worksheets[1].Shapes[3].Fill.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[3].Fill.UserPicture($LocalIcons[2])
$workbook.Worksheets[1].Shapes[3].Fill.TextureTile = $msoFalse
$workbook.Worksheets[1].Shapes[3].Fill.RotateWithObject = $msoTrue
$workbook.Worksheets[1].Shapes[3].Line.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[3].Line.Weight = 3
$workbook.Worksheets[1].Shapes[3].Reflection.Type = $msoReflectionType1

$workbook.Worksheets[1].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
$workbook.Worksheets[1].Shapes[4].Left = (Get-LeftByColumn "B")
$workbook.Worksheets[1].Shapes[4].Top  = (Get-TopByRow 26)
$workbook.Worksheets[1].Shapes[4].Width  = 72
$workbook.Worksheets[1].Shapes[4].Height = 72
$workbook.Worksheets[1].Shapes[4].Fill.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[4].Fill.UserPicture($LocalIcons[3])
$workbook.Worksheets[1].Shapes[4].Fill.TextureTile = $msoFalse
$workbook.Worksheets[1].Shapes[4].Fill.RotateWithObject = $msoTrue
$workbook.Worksheets[1].Shapes[4].Line.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[4].Line.Weight = 3
$workbook.Worksheets[1].Shapes[4].Reflection.Type = $msoReflectionType1

$workbook.Worksheets[1].Shapes.AddShape($msoShapeFlowchartConnector, 1, 1, 1, 1)
$workbook.Worksheets[1].Shapes[5].Left = (Get-LeftByColumn "B")
$workbook.Worksheets[1].Shapes[5].Top  = (Get-TopByRow 34)
$workbook.Worksheets[1].Shapes[5].Width  = 72
$workbook.Worksheets[1].Shapes[5].Height = 72
$workbook.Worksheets[1].Shapes[5].Fill.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[5].Fill.UserPicture($LocalIcons[4])
$workbook.Worksheets[1].Shapes[5].Fill.TextureTile = $msoFalse
$workbook.Worksheets[1].Shapes[5].Fill.RotateWithObject = $msoTrue
$workbook.Worksheets[1].Shapes[5].Line.Visible = $msoTrue
$workbook.Worksheets[1].Shapes[5].Line.Weight = 3
$workbook.Worksheets[1].Shapes[5].Reflection.Type = $msoReflectionType1

$Workbook.Worksheets[1].Shapes.AddShape(73, 1, 1, 1, 1)
$Workbook.Worksheets[1].Shapes[1].Left = 72
$Workbook.Worksheets[1].Shapes[1].Top  = 14
$Workbook.Worksheets[1].Shapes[1].Width  = 72
$Workbook.Worksheets[1].Shapes[1].Height = 72
$Workbook.Worksheets[1].Shapes[1].Fill.Visible = -1
$Workbook.Worksheets[1].Shapes[1].Fill.UserPicture("C:\Users\royco\Documents\Test\CareTime_1x1.png")
$Workbook.Worksheets[1].Shapes[1].Fill.TextureTile = 0
$Workbook.Worksheets[1].Shapes[1].Fill.RotateWithObject = -1
$Workbook.Worksheets[1].Shapes[1].Line.Visible = -1
$Workbook.Worksheets[1].Shapes[1].Line.Weight = 3
$Workbook.Worksheets[1].Shapes[1].Reflection.Type = 1

$Workbook.Worksheets[1].Shapes.AddShape(73, 1, 1, 1, 1)
$Workbook.Worksheets[1].Shapes[2].Left = 72
$Workbook.Worksheets[1].Shapes[2].Top  = 128
$Workbook.Worksheets[1].Shapes[2].Width  = 72
$Workbook.Worksheets[1].Shapes[2].Height = 72
$Workbook.Worksheets[1].Shapes[2].Fill.Visible = -1
$Workbook.Worksheets[1].Shapes[2].Fill.UserPicture("C:\Users\royco\Documents\Test\Radiation_Oncology_1x1.png")
$Workbook.Worksheets[1].Shapes[2].Fill.TextureTile = 0
$Workbook.Worksheets[1].Shapes[2].Fill.RotateWithObject = -1
$Workbook.Worksheets[1].Shapes[2].Line.Visible = -1
$Workbook.Worksheets[1].Shapes[2].Line.Weight = 3
$Workbook.Worksheets[1].Shapes[2].Reflection.Type = 1

Workbook.Worksheets[1].Shapes.AddShape(73, 1, 1, 1, 1)
Workbook.Worksheets[1].Shapes[3].Left = 72
Workbook.Worksheets[1].Shapes[3].Top  = 242
Workbook.Worksheets[1].Shapes[3].Width  = 72
Workbook.Worksheets[1].Shapes[3].Height = 72
Workbook.Worksheets[1].Shapes[3].Fill.Visible = -1
Workbook.Worksheets[1].Shapes[3].Fill.UserPicture(C:\Users\royco\Documents\Test\Cardiology_1x1.png)
Workbook.Worksheets[1].Shapes[3].Fill.TextureTile = 0
Workbook.Worksheets[1].Shapes[3].Fill.RotateWithObject = -1
Workbook.Worksheets[1].Shapes[3].Line.Visible = -1
Workbook.Worksheets[1].Shapes[3].Line.Weight = 3
Workbook.Worksheets[1].Shapes[3].Reflection.Type = 1

Workbook.Worksheets[1].Shapes.AddShape(73, 1, 1, 1, 1)
Workbook.Worksheets[1].Shapes[4].Left = 72
Workbook.Worksheets[1].Shapes[4].Top  = 356
Workbook.Worksheets[1].Shapes[4].Width  = 72
Workbook.Worksheets[1].Shapes[4].Height = 72
Workbook.Worksheets[1].Shapes[4].Fill.Visible = -1
Workbook.Worksheets[1].Shapes[4].Fill.UserPicture(C:\Users\royco\Documents\Test\MDN_Glucose_1x1.png)
Workbook.Worksheets[1].Shapes[4].Fill.TextureTile = 0
Workbook.Worksheets[1].Shapes[4].Fill.RotateWithObject = -1
Workbook.Worksheets[1].Shapes[4].Line.Visible = -1
Workbook.Worksheets[1].Shapes[4].Line.Weight = 3
Workbook.Worksheets[1].Shapes[4].Reflection.Type = 1

Workbook.Worksheets[1].Shapes.AddShape(73, 1, 1, 1, 1)
Workbook.Worksheets[1].Shapes[5].Left = 72
Workbook.Worksheets[1].Shapes[5].Top  = 470
Workbook.Worksheets[1].Shapes[5].Width  = 72
Workbook.Worksheets[1].Shapes[5].Height = 72
Workbook.Worksheets[1].Shapes[5].Fill.Visible = -1
Workbook.Worksheets[1].Shapes[5].Fill.UserPicture(C:\Users\royco\Documents\Test\MyDining_1x1.png)
Workbook.Worksheets[1].Shapes[5].Fill.TextureTile = 0
Workbook.Worksheets[1].Shapes[5].Fill.RotateWithObject = -1
Workbook.Worksheets[1].Shapes[5].Line.Visible = -1
Workbook.Worksheets[1].Shapes[5].Line.Weight = 3
Workbook.Worksheets[1].Shapes[5].Reflection.Type = 1



#>
