#---------------------------------------------------------[Initialisations]--------------------------------------------------------
# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$LineLocation = @(20,50,80,110,140,170,200,230,260)

function Get-ReportNames {
    $ReportNames = @()
    $ReportNames += 'Yearly Patient & Treatment Count'
    $ReportNames += 'Combined Analytics Report'
    $ReportNames += 'Synergy Analytics Report'
    $ReportNames += 'TrueBeam Analytics Report'
    $ReportNames += 'Sims Analytics Report'
    $ReportNames += 'SRS Data'
    $ReportNames += 'Diagnosis Analytics Report (Current Year)'
    $ReportNames += 'Diagnosis Analytics Report (Previous Year)'
    Return $ReportNames
}

function Get-FieldNames {
    $FieldNames = @()
    $FieldNames += 'Report_Start_Date: '
    $FieldNames += '  Report_End_Date: '
    $FieldNames += '     Machine_Name: '
    $FieldNames += '         Location: '
    $FieldNames += '          Holiday: '
    $FieldNames += '          Weekend: '
    Return $FieldNames
}

function Get-Year {
    Return
    Return ([DateTime]::Today).Year
}
function Get-Month {Return ([DateTime]::Today).Month}
function Get-YD1 {
    if((Get-Month) -eq 1){$PM = 1;$YearOffset=1}else{$PM = (Get-Month)-1;;$YearOffset=0}
    Return ([DateTime]::new(((Get-Year)-$YearOffset),1,1).ToString("MM/dd/yyyy"))
}
function Get-PMD1 {
    if((Get-Month) -eq 1){$PM = 1;$YearOffset=1}else{$PM = (Get-Month)-1;;$YearOffset=0}
}
function Get-PME {
    if((Get-Month) -eq 1){$PM = 1;$YearOffset=1}else{$PM = (Get-Month)-1;;$YearOffset=0}
    Return ([DateTime]::new(((Get-Year)-$YearOffset),$PM,[DateTime]::DaysInMonth(((Get-Year)-$YearOffset),$PM))).ToString("MM/dd/yyyy")
}

function Update-DisplayFields ($ReportIndex) {
    $lblField1.Visible = $false ; $txtField1.Visible = $false
    $lblField2.Visible = $false ; $txtField2.Visible = $false
    $lblField3.Visible = $false ; $txtField3.Visible = $false
    $lblField4.Visible = $false ; $txtField4.Visible = $false
    $lblField5.Visible = $false ; $txtField5.Visible = $false
    $lblField6.Visible = $false ; $txtField6.Visible = $false
    if ($ReportIndex -eq 0) {    # Yearly Patient & Treatment Count
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1] ; $lblField2.Visible = $true 
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-YD1)           ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-PME)           ; $txtField2.Visible = $true
    }
    if ($ReportIndex -eq 1) {    # Combined Analytics Report
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2] ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4] ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5] ; $lblField4.Visible = $true 
    }
    if ($ReportIndex -eq 2) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2] ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4] ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5] ; $lblField4.Visible = $true 
    }
    if ($ReportIndex -eq 3) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2] ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4] ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5] ; $lblField4.Visible = $true 
    }
    if ($ReportIndex -eq 4) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[3] ; $lblField2.Visible = $true 
    }
    if ($ReportIndex -eq 5) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1] ; $lblField2.Visible = $true 
    }
    if ($ReportIndex -eq 6) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1] ; $lblField2.Visible = $true 
    }
    if ($ReportIndex -eq 7) {
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0] ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1] ; $lblField2.Visible = $true 
    }
}

function Msg-Box ($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg)
}
#---------------------------------------------------------[Form]--------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                    = New-Object system.Windows.Forms.Form
$Form.ClientSize         = '570,300'
$Form.text               = "UHS 3rd Party Applications Analyst Launcher"
$Form.BackColor          = "#ffffff"
$Form.TopMost            = $true

$lblSelectRpt                 = New-Object system.Windows.Forms.Label
$lblSelectRpt.text            = "Select Report"
$lblSelectRpt.AutoSize        = $true
$lblSelectRpt.width           = 25
$lblSelectRpt.height          = 10
$lblSelectRpt.location = New-Object System.Drawing.Point(20,$LineLocation[0])
$lblSelectRpt.Font            = 'Microsoft Sans Serif,13'

$cboReports                     = New-Object system.Windows.Forms.ComboBox
$cboReports.text                = ""
$cboReports.width               = 300
$cboReports.height              = 25
(Get-ReportNames) | ForEach-Object {[void] $cboReports.Items.Add($_)}
#$cboReports.SelectedIndex       = 0
$cboReports.location   = New-Object System.Drawing.Point(20,$LineLocation[1])
$cboReports.Font                = 'Microsoft Sans Serif,10'
$cboReports.Visible             = $true

$lblField1                      = New-Object system.Windows.Forms.Label
$lblField2                      = New-Object system.Windows.Forms.Label
$lblField3                      = New-Object system.Windows.Forms.Label
$lblField4                      = New-Object system.Windows.Forms.Label
$lblField5                      = New-Object system.Windows.Forms.Label
$lblField6                      = New-Object system.Windows.Forms.Label

$lblField1.AutoSize        = $false
$lblField2.AutoSize        = $false
$lblField3.AutoSize        = $false
$lblField4.AutoSize        = $false
$lblField5.AutoSize        = $false
$lblField6.AutoSize        = $false

$lblField1.width           = 200
$lblField2.width           = 200
$lblField3.width           = 200
$lblField4.width           = 200
$lblField5.width           = 200
$lblField6.width           = 200

$lblField1.height          = 25
$lblField2.height          = 25
$lblField3.height          = 25
$lblField4.height          = 25
$lblField5.height          = 25
$lblField6.height          = 25

$lblField1.Font            = 'Microsoft Sans Serif,13'
$lblField2.Font            = 'Microsoft Sans Serif,13'
$lblField3.Font            = 'Microsoft Sans Serif,13'
$lblField4.Font            = 'Microsoft Sans Serif,13'
$lblField5.Font            = 'Microsoft Sans Serif,13'
$lblField6.Font            = 'Microsoft Sans Serif,13'

$lblField1.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblField2.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblField3.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblField4.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblField5.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblField6.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

$lblField1.text            = ''
$lblField2.text            = ''
$lblField3.text            = ''
$lblField4.text            = ''
$lblField5.text            = ''
$lblField6.text            = ''

$txtField1 = New-Object System.Windows.Forms.TextBox
$txtField2 = New-Object System.Windows.Forms.TextBox
$txtField3 = New-Object System.Windows.Forms.TextBox
$txtField4 = New-Object System.Windows.Forms.TextBox
$txtField5 = New-Object System.Windows.Forms.TextBox
$txtField6 = New-Object System.Windows.Forms.TextBox

$txtField1.multiline           = $false
$txtField2.multiline           = $false
$txtField3.multiline           = $false
$txtField4.multiline           = $false
$txtField5.multiline           = $false
$txtField6.multiline           = $false

$txtField1.width               = 314
$txtField2.width               = 314
$txtField3.width               = 314
$txtField4.width               = 314
$txtField5.width               = 314
$txtField6.width               = 314

$txtField1.height              = 25
$txtField2.height              = 25
$txtField3.height              = 25
$txtField4.height              = 25
$txtField5.height              = 25
$txtField6.height              = 25

$txtField1.location            = New-Object System.Drawing.Point(220,$LineLocation[2])
$txtField2.location            = New-Object System.Drawing.Point(220,$LineLocation[3])
$txtField3.location            = New-Object System.Drawing.Point(220,$LineLocation[4])
$txtField4.location            = New-Object System.Drawing.Point(220,$LineLocation[5])
$txtField5.location            = New-Object System.Drawing.Point(220,$LineLocation[6])
$txtField6.location            = New-Object System.Drawing.Point(220,$LineLocation[7])

$txtField1.Font                = 'Microsoft Sans Serif,10'
$txtField2.Font                = 'Microsoft Sans Serif,10'
$txtField3.Font                = 'Microsoft Sans Serif,10'
$txtField4.Font                = 'Microsoft Sans Serif,10'
$txtField5.Font                = 'Microsoft Sans Serif,10'
$txtField6.Font                = 'Microsoft Sans Serif,10'

$txtField1.Visible             = $true
$txtField2.Visible             = $true
$txtField3.Visible             = $true
$txtField4.Visible             = $true
$txtField5.Visible             = $true
$txtField6.Visible             = $true

$txtField1.Text = '111111111111111111111111111'
$txtField2.Text = '222222222222222222222222222'
$txtField3.Text = '333333333333333333333333333'
$txtField4.Text = '444444444444444444444444444'
$txtField5.Text = '555555555555555555555555555'
$txtField6.Text = '666666666666666666666666666'











$Form.controls.AddRange(@($lblSelectRpt,$cboReports,$lblField1,$lblField2,$lblField3,$lblField4,$lblField5,$lblField6,$txtField1,$txtField2,$txtField3,$txtField4,$txtField5,$txtField6))




$cboReports.Add_SelectedIndexChanged({Update-DisplayFields $cboReports.SelectedIndex})

[void]$Form.ShowDialog()

$Reports = @()
$Report = [PSCustomObject]@{
    
    Active = $true
    ID = if($Active){$Reports.Count}else{}
    Name   = 'Yearly Patient & Treatment Count'
    Path   = '\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\Yearly Patient and Treament count.rpt'
    Fields = [PSCustomObject]@{
        Report_Start_Date = 'YD1'
        Report_End_Date   = 'PME'
    }
}

function Add-Report {
    [CmdletBinding()]
    Param([string]$Name, [string]$Path)
}

$Name = 'Yearly Patient & Treatment Count'
$Path = '\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\Yearly Patient and Treament count.rpt'


$StartDates  = @()
$StartDates += 'YD1'
$StartDates += 'PYD1'
$StartDates += 'PMD1'

$EndDates = @()
$EndDates += 'PME'
$EndDates += 'PYPME'

$Machines = @()
$Machines += 'Synergy+TB'
$Machines += 'Synergy'
$Machines += 'TrueBeam'

$Locations = @()
$Locations += 'Simulator'

$Holiday = @(0,1,2,3)
$Weekend = @(0,1,2,3)





