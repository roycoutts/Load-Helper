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
    Return ([DateTime]::Today).Year
}
function Get-Month {Return ([DateTime]::Today).Month}
function Get-YD1 {
    if((Get-Month) -eq 1){$YR = (Get-Year)-1 ; $MO = 1 ; $DAY = 1}else{$YR = (Get-Year) ; $MO = 1 ; $DAY = 1}
    Return ([DateTime]::new($YR,$MO,$DAY).ToString("MM/dd/yyyy"))
}
function Get-PMD1 {
    if((Get-Month) -eq 1){$YR = (Get-Year)-1 ; $MO = 12 ; $DAY = 1}else{$YR = (Get-Year) ; $MO = (Get-Month)-1 ; $DAY = 1}
    Return ([DateTime]::new($YR,$MO,$DAY).ToString("MM/dd/yyyy"))
}
function Get-PME {
    if((Get-Month) -eq 1){$YR = (Get-Year)-1 ; $MO = 12 ; $DAY = [DateTime]::DaysInMonth($YR,$MO)}else{$YR = (Get-Year) ; $MO = (Get-Month)-1 ; $DAY = [DateTime]::DaysInMonth($YR,$MO)}
    Return ([DateTime]::new($YR,$MO,$DAY).ToString("MM/dd/yyyy"))
}

function Get-PYD1 {
    if((Get-Month) -eq 1){$YR = (Get-Year)-2 ; $MO = 1 ; $DAY = 1}else{$YR = (Get-Year)-1 ; $MO = 1 ; $DAY = 1}
    Return ([DateTime]::new($YR,$MO,$DAY).ToString("MM/dd/yyyy"))
}
function Get-PYPME {
    if((Get-Month) -eq 1){$YR = (Get-Year)-2 ; $MO = 12 ; $DAY = [DateTime]::DaysInMonth($YR,$MO)}else{$YR = (Get-Year)-1 ; $MO = (Get-Month)-1 ; $DAY = [DateTime]::DaysInMonth($YR,$MO)}
    Return ([DateTime]::new($YR,$MO,$DAY).ToString("MM/dd/yyyy"))
}
function Get-MachineName ($Mnemonic) {
    If ($Mnemonic.ToUpper() -eq 'C') {$MachineName = 'Synergy+TB'}
    If ($Mnemonic.ToUpper() -eq 'S') {$MachineName = 'Synergy'}
    If ($Mnemonic.ToUpper() -eq 'T') {$MachineName = 'TrueBeam'}
    Return $MachineName
}

function Update-DisplayFields ($ReportIndex) {
    $lblField1.Visible = $false ; $lblField2.Visible = $false ; $lblField3.Visible = $false ; $lblField4.Visible = $false ; $lblField5.Visible = $false ; $lblField6.Visible = $false
    $txtField1.Visible = $false ; $txtField2.Visible = $false ; $txtField3.Visible = $false ; $txtField4.Visible = $false
    $cboField1.Visible = $false ; $cboField2.Visible = $false ; $cboField3.Visible = $false

    if ($ReportIndex -eq 0) {    # Yearly Patient & Treatment Count
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1]   ; $lblField2.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-YD1)             ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-PME)             ; $txtField2.Visible = $true
    }
    if ($ReportIndex -eq 1) {    # Combined Analytics Report
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2]   ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4]   ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5]   ; $lblField4.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-PMD1)            ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-MachineName 'C') ; $txtField2.Visible = $true
        $txtField3.Location    = New-Object System.Drawing.Point(230,$LineLocation[4]); $txtField3.Text = '0'                   ; $txtField3.Visible = $true
        $txtField4.Location    = New-Object System.Drawing.Point(230,$LineLocation[5]); $txtField4.Text = '0'                   ; $txtField4.Visible = $true
    }
    if ($ReportIndex -eq 2) {    # Synergy Analytics Report
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2]   ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4]   ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5]   ; $lblField4.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-PMD1)            ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-MachineName 'S') ; $txtField2.Visible = $true
        $txtField3.Location    = New-Object System.Drawing.Point(230,$LineLocation[4]); $txtField3.Text = '0'                   ; $txtField3.Visible = $true
        $txtField4.Location    = New-Object System.Drawing.Point(230,$LineLocation[5]); $txtField4.Text = '0'                   ; $txtField4.Visible = $true
    }
    if ($ReportIndex -eq 3) {    # TrueBeam Analytics Report
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[2]   ; $lblField2.Visible = $true 
        $lblField3.location    = New-Object System.Drawing.Point(20,$LineLocation[4]) ; $lblField3.Text = (Get-FieldNames)[4]   ; $lblField3.Visible = $true 
        $lblField4.location    = New-Object System.Drawing.Point(20,$LineLocation[5]) ; $lblField4.Text = (Get-FieldNames)[5]   ; $lblField4.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-PMD1)            ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-MachineName 'T') ; $txtField2.Visible = $true
        $txtField3.Location    = New-Object System.Drawing.Point(230,$LineLocation[4]); $txtField3.Text = '0'                   ; $txtField3.Visible = $true
        $txtField4.Location    = New-Object System.Drawing.Point(230,$LineLocation[5]); $txtField4.Text = '0'                   ; $txtField4.Visible = $true
    }
    if ($ReportIndex -eq 4) {    # Sims Analytics Report
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[3]   ; $lblField2.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-PMD1)            ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = 'Simulator'           ; $txtField2.Visible = $true
    }
    if ($ReportIndex -eq 5) {    # SRS Data
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1]   ; $lblField2.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-YD1)             ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-PME)             ; $txtField2.Visible = $true
    }
    if ($ReportIndex -eq 6) {    # Diagnosis Analytics Report (Current Year)
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1]   ; $lblField2.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-YD1)             ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-PME)             ; $txtField2.Visible = $true
    }
    if ($ReportIndex -eq 7) {    # Diagnosis Analytics Report (Previous Year)
        # LEFT COLUMN
        $lblField1.location    = New-Object System.Drawing.Point(20,$LineLocation[2]) ; $lblField1.Text = (Get-FieldNames)[0]   ; $lblField1.Visible = $true
        $lblField2.location    = New-Object System.Drawing.Point(20,$LineLocation[3]) ; $lblField2.Text = (Get-FieldNames)[1]   ; $lblField2.Visible = $true 
        # RIGHT COLUMN
        $txtField1.Location    = New-Object System.Drawing.Point(230,$LineLocation[2]); $txtField1.Text = (Get-PYD1)            ; $txtField1.Visible = $true
        $txtField2.Location    = New-Object System.Drawing.Point(230,$LineLocation[3]); $txtField2.Text = (Get-PYPME)           ; $txtField2.Visible = $true
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

# LEFT COLUMN
$lblField1 = New-Object system.Windows.Forms.Label
$lblField2 = New-Object system.Windows.Forms.Label
$lblField3 = New-Object system.Windows.Forms.Label
$lblField4 = New-Object system.Windows.Forms.Label
$lblField5 = New-Object system.Windows.Forms.Label
$lblField6 = New-Object system.Windows.Forms.Label

$lblField1.AutoSize = $false ; $lblField1.width = 200 ; $lblField1.height = 25 ; $lblField1.Font = 'Microsoft Sans Serif,13' ; $lblField1.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField1.text = ''
$lblField2.AutoSize = $false ; $lblField2.width = 200 ; $lblField2.height = 25 ; $lblField2.Font = 'Microsoft Sans Serif,13' ; $lblField2.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField2.text = ''
$lblField3.AutoSize = $false ; $lblField3.width = 200 ; $lblField3.height = 25 ; $lblField3.Font = 'Microsoft Sans Serif,13' ; $lblField3.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField3.text = ''
$lblField4.AutoSize = $false ; $lblField4.width = 200 ; $lblField4.height = 25 ; $lblField4.Font = 'Microsoft Sans Serif,13' ; $lblField4.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField4.text = ''
$lblField5.AutoSize = $false ; $lblField5.width = 200 ; $lblField5.height = 25 ; $lblField5.Font = 'Microsoft Sans Serif,13' ; $lblField5.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField5.text = ''
$lblField6.AutoSize = $false ; $lblField6.width = 200 ; $lblField6.height = 25 ; $lblField6.Font = 'Microsoft Sans Serif,13' ; $lblField6.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight ; $lblField6.text = ''

# RIGHT COLUMN
$txtField1 = New-Object System.Windows.Forms.TextBox
$txtField2 = New-Object System.Windows.Forms.TextBox
$txtField3 = New-Object System.Windows.Forms.TextBox
$txtField4 = New-Object System.Windows.Forms.TextBox
$cboField1 = New-Object System.Windows.Forms.ComboBox
$cboField2 = New-Object System.Windows.Forms.ComboBox
$cboField3 = New-Object System.Windows.Forms.ComboBox

$txtField1.multiline = $false ; $txtField1.width = 314 ; $txtField1.height = 25 ; $txtField1.Font = 'Microsoft Sans Serif,10' ; $txtField1.Visible = $false ; $txtField1.Text = '111111111111111111111111111'
$txtField2.multiline = $false ; $txtField2.width = 314 ; $txtField2.height = 25 ; $txtField2.Font = 'Microsoft Sans Serif,10' ; $txtField2.Visible = $false ; $txtField2.Text = '222222222222222222222222222'
$txtField3.multiline = $false ; $txtField3.width = 314 ; $txtField3.height = 25 ; $txtField3.Font = 'Microsoft Sans Serif,10' ; $txtField3.Visible = $false ; $txtField3.Text = '333333333333333333333333333'
$txtField4.multiline = $false ; $txtField4.width = 314 ; $txtField4.height = 25 ; $txtField4.Font = 'Microsoft Sans Serif,10' ; $txtField4.Visible = $false ; $txtField4.Text = '444444444444444444444444444'
$cboField1.Width = 300 ; $cboField1.Height = 25 ; $cboField1.Font = 'Microsoft Sans Serif,10' ; $cboField1.Visible = $false ; $cboField1.text = ''
$cboField2.Width = 300 ; $cboField2.Height = 25 ; $cboField2.Font = 'Microsoft Sans Serif,10' ; $cboField2.Visible = $false ; $cboField2.text = ''
$cboField3.Width = 300 ; $cboField3.Height = 25 ; $cboField3.Font = 'Microsoft Sans Serif,10' ; $cboField3.Visible = $false ; $cboField3.text = ''

$Form.controls.AddRange(@($lblSelectRpt,$cboReports,$lblField1,$lblField2,$lblField3,$lblField4,$lblField5,$lblField6,$txtField1,$txtField2,$txtField3,$txtField4,$cboField1,$cboField2,$cboField3))

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







