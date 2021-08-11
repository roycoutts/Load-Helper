$ScriptPath = $PSCommandPath
$test=$true
$inipath = 'C:\Temp\NewMosaiqMerge.ini'
$defaultDirectory = '\\wmhfilesrv\GROUPSHARES\PACS\Andrea and Kelly reports\MR MERGES'
$Global:ExcelPW = ''

$Label1Prefix = '     File Source: '
$Label2Prefix = '        Password: '
$Label3Prefix = 'File In Progress: '
$Label4Prefix = '   File Complete: '
$Label5Prefix = '         Records: '



#$data = Import-Excel $pathToExcelFile -Password 'pass'


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

function Return-Password ($pwd) {
    $Global:ExcelPW = $pwd
    Return $Global:ExcelPW
}

function Prompt-Password {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # INITIALIZE FORM
    $FormPW      = New-Object System.Windows.Forms.Form
    $FormPW.Size = New-Object System.Drawing.Size(370,90)
    $FormPW.Font = New-Object System.Drawing.Font("Century Gothic",10)
    $FormPW.Text = 'Enter Excel Spreadsheet Password'
    $FormPW.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent

    $TextBoxPW = New-Object System.Windows.Forms.TextBox
    $TextBoxPW.PasswordChar = '*'
    $TextBoxPW.Size = New-Object System.Drawing.Size(200,20)
    $TextBoxPW.Location = New-Object System.Drawing.Point(10,20)
    $ButtonPW = New-Object System.Windows.Forms.Button
    $ButtonPW.Text = 'OK'
    $ButtonPW.Size = New-Object System.Drawing.Size(75,24)
    $ButtonPW.Location = New-Object System.Drawing.Point(220,20)
    
    $FormPW.AcceptButton = $ButtonPW
    $FormPW.controls.AddRange(@($TextBoxPW,$ButtonPW))
    $ButtonPW.Add_Click({ Return-Password $TextBoxPW.Text; $FormPW.Close() })
    $FormPW.ShowDialog()

    Update-Content 'Password' $Global:ExcelPW
}


function New-Ini {
    if ($test) {} else {$IniPath = $PSCommandPath -replace '.ps1','.ini'}
    $content = @("FileSource=","Password=","FileInProgress=","FileComplete=","Records=")
    Set-Content -Path $IniPath -Value $content -Force
}

function Load-Ini {
    if ($test) {} else {$IniPath = $PSCommandPath -replace '.ps1','.ini'}
    if (!(Test-Path -Path $IniPath)) { 
        New-Ini
    }
    if (Test-Path -Path $IniPath) { 
        $IniContent = Get-Content -Path $IniPath 
        $Content = [PSCustomObject]@{}
        $IniContent | ForEach ({  
            $k = $PSItem.Split('=')[0]
            $v = $PSItem.Split('=')[1]
            $Content | Add-Member -MemberType NoteProperty -Name $k -Value $v
        })
    }
    Return $Content
}

function Update-Ini ($content) {
    $NewContent = @()
    $NewContent += "FileSource="     + $content.FileSource
    $NewContent += "Password="       + $content.Password
    $NewContent += "FileInProgress=" + $content.FileInProgress
    $NewContent += "FileComplete="   + $content.FileComplete
    $NewContent += "Records="        + $content.Records           
    Set-Content -Path $IniPath -Value $NewContent -Force
}

function Open-FileDialog {
    $OpenFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog1.InitialDirectory = $defaultDirectory
    $ShowDialog = $openfiledialog1.ShowDialog()
    If ($ShowDialog -eq "OK") {$FilePath = $openFileDialog1.FileName} else {$FilePath = ''}
    Return $FilePath
}

function Update-Content ($key,$update) {
    # Load Ini
    $Content = Load-Ini

    if ($key -eq 'FileSource') {$Content.FileSource = $update; Set-LabelText 1 $update; Update-Ini $Content}
    if ($key -eq 'Password')   {$Content.Password   = $update; Set-LabelText 2 $update; Update-Ini $Content}

}

function New-FileSource {
    # Create Blank Ini
    New-Ini
    # Select File Path to Excel Spreadsheet
    $FilePath = Open-FileDialog
    Update-Content 'FileSource' (Get-ChildItem -Path $FilePath).FullName
    Get-PasswordProtectedStatus
    Resize-Controls
}

function Get-PasswordProtectedStatus {
    $path = $Label1.Text.Replace($Label1Prefix,'') 
    Import-Excel -Path $path -ErrorAction SilentlyContinue -ErrorVariable pwd
    Write-host $pwd.Count
    if ($pwd.Count -eq 0) {
        Update-Content 'Password' 'None'
    } else {
        if ($pwd[0].ToString().Contains('If the file is encrypted')) {
            Update-Content 'Password' 'Protected'
        }     
    }
    Resize-Controls
}

function Set-LabelText ($num,$txt) {
    if ($num -eq 1) {$Label1.Text = $Label1Prefix + $txt}
    if ($num -eq 2) {$Label2.Text = $Label2Prefix + $txt}
    if ($num -eq 3) {$Label3.Text = $Label3Prefix + $txt}
    if ($num -eq 4) {$Label4.Text = $Label4Prefix + $txt}
    if ($num -eq 5) {$Label5.Text = $Label5Prefix + $txt}
}

function Resize-Controls {
    # SET WIDTH TO WIDEST BUTTON / LABEL
    $ButtonPreferredWidth  = ($Button1.PreferredSize.Width, $Button2.PreferredSize.Width, $Button3.PreferredSize.Width, $Button4.PreferredSize.Width, $Button5.PreferredSize.Width  | measure -Maximum).Maximum
    $LabelPreferredWidth  = ($Label1.PreferredSize.Width, $Label2.PreferredSize.Width, $Label3.PreferredSize.Width, $Label4.PreferredSize.Width, $Label5.PreferredSize.Width  | measure -Maximum).Maximum
    # SET HEIGHT TO TALLEST BUTTON OR LABEL
    $ButtonPreferredHeight = ($Button1.PreferredSize.Height,$Button2.PreferredSize.Height,$Button3.PreferredSize.Height,$Button4.PreferredSize.Height,$Button5.PreferredSize.Height | measure -Maximum).Maximum
    $LabelPreferredHeight  = ($Label1.PreferredSize.Height, $Label2.PreferredSize.Height, $Label3.PreferredSize.Height, $Label4.PreferredSize.Height, $Label5.PreferredSize.Height  | measure -Maximum).Maximum
    $PreferredHeight = ($ButtonPreferredHeight,$LabelPreferredHeight | measure -Maximum).Maximum
    # SET SIZE
    $Button1.Size = New-Object System.Drawing.Size($ButtonPreferredWidth,$PreferredHeight)
    $Button2.Size = New-Object System.Drawing.Size($ButtonPreferredWidth,$PreferredHeight)
    $Button3.Size = New-Object System.Drawing.Size($ButtonPreferredWidth,$PreferredHeight)
    $Button4.Size = New-Object System.Drawing.Size($ButtonPreferredWidth,$PreferredHeight)
    $Button5.Size = New-Object System.Drawing.Size($ButtonPreferredWidth,$PreferredHeight)
    $Label1.Size  = New-Object System.Drawing.Size($LabelPreferredWidth, $PreferredHeight)
    $Label2.Size  = New-Object System.Drawing.Size($LabelPreferredWidth, $PreferredHeight)
    $Label3.Size  = New-Object System.Drawing.Size($LabelPreferredWidth, $PreferredHeight)
    $Label4.Size  = New-Object System.Drawing.Size($LabelPreferredWidth, $PreferredHeight)
    $Label5.Size  = New-Object System.Drawing.Size($LabelPreferredWidth, $PreferredHeight)
    # SET GRID POINTS
    $x0 = 10
    $x1 = $x0 + $ButtonPreferredWidth + 10
    $y0 = 25
    $y1 = $y0 + $PreferredHeight
    $y2 = $y1 + $PreferredHeight
    $y3 = $y2 + $PreferredHeight
    $y4 = $y3 + $PreferredHeight
    # SET LOCATION
    $Button1.Location = New-Object System.Drawing.Point($x0,$y0)
    $Button2.Location = New-Object System.Drawing.Point($x0,$y1)
    $Button3.Location = New-Object System.Drawing.Point($x0,$y2)
    $Button4.Location = New-Object System.Drawing.Point($x0,$y3)
    $Button5.Location = New-Object System.Drawing.Point($x0,$y4)
    $Label1.Location  = New-Object System.Drawing.Point($x1,$y0)
    $Label2.Location  = New-Object System.Drawing.Point($x1,$y1)
    $Label3.Location  = New-Object System.Drawing.Point($x1,$y2)
    $Label4.Location  = New-Object System.Drawing.Point($x1,$y3)
    $Label5.Location  = New-Object System.Drawing.Point($x1,$y4)

    $GroupBox1.Size = $GroupBox1.PreferredSize
    $Form.size = $Form.PreferredSize
}

$Content = Load-Ini



# INITIALIZE FORM
$Form      = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(300,200)
$Form.Font = New-Object System.Drawing.Font("Century Gothic",10)

# INITIALIZE CONTROLS
$Button1 = New-Object System.Windows.Forms.Button
$Button2 = New-Object System.Windows.Forms.Button
$Button3 = New-Object System.Windows.Forms.Button
$Button4 = New-Object System.Windows.Forms.Button
$Button5 = New-Object System.Windows.Forms.Button
$Label1  = New-Object System.Windows.Forms.Label
$Label2  = New-Object System.Windows.Forms.Label
$Label3  = New-Object System.Windows.Forms.Label
$Label4  = New-Object System.Windows.Forms.Label
$Label5  = New-Object System.Windows.Forms.Label
# SET FONT
$Button1.Font = New-Object System.Drawing.Font("Consolas",10)
$Button2.Font = New-Object System.Drawing.Font("Consolas",10)
$Button3.Font = New-Object System.Drawing.Font("Consolas",10)
$Button4.Font = New-Object System.Drawing.Font("Consolas",10)
$Button5.Font = New-Object System.Drawing.Font("Consolas",10)
$Label1.Font  = New-Object System.Drawing.Font("Consolas",10)
$Label2.Font  = New-Object System.Drawing.Font("Consolas",10)
$Label3.Font  = New-Object System.Drawing.Font("Consolas",10)
$Label4.Font  = New-Object System.Drawing.Font("Consolas",10)
$Label5.Font  = New-Object System.Drawing.Font("Consolas",10)
# SET TEXT
$Button1.Text = 'New'
$Button2.Text = 'Remove'
$Button3.Text = 'Start'
$Button4.Text = 'Finish'
$Button5.Text = 'New'
$Label1.Text  = $Label1Prefix + $Content.FileSource
$Label2.Text  = $Label2Prefix + $Content.Password
$Label3.Text  = $Label3Prefix + $Content.FileInProgress
$Label4.Text  = $Label4Prefix + $Content.FileComplete
$Label5.Text  = $Label5Prefix + $Content.Records
# SET TEXT ALIGNMENT
$Label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label2.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label3.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label4.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$Label5.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft




$GroupBox1 = New-Object System.Windows.Forms.GroupBox
$GroupBox1.Text = 'Source File'
$GroupBox1.Location = New-Object System.Drawing.Point(10,10)
$GroupBox1.Controls.AddRange(@($Button1,$Button2,$Button3,$Button4,$Button5,$Label1,$Label2,$Label3,$Label4,$Label5))

$Form.controls.AddRange(@($GroupBox1))



$Button1.Add_Click({ New-FileSource })
$Button2.Add_Click({ Prompt-Password })


Resize-Controls

$form.ShowDialog()



