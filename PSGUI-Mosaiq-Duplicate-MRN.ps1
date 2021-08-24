##########################################################################
## WPF MOSAIQ Duplicate MRN Checker                                     ##
## v1.0                                                                 ##
## Author: Roy Coutts                                                   ##
## Released: 26-Aug-2021                                                ##
##########################################################################

<#
.Synopsis
   Exposes the properties, methods and events of the built-in WPF controls commonly used in WPF Windows desktop applications
.Notes
   Do not run from an existing PowerShell console session as the script will close it.  Right-click the script and run with PowerShell.
#>

#region UserInterface
# Import AutoItX in order to use commands
Import-Module 'F:\UTILITIES\AutoItX\AutoItX.psd1'
# Load Assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Load Associated XAML Form
if ([Environment]::MachineName -eq 'PC11418') { 
    $path = ($pscommandpath).ToString().Replace('.ps1','.xaml') 
} else {
    $path = 'C:\Users\royco\Documents\Test-Merge\Merge.xaml'
}

# Define XAML code
[xml]$xaml = (Get-Content -Path $path) -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'

# Load XAML elements into a hash table
$script:hash = [hashtable]::Synchronized(@{})
$hash.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $hash.$($_.Name) = $hash.Window.FindName($_.Name)
}
#endregion

#region PopulateInitialData
$CPN   = 'Current Patient Name'
$CPMRN = 'Current Patient MRN'
$SPN   = 'Source Patient Name'
$SPMRN = 'Source Patient MRN'
$ML    = 'MatchLevel'
$FMQ   = 'FoundInMosaiq'
$Global:Records = @()
$Global:ActiveRecord = 0
$OpenFileInitialDirectory = "\\wmhfilesrv\\GROUPSHARES\\PACS\\Andrea and Kelly reports\\MR MERGES\\"

#endregion

#region Functions
function Show-Msg ($msg) {[System.Windows.Forms.MessageBox]::show($msg)}

function Unblock-FileProtection ($data) {
    $encryptFound = 0
    $encrypt = 'http://schemas.microsoft.com/office/2006/encryption'
    $data | ForEach-Object ({ if ($PSItem.IndexOf($encrypt) -gt -1) {$encryptFound++} })
    Update-Display 'PROTECTION' $encryptFound
    Return $encryptFound
}

function Open-File {
    # Get FilePath
    $filepath = ''
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $OpenFileInitialDirectory
    $OpenFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.RestoreDirectory = $true
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq 'OK') {$filepath = $OpenFileDialog.FileName}
    # Get Directory and Update SourceFilePath
    $directory = (Get-ChildItem -Path $filepath).DirectoryName
    Update-Display 'SOURCEFILEPATH' $directory
    # Get FileName and Update SourceFileName
    $filename = (Get-ChildItem -Path $filepath).Name
    Update-Display 'SOURCEFILENAME' $filename
    # Check if File is Password Protected
    $IsProtected = Unblock-FileProtection (Get-Content -Path $filepath)
    # Double Check Enter Password is not visible
    if ($IsProtected -eq $true)  {Update-Display 'ENTERPASSWORD' 'VISIBLE'}
    if ($IsProtected -eq $false) {Update-Display 'ENTERPASSWORD' 'COLLAPSED'}
}

function Get-ExcelPassword {
    # Symbol: Unlocked Lock
    $symbol = $Hash.FileProtectionButton.Content
    if ($symbol -eq [char]208) { 
        # Do Nothing - File is already unlocked 
    }
    # Symbol: Locked Lock
    if ($symbol -eq [char]207) {
        # Make Enter Password Visible
        Update-Display 'ENTERPASSWORD' 'VISIBLE'
    }
}

function Get-EnteredPassword {
    # Get Password
    $pw = $Hash.EnterPasswordInputTextBox.Text
    Update-Display 'ENTERPASSWORD' 'COLLAPSED'
    # Build Filepath to Excel Sheet
    $dn = $Hash.SourceFilePathLabel.Content
    $fn = $Hash.SourceFileNameLabel.Content
    $filepath = Join-Path -Path $dn -ChildPath $fn
    # Attempt Pulling Protected Data with Password
    $data = ''
    # Using Try/Catch since using the wrong password will freeze or crash the script
    try
    {
    # Script will 'Continue' past error and save error results in $excelerror
    $data = Import-Excel -Path $filepath -Password $pw -ErrorAction Continue -ErrorVariable excelerror
    }
    # This is the type of error that is thrown
    catch [System.Management.Automation.MethodInvocationException]
    {
    $msg = $excelerror[0].ToString()
    [System.Windows.Forms.MessageBox]::Show($msg)
    }
    if ($data) {
        $NewFileName = "COPY-" + ((Get-ChildItem -Path $filepath).Name)
        $NewFilePath = Join-Path -Path ((Get-ChildItem -Path $filepath).DirectoryName) -ChildPath $NewFileName
        Export-Excel -Path $NewFilePath -InputObject $data
        Update-Display 'SOURCEFILEPATH' ((Get-ChildItem -Path $filepath).DirectoryName)
        Update-Display 'SOURCEFILENAME' $NewFileName
        Update-Display 'PROTECTION' $false
    }
    if ($excelerror) {
        Update-Display 'INVALIDPASSWORD' 'VISIBLE'
    }
}

function Update-Display ($display, $content) {
    if ($display -eq 'SOURCEFILEPATH')    {$Hash.SourceFilePathLabel.Content    = $content}
    if ($display -eq 'SOURCEFILENAME')    {$Hash.SourceFileNameLabel.Content    = $content}
    if ($display -eq 'RECORDSCOUNT')      {$Hash.RecordsCountLabel.Content      = $content}
    if ($display -eq 'FULLMATCHCOUNT')    {$Hash.FullMatchCountLabel.Content    = $content}
    if ($display -eq 'PARTIALMATCHCOUNT') {$Hash.PartialMatchCountLabel.Content = $content}
    if ($display -eq 'MISMATCHCOUNT')     {$Hash.MisMatchCountLabel.Content     = $content}
    if ($display -eq 'ACTIVERECORD')      {$Hash.RecordCountLabel.Content       = $content}
    if ($display -eq 'CURRENTPATIENT')    {$Hash.CurrentNameLabel.Content       = $content}
    if ($display -eq 'CURRENTMRN')        {$Hash.CurrentMRNLabel.Content        = $content}
    if ($display -eq 'SOURCEPATIENT')     {$Hash.SourceNameLabel.Content        = $content}
    if ($display -eq 'SOURCEMRN')         {$Hash.SourceMRNLabel.Content         = $content}
    if ($display -eq 'MATCHCOLORS')       {
                                           if ($content -eq 'FullMatch')    {$Hash.CurrentNameLabel.Background = "LawnGreen"
                                                                             $Hash.SourceNameLabel.Background  = "LawnGreen"}
                                           if ($content -eq 'PartialMatch') {$Hash.CurrentNameLabel.Background = "Yellow"
                                                                             $Hash.SourceNameLabel.Background  = "Yellow"}
                                           if ($content -eq 'MisMatch')     {$Hash.CurrentNameLabel.Background = "Red"
                                                                             $Hash.SourceNameLabel.Background  = "Red"}
    }
    if ($display -eq 'PREVIOUSRECORD')    {$Hash.RecordCountLabel.Content       = $content}
    if ($display -eq 'NEXTRECORD')        {$Hash.RecordCountLabel.Content       = $content}
    if ($display -eq 'PROTECTION')        {
                                           if ($content -gt 0) {
                                               $Hash.FileProtectionBorder.BorderBrush = "Red"
                                               $Hash.FileProtectionButton.Background  = "Red"
                                               $Hash.FileProtectionButton.Content     = [char]207     # LOCK SYMBOL
                                           } else {
                                               $Hash.FileProtectionBorder.BorderBrush = "DodgerBlue"
                                               $Hash.FileProtectionButton.Background  = "DodgerBlue"
                                               $Hash.FileProtectionButton.Content     = [char]208     # UNLOCK SYMBOL
                                           }
    }
    if ($display -eq 'ENTERPASSWORD')      {
                                            if ($content -eq 'VISIBLE') {
                                                $Hash.EnterPasswordLabel.Content               = "ENTER PASSWORD"
                                                $Hash.EnterPasswordWrapInputBorder.Visibility  = "Visible"
                                                $Hash.EnterPasswordOKButton.Visibility         = "Visible"
                                                $Hash.EnterPasswordOKBorder.Visibility         = "Visible"
                                                $Hash.EnterPasswordLabel.Visibility            = "Visible"
                                                $Hash.EnterPasswordBorder.Visibility           = "Visible"
                                                $Hash.EnterPasswordInputTextBox.Visibility     = "Visible"
                                                $Hash.EnterPasswordInputBorder.Visibility      = "Visible"
                                                $Hash.EnterPasswordWrapInputBorder.BorderBrush = "Yellow"
                                                $Hash.EnterPasswordOKButton.Background         = "Yellow"
                                                $Hash.EnterPasswordOKBorder.BorderBrush        = "Yellow"
                                                $Hash.EnterPasswordLabel.Background            = "Yellow"
                                                $Hash.EnterPasswordInputTextBox.Foreground     = "Yellow"
                                                $Hash.EnterPasswordInputBorder.BorderBrush     = "Yellow"
                                           } else {
                                                $Hash.EnterPasswordLabel.Content               = "ENTER PASSWORD"
                                                $Hash.EnterPasswordWrapInputBorder.Visibility  = "Collapsed"
                                                $Hash.EnterPasswordOKButton.Visibility         = "Collapsed"
                                                $Hash.EnterPasswordOKBorder.Visibility         = "Collapsed"
                                                $Hash.EnterPasswordLabel.Visibility            = "Collapsed"
                                                $Hash.EnterPasswordBorder.Visibility           = "Collapsed"
                                                $Hash.EnterPasswordInputTextBox.Visibility     = "Collapsed"
                                                $Hash.EnterPasswordInputBorder.Visibility      = "Collapsed"
                                           }
    }
    if ($display -eq 'INVALIDPASSWORD')    {
                                            if ($content -eq 'VISIBLE') {
                                                $Hash.EnterPasswordLabel.Content               = "INVALID PASSWORD"
                                                $Hash.EnterPasswordWrapInputBorder.Visibility  = "Visible"
                                                $Hash.EnterPasswordOKButton.Visibility         = "Visible"
                                                $Hash.EnterPasswordOKBorder.Visibility         = "Visible"
                                                $Hash.EnterPasswordLabel.Visibility            = "Visible"
                                                $Hash.EnterPasswordBorder.Visibility           = "Visible"
                                                $Hash.EnterPasswordInputTextBox.Visibility     = "Visible"
                                                $Hash.EnterPasswordInputBorder.Visibility      = "Visible"
                                                $Hash.EnterPasswordWrapInputBorder.BorderBrush = "Red"
                                                $Hash.EnterPasswordOKButton.Background         = "Red"
                                                $Hash.EnterPasswordOKBorder.BorderBrush        = "Red"
                                                $Hash.EnterPasswordLabel.Background            = "Red"
                                                $Hash.EnterPasswordInputTextBox.Foreground     = "Red"
                                                $Hash.EnterPasswordInputBorder.BorderBrush     = "Red"
                                           } else {
                                                $Hash.EnterPasswordLabel.Content               = "INVALID PASSWORD"
                                                $Hash.EnterPasswordWrapInputBorder.Visibility  = "Collapsed"
                                                $Hash.EnterPasswordOKButton.Visibility         = "Collapsed"
                                                $Hash.EnterPasswordOKBorder.Visibility         = "Collapsed"
                                                $Hash.EnterPasswordLabel.Visibility            = "Collapsed"
                                                $Hash.EnterPasswordBorder.Visibility           = "Collapsed"
                                                $Hash.EnterPasswordInputTextBox.Visibility     = "Collapsed"
                                                $Hash.EnterPasswordInputBorder.Visibility      = "Collapsed"
                                           }
    }
}

function Update-ProgressBars {
    # Create Variable to Track ProgressBar Values
    $MisMatchProgress     = @{FoundInMosaiq = [int]0; MisMatch     = [int]0; Progress = [Single]0}
    $PartialMatchProgress = @{FoundInMosaiq = [int]0; PartialMatch = [int]0; Progress = [Single]0}
    $FullMatchProgress    = @{FoundInMosaiq = [int]0; FullMatch    = [int]0; Progress = [Single]0}
    $TotalProgress        = @{FoundInMosaiq = [int]0; Total        = [int]0; Progress = [Single]0}
    # Gather MisMatch Stats
    $MisMatchProgress.FoundInMosaiq = ($Global:Records.FoundInMosaiq -join '' -split 'MisMatch'         | Measure-Object | Select-Object -ExpandProperty Count) - 1
    $MisMatchProgress.MisMatch      = ($Global:Records.MatchLevel -join '' -split 'MisMatch'            | Measure-Object | Select-Object -ExpandProperty Count) - 1
    if ($MisMatchProgress.MisMatch -eq 0) {$MisMatchProgress.Progress = 0} else {$MisMatchProgress.Progress =  ($MisMatchProgress.FoundInMosaiq / $MisMatchProgress.MisMatch) * 100}
    # Gather PartialMatch Stats
    $PartialMatchProgress.FoundInMosaiq = ($Global:Records.FoundInMosaiq -join '' -split 'PartialMatch' | Measure-Object | Select-Object -ExpandProperty Count) - 1
    $PartialMatchProgress.PartialMatch  = ($Global:Records.MatchLevel -join '' -split 'PartialMatch'    | Measure-Object | Select-Object -ExpandProperty Count) - 1 
    if ($PartialMatchProgress.PartialMatch -eq 0) {$PartialMatchProgress.Progress = 0} else {$PartialMatchProgress.Progress = ($PartialMatchProgress.FoundInMosaiq / $PartialMatchProgress.PartialMatch) * 100}
    # Gather FullMatch Stats
    $FullMatchProgress.FoundInMosaiq = ($Global:Records.FoundInMosaiq -join '' -split 'FullMatch'       | Measure-Object | Select-Object -ExpandProperty Count) - 1
    $FullMatchProgress.FullMatch     = ($Global:Records.MatchLevel -join '' -split 'FullMatch'          | Measure-Object | Select-Object -ExpandProperty Count) - 1
    if ($FullMatchProgress.FullMatch -eq 0) {$FullMatchProgress.Progress = 0} else {$FullMatchProgress.Progress = ($FullMatchProgress.FoundInMosaiq / $FullMatchProgress.FullMatch) * 100}
    # Gather Total Stats
    $TotalProgress.FoundInMosaiq = ([int]$MisMatchProgress.FoundInMosaiq + [int]$PartialMatchProgress.FoundInMosaiq + [int]$FullMatchProgress.FoundInMosaiq)
    $TotalProgress.Total         = ([int]$MisMatchProgress.MisMatch      + [int]$PartialMatchProgress.PartialMatch  + [int]$FullMatchProgress.FullMatch)
    if ($TotalProgress.Total -eq 0) {$TotalProgress.Progress = 0} else {$TotalProgress.Progress = ($TotalProgress.FoundInMosaiq / $TotalProgress.Total) * 100}
    # Update Values to Progress Bars
    $Hash.MisMatchProgressBar.Value     = $MisMatchProgress.Progress
    $Hash.PartialMatchProgressBar.Value = $PartialMatchProgress.Progress
    $Hash.FullMatchProgressBar.Value    = $FullMatchProgress.Progress
    $Hash.RecordsProgressBar.Value      = $TotalProgress.Progress
}

function Get-RecordsCount {Update-Display 'RECORDSCOUNT' $Global:Records.Count}

function Get-MatchLevels {
    $MatchLevels = @("MisMatch","PartialMatch","FullMatch")
    (0..($Global:Records.Count-1)) | ForEach-Object({ 
        $match = 0
        $CLastName  = $Global:Records[$PSItem].$CPN.Split(',')[0].Trim()
        $SLastName  = $Global:Records[$PSItem].$SPN.Split(',')[0].Trim()
        $CFirstName = $Global:Records[$PSItem].$CPN.Split(',')[1].Trim()
        $SFirstName = $Global:Records[$PSItem].$SPN.Split(',')[1].Trim()
        if ($CLastName -eq $SLastName) {$match++; if ($CFirstName -eq $SFirstName) {$match++}}
        $Global:Records[$PSItem].MatchLevel = $MatchLevels[$match]
    })
}
function Get-FullMatchCount    {Update-Display 'FULLMATCHCOUNT' $Global:Records.MatchLevel.Where({ $_ -eq 'FullMatch' }).Count}

function Get-PartialMatchCount {Update-Display 'PARTIALMATCHCOUNT' $Global:Records.MatchLevel.Where({ $_ -eq 'PartialMatch' }).Count}

function Get-MisMatchCount     {Update-Display 'MISMATCHCOUNT' $Global:Records.MatchLevel.Where({ $_ -eq 'MisMatch' }).Count}

function Get-ActiveRecord      {Update-Display 'ACTIVERECORD' ($Global:ActiveRecord + 1)}

function Get-CurrentPatient    {
                                Update-Display 'CURRENTPATIENT' $Global:Records[$Global:ActiveRecord].$CPN
                                Update-Display 'CURRENTMRN' $Global:Records[$Global:ActiveRecord].$CPMRN 
}

function Get-SourcePatient     {
                                Update-Display 'SOURCEPATIENT' $Global:Records[$Global:ActiveRecord].$SPN
                                Update-Display 'SOURCEMRN' $Global:Records[$Global:ActiveRecord].$SPMRN 
}

function Get-MatchColors       {Update-Display 'MATCHCOLORS' $Global:Records[$Global:ActiveRecord].$ML}

function Get-PreviousRecord    {if ($Global:ActiveRecord -eq 0) {$Global:ActiveRecord = $Global:Records.GetUpperBound(0)} else {$Global:ActiveRecord--}
                                Update-Display 'PREVIOUSRECORD' $Global:ActiveRecord
                                Get-FullMatchCount
                                Get-PartialMatchCount
                                Get-MisMatchCount
                                Get-ActiveRecord
                                Get-CurrentPatient
                                Get-SourcePatient
                                Get-MatchColors
                                Update-ProgressBars
}

function Get-NextRecord        {if ($Global:ActiveRecord -eq $Global:Records.GetUpperBound(0)) {$Global:ActiveRecord = 0} else {$Global:ActiveRecord++}
                                Update-Display 'NEXTRECORD' $Global:ActiveRecord
                                Get-FullMatchCount
                                Get-PartialMatchCount
                                Get-MisMatchCount
                                Get-ActiveRecord
                                Get-CurrentPatient
                                Get-SourcePatient
                                Get-MatchColors
                                Update-ProgressBars
}

function Set-InMosaiq ($found) {
    $Result = @($found,$Global:Records[$Global:ActiveRecord].$ML)
    $Global:Records[$Global:ActiveRecord].$FMQ = $Result
    Get-NextRecord
}

function Read-Records {
    $dn = $Hash.SourceFilePathLabel.Content
    $fn = $Hash.SourceFileNameLabel.Content
    $filepath = Join-Path -Path $dn -ChildPath $fn
    $Global:Records = ''
    $Global:Records = Import-Excel -Path $filepath
    # Add New Columns for Tracking and Statistics
    $Global:Records | Add-Member -MemberType NoteProperty -Name MatchLevel -Value ''
    $Global:Records | Add-Member -MemberType NoteProperty -Name FoundInMosaiq -Value ''
    # Initial Read to Populate Form
    Get-RecordsCount
    Get-MatchLevels
    Get-FullMatchCount
    Get-PartialMatchCount
    Get-MisMatchCount
    Get-ActiveRecord
    Get-CurrentPatient
    Get-SourcePatient
    Get-MatchColors
    Update-ProgressBars
}

function Show-Process ($Process, [Switch]$Maximize) {
    # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindowasync
    # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
    #
    # SW_HIDE
    #     0    Hides the window and activates another window.
    # SW_SHOWNORMAL / SW_NORMAL
    #     1    Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
    # SW_SHOWMINIMIZED
    #     2    Activates the window and displays it as a minimized window.
    # SW_SHOWMAXIMIZED / SW_MAXIMIZE
    #     3    Activates the window and displays it as a maximized window.
    # SW_SHOWNOACTIVATE
    #     4    Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
    # SW_SHOW
    #     5    Activates the window and displays it in its current size and position.
    # SW_MINIMIZE
    #     6    Minimizes the specified window and activates the next top-level window in the Z order.
    # SW_SHOWMINNOACTIVE
    #     7    Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
    # SW_SHOWNA
    #     8    Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
    # SW_RESTORE
    #     9    Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    # SW_SHOWDEFAULT
    #    10    Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    # SW_FORCEMINIMIZE
    #    11    Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.

    $sig = '
        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
    '
    if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
    $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
    $hwnd = $Process.MainWindowHandle
    $null = $type::ShowWindowAsync($hwnd, $Mode)
    $null = $type::SetForegroundWindow($hwnd)
}

function Move-Process ($Process) {
    $Title = $Process.MainWindowTitle
    $X = 0
    $Y = 0

    Move-AU3Win -Title $Title -X $X -Y $Y
}

function Resize-Process ($Process) {
    $ScreenWidth  = ([System.Windows.Forms.Screen]::PrimaryScreen).WorkingArea.Width
    $ScreenHeight = ([System.Windows.Forms.Screen]::PrimaryScreen).WorkingArea.Height
    $HalfScreen = [math]::Round( ($ScreenWidth / 2) )
    Move-AU3Win -Title $Process.MainWindowTitle -Width $HalfScreen -Height $ScreenHeight
}

function Run-Report {
    Show-Process -Process (GET-PROCESS -Name crw32) ; Start-Sleep -Milliseconds 250
    $WinHandle = (GET-PROCESS -Name crw32).MainWindowHandle
    $CtrlHandle = Get-AU3ControlHandle -WinHandle $WinHandle -Control "[Class:AfxFrameOrView80u;Instance:1]"
    Send-AU3ControlKey -WinHandle $WinHandle -ControlHandle $CtrlHandle -Key "^r"
}

function Enter-Values {
    Show-Process -Process (GET-PROCESS -Name crw32) ; Start-Sleep -Milliseconds 250
    $EnterValuesWinHandle = Get-AU3WinHandle -Title "Enter Values"
    $EnterValuesCtrlHandle = Get-AU3ControlHandle -WinHandle 13573050 -Control "[Class:Internet Explorer_Server;Instance:1]"
    Invoke-AU3ControlClick -WinHandle $EnterValuesWinHandle -ControlHandle $EnterValuesCtrlHandle -X 116 -Y 64
    Send-AU3ControlKey -WinHandle $EnterValuesWinHandle -ControlHandle $EnterValuesCtrlHandle -Key "smith"
}

if ([System.Diagnostics.Process]::GetProcessesByName("crw32")) {
    # Do Nothing - Crystal Reports is already running
} else {
    # Start Crystal Reports and Load Report
    . ($ReportPath)
    Start-Sleep -Milliseconds 5000
}

Show-Process -Process (GET-PROCESS -Name crw32)                   ; Start-Sleep -Milliseconds 1000
Move-Process -Process (GET-PROCESS -Name crw32)                   ; Start-Sleep -Milliseconds 1000 
Resize-Process -Process (GET-PROCESS -Name crw32)                 ; Start-Sleep -Milliseconds 1000

Run-Report

Start-Process
#endregion

#region Event Handling
# Button Events
$Hash.LoadFileButton.Add_Click({                                             Open-File })
$Hash.EnterPasswordOKButton.Add_Click({                                      Get-EnteredPassword })
$hash.EnterPasswordInputTextBox.Add_KeyDown({ if ($PSItem.Key -eq 'Return') {Get-EnteredPassword} })
$Hash.ReadRecordsButton.Add_Click({                                          Read-Records })
$Hash.PreviousRecordButton.Add_Click({                                       Get-PreviousRecord })
$Hash.NextRecordButton.Add_Click({                                           Get-NextRecord })
$Hash.InMosaiqYesButton.Add_Click({                                          Set-InMosaiq $true })
$Hash.InMosaiqNoButton.Add_Click({                                           Set-InMosaiq $false })

$Hash.FileProtectionButton.Add_Click({})
$Hash.RenameFileButton.Add_Click({})
$Hash.QueryMosaiqButton.Add_Click({})

# Mouse Hover Events
$Hash.LoadFileButton.Add_MouseEnter({        $Hash.LoadFileButton.Cursor =        [System.Windows.Input.Cursors]::Hand })
$Hash.LoadFileButton.Add_MouseLeave({        $Hash.LoadFileButton.Cursor =        [System.Windows.Input.Cursors]::None })
$Hash.FileProtectionButton.Add_MouseEnter({  $Hash.FileProtectionButton.Cursor =  [System.Windows.Input.Cursors]::Hand })
$Hash.FileProtectionButton.Add_MouseLeave({  $Hash.FileProtectionButton.Cursor =  [System.Windows.Input.Cursors]::None })
$Hash.RenameFileButton.Add_MouseEnter({      $Hash.RenameFileButton.Cursor =      [System.Windows.Input.Cursors]::Hand })
$Hash.RenameFileButton.Add_MouseLeave({      $Hash.RenameFileButton.Cursor =      [System.Windows.Input.Cursors]::None })
$Hash.ReadRecordsButton.Add_MouseEnter({     $Hash.ReadRecordsButton.Cursor =     [System.Windows.Input.Cursors]::Hand })
$Hash.ReadRecordsButton.Add_MouseLeave({     $Hash.ReadRecordsButton.Cursor =     [System.Windows.Input.Cursors]::None })
$Hash.QueryMosaiqButton.Add_MouseEnter({     $Hash.QueryMosaiqButton.Cursor =     [System.Windows.Input.Cursors]::Hand })
$Hash.QueryMosaiqButton.Add_MouseLeave({     $Hash.QueryMosaiqButton.Cursor =     [System.Windows.Input.Cursors]::None })
$Hash.PreviousRecordButton.Add_MouseEnter({  $Hash.PreviousRecordButton.Cursor =  [System.Windows.Input.Cursors]::Hand })
$Hash.PreviousRecordButton.Add_MouseLeave({  $Hash.PreviousRecordButton.Cursor =  [System.Windows.Input.Cursors]::None })
$Hash.NextRecordButton.Add_MouseEnter({      $Hash.NextRecordButton.Cursor =      [System.Windows.Input.Cursors]::Hand })
$Hash.NextRecordButton.Add_MouseLeave({      $Hash.NextRecordButton.Cursor =      [System.Windows.Input.Cursors]::None })
$Hash.InMosaiqYesButton.Add_MouseEnter({     $Hash.InMosaiqYesButton.Cursor =     [System.Windows.Input.Cursors]::Hand })
$Hash.InMosaiqYesButton.Add_MouseLeave({     $Hash.InMosaiqYesButton.Cursor =     [System.Windows.Input.Cursors]::None })
$Hash.InMosaiqNoButton.Add_MouseEnter({      $Hash.InMosaiqNoButton.Cursor =      [System.Windows.Input.Cursors]::Hand })
$Hash.InMosaiqNoButton.Add_MouseLeave({      $Hash.InMosaiqNoButton.Cursor =      [System.Windows.Input.Cursors]::None })
$Hash.EnterPasswordOKButton.Add_MouseEnter({ $Hash.EnterPasswordOKButton.Cursor = [System.Windows.Input.Cursors]::Hand })
$Hash.EnterPasswordOKButton.Add_MouseLeave({ $Hash.EnterPasswordOKButton.Cursor = [System.Windows.Input.Cursors]::None })

#endregion

#region Display the UI
# Display Window
# If code is running in ISE, use ShowDialog()...
if ($psISE)
{
    $null = $Hash.window.Dispatcher.InvokeAsync{$Hash.Window.ShowDialog()}.Wait()
}
# ...otherwise run as an application
Else
{
    # Make PowerShell Disappear
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -Name Win32ShowWindowAsync -Namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
 
    $app = New-Object -TypeName Windows.Application
    $app.Run($Hash.Window)
}
#endregion
