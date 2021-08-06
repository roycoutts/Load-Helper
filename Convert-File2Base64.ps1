Using Namespace System.Windows.Window
Add-Type -AssemblyName PresentationFramework

$Window = New-Object System.Windows.Window
$Window.Title = "Convert File2Base64"
$Window.Height = "450"
$Window.Width = "800"
$Window.WindowStyle = "ToolWindow"

$Grid = New-Object System.Windows.Controls.Grid
$txtFilePath = New-Object System.Windows.Controls.TextBox
$txtBase64   = New-Object System.Windows.Controls.TextBox
$btnGetFile  = New-Object System.Windows.Controls.Button
$btnSaveFile = New-Object System.Windows.Controls.Button


$Grid.Background = "#FF191919"

$txtFilePath.Name = "txtFilePath" 
$txtFilePath.HorizontalAlignment="Left" 
$txtFilePath.Height="30" 
$txtFilePath.Margin="10,10,0,0" 
$txtFilePath.TextWrapping="Wrap" 
$txtFilePath.Text="Select File to Covert to Base64" 
$txtFilePath.VerticalAlignment="Top" 
$txtFilePath.Width="500" 
$txtFilePath.Background="#FF333337" 
$txtFilePath.Foreground="White" 
$txtFilePath.BorderBrush="#FF434346" 
$txtFilePath.SelectionBrush="#FF3399FF" 
$txtFilePath.FontSize="16" 
$txtFilePath.Padding="10,0,0,0"

$btnGetFile.Name = "btnGetFile" 
$btnGetFile.Content="•••" 
$btnGetFile.HorizontalAlignment="Left" 
$btnGetFile.Margin="515,10,0,0" 
$btnGetFile.VerticalContentAlignment="Center" 
$btnGetFile.VerticalAlignment="Top" 
$btnGetFile.Width="75" 
$btnGetFile.Background="#FF333337" 
$btnGetFile.Foreground="White" 
$btnGetFile.BorderBrush="#FF434346" 
$btnGetFile.FontSize="20" 
$btnGetFile.Height="30"

$txtBase64.Name = "txtBase64"
$txtBase64.HorizontalAlignment="Left" 
$txtBase64.Height="365" 
$txtBase64.Margin="10,45,0,0" 
$txtBase64.TextWrapping="Wrap" 
$txtBase64.Text="" 
$txtBase64.VerticalAlignment="Top" 
$txtBase64.Width="772" 
$txtBase64.Background="#FF333337" 
$txtBase64.Foreground="#FF0E82E2" 
$txtBase64.BorderBrush="#FF434346" 
$txtBase64.SelectionBrush="#FF3399FF" 
$txtBase64.FontFamily="Consolas" 
$txtBase64.FontSize="6"

$btnSaveFile.Name = "btnSaveFile" 
$btnSaveFile.Content="Save to File" 
$btnSaveFile.HorizontalAlignment="Left" 
$btnSaveFile.Margin="595,10,0,0" 
$btnSaveFile.VerticalContentAlignment="Center" 
$btnSaveFile.VerticalAlignment="Top" 
$btnSaveFile.Width="187" 
$btnSaveFile.Background="#FF333337" 
$btnSaveFile.Foreground="White" 
$btnSaveFile.BorderBrush="#FF434346" 
$btnSaveFile.FontSize="20" 
$btnSaveFile.Height="30"


$Grid.Children.Add($txtFilePath)
$Grid.Children.Add($txtBase64)
$Grid.Children.Add($btnGetFile)
$Grid.Children.Add($btnSaveFile)
$Window.AddChild($Grid)

function Get-File {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = "c:\\"
    $OpenFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 2
    $OpenFileDialog.RestoreDirectory = $true
    $ShowDialog = $OpenFileDialog.ShowDialog()
    if ($ShowDialog -eq "OK") {$txtFilePath.Text = $OpenFileDialog.FileName}
}

function Convert-FileToBase64 {
    $file           = $txtFilePath.Text
    $fileAsBase64   = [convert]::ToBase64String((Get-Content -path $file -Encoding byte))
    $txtBase64.Text = $fileAsBase64
}

function Save-File {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.AddExtension = (Get-ChildItem -Path ($txtFilePath.Text)).Extension
    $SaveFileDialog.DefaultExt = (Get-ChildItem -Path ($txtFilePath.Text)).Extension
    $SaveFileDialog.Filter = "{0} files (*.{0})|*.{0}|All files (*.*)|*.*" -f $SaveFileDialog.DefaultExt  
    $SaveFileDialog.FilterIndex = 2 
    $SaveFileDialog.RestoreDirectory = $true 
    $ShowDialog = $SaveFileDialog.ShowDialog()
    if ($ShowDialog -eq "OK") {
        $b64 = $txtBase64.Text
        $filename = $SaveFileDialog.FileName
        $bytes = [Convert]::FromBase64String($b64)
        [System.IO.File]::WriteAllBytes($filename, $bytes)
    }
    Write-Host $ShowDialog
    Write-Host $SaveFileDialog.FileName
    

}

$btnGetFile.Add_Click({ Get-File; Convert-FileToBase64 })
$btnSaveFile.Add_Click({ Save-File })

$Window.ShowDialog()
