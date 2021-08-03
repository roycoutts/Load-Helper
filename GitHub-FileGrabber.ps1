Add-Type -AssemblyName PresentationFramework

$LoadRepos_Url = 'https://api.github.com/users/{owner}/repos'
$LoadReposFiles_Url = 'https://api.github.com/repos/{owner}/{repos}/contents'
$Global:Repos = ''
$Global:ReposFiles = ''
$Global:FileContent = ''

function Load-Wrapper 
{
Return @"
<Window Name="Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="GitHub File Grabber" Height="450" Width="800">
    <Grid>
        <TextBox Name="txtOwner" HorizontalAlignment="Left" Height="23" Margin="10,10,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" HorizontalContentAlignment="Center"/>
        <Button Name="btnLoadRepos" Content="Load Repositories for Owner" HorizontalAlignment="Left" Margin="135,10,0,0" VerticalAlignment="Top" Width="200" IsEnabled="False"/>
        <ComboBox Name="cboRepos" HorizontalAlignment="Left" Margin="10,38,0,0" VerticalAlignment="Top" Width="220" RenderTransformOrigin="-1.18,-2.132" IsEnabled="False"/>
        <Button Name="btnLoadReposFiles" Content="Load Repos Files" HorizontalAlignment="Left" Margin="235,38,0,0" VerticalAlignment="Top" Width="100" IsEnabled="False"/>
        <ComboBox Name="cboFiles" HorizontalAlignment="Left" Margin="10,65,0,0" VerticalAlignment="Top" Width="220" IsEnabled="False"/>
        <Button Name="btnGetFile" Content="Get File" HorizontalAlignment="Left" Margin="235,65,0,0" VerticalAlignment="Top" Width="100" IsEnabled="False"/>
        <TextBox Name="txtFileContent" Margin="10,92,0,0" TextWrapping="Wrap" IsEnabled="False"/>
        <WebBrowser Name="webFileContent" Margin="351,10,10,10"/>
    </Grid>
</Window>
"@
}

$code = (Load-Wrapper)
[xml]$xaml = $code

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $reader )

$window.ToolTip = "This is a window. Cool, isn't it?"

$txtOwner          = $Window.FindName("txtOwner")
$btnLoadRepos      = $Window.FindName("btnLoadRepos")
$cboRepos          = $Window.FindName("cboRepos")
$btnLoadReposFiles = $Window.FindName("btnLoadReposFiles")
$cboFiles          = $Window.FindName("cboFiles")
$btnGetFile        = $Window.FindName("btnGetFile")
$txtFileContent    = $Window.FindName("txtFileContent")
$webFileContent    = $Window.FindName("webFileContent")

$txtOwner.Add_TextChanged({
    if ($txtOwner.Text.Length -eq 0) {$btnLoadRepos.IsEnabled = $false} else {$btnLoadRepos.IsEnabled = $true}
})

$btnLoadRepos.Add_Click({
    $cboRepos.items.Clear()
    $cboRepos.IsEnabled = $false
    $btnLoadReposFiles.IsEnabled = $false
    $cboFiles.items.Clear()
    $cboFiles.IsEnabled = $false
    $btnGetFile.IsEnabled = $false
    $url = $LoadRepos_Url -replace '{owner}',$txtOwner.Text
    $Global:Repos = (Invoke-WebRequest -uri $url) | ConvertFrom-Json
    $Global:Repos.ForEach{ $cboRepos.Items.Add( ($psitem).full_name ) }
    $cboRepos.IsEnabled = $true
    })

$cboRepos.Add_SelectionChanged({
    $btnLoadReposFiles.IsEnabled = $true
})

$btnLoadReposFiles.Add_Click({
    $cboFiles.items.Clear()
    $cboFiles.IsEnabled = $false
    $btnGetFile.IsEnabled = $false
    $url = $LoadReposFiles_Url -replace '{owner}/{repos}',$cboRepos.SelectedItem
    $Global:ReposFiles = (Invoke-WebRequest -uri $url) | ConvertFrom-Json
    $Global:ReposFiles.ForEach{ $cboFiles.Items.Add( ($PSItem).name ) }
    $cboFiles.IsEnabled = $true
})

$cboFiles.Add_SelectionChanged({
    $btnGetFile.IsEnabled = $true
})

$btnGetFile.Add_Click({
    
    $url = $Global:ReposFiles[($cboFiles.SelectedIndex)].download_url
    $Global:FileContent = (Invoke-WebRequest -uri $url) 
    $txtFileContent.Text = $Global:FileContent.Content
    $txtFileContent.IsEnabled = $true

    if ($Global:ReposFiles[0].name.EndsWith('.ps1')) {
        $Url = "http://hilite.me/api"
        $code = $Global:FileContent.Content
        $CodeCA = ($code -join [char]10)
        $Good = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-!*()_'.~"
        $html = "code="
        $CodeCA[0..($CodeCA.Length - 1)] | ForEach({ if ($Good.Contains($_)) {$html += $_ } else {$html += "%" + [convert]::ToString([int][char]$_,16).PadLeft(2,"0").ToUpper()} })
        $html += '&lexer=powershell&style=native'
        $a = Invoke-WebRequest -Uri $url -Method Post -Body code=$html
        $webfilecontent.NavigateToString($a.Content)
        $Window.Width = $Window.DesiredSize.Width
        $window.Height = $Window.DesiredSize.Height
    }
})

$Window.ShowDialog()
