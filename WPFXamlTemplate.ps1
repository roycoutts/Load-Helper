###############################################################################################################
#Description  :  Powershell WPF - Add/Remove Items From List
#Author : Florian Clisson
###############################################################################################################

#Xamal Loader section has been wrote by stephen owen :
#https://github.com/1RedOne/PowerShell_XAML/

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window Name="Window1" x:Class="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfTemplate"
        mc:Ignorable="d"
        Title="MainWindow" Height="500" Width="800">
    <Grid Name="Grid1">
        <Menu Name="Menu1" HorizontalAlignment="Left" Height="36" VerticalAlignment="Top" Margin="10,10,10,385">
            <MenuItem Name="File1" Header="_File">
                <MenuItem Name="Exit1" Header="E_xit"></MenuItem>
            </MenuItem>
            <MenuItem Name="Help1" Header="Help">
                <MenuItem Name="About1" Header="About"></MenuItem>
            </MenuItem>
        </Menu>
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="90" Margin="10,50,0,0" VerticalAlignment="Top" Width="100">
            <StackPanel HorizontalAlignment="Left" Margin="0,0,0,0" VerticalAlignment="Top">
                <CheckBox Name="CheckBox1" Content="CheckBox1" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                <CheckBox Name="CheckBox2" Content="CheckBox2" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                <CheckBox Name="CheckBox3" Content="CheckBox3" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
            </StackPanel>
        </Border>
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="90" Margin="120,50,0,0" VerticalAlignment="Top" Width="110">
            <StackPanel HorizontalAlignment="Left" Margin="0,0,0,0" VerticalAlignment="Top">
                <RadioButton Name="RadioButton1" Content="RadioButton1" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                <RadioButton Name="RadioButton2" Content="RadioButton2" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                <RadioButton Name="RadioButton3" Content="RadioButton3" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
            </StackPanel>
        </Border>
        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="50" Margin="10,150,0,0" VerticalAlignment="Top" Width="220">
            <Label Name="Label1" Content="Label1" VerticalAlignment="Top" HorizontalAlignment="Stretch" HorizontalContentAlignment="Center" Margin="10,10,10,0"/>
            
        </Border>
        <ComboBox Name="ComboBox1" HorizontalAlignment="Left" Margin="10,205,0,0" VerticalAlignment="Top" Width="220" Height="40">
            
        </ComboBox>
        <TextBox HorizontalAlignment="Left" Height="50" Margin="10,250,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="220"/>
        <RichTextBox HorizontalAlignment="Left" Height="100" Margin="10,305,0,0" VerticalAlignment="Top" Width="220">
            <FlowDocument>
                <Paragraph>
                    <Run Text="RichTextBox"/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button Content="Button" HorizontalAlignment="Left" Margin="10,410,0,0" VerticalAlignment="Top" Width="75"/>

    </Grid>
</Window>

"@       

function Select-Item {
    $WPFLabel1.Content = $WPFComboBox1.SelectedItem
} # Select-Item

function Exit-Form {
    $Form.Close()
} # Exit-Form

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
 

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

 
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {$Form = [Windows.Markup.XamlReader]::Load( $reader )}
catch {Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================


$xaml.SelectNodes("//*[@Name]") | Where-Object {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}


Function Get-FormVariables {
    if ($global:ReadmeDisplay -ne $true) {Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true}
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
    #get-variable 
}

#Display WPF Variable (You can comment this function once script finished)
Get-FormVariables

$WPFCheckBox1.IsChecked = $true
$WPFRadioButton1.IsChecked = $true
$FontArray = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
$WPFComboBox1.ItemsSource = $FontArray
$WPFComboBox1.Add_SelectionChanged({ Select-Item })
$WPFExit1.Add_Click({ Exit-Form })



#Display powershell GUI
$Form.ShowDialog() | out-null


<######################################################################################################
Name                           Value                                                                                                                                                          
----                           -----                                                                                                                                                          
WPFAbout1                      System.Windows.Controls.MenuItem Header:About Items.Count:0                                                                                                    
WPFCheckBox1                   System.Windows.Controls.CheckBox Content:CheckBox1 IsChecked:False                                                                                             
WPFCheckBox2                   System.Windows.Controls.CheckBox Content:CheckBox2 IsChecked:False                                                                                             
WPFCheckBox3                   System.Windows.Controls.CheckBox Content:CheckBox3 IsChecked:False                                                                                             
WPFComboBox1                   System.Windows.Controls.ComboBox Items.Count:0                                                                                                                 
WPFExit1                       System.Windows.Controls.MenuItem Header:E_xit Items.Count:0                                                                                                    
WPFFile1                       System.Windows.Controls.MenuItem Header:_File Items.Count:1                                                                                                    
WPFGrid1                       System.Windows.Controls.Grid                                                                                                                                   
WPFHelp1                       System.Windows.Controls.MenuItem Header:Help Items.Count:1                                                                                                     
WPFLabel1                      System.Windows.Controls.Label: Label1                                                                                                                          
WPFMenu1                       System.Windows.Controls.Menu Items.Count:2                                                                                                                     
WPFRadioButton1                System.Windows.Controls.RadioButton Content:RadioButton1 IsChecked:False                                                                                       
WPFRadioButton2                System.Windows.Controls.RadioButton Content:RadioButton2 IsChecked:False                                                                                       
WPFRadioButton3                System.Windows.Controls.RadioButton Content:RadioButton3 IsChecked:False  
########################################################################################################>


