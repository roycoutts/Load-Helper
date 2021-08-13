<#
.SYNOPSIS
    This is a WinForms GUI Template with Common Controls
 
.DESCRIPTION
    Shows dialog box Form1
 
.PARAMETER FormOwner
    The window that owns Form1.
 
.PARAMETER KeepResults
    Places the dialog response code in the pipeline.
 
.OUTPUTS
    None by default
 
    System.Windows.Forms.DialogResult if -KeepResults is used
#>
[CmdletBinding()]
param(
    [PSObject]$FormOwner, [Switch]$KeepResults = $false
)
  
## Add WinForm and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 ## Enable visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()
  
function Show-Form1
{
    param($FormOwner = $null, [Switch]$KeepResults = $false)
 
    $ReturnCode = $Form1.ShowDialog($null)
    if ($KeepResults)
    {
        $ReturnCode
    }
} # Show-Form1
  
function Build-Form1
{
 	## Form1 Form ##
	$Script:Form1              = New-Object System.Windows.Forms.Form
	$Form1.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
	$Form1.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::Font
	$Form1.ClientSize          = New-Object System.Drawing.Size(784, 561)
	$Form1.MainMenuStrip       = $Form1.MenuStrip1
	$Form1.Name                = "Form1"
	$Form1.Text                = "Form1"

	## Form1 Controls ##
	$MenuStrip1             = New-Object System.Windows.Forms.MenuStrip
	$FileToolStripMenuItem  = New-Object System.Windows.Forms.ToolStripMenuItem
	$ExitToolStripMenuItem  = New-Object System.Windows.Forms.ToolStripMenuItem
	$HelpToolStripMenuItem  = New-Object System.Windows.Forms.ToolStripMenuItem
	$AboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
	$GroupBox1              = New-Object System.Windows.Forms.GroupBox
	$CheckBox1              = New-Object System.Windows.Forms.CheckBox
	$CheckBox2              = New-Object System.Windows.Forms.CheckBox
	$CheckBox3              = New-Object System.Windows.Forms.CheckBox
	$GroupBox2              = New-Object System.Windows.Forms.GroupBox
	$RadioButton1           = New-Object System.Windows.Forms.RadioButton
	$RadioButton2           = New-Object System.Windows.Forms.RadioButton
	$RadioButton3           = New-Object System.Windows.Forms.RadioButton
	$GroupBox3              = New-Object System.Windows.Forms.GroupBox
	$Label1                 = New-Object System.Windows.Forms.Label
	$ComboBox1              = New-Object System.Windows.Forms.ComboBox
	$TextBox1               = New-Object System.Windows.Forms.TextBox
	$RichTextBox1           = New-Object System.Windows.Forms.RichTextBox
	$Button1                = New-Object System.Windows.Forms.Button

	## Form1 Control Constructions ##

	# Control: MenuStrip1
    $MenuStrip1.Items.AddRange(@($FileToolStripMenuItem, $HelpToolStripMenuItem))
	$MenuStrip1.Location = New-Object System.Drawing.Point(0, 0)
	$MenuStrip1.Name     = "MenuStrip1"
	$MenuStrip1.Size     = New-Object System.Drawing.Size(784, 24)
	$MenuStrip1.TabIndex = 0
	$MenuStrip1.Text     = "MenuStrip1"

	# Control: FileToolStripMenuItem
    $FileToolStripMenuItem.DropDownItems.AddRange(@($ExitToolStripMenuItem))
	$FileToolStripMenuItem.Name = "FileToolStripMenuItem"
	$FileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
	$FileToolStripMenuItem.Text = "&File"

	# Control: ExitToolStripMenuItem
	$ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
	$ExitToolStripMenuItem.Size = New-Object System.Drawing.Size(180, 22)
	$ExitToolStripMenuItem.Text = "E&xit"
    $ExitToolStripMenuItem.Add_Click({ Exit-Form1 })

	# Control: HelpToolStripMenuItem
    $HelpToolStripMenuItem.DropDownItems.AddRange(@($AboutToolStripMenuItem))
	$HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
	$HelpToolStripMenuItem.Size = New-Object System.Drawing.Size(44, 20)
	$HelpToolStripMenuItem.Text = "&Help"

	# Control: AboutToolStripMenuItem
	$AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
	$AboutToolStripMenuItem.Size = New-Object System.Drawing.Size(180, 22)
	$AboutToolStripMenuItem.Text = "&About"
	
	# Control: GroupBox1
	$GroupBox1.Location = New-Object System.Drawing.Point(10, 30)
	$GroupBox1.Name     = "GroupBox1"
	$GroupBox1.Size     = New-Object System.Drawing.Size(100, 90)
	$GroupBox1.TabIndex = 1
	$GroupBox1.TabStop  = $False
	$GroupBox1.Text     = "GroupBox1"

	# Control: CheckBox1
	$CheckBox1.AutoSize                = $True
	$CheckBox1.Location                = New-Object System.Drawing.Point(10, 20)
	$CheckBox1.Name                    = "CheckBox1"
	$CheckBox1.Size                    = New-Object System.Drawing.Size(81, 17)
	$CheckBox1.TabIndex                = 0
	$CheckBox1.Text                    = "CheckBox1"
	$CheckBox1.UseVisualStyleBackColor = $True
    $CheckBox1.Checked = $True

	# Control: CheckBox2
	$CheckBox2.AutoSize                = $True
	$CheckBox2.Location                = New-Object System.Drawing.Point(10, 40)
	$CheckBox2.Name                    = "CheckBox2"
	$CheckBox2.Size                    = New-Object System.Drawing.Size(81, 17)
	$CheckBox2.TabIndex                = 1
	$CheckBox2.Text                    = "CheckBox2"
	$CheckBox2.UseVisualStyleBackColor = $True

	# Control: CheckBox3
	$CheckBox3.AutoSize                = $True
	$CheckBox3.Location                = New-Object System.Drawing.Point(10, 60)
	$CheckBox3.Name                    = "CheckBox3"
	$CheckBox3.Size                    = New-Object System.Drawing.Size(81, 17)
	$CheckBox3.TabIndex                = 2
	$CheckBox3.Text                    = "CheckBox3"
	$CheckBox3.UseVisualStyleBackColor = $True

	# Control: GroupBox2
	$GroupBox2.Location = New-Object System.Drawing.Point(120, 30)
	$GroupBox2.Name     = "GroupBox2"
	$GroupBox2.Size     = New-Object System.Drawing.Size(110, 90)
	$GroupBox2.TabIndex = 2
	$GroupBox2.TabStop  = $False
	$GroupBox2.Text     = "GroupBox2"

	# Control: RadioButton1
	$RadioButton1.AutoSize                = $True
	$RadioButton1.Location                = New-Object System.Drawing.Point(10, 20)
	$RadioButton1.Name                    = "RadioButton1"
	$RadioButton1.Size                    = New-Object System.Drawing.Size(90, 17)
	$RadioButton1.TabIndex                = 0
	$RadioButton1.TabStop                 = $True
	$RadioButton1.Text                    = "RadioButton1"
	$RadioButton1.UseVisualStyleBackColor = $True
    $RadioButton1.Checked = $True

	# Control: RadioButton2
	$RadioButton2.AutoSize                = $True
	$RadioButton2.Location                = New-Object System.Drawing.Point(10, 40)
	$RadioButton2.Name                    = "RadioButton2"
	$RadioButton2.Size                    = New-Object System.Drawing.Size(90, 17)
	$RadioButton2.TabIndex                = 1
	$RadioButton2.TabStop                 = $True
	$RadioButton2.Text                    = "RadioButton2"
	$RadioButton2.UseVisualStyleBackColor = $True

	# Control: RadioButton3
	$RadioButton3.AutoSize                = $True
	$RadioButton3.Location                = New-Object System.Drawing.Point(10, 60)
	$RadioButton3.Name                    = "RadioButton3"
	$RadioButton3.Size                    = New-Object System.Drawing.Size(90, 17)
	$RadioButton3.TabIndex                = 2
	$RadioButton3.TabStop                 = $True
	$RadioButton3.Text                    = "RadioButton3"
	$RadioButton3.UseVisualStyleBackColor = $True

	# Control: GroupBox3
	$GroupBox3.Location = New-Object System.Drawing.Point(10, 126)
	$GroupBox3.Name     = "GroupBox3"
	$GroupBox3.Size     = New-Object System.Drawing.Size(220, 280)
	$GroupBox3.TabIndex = 3
	$GroupBox3.TabStop  = $False
	$GroupBox3.Text     = "GroupBox3"

	# Control: Label1
	$Label1.Location  = New-Object System.Drawing.Point(10, 30)
	$Label1.Name      = "Label1"
	$Label1.Size      = New-Object System.Drawing.Size(200, 23)
	$Label1.TabIndex  = 0
	$Label1.Text      = "Label1"
	$Label1.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

	# Control: ComboBox1
	$ComboBox1.FormattingEnabled = $True
	$ComboBox1.Location          = New-Object System.Drawing.Point(10, 60)
	$ComboBox1.Name              = "ComboBox1"
	$ComboBox1.Size              = New-Object System.Drawing.Size(200, 21)
	$ComboBox1.TabIndex          = 1
    $ComboBox1.Items.AddRange((New-Object System.Drawing.Text.InstalledFontCollection).Families.Name)

	# Control: TextBox1
	$TextBox1.Location  = New-Object System.Drawing.Point(10, 90)
	$TextBox1.Multiline = $True
	$TextBox1.Name      = "TextBox1"
	$TextBox1.Size      = New-Object System.Drawing.Size(200, 50)
	$TextBox1.TabIndex  = 2

	# Control: RichTextBox1
	$RichTextBox1.Location = New-Object System.Drawing.Point(10, 150)
	$RichTextBox1.Name     = "RichTextBox1"
	$RichTextBox1.Size     = New-Object System.Drawing.Size(200, 96)
	$RichTextBox1.TabIndex = 3
	$RichTextBox1.Text     = ""

	# Control: Button1
	$Button1.Location                = New-Object System.Drawing.Point(10, 250)
	$Button1.Name                    = "Button1"
	$Button1.Size                    = New-Object System.Drawing.Size(75, 23)
	$Button1.TabIndex                = 4
	$Button1.Text                    = "Button1"
	$Button1.UseVisualStyleBackColor = $True

	## Adding controls to Form Form1
	$Form1.Controls.Add($GroupBox3)
	$Form1.Controls.Add($GroupBox2)
	$Form1.Controls.Add($GroupBox1)
	$Form1.Controls.Add($MenuStrip1)

	## Adding controls to container controls
	$GroupBox1.Controls.Add($CheckBox3)
	$GroupBox1.Controls.Add($CheckBox2)
	$GroupBox1.Controls.Add($CheckBox1)
	$GroupBox2.Controls.Add($RadioButton3)
	$GroupBox2.Controls.Add($RadioButton2)
	$GroupBox2.Controls.Add($RadioButton1)
	$GroupBox3.Controls.Add($Button1)
	$GroupBox3.Controls.Add($RichTextBox1)
	$GroupBox3.Controls.Add($TextBox1)
	$GroupBox3.Controls.Add($ComboBox1)
	$GroupBox3.Controls.Add($Label1)

	## Adding controls as properties to Form Form1
	$Form1 | Add-Member -MemberType NoteProperty -Name MenuStrip1 -Value $MenuStrip1
	$Form1 | Add-Member -MemberType NoteProperty -Name FileToolStripMenuItem -Value $FileToolStripMenuItem
	$Form1 | Add-Member -MemberType NoteProperty -Name ExitToolStripMenuItem -Value $ExitToolStripMenuItem
	$Form1 | Add-Member -MemberType NoteProperty -Name HelpToolStripMenuItem -Value $HelpToolStripMenuItem
	$Form1 | Add-Member -MemberType NoteProperty -Name AboutToolStripMenuItem -Value $AboutToolStripMenuItem
	$Form1 | Add-Member -MemberType NoteProperty -Name GroupBox1 -Value $GroupBox1
	$Form1 | Add-Member -MemberType NoteProperty -Name CheckBox1 -Value $CheckBox1
	$Form1 | Add-Member -MemberType NoteProperty -Name CheckBox2 -Value $CheckBox2
	$Form1 | Add-Member -MemberType NoteProperty -Name CheckBox3 -Value $CheckBox3
	$Form1 | Add-Member -MemberType NoteProperty -Name GroupBox2 -Value $GroupBox2
	$Form1 | Add-Member -MemberType NoteProperty -Name RadioButton1 -Value $RadioButton1
	$Form1 | Add-Member -MemberType NoteProperty -Name RadioButton2 -Value $RadioButton2
	$Form1 | Add-Member -MemberType NoteProperty -Name RadioButton3 -Value $RadioButton3
	$Form1 | Add-Member -MemberType NoteProperty -Name GroupBox3 -Value $GroupBox3
	$Form1 | Add-Member -MemberType NoteProperty -Name Label1 -Value $Label1
	$Form1 | Add-Member -MemberType NoteProperty -Name ComboBox1 -Value $ComboBox1
	$Form1 | Add-Member -MemberType NoteProperty -Name TextBox1 -Value $TextBox1
	$Form1 | Add-Member -MemberType NoteProperty -Name RichTextBox1 -Value $RichTextBox1
	$Form1 | Add-Member -MemberType NoteProperty -Name Button1 -Value $Button1
} # Build-Form1
 
function Exit-Form1
{
    $Form1.Close()
} # Exit-Form

Build-Form1
 Show-Form1

