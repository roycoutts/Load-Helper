<#
.SYNOPSIS
    Converts a .designer.vb file to a PS1 file.
 
.DESCRIPTION
    Parses a Visual Basic file that creates a Windows Form and generates a roughly equivalent PowerShell script.
 
.PARAMETER SourceFile
    Path to the .vb.designer file to be converted.
 
.PARAMETER DestinationScript
    Path to the output .ps1 file. The current directory is assumed if not specified and .ps1 is appended if not present.
    Mutually exclusive with ToPipeline parameter.
 
.PARAMETER ToPipeline
    Outputs the text to the pipeline. If there is nothing else in the pipeline, the script will appear on screen.
    Mutually exclusive with DestinationScript parameter.
 
.PARAMETER ScriptType
    StandAlone = A stand-alone script. When executed, this script will display the form. When the form closes, the script ends.
    DotSourced = A script that will expose the form and its Show-FormName function to the calling scope.
    Inline = A script chunk that resembles the output from DotSourced, but contains no help text or Parameter declarations
 
.PARAMETER ExcludeAssemblies
    Excludes the "Add-Type" lines for the .Net assemblies necessary to display Windows forms (System.Windows.Forms and System.Drawing).
    If you specify this parameter, your script must include these lines or the emitted script will be non-functional.
 
.PARAMETER DoNotEnableVisualStyles
    By default, a directive will be included to enable visual styles. This provides a more consistent experience between what you design
    in Visual Basic and what PowerShell displays. Use this parameter to leave visual styles at their default/current setting.
 
.LINK
 
https://etechgoodness.wordpress.com/2014/05/02/convert-visual-basic-form-to-powershell/
 
http://wp.me/p1pPNH-9r
 
.NOTES
    Updates in 1.05 (May 22, 2015)
    ---------------
    -KeepResults (a parameter of the emitted script) is now a Switch instead of a Boolean
    Controls are not added to container controls until all controls are defined; this corrects issues such as text boxes inside group boxes being placed incorrectly
    Multiline strings convert correctly
    Detection of System.Windows.Forms.x and System.Drawing.x enumerators drastically improved
    Processing of System.Drawing.Font declarations improved
    Processing of cross-control references improved (such as AcceptButton and CancelButton)
    Tree View controls and nodes are fully supported
    Lines containing an apostrophe are not automatically considered a comment
    Improved handling of parameters; positional parameters work as expected
    Visual styles are enabled by default
    Custom colors are properly translated
 
    Updates in 1.04
    ---------------
    Some clean-up of output format
    Controls are added as properties to the form so that they can be manipulated without indexing into $Form1.Controls[]
 
    Script by Eric Siron
    Licensed under a Creative Commons Attribution-ShareAlike 4.0 International License
#>
function ConvertTo-PowerShellForm
{
    [CmdletBinding(DefaultParameterSetName="ToPipeline")]
    param(
        [Parameter(HelpMessage="Enter the .designer.vb file to parse", Mandatory=$true, ValueFromPipeline=$true, Position=1)]
        [String]$SourceFile,
 
        [Parameter(Position=2, ParameterSetName="ToFile")]
        [String]$DestinationScript,
 
        [Parameter(ParameterSetName="ToPipeline")]
        [Switch]$ToPipeline = $false,
 
        [Parameter()]
        [ValidateSet("StandAlone", "DotSourced", "Inline")]
        [String]$ScriptType = "StandAlone",
 
        [Parameter()]
        [Switch]$ExcludeAssemblies = $false,
 
        [Parameter()]
        [Switch]$DoNotEnableVisualStyles = $false
    )
 
    BEGIN {
        $DotSourcedHeaderText = @'
<#
    Call Show-{0} to display {0} as a dialog.
    Use parameter FormOwner to specify another form as the owner of {0}.
    Use parameter KeepResults to return the DialogResult to the calling script.
#>
 
'@
 
        $StandAloneHeaderText = @'
<#
.SYNOPSIS
    Definition script for window {0}.
 
.DESCRIPTION
    Shows dialog box {0}
 
.PARAMETER FormOwner
    The window that owns {0}.
 
.PARAMETER KeepResults
    Places the dialog response code in the pipeline.
 
.OUTPUTS
    None by default
 
    System.Windows.Forms.DialogResult if -KeepResults is used
#>
[CmdletBinding()]
param(
    [PSObject]$FormOwner, [Switch]$KeepResults=$false
)
 
'@
 
    $AssemblyDeclarationsText = @'
 
## Add WinForm and Drawing assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
'@
 
    $VisualStylesText = @'
## Enable visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()
 
'@
 
    $BuildFunctionLeadIn = @'
 
function Build-{0}
{{
 
'@
 
    $BuildFunctionLeadOut = @'
}} # Build-{0}
 
Build-{0}
 
'@
 
    $ShowFunction = @'
 
function Show-{0}
{{
    param($FormOwner = $null, [Switch]$KeepResults=$false)
 
    $ReturnCode = {1}.ShowDialog($null)
    if($KeepResults)
    {{
        $ReturnCode
    }}
}} # Show-{0}
 
'@
    }
 
    PROCESS
    {
        filter CommonConversions
        {
            param([String]$Line)
            $Line = $Line -replace '(?<=( = |\(|, ))New ([A-Za-z0-9\.]*)([\(][\)](?!={)){0,}', 'New-Object $2' # convert "New" to "New-Object", drop any empty parens at the end of the line, leave non-empty parens
            $Line = $Line -replace '(, New-Object)(.*?)(})', ', (New-Object$2})'
            $Line = $Line -replace 'CType\((\d+), Byte\)', '$1' # byte conversions are not necessary
            $Line = $Line -replace '(True|False)$', '$$$0' # " = True" or " = False" ending a line converted to " = $True" or " = "$False"
            $Line = $Line -replace '(System\.)(Windows\.Forms|Drawing)(\.)(?!\S*\()(\S*)(\.)(\S*)', '[$1$2$3$4]::$6'    # changes enumerators. ex: System.Drawing.FontStyle.Bold to [System.Drawing.FontStyle]::Bold. -- (?!\S*\() negative lookahead to find parens to throw out items like System.Drawing.Point(1,2)
            $Line = $Line -replace "\(Me\.", '($$' # drop the $FormName. and leave control variable
            $Line
        }
 
        filter FormConversions
        {
            param([String]$Line)
 
            $Line = $Line -replace "(^\s*?|(?<= = ))Me\.", "$FormVar." # lines that start with Me. replaced with $FormName.; any starting whitespace removed
            $Line
        }
 
        filter ControlConversions
        {
            param([String]$Line)
            $Line = $Line -replace '^\s*?Me\.', '$$' # lines that start with "Me." shortened to the control's variable name; any starting whitespace removed
            if($Line -match '(\.AddRange\(New-Object \S*?\ ){(.*)}')# searching for AddRange
            {
                $Line = $Line -replace "(\.AddRange\(New-Object \S*?\ ){(.*)}", '.AddRange(@($2)'   # VB's AddRange to PowerShell's AddRange
                ## AddRange in this context usually only adds text items or controls; text items should be ignored, but control items need to have a dollar-sign prepended
                $RangeItemsIn = $Matches[2].Split(',')
                $RangeItemsOut = @()
                foreach($RangeItem in $RangeItemsIn)
                {
                    $RangeItem = $RangeItem.Trim()
                    if($RangeItem[0] -ne '"') { $RangeItem = '$' + $RangeItem }
                    $RangeItemsOut += ,($RangeItem)
                }
                $Line = $Line -replace $Matches[2], [String]::Join(', ', $RangeItemsOut)
            }
            $Line = $Line -replace '& _$', '+'  # convert strings broken across multiple code lines
            $Line = $Line -replace '(System.Drawing.Font\(.*)(!)', '$1' # remove exclamation marks from font definitions
            $Line = $Line -replace 'System\.Drawing\.Color\.FromArgb[^\d]*(\d+)[^\d]*(\d+)[^\d]*(\d+)[^\d]*', '[System.Drawing.Color]::FromArgb($1, $2, $3)'    # Custom colors
            if($Line.Contains('[System.Windows.Forms.AnchorStyles]::'))
            {
                $AnchorPoints = @()
                if($Line.Contains('Top')) { $AnchorPoints += '[System.Windows.Forms.AnchorStyles]::Top' }
                if($Line.Contains('Bottom')) { $AnchorPoints += '[System.Windows.Forms.AnchorStyles]::Bottom' }
                if($Line.Contains('Left')) { $AnchorPoints += '[System.Windows.Forms.AnchorStyles]::Left' }
                if($Line.Contains('Right')) { $AnchorPoints += '[System.Windows.Forms.AnchorStyles]::Right' }
                $AnchorSettings = [String]::Join(' -bor ', $AnchorPoints)
                if($Line -match '^\s*\$.*\.Anchor = ')
                {
                    $Line = $Matches[0] + $AnchorSettings
                }
                else
                {
                    $Line = "`t`t -bor " + $AnchorSettings
                }
            }
            if($Line -match '(\S*) As System.Windows.Forms.TreeNode')
            {
                $Script:TreeNodeVars += $Matches[1]
                $Line = $Line -replace '^\s*Dim (\S*) As System.Windows.Forms.TreeNode', '$$$1'
                if($Line -match '\(New-Object System.Windows.Forms.TreeNode {([^}]*)}\)')
                {
                    $SubNodesIn = $Matches[1].Split(',')
                    $SubNodesOut = @()
                    foreach($SubNode in $SubNodesIn)
                    {
                        $SubNode = $SubNode.Trim()
                        if($SubNode[0] -ne '"') { $SubNode = '$' + $SubNode }
                        $SubNodesOut += ,($SubNode)
                    }
                    $SubNodesOutLine = [String]::Join(', ', $SubNodesOut)
                    $Line = $Line -replace '\(New-Object System.Windows.Forms.TreeNode {([^}]*)}\)', "@($SubNodesOutLine)"
                }
            }
            $Line
        }
 
        function ConvertTo-PowerShellForm
        {
            if(-not(Test-Path -Path $SourceFile))
            {
                throw "$SourceFile not found"
            }
            $RawFileData = Get-Content $SourceFile
 
            if($RawFileData[0] -ne "<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _")
            {
                throw "$SourceFile does not appear to be a valid VB Designer file"
            }
 
            $FormShortName = $RawFileData[1] -replace "Partial Class "
            New-Variable -Name FormVar -Value "`$$FormShortName"
 
            $OutputFormDefinition = @("`t## $FormShortName Form ##", "`t`$Script:$FormShortName = New-Object System.Windows.Forms.Form")
            $OutputControlDeclarations = @("", "`t## $FormShortName Controls ##")
            $OutputControlDefinitions = @("", "`t## $FormShortName Control Constructions ##")
            $OutputControlAddToForm = @("", "`t## Adding controls to Form $FormShortName")
            $OutputControlAddToContainer = @("", "`t## Adding controls to container controls")
            $ControlVars = @("", "`t## Adding controls as properties to Form $FormShortName")
            $Script:TreeNodeVars = @()
            $TreeNodeSettings = @("", "`t## Configuring tree nodes")
            $TextToWatchForNext = "Private Sub InitializeComponent()"
            $InInitComponent = $false
            $InControlDeclarations = $false
            $InControlDefinitions = $false
            $InFormDefinition = $false
            $CurrentControl = $null
 
            foreach($CodeLine in $RawFileData)
            {
                if($CodeLine -notmatch $TextToWatchForNext)
                {
                    if($InInitComponent -and $InControlDeclarations)
                    {
                        if($CodeLine -match "New System.Windows.Forms")
                        {
                            $ControlDeclaration = CommonConversions -Line $CodeLine
                            $ControlDeclaration = ControlConversions -Line $ControlDeclaration
                            $OutputControlDeclarations += ("`t" + $ControlDeclaration)
                        }
                    }
                    elseif($InControlDefinitions)
                    {
                        if($CodeLine -match "'$FormShortName") #vb sets up the form after all the controls, and prepends it with some comments, one of which is in the format 'FormName
                        {
                            $InControlDeclarations = $false
                            $InControlDefinitions = $false
                            $InFormDefinition = $true
                        }
                        elseif($CodeLine -match "^\s*'.*$")
                        {
                            # this line is commented out (but isn't the form's name, that was caught in a previous check), so just eat it
                        }
                        else
                        {
                            if($CurrentControl -eq $null -and $CodeLine -match "(?<=\.)(.*?)(?=\.)")
                            {
                                $CurrentControl = $Matches[1]
                                $ControlVars += "`t`$$FormShortName | Add-Member -MemberType NoteProperty -Name $CurrentControl -Value `$$CurrentControl"
                                $OutputControlDefinitions += @("", "`t# Control: $CurrentControl")
                            }
                            $ControlDefinition = CommonConversions -Line $CodeLine
                            $ControlDefinition = ControlConversions -Line $ControlDefinition
                            if($ControlDefinition.Contains("$CurrentControl.Controls.Add("))
                            {
                                $OutputControlAddToContainer += ("`t" + $ControlDefinition)
                            }
                            else
                            {
                                if($ControlDefinition -match '^\s*-bor')
                                {   # if a line starts with -bor, it is a continuation of the previous line. PS can't parse the break unless the -bor finishes the line. The easiest way to handle this is to just make it one big line.
                                    $OutputControlDefinitions[($OutputControlDefinitions.Length - 1)] += ($ControlDefinition -replace '\t+?', '')
                                }
                                else
                                {
                                    foreach($TreeNode in $TreeNodeVars)
                                    {
                                        if($ControlDefinition -match "^\s*?($TreeNode)")
                                        {
                                            $ControlDefinition = $ControlDefinition -replace "^\s*?($TreeNode)", "`$$TreeNode"
                                        }
                                    }
                                    $OutputControlDefinitions += ("`t" + $ControlDefinition)
                                }
                            }
                        }
                    }
                    elseif($InFormDefinition)
                    {
                        if(($CodeLine -notmatch "=") -and $CodeLine -match "\.Add\(")
                        {
                            $ControlAddition = CommonConversions -Line $CodeLine
                            $ControlAddition = FormConversions -Line $ControlAddition
                            $OutputControlAddToForm += ("`t" + $ControlAddition)
                        }
                        elseif($CodeLine -match "=")
                        {
                            $FormDefinition = CommonConversions -Line $CodeLine
                            $FormDefinition = FormConversions -Line $FormDefinition
                            if($FormDefinition -match "\.SizeF\(")
                            {
                                $FormDefinition = $FormDefinition -replace "!"
                            }
                            $OutputFormDefinition += ("`t" + $FormDefinition)
                        }
                    }
                }
                else
                {
                    if($TextToWatchForNext -eq "Private Sub InitializeComponent()")
                    {
                        $InInitComponent = $true
                        $InControlDeclarations = $true
                        $TextToWatchForNext = "'$"
                    }
                    elseif($TextToWatchForNext -eq "'$")
                    {
                        if($InControlDeclarations)
                        {
                            $InControlDeclarations = $false
                            $InControlDefinitions = $true
                        }
                        elseif($InControlDefinitions)
                        {
                            $CurrentControl = $null
                        }
                    }
                }
            }
 
            $OutputText = New-Object System.Text.StringBuilder("")
            switch($ScriptType)
            {
                "StandAlone" {
                    $OutputText.AppendFormat($StandAloneHeaderText, $FormShortName) | Out-Null
                }
                "DotSourced" {
                    $OutputText.AppendFormat($DotSourcedHeaderText, $FormShortName) | Out-Null
                }
                default {
                }
            }
 
            if(-not($ExcludeAssemblies))
            {
                $OutputText.Append($AssemblyDeclarationsText) | Out-Null
            }
            if(-not($DoNotEnableVisualStyles))
            {
                $OutputText.Append($VisualStylesText) | Out-Null
            }
 
            $OutputText.AppendFormat($ShowFunction, $FormShortName, $FormVar) | Out-Null
            $OutputText.Append(($BuildFunctionLeadIn -f $FormShortName, "")) | Out-Null
 
            $OutputFormDefinition | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            $OutputControlDeclarations | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            $OutputControlDefinitions | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            $OutputControlAddToForm | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            $OutputControlAddToContainer | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            $ControlVars | ForEach-Object { $OutputText.AppendLine($_) | Out-Null }
            Write-Host $ControlList
 
            $OutputText.AppendFormat($BuildFunctionLeadOut, $FormShortName) | Out-Null
 
            if($ScriptType -eq "StandAlone")
            {
                $OutputText.AppendLine("Show-$FormShortName") | Out-Null
            }
 
            if(-not([String]::IsNullOrEmpty($DestinationScript)))
            {
                if([String]::IsNullOrEmpty((Split-Path -Path $DestinationScript)))
                {
                    $DestinationScript = ".\" + $DestinationScript
                }
                if($DestinationScript -notmatch "\.ps1$")
                {
                    $DestinationScript += ".ps1"
                }
                Set-Content -Path $DestinationScript -Value $OutputText.ToString() # We did enough checking, let PS handle any invalid path problems from here
            }
            else
            {
                $OutputText.ToString()
            }
        }
 
        ConvertTo-PowerShellForm
    }
 
    END {}
}
