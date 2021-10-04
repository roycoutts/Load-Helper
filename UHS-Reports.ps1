
#region Load Init Variables
    # Load Date/Time Variables
    $MonthName = '|January|February|March|April|May|June|July|August|September|October|November|December'.Split('|')
    $currentDate      = [System.DateTime]::Today
    $currentDayOfWeek = $currentDate.DayOfWeek
    $currentMonthName = $MonthName[$currentDate.Month]
    $currentMonth     = $currentDate.Month
    $currentDay       = $currentDate.Day
    $currentYear      = $currentDate.Year
    # Report Field Variables
    $YesNo = @('YES','NO')
    $FirstDayOfYear = ''
    $LastDayOfMonth = ''
    $Simulator = 'Simulator'
    # Load Report Names
    $ReportName = @('Yearly Patient & Treatment Count',
                    'SRS Data',
                    'Combind Analytics Report',
                    'Synergy Analytics Report',
                    'TrueBeam Analytics Report',
                    'Sims Analytics Report',
                    'Diagnosis Analytics Report')
    # Load Paths to Reports
    $ReportPath = @('\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\Yearly Patient and Treament count.rpt',
                    '\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\SRS Patient and Treatment Count.rpt',
                    '\\wmhfilesrv\GROUPSHARES\Ancillary Systems\Mosaiq\Mosaiq Tasks\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2.rpt',
                    '\\wmhfilesrv\GROUPSHARES\Ancillary Systems\Mosaiq\Mosaiq Tasks\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2.rpt',
                    '\\wmhfilesrv\GROUPSHARES\Ancillary Systems\Mosaiq\Mosaiq Tasks\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2.rpt',
                    '\\wmhfilesrv\GROUPSHARES\Ancillary Systems\Mosaiq\Mosaiq Tasks\SIMS_1_Prod_Month_CT_Count_Summary_Enhanced_Mod_2.rpt',
                    '\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\3_Prod_Diag_Tx_Count_Diag_analysis.rpt')
    # Load Machine Names
    $MachineName = @('',
                     '',
                     'Synergy+TB',
                     'Synergy',
                     'TrueBeam',
                     '\\wmhfilesrv\GROUPSHARES\Ancillary Systems\Mosaiq\Mosaiq Tasks\SIMS_1_Prod_Month_CT_Count_Summary_Enhanced_Mod_2.rpt',
                     '\\wmhfilesrv.uhsh.uhs.org\groupshares\Ancillary Systems\Mosaiq\RadOnc Reports\Other\Count_Diagnosis\3_Prod_Diag_Tx_Count_Diag_analysis.rpt')
#endregion Load Init Variables

$ExportFormat = @('Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003)',
                  'Microsoft Excel (97-2003) Data only')

$OutputPath = @('C:\TEMP\Yearly Patient and Treatment count.xls',
                'C:\TEMP\SRS Patient and Treatment Count.xls',
                'C:\TEMP\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2.xls',
                'C:\TEMP\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2-Synergy.xls',
                'C:\TEMP\1_Prod_Month_Tx_Count_Summary_Enhanced_Mod_2-TrueBeam.xls',
                'C:\TEMP\SIMS_1_Prod_Month_CT_Count_Summary_Enhanced_Mod_2.xls',
                'C:\TEMP\3_Prod_Diag_Tx_Count_Diag_analysis-<REPORT-YEAR>.xls')


Function Select-Report {
    $Title = 'Select Report'
    $ReportName | ForEach-Object {$menu=''; $i=0} {$menu += "[$i] : $_`n"; $i++}
    $Message = $menu
    $Choices = (0..($ReportName.Count - 1))
    $DefaultChoice = 0
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "&$($_)", $ReportName[$_]
    }       
    $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice )  
    #Return Read-Choice -Title $Title -Message $Message -Choices $Choices -DefaultChoice $DefaultChoice
}

Function Select-Year {
    $Title = 'Select Report Year'
    $Message = "Default: $currentYear`n"
    $Choices = (($currentYear-1),$currentYear)
    $DefaultChoice = 1
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "$($_.ToString().Substring(0,3))&$($_.ToString().Substring(3))", $Choices[$_]
    }       
    $UserSelection = $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice)  
    Return [int]$Choices[$UserSelection]
}

Function Select-Month {
    $Title = 'Select Report Month'
    $MonthName | ForEach-Object {$menu=''; $i=0} {$menu += "          ["+[char]($i+64)+"] : $_`n"; $i++}
    $Message = "Default: $currentMonthName`n`n"+($menu.Split([char]10)[1..12] -join "`n")
    $Choices = (1..($MonthName.Count - 1))
    $DefaultChoice = $currentMonth
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "&$([char]($_+64))", $MonthName[$_]
    }      
    $UserSelection = $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice)  
    Return $UserSelection
}

Function Select-HolidaysTaken {
    $Title = 'Holidays Taken'
    $Message = "Default: Zero (0) Holidays Taken for $($MonthName[$SelectedMonth]), $SelectedYear`n"
    $Choices = (0..4)
    $DefaultChoice = 0
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "&$($_)", $Choices[$_]
    }       
    $UserSelection = $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice)  
    Return [int]$Choices[$UserSelection]
}

Function Select-WeekendDaysWorked {
    $Title = 'Weekend Days Worked'
    $Message = "Default: Zero (0) Weekend Days Worked for $($MonthName[$SelectedMonth]), $SelectedYear`n"
    $Choices = (0..4)
    $DefaultChoice = 0
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "&$($_)", $Choices[$_]
    }       
    $UserSelection = $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice)  
    Return [int]$Choices[$UserSelection]
}
# Get First Day Of Year (ex: 01/01/2021)
Function Get-FirstDayOfYear {
    $FirstDayOfYear = (Get-Date "1/1/$SelectedYear").ToString("MM/dd/yyyy")
    Return $FirstDayOfYear
}
# Get Last Day Of Month (ex: 10/31/2021)
Function Get-LastDayOfMonth {
    $LastDay        = [DateTime]::DaysInMonth($SelectedYear, $SelectedMonth)
    $LastDayOfMonth = (Get-Date "$SelectedMonth/$LastDay/$SelectedYear").ToString("MM/dd/yyyy")
    Return $LastDayOfMonth
}
# Get First Day Of Month (ex: 10/01/2021)
Function Get-FirstDayOfMonth {
    $FirstDayOfMonth = (Get-Date "$SelectedMonth/1/$SelectedYear").ToString("MM/dd/yyyy")
    Return $FirstDayOfMonth
}
# Get First Day Of Previous Year (ex: 01/01/2020)
Function Get-FirstDayOfPreviousYear {
    $FirstDayOfPreviousYear = (Get-Date "1/1/$SelectedYear").AddYears(-1).ToString("MM/dd/yyyy")
    Return $FirstDayOfPreviousYear
}
# Get Last Day Of Month Of Previous Year (ex: 10/31/2020)
Function Get-LastDayOfMonthOfPreviousYear {
    $PreviousYear                 = $SelectedYear - 1
    $LastDay                      = [DateTime]::DaysInMonth($PreviousYear, $SelectedMonth)
    $LastDayOfMonthOfPreviousYear = (Get-Date "$SelectedMonth/$LastDay/$PreviousYear").ToString("MM/dd/yyyy")
    Return $LastDayOfMonthOfPreviousYear
}

$SelectedYear                 = Select-Year
$SelectedMonth                = Select-Month
$SelectedHolidays             = Select-HolidaysTaken
$SelectedWeekends             = Select-WeekendDaysWorked
$AdjustDays                   = 0 - $SelectedHolidays + $SelectedWeekends
$FirstDayOfYear               = Get-FirstDayOfYear
$LastDayOfMonth               = Get-LastDayOfMonth
$FirstDayOfMonth              = Get-FirstDayOfMonth
$FirstDayOfPreviousYear       = Get-FirstDayOfPreviousYear
$LastDayOfMonthOfPreviousYear = Get-LastDayOfMonthOfPreviousYear
$SelectedReport               = Select-Report

$ReportFields = @()
$ReportFields += '$FirstDayOfYear|$LastDayOfMonth'                             # Yearly Patient & Treatment Count
$ReportFields += '$FirstDayOfYear|$LastDayOfMonth'                             # SRS Data
$ReportFields += '$MachineName[$SelectedReport]|$FirstDayOfMonth|$AdjustDays'  # Combind Analytics Report
$ReportFields += '$MachineName[$SelectedReport]|$FirstDayOfMonth|$AdjustDays'  # Synergy Analytics Report
$ReportFields += '$MachineName[$SelectedReport]|$FirstDayOfMonth|$AdjustDays'  # TrueBeam Analytics Report
$ReportFields += '$Simulator|$FirstDayOfMonth'                                 # Sims Analytics Report
$ReportFields += '$FirstDayOfPreviousYear|$LastDayOfMonthOfPreviousYear'       # Diagnosis Analytics Report

Function Launch-Report {
    $Title = 'Launch Report'
    $Message = "Launch Report?`n`n          $($ReportName[$SelectedReport])`n`n                    [0] YES`n                    [1] NO"
    $Choices = (0..($YesNo.Count - 1))
    $DefaultChoice = 0
    [System.Management.Automation.Host.ChoiceDescription[]]$Poss = $Choices | ForEach-Object {            
        New-Object System.Management.Automation.Host.ChoiceDescription "&$($_)", $Choices[$_]
    }       
    $UserSelection = $Host.UI.PromptForChoice($Title, $Message, $Poss, $DefaultChoice)  
    Return [int]$Choices[$UserSelection]
}

[System.Windows.Forms.MessageBox]::Show("Report Launched")

