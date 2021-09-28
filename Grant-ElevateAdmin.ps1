# Grant-ElevateAdmin
# Self-elevate the script if required
# Simply add this snippet at the beginning of a script that requires elevation to run properly. 
# It works by starting a new elevated PowerShell window and then re-executes the script in this new window, if necessary. 
# If User Account Control (UAC) is enabled, you will get a UAC prompt. 
# If the script is already running in an elevated PowerShell session or UAC is disabled, the script will run normally. 
# This code also allows you to right-click the script in File Explorer and select "Run with PowerShell".

# The first line checks to see if the script is already running in an elevated environment. 
# This would occur if PowerShell is running as Administrator or UAC is disabled. 
# If it is, the script will continue to run normally in that process.

# The second line checks to see if the Windows operating system build number is 6000 (Windows Vista) or greater. 
# Earlier builds did not support Run As elevation.

# The third line retrieves the command line used to run the original script, including any arguments.

# Finally, the fourth line starts a new elevated PowerShell process where the script runs again. 
# Once the script terminates, the elevated PowerShell window closes.

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
        Exit
    }
}
