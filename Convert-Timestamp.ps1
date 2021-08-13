# Convert-Timestamp.ps1

<#
.SYNOPSIS
    Convert UNIX or MEDITECH Timestamp to Date/Time.

.DESCRIPTION
    Script will accept either a Unix Timestamp or Meditech Timestamp and return the DateTime

.PARAMETER -MEDITECH_Timestamp
    Identifies the timestamp is from MEDITECH and should be converted using MEDITECH's Epoch Date

.PARAMETER -UNIX_Timestamp
    Identifies the timestamp is from UNIX and should be converted using UNIX's Epoch Date

.INPUTS
    Input will always be the number of seconds since the Epoch Date

.OUTPUTS
    Script returns a [DateTime] object with all the advantages associated with [DateTime]

.NOTES
  Version:        1.0
  Author:         Roy Coutts
  Creation Date:  08/12/2021
  Purpose/Change: Redeveloped an old version of this solution
  
.EXAMPLE
    Example Using Unix Timestamp
    PS:> Convert-Timestamp -UNIX_Timestamp 987654321
         Thursday, April 19, 2001 12:25:21 AM

    Example Using Meditech Timestamp
    PS:> Convert-Timestamp -MEDITECH_Timestamp 987654321
         Saturday, June 18, 2011 12:25:21 AM
#>

function Convert-Timestamp {
    Param(
        [Parameter(Mandatory,ParameterSetName="MEDITECH")]$MEDITECH_Timestamp,
        [Parameter(Mandatory,ParameterSetName="UNIX")]    $UNIX_Timestamp
    )
    # Set UNIX/MEDITECH Epoch Dates
    $Epoch_UNIX     = '1/1/1970'
    $Epoch_MEDITECH = '3/1/1980'

    if ($MEDITECH_Timestamp) { $DateTime = [TimeZone]::CurrentTimeZone.ToLocalTime(([DateTime]$Epoch_MEDITECH).AddSeconds($MEDITECH_Timestamp)) }
    if ($UNIX_Timestamp)     { $DateTime = [TimeZone]::CurrentTimeZone.ToLocalTime(([DateTime]$Epoch_UNIX).AddSeconds($UNIX_Timestamp)) }

    Return $DateTime
}
