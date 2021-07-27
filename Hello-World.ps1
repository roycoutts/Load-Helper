Write-Host "Hello World"
# See https://developer.github.com/v3/pulls/#list-pull-requests

$endpoint = "https://api.github.com/repos/[Owner]/[Repo]/pulls?state=[state]"

function Get-BasicAuthCreds {
    param([string]$Username,[string]$Password)
    $AuthString = "{0}:{1}" -f $Username,$Password
    $AuthBytes  = [System.Text.Encoding]::Ascii.GetBytes($AuthString)
    return [Convert]::ToBase64String($AuthBytes)
}

$BasicCreds = Get-BasicAuthCreds -Username "[YourUser]" -Password "[YourPassword]"
$val = Invoke-WebRequest -Uri $endpoint -Headers @{"Authorization"="Basic $BasicCreds"}
$json = $val | ConvertFrom-JSON

foreach($obj in $json)
{
    Write-Host "Pull request: #" + $obj.number
    Write-Host "Title: " + $obj.title
    Write-Host "Url: " + $obj.url
    
    $releaseNotes = $releaseNotes + "Body: "
    $obj.body.Split("`n") | ForEach { 
        # ignore comments from issue templates
        if($_.Trim().StartsWith("<!---") -eq $FALSE)
        {
            $releaseNotes = $releaseNotes + $_ + "`n" 
        }
     }   
    $releaseNotes = $releaseNotes + "`n"
    
    Write-Host "User: " + $obj.user.login
    Write-Host ""
}
