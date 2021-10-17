function Assert-FolderExists
{
  <#
    .SYNOPSIS
    Makes sure the specified folder(s) exist

    .DESCRIPTION
    If a folder does not exist, it will be created.

    .EXAMPLE
    ($Path = 'C:\test') | Assert-PsOneFolderExists
    Makes sure the folder c:\test exists. If it is still missing, it will be created.

    .EXAMPLE
    'C:\test','c:\test2' | Assert-PsOneFolderExists
    Makes sure the folders. If a folder is still missing, it will be created.

    .EXAMPLE
    Assert-PsOneFolderExists -Path 'C:\test','c:\test2'
    Makes sure the folders. If a folder is still missing, it will be created.

    .LINK
    https://powershell.one
  #>

  Param
  (
    [Parameter(Mandatory,HelpMessage='Path to folder that must exist',ValueFromPipeline)]
    [string[]]
    $Path
  )
  
  Process
  {
    ForEach($_ in $Path)
    {
      $FolderExists = Test-Path -Path $_ -PathType Container
      if (!$FolderExists) { 
        Write-Warning -Message "$_ did not exist. Folder created."
        $null = New-Item -Path $_ -ItemType Directory 
      }
    }
  }
}
