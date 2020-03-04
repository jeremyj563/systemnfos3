# Purpose: Delete the destination folder that was created by ChocolateyInstall.ps1.
#          This is only necessary since the free version of chocolatey doesn't
#          support custom install paths.

# Dot source the init file
. "$PSScriptRoot\init.ps1"

# Delete destination folder if it exists
If (Test-Path -Path $destPath) {
    Write-Host ("`nRemoving directory: {0}`n" -f $destPath) -ForegroundColor Green
    [System.IO.Directory]::Delete($destPath, $true)
}