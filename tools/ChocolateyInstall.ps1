# Purpose: Copy install folder to production Applications folder since the
#          free version of chocolatey doesn't support custom install paths.

# Dot source the init file
. "$PSScriptRoot\init.ps1"

# Delete previous version folder if it exists (eg. 'choco upgrade')
if (Test-Path -Path $preVerPath) {
    Write-Host ("`nRemoving previous version directory: {0}`n" -f $preVerPath) -ForegroundColor Green
    [System.IO.Directory]::Delete("$preVerPath", $true)
}

# Move previous version files to previous version folder
Write-Host ("Moving previous version files to: {0}`n" -f $preVerPath) -ForegroundColor Green
New-Item -Path "$preVerPath" -ItemType Directory -Force | Out-Null
Copy-Item -Path "$destPath\*" -Destination "$preVerPath" -Exclude "$preVerDir" -Force | Out-Null
Get-ChildItem -Path "$destPath" -Exclude "$preVerDir" | Remove-Item -Recurse -Force

# Copy new version files to destination folder
Write-Host ("Copying new version files: {0} -> {1}`n" -f $srcPath, $destPath) -ForegroundColor Green
Copy-Item -Path "$srcPath\*" -Destination "$destPath" -Force -Exclude ".chocolateyPending", "tools" | Out-Null

# Create ignore files so that chocolatey doesn't shim the binaries
foreach ($exe in $exesToIgnore) {
    $null > "$exe.ignore"
}

# Set ACL entries for all installed files
Write-Host ("Applying access control entries`n") -ForegroundColor Green
$acl = Get-Acl -Path $destPath
Get-ChildItem -Path $destPath -Recurse -Force | Set-Acl -AclObject $acl

if ($params['LAUNCH'] -eq $true) {
    # The LAUNCH package parameter was set so launch the application
    Write-Host ("Launching {0}`n" -f $appName) -ForegroundColor Green
    Start-Process -FilePath "$destPath\$appName.exe"
}