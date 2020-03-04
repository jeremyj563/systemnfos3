# Package parameters
$params = Get-PackageParameters
if (!$params['LAUNCH']) {$params['LAUNCH'] = $false}

# Global variables
$appName = "systemnfos3"
$repoName = "it-systemnfos3"
$srcPath = "C:\ProgramData\chocolatey\lib\{0}" -f $repoName
$destPath = "C:\IT\{0}" -f $appName
$preVerDir = "previous_version"
$preVerPath = "{0}\{1}" -f $destPath, $preVerDir
[String[]] $exesToIgnore = "$srcPath\$appName.exe" #, $srcPath\<NextExecutableInArray>.exe

# Close any open instances of the application
if (Get-Process -Name "$appName" -ErrorAction SilentlyContinue) {
    Write-Host ("`nAttempting to stop all running processes for application: {0}" -f $appName) -ForegroundColor Green
    Stop-Process -Name "$appName" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    if (Get-Process -Name "$appName" -ErrorAction SilentlyContinue) {
        throw "INSTALLATION ABORTED: $appName is still running."
    }
}