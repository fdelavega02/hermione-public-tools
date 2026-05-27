param(
    [string]$TaskName = "OutlookLocalInboxSync",
    [string]$RunAt = "07:00"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot

$wslMatch = $null
if ($ProjectRoot -match '^\\\\wsl\.localhost\\([^\\]+)\\(.+)$') {
    $wslMatch = @{ Distro = $Matches[1]; Path = '/' + ($Matches[2] -replace '\\', '/') }
} elseif ($ProjectRoot -match '^\\\\wsl\$\\([^\\]+)\\(.+)$') {
    $wslMatch = @{ Distro = $Matches[1]; Path = '/' + ($Matches[2] -replace '\\', '/') }
} else {
    throw "This script expects to be run from the WSL-backed project path."
}

$LinuxProjectRoot = $wslMatch.Path
$Distro = $wslMatch.Distro
$LinuxConfigPath = "$LinuxProjectRoot/config.json"
$BashCommand = "cd '$LinuxProjectRoot' || exit 1; if [ ! -f '$LinuxConfigPath' ]; then echo 'config.json is missing. Copy config.example.json to config.json first.'; exit 1; fi; mkdir -p output; node sync-inbox.mjs >> output/scheduled-run.log 2>&1"
$QuotedBashCommand = $BashCommand.Replace('"', '\"')

$trigger = New-ScheduledTaskTrigger -Daily -DaysInterval 2 -At $RunAt
$action = New-ScheduledTaskAction -Execute "wsl.exe" -Argument "-d $Distro bash -lc `"$QuotedBashCommand`""
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -Principal $principal `
    -Force | Out-Null

Write-Host "Scheduled task '$TaskName' created."
Write-Host "It will run every 2 days at $RunAt using WSL distro '$Distro'."
Write-Host "Project root: $LinuxProjectRoot"
Write-Host "If Outlook asks for sign-in again, rerun 'npm run auth' manually to refresh saved login state."
