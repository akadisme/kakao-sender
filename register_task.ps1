# Registers a daily task for KakaoSender.exe (works from any folder/drive, e.g. C:\KakaoSender or D:\KakaoSender)
# To change the time, edit the line below (24h format, e.g. "09:30", "14:00")
$time = "11:00"

$exe = Join-Path $PSScriptRoot "KakaoSender.exe"
if (-not (Test-Path $exe)) { Write-Host "KakaoSender.exe not found in $PSScriptRoot"; Read-Host "Press Enter to close"; exit 1 }

$action    = New-ScheduledTaskAction -Execute $exe -Argument "--auto" -WorkingDirectory $PSScriptRoot
$trigger   = New-ScheduledTaskTrigger -Daily -At $time
# No time limit, run on battery, run as soon as possible if the time was missed
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero)
# Highest = run as administrator (needed because KakaoTalk runs as administrator)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName "KakaoOpenChatSender" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host ""
Write-Host "OK - Task 'KakaoOpenChatSender' registered: every day $time (admin, no time limit)"
Write-Host "Folder: $PSScriptRoot"
Read-Host "Press Enter to close"
