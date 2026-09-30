# Registers a daily 10:00 task for KakaoSender.exe (same folder as this script)
$exe = Join-Path $PSScriptRoot "KakaoSender.exe"
if (-not (Test-Path $exe)) { Write-Host "KakaoSender.exe not found in $PSScriptRoot"; exit 1 }

$action    = New-ScheduledTaskAction -Execute $exe -WorkingDirectory $PSScriptRoot
$trigger   = New-ScheduledTaskTrigger -Daily -At "10:00"
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName "KakaoOpenChatSender" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host ""
Write-Host "OK - Task 'KakaoOpenChatSender' registered: every day 10:00"
