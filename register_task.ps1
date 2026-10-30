# Registers KakaoSender.exe to run every 30 minutes, all day.
# The program itself decides which rooms to send, using the "발송시간" column in the Google Sheet.
# Works from any folder/drive (e.g. C:\KakaoSender or D:\KakaoSender)
$intervalMinutes = 30

$exe = Join-Path $PSScriptRoot "KakaoSender.exe"
if (-not (Test-Path $exe)) { Write-Host "KakaoSender.exe not found in $PSScriptRoot"; Read-Host "Press Enter to close"; exit 1 }

$action  = New-ScheduledTaskAction -Execute $exe -Argument "--auto" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Daily -At "00:00"
$repeat  = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Minutes $intervalMinutes) -RepetitionDuration (New-TimeSpan -Hours 23 -Minutes 59)
$trigger.Repetition = $repeat.Repetition
# No time limit, run on battery, run as soon as possible if a time was missed
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -DontStopOnIdleEnd -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero)
# Highest = run as administrator (needed because KakaoTalk runs as administrator)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName "KakaoOpenChatSender" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host ""
Write-Host "OK - Task 'KakaoOpenChatSender' registered: checks every $intervalMinutes minutes (admin, no time limit)"
Write-Host "Send times are set per room in the Google Sheet column: 발송시간"
Write-Host "Folder: $PSScriptRoot"
Read-Host "Press Enter to close"
