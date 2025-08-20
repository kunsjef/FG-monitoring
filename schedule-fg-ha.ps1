# schedule-fg-ha.ps1
# Run in an elevated PowerShell

[CmdletBinding()]
param(
  [int]$EveryMinutes = 1,
  [string]$MonitorPath = "C:\Scripts\fg-ha-monitor.ps1",
  [string]$StatsPath   = "C:\Scripts\fg-ha-stats.ps1"
)

function New-RepeatingTask {
  param(
    [Parameter(Mandatory)] [string] $TaskName,
    [Parameter(Mandatory)] [string] $ScriptPath,
    [int] $EveryMinutes = 1
  )
  if (-not (Test-Path $ScriptPath)) { throw "Script not found: $ScriptPath" }

  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

  $psArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
  $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $psArgs

  $start   = (Get-Date).AddMinutes(1)
  $trigger = New-ScheduledTaskTrigger -Once -At $start `
             -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes) `
             -RepetitionDuration (New-TimeSpan -Days 3650)

  try { $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest }
  catch { $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest }

  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
}

New-RepeatingTask -TaskName "FG-HA-Monitor" -ScriptPath $MonitorPath -EveryMinutes $EveryMinutes
New-RepeatingTask -TaskName "FG-HA-Stats"   -ScriptPath $StatsPath   -EveryMinutes $EveryMinutes

# kick them once now (optional)
Start-ScheduledTask -TaskName "FG-HA-Monitor" -ErrorAction SilentlyContinue
Start-ScheduledTask -TaskName "FG-HA-Stats"   -ErrorAction SilentlyContinue
