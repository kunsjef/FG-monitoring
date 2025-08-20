# ================= About=====================
# This is a powershell script that will fetch important metrics from a Fantasy Grounds
# and publish these to a MQTT broker. The script will try to get these metrics:
# * Number of active players
# * Names of active players
# * Is the FG application running?
# * Last player event (connects, disconnects ++)
# * Check FG port (normally 1802/UDP)

# ================= Settings =================
$broker       = "<MQTT IP/hostname>" # Or HA-server if using MQTT HA addon
$port         = 1883                 # 8883 if TLS (then set $useTls = $true)
$useTls       = $false
$mqttUser     = "<MQTT username>"    # fill if broker requires auth
$mqttPass     = "<MQTT password>"


$topicBase    = "fg/server"
$discovery    = "homeassistant"      # HA discovery prefix
$deviceId     = "fg_server"
$deviceName   = "Fantasy Grounds Server"

$FGRoot       = "C:\FG"
$FGPort       = 1802                 # FGU uses UDP 1802
$MosquittoPub = "C:\Program Files\mosquitto\mosquitto_pub.exe"

$processNames = @("FantasyGrounds","FGUpdaterEngine","FantasyGroundsUpdater")

$logPaths     = @("$FGRoot\console.log", "$FGRoot\network.log") | Where-Object { Test-Path $_ }
#$logPaths = @("$FGRoot\console.log") | Where-Object { Test-Path $_ }

# ---- Expiry: entities go Unavailable if no updates in this window ----
$expire = 180  # seconds


function To-Hashtable($obj) {
  if ($null -eq $obj) { return @{} }
  if ($obj -is [hashtable]) { return $obj }
  $ht = @{}
  foreach ($p in $obj.PSObject.Properties) { $ht[$p.Name] = $p.Value }
  return $ht
}

# ============= Get players from logs =============
# Build current player set from the last X lines of the logs (robust, idempotent)
function Get-PlayersFromLogs {
  param(
    [string[]] $Logs,
    [int] $Tail = 2000
  )

  # Regexes as here-strings (no escaping headaches)
  $rxConnect = [regex]@'
(?i)['"]\s*(?<name>[^'"]+?)\s*['"]\s+(?:connected|player\s+connected|client\s+connected|joined|login|authenticated)\b
'@

  $rxDisconnect = [regex]@'
(?i)['"]\s*(?<name>[^'"]+?)\s*['"]\s+(?:disconnected|left|disconnect(?:ed)?|connection\s+closed)\b
'@

  # simple set
  $set = @{}
  foreach ($log in $Logs) {
    if (-not (Test-Path $log)) { continue }
    # Get-Content auto-detects UTF-16; -Tail is efficient
    $lines = Get-Content -Path $log -Tail $Tail
    foreach ($line in $lines) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      $m = $rxConnect.Match($line)
      if ($m.Success) { $set[$m.Groups['name'].Value.Trim()] = $true; continue }
      $m = $rxDisconnect.Match($line)
      if ($m.Success) { [void]$set.Remove($m.Groups['name'].Value.Trim()) }
    }
  }
  ,(@($set.Keys | Sort-Object))
}

# ============= Get number of connected =============
# Parse log lines and update connected players
function Apply-PlayerEvents([object]$st, [string[]]$lines) {
  if (-not $lines) { return }
  $rxConnect    = [regex]"'\s*(?<name>[^']+)\s*'\s+connected"
  $rxDisconnect = [regex]"'\s*(?<name>[^']+)\s+(?:disconnected|left|disconnect(?:ed)?|connection\s+closed)"
  foreach ($line in $lines) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $m = $rxConnect.Match($line)
    if ($m.Success) {
      $n = $m.Groups['name'].Value.Trim()
      if ($n -and -not ($st.players -contains $n)) { $st.players += $n }
      continue
    }
    $m = $rxDisconnect.Match($line)
    if ($m.Success) {
      $n = $m.Groups['name'].Value.Trim()
      if ($n) { $st.players = @($st.players | Where-Object { $_ -ne $n }) }
    }
  }
}

function Get-TextEncoding($path) {
  $fs = [System.IO.File]::Open($path,'Open','Read','ReadWrite')
  try {
    $b = New-Object byte[] 3
    $null = $fs.Read($b,0,3)
  } finally { $fs.Close() }
  if ($b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode } # UTF-16 LE
  if ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8 } # UTF-8 BOM
  return [System.Text.Encoding]::UTF8  # default: UTF-8 (no BOM)
}

# Read only the new bytes appended to a log since last run, update file position
function Process-LogDelta([object]$st, [string]$logPath) {
  if (-not (Test-Path $logPath)) { return }
  $fileInfo = Get-Item $logPath
  $len = [int64]$fileInfo.Length
  if (-not $st.pos.ContainsKey($logPath)) { $st.pos[$logPath] = 0 }
  $pos = [int64]$st.pos[$logPath]
  if ($pos -gt $len) { $pos = 0 }  # truncated/rotated

  $fs = [System.IO.File]::Open($logPath, 'Open', 'Read', 'ReadWrite')
  try {
    $fs.Seek($pos, 'Begin') | Out-Null
    $enc = Get-TextEncoding $logPath
    # IMPORTANT: pass detectEncodingFromByteOrderMarks:$false since we're mid-file
    $sr = New-Object System.IO.StreamReader($fs, $enc, $false)
    $new = $sr.ReadToEnd()
    $sr.Close()
  } finally { $fs.Close() }

  if ($new) {
    $lines = $new -split "(`r`n|`n)"
    Apply-PlayerEvents $st $lines
  }
  $st.pos[$logPath] = $len
}

# First-run seeding: replay last N lines to build initial player set
function Seed-FromTail([object]$st, [string]$logPath, [int]$tail = 500) {
  if (-not (Test-Path $logPath)) { return }
  $lines = Get-Content $logPath -Tail $tail
  Apply-PlayerEvents $st $lines
  $st.pos[$logPath] = (Get-Item $logPath).Length
}

# --------------- MQTT helpers --------------------
function PubRaw($topic, $msg, [switch]$Retain) {
  if (-not (Test-Path $MosquittoPub)) { throw "mosquitto_pub not found at: $MosquittoPub" }
  $args = @("-h", $broker, "-p", $port.ToString(), "-t", $topic)
  if ($mqttUser) { $args += @("-u", $mqttUser, "-P", $mqttPass) }
  if ($Retain)   { $args += "-r" }
  if ($useTls)   { $args += "-s" }

  if ([string]::IsNullOrEmpty([string]$msg)) {
    # publish a zero-length payload safely
    $args += "-n"
  } else {
    $args += @("-m", [string]$msg)
  }

  & "$MosquittoPub" @args | Out-Null
}

function Pub($subtopic, $msg, [switch]$Retain) {
  PubRaw "$topicBase/$subtopic" $msg $Retain
}

function PubJson($topic, $obj, [switch]$Retain) {
  if (-not (Test-Path $MosquittoPub)) { throw "mosquitto_pub not found at: $MosquittoPub" }
  $json = $obj | ConvertTo-Json -Compress

  # Write JSON as UTF-8 *without* BOM (avoid the leading ﻿ character)
  $tmp = [System.IO.Path]::GetTempFileName()
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($tmp, $json, $utf8NoBom)

  $args = @("-h", $broker, "-p", $port.ToString(), "-t", $topic, "-f", $tmp)
  if ($mqttUser) { $args += @("-u", $mqttUser, "-P", $mqttPass) }
  if ($Retain)   { $args += "-r" }
  if ($useTls)   { $args += "-s" }
  & "$MosquittoPub" @args | Out-Null
  Remove-Item $tmp -Force
}

# --------------- MQTT Discovery (retained) ---------------
$dev = @{
  identifiers  = @($deviceId)
  name         = $deviceName
  manufacturer = "SmiteWorks"
  model        = "Windows Server"
  sw_version   = "Fantasy Grounds Unity"
}

# Prebuild config topics (short lines; no wrapping)
$topicCfgApp     = "$discovery/binary_sensor/$deviceId/app_running/config"
$topicCfgCount   = "$discovery/sensor/$deviceId/player_count/config"
$topicCfgNames   = "$discovery/sensor/$deviceId/player_names/config"
$topicCfgEvent   = "$discovery/sensor/$deviceId/last_event/config"
$topicCfgListen  = "$discovery/binary_sensor/$deviceId/port_listening/config"

$cfg_app = @{
  name               = "FG App Running"
  unique_id          = "${deviceId}_app_running"
  state_topic        = "$topicBase/app_running"
  payload_on         = "true"
  payload_off        = "false"
  device_class       = "connectivity"
  availability_topic = "$topicBase/availability"
  expire_after       = $expire
  device             = $dev
}
PubJson $topicCfgApp $cfg_app -Retain

$cfg_count = @{
  name               = "FG Player Count"
  unique_id          = "${deviceId}_player_count"
  state_topic        = "$topicBase/player_count"
  unit_of_measurement= "players"
  icon               = "mdi:account-group"
  state_class        = "measurement"
  availability_topic = "$topicBase/availability"
  expire_after       = $expire
  device             = $dev
}
PubJson $topicCfgCount $cfg_count -Retain

$cfg_names = @{
  name               = "FG Player Names"
  unique_id          = "${deviceId}_player_names"
  state_topic        = "$topicBase/player_names"
  icon               = "mdi:account-multiple-outline"
  availability_topic = "$topicBase/availability"
  expire_after       = $expire
  device             = $dev
}
PubJson $topicCfgNames $cfg_names -Retain

$cfg_event = @{
  name               = "FG Last Player Event"
  unique_id          = "${deviceId}_last_event"
  state_topic        = "$topicBase/last_event"
  icon               = "mdi:account-clock-outline"
  availability_topic = "$topicBase/availability"
  #expire_after       = $expire
  device             = $dev
}
PubJson $topicCfgEvent $cfg_event -Retain

$cfg_listen = @{
  name               = "FG Port Listening"
  unique_id          = "${deviceId}_port_listening"
  state_topic        = "$topicBase/port_listening"
  payload_on         = "true"
  payload_off        = "false"
  icon               = "mdi:lan-pending"
  availability_topic = "$topicBase/availability"
  expire_after       = $expire
  device             = $dev
}
PubJson $topicCfgListen $cfg_listen -Retain

# Availability (retained)
Pub "availability" "online" -Retain

# --------------- Measurements ---------------

# 1) App running?
$running = $false
try {
  $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -in $processNames -or ($_.Path -and $_.Path -like "$FGRoot*")
  }
  $running = ($procs.Count -gt 0)
} catch { $running = $false }
Pub "app_running" ($running.ToString().ToLower()) -Retain

# 2) Player count & names via logs (UDP-friendly)
# --- Players (recomputed each run from recent log lines) ---
$players = Get-PlayersFromLogs -Logs $logPaths -Tail 2000
$names   = ($players -join ", ")
Pub "player_names" $names -Retain
Pub "player_count" ([string]$players.Count) -Retain


# 3) Port listening? (UDP + TCP + netstat fallback)
$portListening = $false
try {
  $udpBound = (Get-NetUDPEndpoint -LocalPort $FGPort -ErrorAction Stop | Measure-Object).Count
  if ($udpBound -gt 0) { $portListening = $true }
} catch {}
if (-not $portListening) {
  try {
    $tcpListen = (Get-NetTCPConnection -LocalPort $FGPort -State Listen -ErrorAction Stop | Measure-Object).Count
    if ($tcpListen -gt 0) { $portListening = $true }
  } catch {}
}
if (-not $portListening) {
  $ns = (& netstat -ano -p tcp) 2>$null
  if ($ns) {
    $rx = "[:\.]$FGPort\s+.*LISTENING"
    if ($ns | Select-String -Pattern $rx) { $portListening = $true }
  }
  Pub "player_names" "" -Retain
  Pub "player_count" "0" -Retain
}
Pub "port_listening" ($portListening.ToString().ToLower()) -Retain

# 4) Last player event (persistent)
function Get-LastEventFromLogs {
  param(
    [string[]] $Logs,
    [int] $Tail = 2000
  )

  if (-not $Logs -or $Logs.Count -eq 0) { return $null }

  # Collect recent lines (order preserved: oldest -> newest)
  $buf = @()
  foreach ($log in $Logs) {
    if (Test-Path $log) {
      $buf += (Get-Content -Path $log -Tail $Tail)
    }
  }
  if ($buf.Count -eq 0) { return $null }

  # Regexes (here-strings avoid quote escaping issues)
  $rxNamed = [regex]@'
(?i)['"]\s*([^'"]+?)\s*['"]\s+(?:connected|disconnected|left)\b
'@
  $rxGeneric = [regex]@'
(?i)(client\s+connected|player\s+connected|connection\s+(?:from|accepted)|waiting\s+for\s+authorization|joined|login|authenticated|disconnect(?:ed)?|left|connection\s+closed)
'@

  # Walk from the end (newest -> oldest), prefer named; remember the first generic we see
  $candidateGeneric = $null
  for ($i = $buf.Count - 1; $i -ge 0; $i--) {
    $line = $buf[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($rxNamed.IsMatch($line)) { return $line }
    if (-not $candidateGeneric -and $rxGeneric.IsMatch($line)) { $candidateGeneric = $line }
  }
  return $candidateGeneric
}

$evtLine = Get-LastEventFromLogs -Logs @("$FGRoot\console.log") -Tail 2000
if ($evtLine) { Pub "last_event" $evtLine -Retain }
