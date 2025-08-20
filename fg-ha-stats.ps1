# ================= About=====================
# This is a powershell script that will fetch important metrics from a Windows Server
# and publish these to a MQTT broker. The script will try to get these metrics:
# * CPU
# * Memory
# * Disks
# * NVIDIA GPU (using nvidia-smi)

# ================= Settings =================
$broker       = "<MQTT IP/hostname>" # Or HA-server if using MQTT HA addon
$port         = 1883                 # 8883 if TLS (then set $useTls = $true)
$useTls       = $false
$mqttUser     = "<MQTT username>"    # fill if broker requires auth
$mqttPass     = "<MQTT password>"

$topicBase    = "fg/server"          # reuse same base as your FG script
$discovery    = "homeassistant"      # HA discovery prefix
$deviceId     = "fg_server"
$deviceName   = "Fantasy Grounds Server"

$MosquittoPub = "C:\Program Files\mosquitto\mosquitto_pub.exe"

# Entities go Unavailable if no update within this window
$expire       = 180  # seconds

# ================= Helpers =================
function PubRaw($topic, $msg, [switch]$Retain) {
  if (-not (Test-Path $MosquittoPub)) { throw "mosquitto_pub not found at: $MosquittoPub" }
  $args = @("-h", $broker, "-p", $port.ToString(), "-t", $topic, "-m", $msg)
  if ($mqttUser) { $args += @("-u", $mqttUser, "-P", $mqttPass) }
  if ($Retain)   { $args += "-r" }
  if ($useTls)   { $args += "-s" }
  & "$MosquittoPub" @args | Out-Null
}
function Pub($subtopic, $msg, [switch]$Retain) { PubRaw "$topicBase/$subtopic" $msg $Retain }

# Publish JSON as UTF-8 **without BOM** (HA discovery is picky)
function PubJson($topic, $obj, [switch]$Retain) {
  if (-not (Test-Path $MosquittoPub)) { throw "mosquitto_pub not found at: $MosquittoPub" }
  $json = $obj | ConvertTo-Json -Compress
  $tmp  = [System.IO.Path]::GetTempFileName()
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($tmp, $json, $utf8NoBom)
  $args = @("-h", $broker, "-p", $port.ToString(), "-t", $topic, "-f", $tmp)
  if ($mqttUser) { $args += @("-u", $mqttUser, "-P", $mqttPass) }
  if ($Retain)   { $args += "-r" }
  if ($useTls)   { $args += "-s" }
  & "$MosquittoPub" @args | Out-Null
  Remove-Item $tmp -Force
}
function SanitizeId($s) { ($s -replace '[^A-Za-z0-9]+','_').Trim('_').ToLower() }

# Find nvidia-smi robustly (handles System32, Program Files, PATH)
function Find-NvSmi {
  $candidates = @(
    "$env:SystemRoot\System32\nvidia-smi.exe",
    "$env:ProgramW6432\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
    "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
    "$env:ProgramFiles(x86)\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
  )
  try {
    $cmd = Get-Command nvidia-smi.exe -ErrorAction Stop
    if ($cmd -and (Test-Path $cmd.Source)) { $candidates = ,$cmd.Source + $candidates }
  } catch {}
  foreach ($p in $candidates) { if ($p -and (Test-Path $p)) { return $p } }
  return $null
}

# Device block (shared)
$dev = @{
  identifiers  = @($deviceId)
  name         = $deviceName
  manufacturer = "Custom"
  model        = "Windows Server"
  sw_version   = "Stats via MQTT"
}

# Availability (retained)
Pub "availability" "online" -Retain

# ============= CPU (locale-proof via WMI/CIM) =============
$cpuPct = $null
try {
  $cpuPct = [int](Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
} catch {
  try {
    $cpuPct = [int]((Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average)
  } catch { $cpuPct = $null }
}
if ($cpuPct -lt 0) { $cpuPct = 0 }
if ($cpuPct -gt 100) { $cpuPct = 100 }

$cfgCpu = @{
  name               = "FG CPU %"
  unique_id          = "${deviceId}_cpu_percent"
  state_topic        = "$topicBase/stats/cpu_percent"
  unit_of_measurement= "%"
  icon               = "mdi:cpu-64-bit"
  state_class        = "measurement"
  availability_topic = "$topicBase/availability"
  expire_after       = $expire
  device             = $dev
}
PubJson "$discovery/sensor/$deviceId/cpu_percent/config" $cfgCpu -Retain
if ($cpuPct -ne $null) { Pub "stats/cpu_percent" ([string]$cpuPct) -Retain }

# ============= Memory =============
$memUsedPct = $null; $memUsedMB = $null; $memTotalMB = $null
try {
  $os = Get-CimInstance Win32_OperatingSystem
  $totalKB = [int64]$os.TotalVisibleMemorySize
  $freeKB  = [int64]$os.FreePhysicalMemory
  $usedKB  = $totalKB - $freeKB
  $memUsedPct  = [math]::Round(($usedKB / [double]$totalKB) * 100, 0)
  $memUsedMB   = [math]::Round($usedKB / 1024.0, 0)
  $memTotalMB  = [math]::Round($totalKB / 1024.0, 0)
} catch {}

$cfgMemPct = @{
  name="FG Memory %"; unique_id="${deviceId}_memory_percent"; state_topic="$topicBase/stats/memory_percent"
  unit_of_measurement="%"; icon="mdi:memory"; state_class="measurement"; availability_topic="$topicBase/availability"
  expire_after=$expire; device=$dev
}
$cfgMemUsed = @{
  name="FG Memory Used"; unique_id="${deviceId}_memory_used_mb"; state_topic="$topicBase/stats/memory_used_mb"
  unit_of_measurement="MB"; icon="mdi:memory"; state_class="measurement"; availability_topic="$topicBase/availability"
  expire_after=$expire; device=$dev
}
$cfgMemTotal = @{
  name="FG Memory Total"; unique_id="${deviceId}_memory_total_mb"; state_topic="$topicBase/stats/memory_total_mb"
  unit_of_measurement="MB"; icon="mdi:memory"; state_class="measurement"; availability_topic="$topicBase/availability"
  expire_after=$expire; device=$dev
}
PubJson "$discovery/sensor/$deviceId/memory_percent/config"  $cfgMemPct  -Retain
PubJson "$discovery/sensor/$deviceId/memory_used_mb/config"  $cfgMemUsed -Retain
PubJson "$discovery/sensor/$deviceId/memory_total_mb/config" $cfgMemTotal -Retain
if ($memUsedPct -ne $null) { Pub "stats/memory_percent"  ([string]$memUsedPct)  -Retain }
if ($memUsedMB  -ne $null) { Pub "stats/memory_used_mb"  ([string]$memUsedMB)   -Retain }
if ($memTotalMB -ne $null) { Pub "stats/memory_total_mb" ([string]$memTotalMB)  -Retain }

# ============= Disks (fixed drives) =============
try {
  $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
  foreach ($d in $disks) {
    if ($d.Size -and $d.FreeSpace -ne $null) {
      $letter  = ($d.DeviceID -replace ':','').ToLower()
      $freePct = [math]::Round( ([double]$d.FreeSpace / [double]$d.Size) * 100, 0 )
      $freeGB  = [math]::Round( [double]$d.FreeSpace / 1GB, 1 )

      $uidPct = "${deviceId}_disk_${letter}_free_percent"
      $uidGB  = "${deviceId}_disk_${letter}_free_gb"
      $topicCfgPct = "$discovery/sensor/$deviceId/disk_${letter}_free_percent/config"
      $topicCfgGB  = "$discovery/sensor/$deviceId/disk_${letter}_free_gb/config"

      $cfgPct = @{
        name="FG Disk $($d.DeviceID) Free %"; unique_id=$uidPct; state_topic="$topicBase/stats/disk/$letter/free_percent"
        unit_of_measurement="%"; icon="mdi:harddisk"; state_class="measurement"; availability_topic="$topicBase/availability"
        expire_after=$expire; device=$dev
      }
      $cfgGB = @{
        name="FG Disk $($d.DeviceID) Free GB"; unique_id=$uidGB; state_topic="$topicBase/stats/disk/$letter/free_gb"
        unit_of_measurement="GB"; icon="mdi:harddisk"; state_class="measurement"; availability_topic="$topicBase/availability"
        expire_after=$expire; device=$dev
      }

      PubJson $topicCfgPct $cfgPct -Retain
      PubJson $topicCfgGB  $cfgGB  -Retain
      Pub "stats/disk/$letter/free_percent" ([string]$freePct) -Retain
      Pub "stats/disk/$letter/free_gb"      ([string]$freeGB)  -Retain
    }
  }
} catch {}

# ============= NVIDIA GPU (auto-locate nvidia-smi) =============
function Find-NvSmi {
  $candidates = @(
    "$env:SystemRoot\System32\nvidia-smi.exe",
    "$env:ProgramW6432\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
    "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
    "$env:ProgramFiles(x86)\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
  )
  try {
    $cmd = Get-Command nvidia-smi.exe -ErrorAction Stop
    if ($cmd -and (Test-Path $cmd.Source)) { $candidates = ,$cmd.Source + $candidates }
  } catch {}
  foreach ($p in $candidates) { if ($p -and (Test-Path $p)) { return $p } }
  return $null
}

$NvSmi = Find-NvSmi
if ($NvSmi) {
  try {
    # One query for everything (util, VRAM, power)
    $lines = & $NvSmi --query-gpu=name,utilization.gpu,memory.total,memory.used,power.draw,power.limit --format=csv,noheader,nounits
    $idx = 0
    foreach ($line in $lines) {
      $parts = ($line -split ',').ForEach({ $_.Trim() })
      if ($parts.Count -lt 6) { continue }

      $name    = $parts[0]
      $util    = [int]$parts[1]
      $memTot  = [int]$parts[2]      # MB
      $memUse  = [int]$parts[3]      # MB
      $pDrawS  = $parts[4]           # may be "N/A"
      $pLimitS = $parts[5]           # may be "N/A"

      $powerW   = $null
      $pLimitW  = $null
      if ($pDrawS  -and $pDrawS  -notin @("N/A","NaN")) { $powerW  = [math]::Round([double]$pDrawS, 0) }
      if ($pLimitS -and $pLimitS -notin @("N/A","NaN")) { $pLimitW = [math]::Round([double]$pLimitS, 0) }

      $id   = SanitizeId(("gpu${idx}_" + $name))
      $base = "$topicBase/stats/gpu/$id"

      # ---- Discovery (retained) ----
      $cfgGpuUtil = @{
        name="FG GPU${idx} Util % ($name)"; unique_id="${deviceId}_${id}_util_percent"
        state_topic="$base/util_percent"; unit_of_measurement="%"; icon="mdi:chart-bell-curve"
        state_class="measurement"; availability_topic="$topicBase/availability"; expire_after=$expire; device=$dev
      }
      $cfgVramUsed = @{
        name="FG GPU${idx} VRAM Used"; unique_id="${deviceId}_${id}_vram_used_mb"
        state_topic="$base/vram_used_mb"; unit_of_measurement="MB"; icon="mdi:memory"
        state_class="measurement"; availability_topic="$topicBase/availability"; expire_after=$expire; device=$dev
      }
      $cfgVramTot = @{
        name="FG GPU${idx} VRAM Total"; unique_id="${deviceId}_${id}_vram_total_mb"
        state_topic="$base/vram_total_mb"; unit_of_measurement="MB"; icon="mdi:memory"
        state_class="measurement"; availability_topic="$topicBase/availability"; expire_after=$expire; device=$dev
      }
      PubJson "$discovery/sensor/$deviceId/${id}_util_percent/config"  $cfgGpuUtil  -Retain
      PubJson "$discovery/sensor/$deviceId/${id}_vram_used_mb/config"  $cfgVramUsed -Retain
      PubJson "$discovery/sensor/$deviceId/${id}_vram_total_mb/config" $cfgVramTot  -Retain

      # NEW: GPU Power draw (W)
      $cfgPower = @{
        name="FG GPU${idx} Power (W) ($name)"; unique_id="${deviceId}_${id}_power_w"
        state_topic="$base/power_w"; unit_of_measurement="W"; device_class="power"; state_class="measurement"
        availability_topic="$topicBase/availability"; expire_after=$expire; device=$dev
      }
      PubJson "$discovery/sensor/$deviceId/${id}_power_w/config" $cfgPower -Retain

      # NEW: GPU Power limit (W) – optional (only if reported)
      if ($pLimitW -ne $null) {
        $cfgPowerLimit = @{
          name="FG GPU${idx} Power Limit (W) ($name)"; unique_id="${deviceId}_${id}_power_limit_w"
          state_topic="$base/power_limit_w"; unit_of_measurement="W"; state_class="measurement"
          availability_topic="$topicBase/availability"; expire_after=$expire; device=$dev
        }
        PubJson "$discovery/sensor/$deviceId/${id}_power_limit_w/config" $cfgPowerLimit -Retain
      }

      # ---- States ----
      Pub "stats/gpu/$id/util_percent"  ([string]$util)   -Retain
      Pub "stats/gpu/$id/vram_used_mb"  ([string]$memUse) -Retain
      Pub "stats/gpu/$id/vram_total_mb" ([string]$memTot) -Retain
      if ($powerW  -ne $null) { Pub "stats/gpu/$id/power_w"       ([string]$powerW)  -Retain }
      if ($pLimitW -ne $null) { Pub "stats/gpu/$id/power_limit_w" ([string]$pLimitW) -Retain }

      $idx++
    }
  } catch {
    # ignore; entities will expire if no updates
  }
}
# else: nvidia-smi not found; skip GPU metrics
