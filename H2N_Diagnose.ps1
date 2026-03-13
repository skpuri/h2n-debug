# ============================================================
# Hand2Note 4 - Diagnostic Script
# Run as Administrator for best results
# Outputs: H2N_DiagnosticReport.html on your Desktop
# ============================================================

$ErrorActionPreference = "SilentlyContinue"
$report = New-Object System.Collections.ArrayList
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$desktop = [Environment]::GetFolderPath("Desktop")
$reportPath = Join-Path $desktop "H2N_DiagnosticReport.html"

function Add-Section($title) {
    [void]$script:report.Add("<h2>$title</h2>")
}

function Add-Row($label, $value, $status) {
    if (-not $status) { $status = "ok" }
    $color = "#d4edda"
    if ($status -eq "warn") { $color = "#fff3cd" }
    if ($status -eq "bad")  { $color = "#f8d7da" }
    [void]$script:report.Add("<tr style='background:$color'><td><b>$label</b></td><td>$value</td></tr>")
}

function Open-Table {
    [void]$script:report.Add("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;margin-bottom:20px'>")
}

function Close-Table {
    [void]$script:report.Add("</table>")
}

function Add-Raw($html) {
    [void]$script:report.Add($html)
}

Write-Host ""
Write-Host "[H2N Diagnostics] Starting... please wait" -ForegroundColor Cyan
Write-Host ""

# --- 1. SYSTEM INFO ---
Add-Section "System Information"
Open-Table

$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$cs = Get-CimInstance Win32_ComputerSystem

Add-Row "Hostname" $env:COMPUTERNAME
Add-Row "OS" "$($os.Caption) $($os.OSArchitecture)"
Add-Row "OS Build" $os.BuildNumber
Add-Row "Last Boot" $os.LastBootUpTime
Add-Row "CPU" $cpu.Name
$coreStr = "$($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) logical"
Add-Row "CPU Cores" $coreStr
$ramStr = "$([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"
Add-Row "Total RAM" $ramStr

Close-Table

# --- 2. MEMORY ---
Add-Section "Memory"
Open-Table

$totalMB = [math]::Round($os.TotalVisibleMemorySize / 1024)
$freeMB = [math]::Round($os.FreePhysicalMemory / 1024)
$usedMB = $totalMB - $freeMB
$usedPct = 0
if ($totalMB -gt 0) { $usedPct = [math]::Round(($usedMB / $totalMB) * 100) }

$memStatus = "ok"
if ($freeMB -lt 1024) { $memStatus = "bad" }
elseif ($freeMB -lt 2048) { $memStatus = "warn" }

Add-Row "Total RAM" "$totalMB MB"
Add-Row "Used RAM" "$usedMB MB ($usedPct%)" $memStatus
Add-Row "Free RAM" "$freeMB MB" $memStatus

$pageFile = Get-CimInstance Win32_PageFileUsage | Select-Object -First 1
if ($pageFile) {
    $pfStr = "$($pageFile.CurrentUsage) MB / $($pageFile.AllocatedBaseSize) MB"
    Add-Row "Page File Usage" $pfStr
}

Close-Table

# --- 3. DISK SPACE ---
Add-Section "Disk Space"
Open-Table

Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -ne $null } | ForEach-Object {
    $totalGB = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
    $freeGB = [math]::Round($_.Free / 1GB, 1)
    $usedPct2 = 0
    if ($totalGB -gt 0) { $usedPct2 = [math]::Round((($totalGB - $freeGB) / $totalGB) * 100) }
    $diskStatus = "ok"
    if ($freeGB -lt 5) { $diskStatus = "bad" }
    elseif ($freeGB -lt 20) { $diskStatus = "warn" }
    $diskStr = "Free: $freeGB GB / $totalGB GB ($usedPct2% used)"
    Add-Row "Drive $($_.Name):" $diskStr $diskStatus
}

Close-Table

# --- 4. GPU / DISPLAY ---
Add-Section "GPU and Display"
Open-Table

$gpus = Get-CimInstance Win32_VideoController
foreach ($gpu in $gpus) {
    $gpuRam = "Unknown"
    if ($gpu.AdapterRAM) { $gpuRam = "$([math]::Round($gpu.AdapterRAM / 1MB)) MB" }
    Add-Row "GPU" $gpu.Name
    Add-Row "Driver Version" $gpu.DriverVersion
    Add-Row "Driver Date" $gpu.DriverDate
    Add-Row "VRAM" $gpuRam
    $gpuStatus = "ok"
    if ($gpu.Status -ne "OK") { $gpuStatus = "bad" }
    Add-Row "Status" $gpu.Status $gpuStatus
}

Close-Table

# --- 5. .NET RUNTIMES ---
Add-Section ".NET Runtimes Installed"
Open-Table

$netFxKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"
$netVersions = Get-ChildItem $netFxKey -Recurse | Get-ItemProperty -Name Version -ErrorAction SilentlyContinue |
    Where-Object { $_.Version } | Select-Object -ExpandProperty Version | Sort-Object -Unique
if ($netVersions) {
    Add-Row ".NET Framework" ($netVersions -join ", ")
} else {
    Add-Row ".NET Framework" "None detected" "warn"
}

$dotnetCmd = Get-Command dotnet -ErrorAction SilentlyContinue
if ($dotnetCmd) {
    $runtimes = & dotnet --list-runtimes 2>$null
    if ($runtimes) {
        $runtimeList = $runtimes | Where-Object { $_ -match "Microsoft\." }
        $runtimeStr = $runtimeList -join "<br>"
        Add-Row ".NET Runtimes" $runtimeStr
    }
} else {
    Add-Row ".NET Core/5+" "dotnet CLI not in PATH" "warn"
}

Close-Table

# --- 6. HAND2NOTE INSTALL ---
Add-Section "Hand2Note 4 Installation"
Open-Table

$h2nPaths = @(
    "$env:LOCALAPPDATA\Hand2Note4",
    "$env:APPDATA\Hand2Note4",
    "C:\Program Files\Hand2Note4",
    "C:\Program Files (x86)\Hand2Note4",
    "$env:LOCALAPPDATA\Programs\Hand2Note4"
)

$h2nFound = $false
foreach ($path in $h2nPaths) {
    if (Test-Path $path) {
        $h2nFound = $true
        Add-Row "Install Path" $path
        $exe = Get-ChildItem $path -Filter "*.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exe) {
            Add-Row "Executable" $exe.FullName
            $exeSize = "$([math]::Round($exe.Length / 1MB, 1)) MB"
            Add-Row "Exe Size" $exeSize
            Add-Row "Last Modified" $exe.LastWriteTime
        }
        $dbFiles = Get-ChildItem $path -Filter "*.db" -Recurse -ErrorAction SilentlyContinue
        foreach ($db in $dbFiles) {
            $dbMB = [math]::Round($db.Length / 1MB, 1)
            $dbStatus = "ok"
            if ($dbMB -gt 5000) { $dbStatus = "warn" }
            Add-Row "Database: $($db.Name)" "$dbMB MB" $dbStatus
        }
    }
}

if (-not $h2nFound) {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($regPath in $regPaths) {
        Get-ChildItem $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($app.DisplayName -like "*Hand2Note*") {
                Add-Row "Registry Entry" $app.DisplayName
                Add-Row "Install Location" $app.InstallLocation
                Add-Row "Version" $app.DisplayVersion
                $h2nFound = $true
            }
        }
    }
}

if (-not $h2nFound) {
    Add-Row "Hand2Note 4" "Not found in common locations" "warn"
}

Close-Table

# --- 7. RUNNING PROCESSES ---
Add-Section "Relevant Running Processes"
Open-Table

$targetNames = @("Hand2Note", "h2n", "dotnet", "conhost", "WerFault")
$allProcs = Get-Process | Where-Object {
    $pName = $_.ProcessName.ToLower()
    foreach ($t in $targetNames) {
        if ($pName -like "*$($t.ToLower())*") { return $true }
    }
    return $false
}

$heavyProcs = Get-Process | Sort-Object CPU -Descending | Select-Object -First 10

Add-Raw "<tr style='background:#e2e3e5'><td><b>Process</b></td><td><b>PID / CPU / RAM</b></td></tr>"

foreach ($p in $allProcs) {
    $ram = [math]::Round($p.WorkingSet64 / 1MB, 1)
    $cpuVal = [math]::Round($p.CPU, 1)
    $procStr = "PID: $($p.Id) | CPU: $($cpuVal)s | RAM: $($ram) MB"
    Add-Row $p.ProcessName $procStr
}

Add-Raw "<tr><td colspan='2'><b>Top 10 CPU processes at time of scan:</b></td></tr>"
foreach ($p in $heavyProcs) {
    $ram = [math]::Round($p.WorkingSet64 / 1MB, 1)
    $cpuVal = [math]::Round($p.CPU, 1)
    $procStr = "PID: $($p.Id) | CPU: $($cpuVal)s | RAM: $($ram) MB"
    Add-Row $p.ProcessName $procStr
}

Close-Table

# --- 8. WINDOWS EVENT LOG ERRORS ---
Add-Section "Windows Event Log - Application Errors (Last 7 Days)"
Open-Table

$since = (Get-Date).AddDays(-7)
$events = Get-WinEvent -FilterHashtable @{
    LogName   = 'Application'
    Level     = 2
    StartTime = $since
} -MaxEvents 100 -ErrorAction SilentlyContinue

$h2nEvents = $events | Where-Object {
    ($_.Message -match "Hand2Note") -or
    ($_.Message -match "h2n") -or
    ($_.Message -match "dotnet") -or
    ($_.Message -match "\.NET") -or
    ($_.Message -match "clr") -or
    ($_.Message -match "crash") -or
    ($_.ProviderName -match "Hand2Note") -or
    ($_.ProviderName -match "\.NET")
}

if ($h2nEvents) {
    foreach ($e in ($h2nEvents | Select-Object -First 20)) {
        $msgLines = ($e.Message -split "`n") | Select-Object -First 3
        $msgText = ($msgLines | ForEach-Object { $_.Trim() }) -join " | "
        $timeStr = $e.TimeCreated.ToString('MM-dd HH:mm')
        $provStr = $e.ProviderName
        Add-Row "$timeStr [$provStr]" $msgText "bad"
    }
} else {
    Add-Row "No Hand2Note/CLR errors found" "in last 7 days" "ok"
}

$sysEvents = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Level     = 2
    StartTime = $since
} -MaxEvents 200 -ErrorAction SilentlyContinue

$displayEvents = $sysEvents | Where-Object {
    ($_.Message -match "display") -or
    ($_.Message -match "video") -or
    ($_.Message -match "driver") -or
    ($_.Message -match "freeze") -or
    ($_.Message -match "hang") -or
    ($_.Message -match "timeout") -or
    ($_.Message -match "black")
}

if ($displayEvents) {
    Add-Raw "<tr><td colspan='2'><b>System Events - Display/Driver Issues:</b></td></tr>"
    foreach ($e in ($displayEvents | Select-Object -First 10)) {
        $msgFirst = ($e.Message -split "`n")[0].Trim()
        $timeStr = $e.TimeCreated.ToString('MM-dd HH:mm')
        $provStr = $e.ProviderName
        Add-Row "$timeStr [$provStr]" $msgFirst "bad"
    }
}

Close-Table

# --- 9. STARTUP + LAUNCH TEST ---
Add-Section "Hand2Note Launch Monitor (30s observation)"
Open-Table

$h2nExe = $null
foreach ($path in $h2nPaths) {
    $found = Get-ChildItem $path -Filter "Hand2Note*.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { $h2nExe = $found.FullName; break }
}

if ($h2nExe) {
    Add-Row "Attempting launch" $h2nExe
    Close-Table

    Write-Host "[H2N Diagnostics] Launching Hand2Note 4 and monitoring for 30 seconds..." -ForegroundColor Yellow

    $proc = Start-Process $h2nExe -PassThru -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    Add-Section "Hand2Note Process Monitor (sampled over 30s)"
    Open-Table
    Add-Raw "<tr style='background:#e2e3e5'><td><b>Time</b></td><td><b>CPU / RAM / Status / Threads</b></td></tr>"

    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 5
        $elapsed = ($i + 1) * 5
        $h2nProc = Get-Process -Name "Hand2Note*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($h2nProc) {
            $ram = [math]::Round($h2nProc.WorkingSet64 / 1MB, 1)
            $cpuVal = [math]::Round($h2nProc.CPU, 2)
            $threadCount = $h2nProc.Threads.Count
            $statusColor = "ok"
            $statusText = "Responding"
            if (-not $h2nProc.Responding) {
                $statusColor = "bad"
                $statusText = "NOT RESPONDING"
            }
            $monStr = "CPU: $($cpuVal)s | RAM: $($ram) MB | Status: $statusText | Threads: $threadCount"
            Add-Row "T+$($elapsed)s" $monStr $statusColor
        } else {
            Add-Row "T+$($elapsed)s" "Process not found (may have crashed or not started)" "bad"
        }
    }
    Close-Table

    Stop-Process -Name "Hand2Note*" -Force -ErrorAction SilentlyContinue
} else {
    Close-Table
    Add-Section "Hand2Note Launch Monitor"
    Open-Table
    Add-Row "Skipped" "Hand2Note 4 executable not found - cannot auto-launch" "warn"
    Close-Table
}

# --- 10. STARTUP ITEMS ---
Add-Section "Startup Programs (potential conflicts)"
Open-Table

$startupItems = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location
foreach ($item in $startupItems) {
    $isAV = $item.Name -match "antivirus|defender|norton|avast|avg|kaspersky|bitdefender|malware|eset|mcafee"
    $startStatus = "ok"
    if ($isAV) { $startStatus = "warn" }
    $startStr = "$($item.Command) [$($item.Location)]"
    Add-Row $item.Name $startStr $startStatus
}

Close-Table

# --- 11. ANTIVIRUS ---
Add-Section "Security Software"
Open-Table

$av = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
if ($av) {
    foreach ($a in $av) {
        $avStr = "$($a.displayName) - State: $($a.productState)"
        Add-Row "Antivirus" $avStr "warn"
    }
} else {
    Add-Row "Antivirus" "None detected or WMI access denied" "ok"
}

Close-Table

# --- 12. CRASH DUMPS ---
Add-Section "Crash Dumps and Temp Files"
Open-Table

$dumpPaths = @(
    "$env:LOCALAPPDATA\CrashDumps",
    "C:\Windows\Minidump",
    "$env:TEMP"
)

foreach ($dumpPath in $dumpPaths) {
    if (Test-Path $dumpPath) {
        $dumps = Get-ChildItem $dumpPath -Filter "*.dmp" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 5
        foreach ($d in $dumps) {
            $dumpSize = [math]::Round($d.Length / 1MB, 1)
            $dumpStr = "$($d.FullName) - $($d.LastWriteTime) - $dumpSize MB"
            Add-Row "Crash Dump" $dumpStr "bad"
        }
        if ($dumpPath -eq $env:TEMP) {
            $tempSize = (Get-ChildItem $env:TEMP -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
            $tempMB = [math]::Round($tempSize / 1MB)
            $tempStatus = "ok"
            if ($tempMB -gt 5000) { $tempStatus = "warn" }
            Add-Row "Temp Folder Size" "$tempMB MB" $tempStatus
        }
    }
}

Close-Table

# --- BUILD REPORT ---
$reportBody = $report -join "`n"

$html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Hand2Note 4 - Diagnostic Report</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; background: #f5f5f5; padding: 20px; color: #333; }
    h1   { background: #1a1a2e; color: white; padding: 15px 20px; border-radius: 8px; }
    h2   { color: #1a1a2e; border-bottom: 2px solid #1a1a2e; padding-bottom: 5px; margin-top: 30px; }
    table { font-size: 13px; }
    td   { padding: 6px 10px; }
    .meta { color: #666; font-size: 12px; margin-bottom: 30px; }
  </style>
</head>
<body>
  <h1>Hand2Note 4 - Diagnostic Report</h1>
  <p class="meta">Generated: $timestamp | Machine: $env:COMPUTERNAME | User: $env:USERNAME</p>
  $reportBody
  <p class="meta" style="margin-top:40px">End of report. Share this file with your support contact.</p>
</body>
</html>
"@

$html | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host ""
Write-Host "[H2N Diagnostics] Done!" -ForegroundColor Green
Write-Host "[H2N Diagnostics] Report saved to: $reportPath" -ForegroundColor Cyan
Write-Host "[H2N Diagnostics] Opening report in browser..." -ForegroundColor Cyan
Start-Process $reportPath
