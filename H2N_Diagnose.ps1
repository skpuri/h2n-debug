# ============================================================
# Hand2Note 4 - Diagnostic Script
# Run as Administrator for best results
# Outputs: H2N_DiagnosticReport.html on your Desktop
# ============================================================

$ErrorActionPreference = "SilentlyContinue"
$report = @()
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$desktop = [Environment]::GetFolderPath("Desktop")
$reportPath = "$desktop\H2N_DiagnosticReport.html"

function Add-Section($title) {
    $script:report += "<h2>$title</h2>"
}

function Add-Row($label, $value, $status = "ok") {
    $color = switch ($status) {
        "warn"  { "#fff3cd" }
        "bad"   { "#f8d7da" }
        default { "#d4edda" }
    }
    $script:report += "<tr style='background:$color'><td><b>$label</b></td><td>$value</td></tr>"
}

function Add-Table($open = $true) {
    if ($open) { $script:report += "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;margin-bottom:20px'>" }
    else        { $script:report += "</table>" }
}

function Add-Raw($html) {
    $script:report += $html
}

Write-Host "`n[H2N Diagnostics] Starting... please wait`n" -ForegroundColor Cyan

# ─── 1. SYSTEM INFO ─────────────────────────────────────────
Add-Section "System Information"
Add-Table

$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$cs = Get-CimInstance Win32_ComputerSystem

Add-Row "Hostname"        $env:COMPUTERNAME
Add-Row "OS"              "$($os.Caption) $($os.OSArchitecture)"
Add-Row "OS Build"        $os.BuildNumber
Add-Row "Last Boot"       $os.LastBootUpTime
Add-Row "CPU"             $cpu.Name
Add-Row "CPU Cores"       "$($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) logical"
Add-Row "Total RAM"       "$([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"

Add-Table $false

# ─── 2. MEMORY ──────────────────────────────────────────────
Add-Section "Memory"
Add-Table

$totalMB  = [math]::Round($os.TotalVisibleMemorySize / 1024)
$freeMB   = [math]::Round($os.FreePhysicalMemory / 1024)
$usedMB   = $totalMB - $freeMB
$usedPct  = [math]::Round(($usedMB / $totalMB) * 100)

$memStatus = if ($freeMB -lt 1024) { "bad" } elseif ($freeMB -lt 2048) { "warn" } else { "ok" }

Add-Row "Total RAM"     "$totalMB MB"
Add-Row "Used RAM"      "$usedMB MB ($usedPct%)" $memStatus
Add-Row "Free RAM"      "$freeMB MB" $memStatus

# Virtual memory / page file
$pageFile = Get-CimInstance Win32_PageFileUsage | Select-Object -First 1
if ($pageFile) {
    Add-Row "Page File Usage" "$($pageFile.CurrentUsage) MB / $($pageFile.AllocatedBaseSize) MB"
}

Add-Table $false

# ─── 3. DISK SPACE ──────────────────────────────────────────
Add-Section "Disk Space"
Add-Table

Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -ne $null } | ForEach-Object {
    $totalGB = [math]::Round(($_.Used + $_.Free) / 1GB, 1)
    $freeGB  = [math]::Round($_.Free / 1GB, 1)
    $usedPct = if ($totalGB -gt 0) { [math]::Round((($totalGB - $freeGB) / $totalGB) * 100) } else { 0 }
    $diskStatus = if ($freeGB -lt 5) { "bad" } elseif ($freeGB -lt 20) { "warn" } else { "ok" }
    Add-Row "Drive $($_.Name):" "Free: $freeGB GB / $totalGB GB ($usedPct% used)" $diskStatus
}

Add-Table $false

# ─── 4. GPU / DISPLAY ───────────────────────────────────────
Add-Section "GPU & Display"
Add-Table

$gpus = Get-CimInstance Win32_VideoController
foreach ($gpu in $gpus) {
    $gpuRam = if ($gpu.AdapterRAM) { "$([math]::Round($gpu.AdapterRAM / 1MB)) MB" } else { "Unknown" }
    Add-Row "GPU"            $gpu.Name
    Add-Row "Driver Version" $gpu.DriverVersion
    Add-Row "Driver Date"    $gpu.DriverDate
    Add-Row "VRAM"           $gpuRam
    Add-Row "Status"         $gpu.Status (if ($gpu.Status -ne "OK") { "bad" } else { "ok" })
}

Add-Table $false

# ─── 5. .NET RUNTIMES ───────────────────────────────────────
Add-Section ".NET Runtimes Installed"
Add-Table

# .NET Framework
$netFxKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"
$netVersions = Get-ChildItem $netFxKey -Recurse | Get-ItemProperty -Name Version -ErrorAction SilentlyContinue |
    Where-Object { $_.Version } | Select-Object -ExpandProperty Version | Sort-Object -Unique
if ($netVersions) {
    Add-Row ".NET Framework" ($netVersions -join ", ")
} else {
    Add-Row ".NET Framework" "None detected" "warn"
}

# .NET Core / .NET 5+
$dotnetCmd = Get-Command dotnet -ErrorAction SilentlyContinue
if ($dotnetCmd) {
    $runtimes = & dotnet --list-runtimes 2>$null
    if ($runtimes) {
        $runtimeStr = ($runtimes | Where-Object { $_ -match "Microsoft\." }) -join "<br>"
        Add-Row ".NET Runtimes" $runtimeStr
    }
} else {
    Add-Row ".NET Core/5+" "dotnet CLI not in PATH" "warn"
}

Add-Table $false

# ─── 6. HAND2NOTE INSTALL ───────────────────────────────────
Add-Section "Hand2Note 4 Installation"
Add-Table

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
            Add-Row "Exe Size" "$([math]::Round($exe.Length / 1MB, 1)) MB"
            Add-Row "Last Modified" $exe.LastWriteTime
        }
        # Check database size
        $dbFiles = Get-ChildItem $path -Filter "*.db" -Recurse -ErrorAction SilentlyContinue
        foreach ($db in $dbFiles) {
            $dbMB = [math]::Round($db.Length / 1MB, 1)
            $dbStatus = if ($dbMB -gt 5000) { "warn" } else { "ok" }
            Add-Row "Database: $($db.Name)" "$dbMB MB" $dbStatus
        }
    }
}

if (-not $h2nFound) {
    # Try registry
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

Add-Table $false

# ─── 7. RUNNING PROCESSES (H2N + Related) ───────────────────
Add-Section "Relevant Running Processes"
Add-Table

$targetProcs = @("Hand2Note", "h2n", "dotnet", "conhost", "WerFault")
$allProcs = Get-Process | Where-Object {
    $name = $_.ProcessName.ToLower()
    $targetProcs | Where-Object { $name -like "*$($_.ToLower())*" }
}

# CPU-heavy processes
$heavyProcs = Get-Process | Sort-Object CPU -Descending | Select-Object -First 10

Add-Raw "<tr style='background:#e2e3e5'><td><b>Process</b></td><td><b>PID | CPU | RAM</b></td></tr>"

foreach ($p in $allProcs) {
    $ram = [math]::Round($p.WorkingSet64 / 1MB, 1)
    Add-Row $p.ProcessName "PID: $($p.Id) | CPU: $($p.CPU)s | RAM: $ram MB"
}

Add-Raw "<tr><td colspan='2'><b>Top 10 CPU processes at time of scan:</b></td></tr>"
foreach ($p in $heavyProcs) {
    $ram = [math]::Round($p.WorkingSet64 / 1MB, 1)
    Add-Row $p.ProcessName "PID: $($p.Id) | CPU: $([math]::Round($p.CPU, 1))s | RAM: $ram MB"
}

Add-Table $false

# ─── 8. WINDOWS EVENT LOG ERRORS ────────────────────────────
Add-Section "Windows Event Log — Application Errors (Last 7 Days)"
Add-Table

$since = (Get-Date).AddDays(-7)
$events = Get-WinEvent -FilterHashtable @{
    LogName   = 'Application'
    Level     = 2  # Error
    StartTime = $since
} -MaxEvents 100 -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "Hand2Note|h2n|dotnet|\.NET|clr|crash" -or $_.ProviderName -match "Hand2Note|\.NET" }

if ($events) {
    foreach ($e in $events | Select-Object -First 20) {
        $msg = $e.Message -replace "<[^>]+>", "" | Select-String "." | Select-Object -First 3 | ForEach-Object { $_.Line.Trim() }
        Add-Row "$($e.TimeCreated.ToString('MM-dd HH:mm')) [$($e.ProviderName)]" ($msg -join " | ") "bad"
    }
} else {
    Add-Row "No Hand2Note/CLR errors found" "in last 7 days" "ok"
}

# Also check System log for display/freeze issues
$sysEvents = Get-WinEvent -FilterHashtable @{
    LogName   = 'System'
    Level     = 2
    StartTime = $since
} -MaxEvents 200 -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "display|video|driver|freeze|hang|timeout|black" }

if ($sysEvents) {
    Add-Raw "<tr><td colspan='2'><b>System Events — Display/Driver Issues:</b></td></tr>"
    foreach ($e in $sysEvents | Select-Object -First 10) {
        $msg = ($e.Message -split "`n")[0].Trim()
        Add-Row "$($e.TimeCreated.ToString('MM-dd HH:mm')) [$($e.ProviderName)]" $msg "bad"
    }
}

Add-Table $false

# ─── 9. STARTUP + LAUNCH TEST ───────────────────────────────
Add-Section "Hand2Note Launch Monitor (30s observation)"
Add-Table

$h2nExe = $null
foreach ($path in $h2nPaths) {
    $found = Get-ChildItem $path -Filter "Hand2Note*.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { $h2nExe = $found.FullName; break }
}

if ($h2nExe) {
    Add-Row "Attempting launch" $h2nExe
    Add-Table $false

    Write-Host "[H2N Diagnostics] Launching Hand2Note 4 and monitoring for 30 seconds..." -ForegroundColor Yellow

    $proc = Start-Process $h2nExe -PassThru -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    Add-Section "Hand2Note Process Monitor (sampled over 30s)"
    Add-Table
    Add-Raw "<tr style='background:#e2e3e5'><td><b>Time</b></td><td><b>CPU | RAM | Status | Threads</b></td></tr>"

    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 5
        $h2nProc = Get-Process -Name "Hand2Note*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($h2nProc) {
            $ram = [math]::Round($h2nProc.WorkingSet64 / 1MB, 1)
            $cpu = [math]::Round($h2nProc.CPU, 2)
            $status = if ($h2nProc.Responding) { "Responding" } else { "NOT RESPONDING" }
            $statusColor = if ($h2nProc.Responding) { "ok" } else { "bad" }
            Add-Row "T+$((($i+1)*5))s" "CPU: $cpu s | RAM: $ram MB | Status: $status | Threads: $($h2nProc.Threads.Count)" $statusColor
        } else {
            Add-Row "T+$((($i+1)*5))s" "Process not found (may have crashed or not started)" "bad"
        }
    }
    Add-Table $false

    # Kill it after observation
    Stop-Process -Name "Hand2Note*" -Force -ErrorAction SilentlyContinue

} else {
    Add-Table $false
    Add-Section "Hand2Note Launch Monitor"
    Add-Table
    Add-Row "Skipped" "Hand2Note 4 executable not found — cannot auto-launch" "warn"
    Add-Table $false
}

# ─── 10. STARTUP ITEMS ──────────────────────────────────────
Add-Section "Startup Programs (potential conflicts)"
Add-Table

$startupItems = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location
foreach ($item in $startupItems) {
    $isAV = $item.Name -match "antivirus|defender|norton|avast|avg|kaspersky|bitdefender|malware|eset|mcafee"
    $status = if ($isAV) { "warn" } else { "ok" }
    Add-Row $item.Name "$($item.Command) [$($item.Location)]" $status
}

Add-Table $false

# ─── 11. ANTIVIRUS ──────────────────────────────────────────
Add-Section "Security Software"
Add-Table

$av = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
if ($av) {
    foreach ($a in $av) {
        Add-Row "Antivirus" "$($a.displayName) — State: $($a.productState)" "warn"
    }
} else {
    Add-Row "Antivirus" "None detected or WMI access denied" "ok"
}

Add-Table $false

# ─── 12. TEMP / CRASH DUMPS ─────────────────────────────────
Add-Section "Crash Dumps & Temp Files"
Add-Table

$dumpPaths = @(
    "$env:LOCALAPPDATA\CrashDumps",
    "C:\Windows\Minidump",
    "$env:TEMP"
)

foreach ($dumpPath in $dumpPaths) {
    if (Test-Path $dumpPath) {
        $dumps = Get-ChildItem $dumpPath -Filter "*.dmp" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 5
        foreach ($d in $dumps) {
            Add-Row "Crash Dump" "$($d.FullName) — $($d.LastWriteTime) — $([math]::Round($d.Length/1MB,1)) MB" "bad"
        }
        # Temp folder size
        if ($dumpPath -eq $env:TEMP) {
            $tempSize = (Get-ChildItem $env:TEMP -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
            $tempMB = [math]::Round($tempSize / 1MB)
            $tempStatus = if ($tempMB -gt 5000) { "warn" } else { "ok" }
            Add-Row "Temp Folder Size" "$tempMB MB" $tempStatus
        }
    }
}

Add-Table $false

# ─── BUILD REPORT ───────────────────────────────────────────
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
  <h1>🔍 Hand2Note 4 — Diagnostic Report</h1>
  <p class="meta">Generated: $timestamp | Machine: $env:COMPUTERNAME | User: $env:USERNAME</p>
  $($report -join "`n")
  <p class="meta" style="margin-top:40px">End of report. Share this file with your support contact.</p>
</body>
</html>
"@

$html | Out-File -FilePath $reportPath -Encoding UTF8
Write-Host "`n[H2N Diagnostics] Done!" -ForegroundColor Green
Write-Host "[H2N Diagnostics] Report saved to: $reportPath" -ForegroundColor Cyan
Write-Host "[H2N Diagnostics] Opening report in browser..." -ForegroundColor Cyan
Start-Process $reportPath
