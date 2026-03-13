# Hand2Note 4 - Diagnostic Tool

Diagnoses why Hand2Note 4 hangs on launch (spinning mouse, black screen, never opens).

## How to Use

1. Download both files to your Windows machine:
   - `RUN_AS_ADMIN.bat`
   - `H2N_Diagnose.ps1`
   (keep them in the same folder)

2. Right-click `RUN_AS_ADMIN.bat` → **Run as Administrator**

3. Wait ~45 seconds (it will briefly launch Hand2Note to observe it)

4. A report called `H2N_DiagnosticReport.html` will open on your Desktop

5. Send the report file back for analysis

## What It Checks

- System specs (CPU, RAM, OS)
- Available memory and disk space
- GPU driver version and status
- .NET runtime versions installed
- Hand2Note install path and database sizes
- Running processes and CPU/RAM usage
- Windows Event Log errors (last 7 days) — app crashes, .NET errors, driver issues
- Live process monitor: launches H2N and watches it for 30 seconds
- Startup programs (antivirus conflicts)
- Crash dump files
- Temp folder size
