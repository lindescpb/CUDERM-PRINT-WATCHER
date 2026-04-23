# ===============================================
#   Cuderm Logis Printer
#   Cuderm is an application built using my free time, time I could've spent with my family.
#   The name represents the initials of my loved ones: 
#   Chantelle, Ullyndyss (junior), Daniel, Eli and Raniyah Maloy ❤️
#   Version 1.2.4
# ===============================================

$Version = "1.2.4"

# Config
$configFolder = "$env:APPDATA\CudermLogisPrinter"
$configPath   = "$configFolder\config.json"

function Load-Config {
    if (Test-Path $configPath) {
        try { return Get-Content $configPath -Raw | ConvertFrom-Json } catch { }
    }
    return $null
}

function Save-Config {
    param($watch, $printed, $log, $printer)
    if (!(Test-Path $configFolder)) { New-Item -ItemType Directory -Path $configFolder -Force | Out-Null }
    @{ WatchFolder = $watch; PrintedFolder = $printed; LogFile = $log; Printer = $printer; LastVersion = $Version } | 
        ConvertTo-Json | Out-File $configPath -Encoding utf8
}

# Single Instance
$mutexName = "CudermLogisPrinterMutex"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)

if (-not $mutex.WaitOne(0, $false)) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Cuderm Logis Printer is already running!`n`nThis appears to be a newer version.`n`nDo you want to close the old version and update now?",
        "Cuderm Logis Printer - Update Available",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -eq "Yes") {
        Get-Process -Name "*cuderm*" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
    } else { exit }
}
$global:mutex = $mutex

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$cfg = Load-Config

# ============== GUI =====================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Cuderm Logis Printer Setup"
$form.Size = New-Object System.Drawing.Size(560, 440)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblWatch = New-Object System.Windows.Forms.Label; $lblWatch.Location = New-Object System.Drawing.Point(20,30); $lblWatch.Size = New-Object System.Drawing.Size(160,20); $lblWatch.Text = "Watch Folder:"; $form.Controls.Add($lblWatch)
$txtWatch = New-Object System.Windows.Forms.TextBox; $txtWatch.Location = New-Object System.Drawing.Point(190,28); $txtWatch.Size = New-Object System.Drawing.Size(310,20); $txtWatch.Text = "C:\install\logisprint"; $form.Controls.Add($txtWatch)
$btnWatch = New-Object System.Windows.Forms.Button; $btnWatch.Location = New-Object System.Drawing.Point(510,27); $btnWatch.Size = New-Object System.Drawing.Size(30,23); $btnWatch.Text = "..."; $btnWatch.Add_Click({ $fbd=New-Object System.Windows.Forms.FolderBrowserDialog; if($fbd.ShowDialog() -eq "OK"){$txtWatch.Text=$fbd.SelectedPath} }); $form.Controls.Add($btnWatch)

$lblPrinted = New-Object System.Windows.Forms.Label; $lblPrinted.Location = New-Object System.Drawing.Point(20,70); $lblPrinted.Size = New-Object System.Drawing.Size(160,20); $lblPrinted.Text = "Printed Folder:"; $form.Controls.Add($lblPrinted)
$txtPrinted = New-Object System.Windows.Forms.TextBox; $txtPrinted.Location = New-Object System.Drawing.Point(190,68); $txtPrinted.Size = New-Object System.Drawing.Size(310,20); $txtPrinted.Text = "C:\install\logisprint\printed"; $form.Controls.Add($txtPrinted)
$btnPrinted = New-Object System.Windows.Forms.Button; $btnPrinted.Location = New-Object System.Drawing.Point(510,67); $btnPrinted.Size = New-Object System.Drawing.Size(30,23); $btnPrinted.Text = "..."; $btnPrinted.Add_Click({ $fbd=New-Object System.Windows.Forms.FolderBrowserDialog; if($fbd.ShowDialog() -eq "OK"){$txtPrinted.Text=$fbd.SelectedPath} }); $form.Controls.Add($btnPrinted)

$lblLog = New-Object System.Windows.Forms.Label; $lblLog.Location = New-Object System.Drawing.Point(20,110); $lblLog.Size = New-Object System.Drawing.Size(160,20); $lblLog.Text = "Log File:"; $form.Controls.Add($lblLog)
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Location = New-Object System.Drawing.Point(190,108); $txtLog.Size = New-Object System.Drawing.Size(340,20); $txtLog.Text = "C:\install\logisprint\print-log.txt"; $form.Controls.Add($txtLog)

$lblPrinter = New-Object System.Windows.Forms.Label; $lblPrinter.Location = New-Object System.Drawing.Point(20,150); $lblPrinter.Size = New-Object System.Drawing.Size(160,20); $lblPrinter.Text = "Printer:"; $form.Controls.Add($lblPrinter)
$cmbPrinter = New-Object System.Windows.Forms.ComboBox; $cmbPrinter.Location = New-Object System.Drawing.Point(190,148); $cmbPrinter.Size = New-Object System.Drawing.Size(340,20); $cmbPrinter.DropDownStyle = "DropDownList"; $form.Controls.Add($cmbPrinter)

$printers = @()
try { $printers = Get-Printer | Select-Object -ExpandProperty Name -ErrorAction Stop }
catch {
    try { $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Select-Object -ExpandProperty Name }
    catch { try { $printers = Get-WmiObject -Class Win32_Printer -ErrorAction Stop | Select-Object -ExpandProperty Name } catch {} }
}

if ($printers.Count -gt 0) {
    $printers | Sort-Object | ForEach-Object { [void]$cmbPrinter.Items.Add($_) }
    $cmbPrinter.SelectedIndex = 0
} else {
    [void]$cmbPrinter.Items.Add("--- No printers found ---")
}

if ($cfg) {
    $txtWatch.Text = $cfg.WatchFolder
    $txtPrinted.Text = $cfg.PrintedFolder
    $txtLog.Text = $cfg.LogFile
    if ($cmbPrinter.Items.Contains($cfg.Printer)) { $cmbPrinter.SelectedItem = $cfg.Printer }
}

$btnStart = New-Object System.Windows.Forms.Button; $btnStart.Location = New-Object System.Drawing.Point(215,210); $btnStart.Size = New-Object System.Drawing.Size(130,35); $btnStart.Text = "Start Cuderm Logis Printer"; $btnStart.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.AcceptButton = $btnStart; $form.Controls.Add($btnStart)

$result = $form.ShowDialog()
if ($result -ne "OK") { $mutex.ReleaseMutex(); exit }

$folder        = $txtWatch.Text.TrimEnd('\')
$printedFolder = $txtPrinted.Text.TrimEnd('\')
$logFile       = $txtLog.Text
$printer       = $cmbPrinter.SelectedItem

Save-Config $folder $printedFolder $logFile $printer

if (!(Test-Path $printedFolder)) { New-Item -ItemType Directory -Path $printedFolder -Force | Out-Null }
if (!(Test-Path (Split-Path $logFile -Parent))) { New-Item -ItemType Directory -Path (Split-Path $logFile -Parent) -Force | Out-Null }

# ============== CLEAN FUNCTION - Simple & Reliable ==============
function Clean-FileContent {
    param ($filePath)

    $lines = Get-Content $filePath -Encoding UTF8
    $cleaned = @()
    $footerStarted = $false

    foreach ($line in $lines) {
        if ($line -match "Signature|Transit") {
            $footerStarted = $true
        }

        if ($footerStarted) {
            $cleaned += $line
            continue
        }

        # Remove blank lines in main content
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $cleaned += $line
        }
    }

    # Add 2 blank lines before Signature / Transit
    if ($footerStarted -and $cleaned.Count -gt 0) {
        $cleaned += "" 
        $cleaned += "" 
    }

    return $cleaned
}

# ============== CORE FUNCTIONS ==============
function Write-Log {
    param($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp | $message" | Out-File -FilePath $logFile -Append -Encoding utf8
}

function Wait-ForFile {
    param($path)
    while ($true) {
        try {
            $stream = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
            $stream.Close()
            break
        } catch { Start-Sleep -Milliseconds 2000 }
    }
}

function Process-File {
    param($file)
    try {
        Write-Log "START | $(Split-Path $file -Leaf)"
        Wait-ForFile $file
        
        $cleanContent = Clean-FileContent $file
        $cleanContent | Out-Printer -Name $printer
        
        Start-Sleep -Seconds 2
        $dest = Join-Path $printedFolder (Split-Path $file -Leaf)
        Move-Item $file $dest -Force
        
        Write-Log "PRINTED | $(Split-Path $file -Leaf)"
    }
    catch {
        Write-Log "ERROR | $(Split-Path $file -Leaf) | $($_.Exception.Message)"
    }
}

# ============== TRAY ICON - Simplified & More Reliable ==============
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Text = "Cuderm Logis Printer is watching your prints :-)"

try {
    $iconPath = "C:\install\cuderm.ico"   # ← UPDATE THIS PATH
    if (Test-Path $iconPath) {
        $trayIcon.Icon = New-Object System.Drawing.Icon($iconPath)
    } else {
        $trayIcon.Icon = [System.Drawing.SystemIcons]::Application
    }
} catch {
    $trayIcon.Icon = [System.Drawing.SystemIcons]::Application
}

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuStatus = New-Object System.Windows.Forms.ToolStripMenuItem; $menuStatus.Text = "Status: Running`nWatching: $folder`nPrinter: $printer"
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem; $menuExit.Text = "Exit Cuderm Logis Printer"

$menuExit.Add_Click({
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    Write-Log "=== CUDERM LOGIS PRINTER STOPPED BY USER ==="
    $global:mutex.ReleaseMutex()
    Start-Sleep -Milliseconds 500
    [System.Environment]::Exit(0)
})

$contextMenu.Items.Add($menuStatus)
$contextMenu.Items.Add($menuExit)
$trayIcon.ContextMenuStrip = $contextMenu
$trayIcon.Visible = $true

$trayIcon.ShowBalloonTip(5000, "Cuderm Logis Printer", "Now watching your Logis prints 🙂`nVersion $Version", [System.Windows.Forms.ToolTipIcon]::Info)

# ============== START WATCHER ==============
Write-Log "=== CUDERM LOGIS PRINTER v$Version STARTED ==="
Write-Log "Watch Folder : $folder"
Write-Log "Printer      : $printer"

# Process existing files
Get-ChildItem -Path $folder -Filter "*.txt" | ForEach-Object { Process-File $_.FullName }

# Watch for new files
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $folder
$watcher.Filter = "*.txt"
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent $watcher Created -Action { 
    $file = $Event.SourceEventArgs.FullPath
    Process-File $file 
} | Out-Null

Write-Host "Cuderm Logis Printer v$Version is running... (copy .txt files to watch folder)"

while ($true) { Start-Sleep -Seconds 5 }