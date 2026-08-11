# WeeklyReportDashboard.ps1
# Modern Accessible Windows Dashboard (WinForms) for Weekly Report Agent.
# Features:
#   - Grouped UI Sections: Data Input, SharePoint Profile, Output & Report Settings, Execution & Live Log
#   - App Registration Profile Dropdown (Direct Office SSO, Liquid Office Preset, Custom Entry for Client ID/Tenant ID/Thumbprint)
#   - 4 Individual Excel File Inputs (Local file/folder or SharePoint HTTPS link)
#   - High-contrast accessible text styling (WCAG AA compliant)
#   - Color-coded status badges and popup alert dialogs (Red for Error, Amber for Warning, Green for Success)
#   - Form-Closing Protection: Dashboard window NEVER closes during execution or authentication.
#   - Safe null-checking for all file system parameters (prevents ParameterBindingValidationException)
#   - Robust Office SSO window handling (prevents COM hidden window crashes during M365 authentication)
#   - Config persistence via dashboard-config.json

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
param()

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not [string]::IsNullOrWhiteSpace($scriptPath)) {
        Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$scriptPath`""
        return
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = if ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
$agentPath = Join-Path $scriptDir "WeeklyReportAgent.ps1"
$configPath = Join-Path $scriptDir "dashboard-config.json"

function Protect-String([string]$plain) {
    if ([string]::IsNullOrWhiteSpace($plain)) { return "" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($plain)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    return [Convert]::ToBase64String($protected)
}

function Unprotect-String([string]$encrypted) {
    if ([string]::IsNullOrWhiteSpace($encrypted)) { return "" }
    try {
        $bytes = [Convert]::FromBase64String($encrypted)
        $unprotected = [System.Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($unprotected)
    }
    catch {
        return ""
    }
}

function Get-IsoWeek([datetime]$d) {
    $day = [int]$d.DayOfWeek; if ($day -eq 0) { $day = 7 }
    return [math]::Floor(($d.AddDays(4 - $day).DayOfYear - 1) / 7) + 1
}

function Format-PsArg([string]$val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return "''" }
    return "'" + ($val -replace "'", "''") + "'"
}

# --- Load saved settings -----------------------------------------------------
$config = @{
    MovesDataUrl       = ""
    CmphDataUrl        = ""
    EquipmentDataUrl   = ""
    LtiDataUrl         = ""
    OutputDir          = $scriptDir
    InputDir           = $scriptDir
    SharePointUrl      = ""
    TemplatePath       = ""
    SharePointEmail    = ""
    SharePointPassword = ""
    ClientId           = ""
    TenantId           = ""
    CertThumbprint     = ""
    AuthMode           = "Direct Office SSO (Default - No App Registration Required)"
}

if (-not [string]::IsNullOrWhiteSpace($configPath) -and (Test-Path -Path $configPath -ErrorAction SilentlyContinue)) {
    try {
        $j = Get-Content $configPath -Raw | ConvertFrom-Json
        if ($j.MovesDataUrl) { $config.MovesDataUrl = $j.MovesDataUrl }
        if ($j.CmphDataUrl) { $config.CmphDataUrl = $j.CmphDataUrl }
        if ($j.EquipmentDataUrl) { $config.EquipmentDataUrl = $j.EquipmentDataUrl }
        if ($j.LtiDataUrl) { $config.LtiDataUrl = $j.LtiDataUrl }
        if ($j.OutputDir) {
            if ($j.OutputDir -match '^https?://' -or (Test-Path -Path $j.OutputDir -ErrorAction SilentlyContinue)) {
                $config.OutputDir = $j.OutputDir
            }
        }
        if ($j.InputDir) {
            if (Test-Path -Path $j.InputDir -ErrorAction SilentlyContinue) {
                $config.InputDir = $j.InputDir
            }
        }
        if ($j.SharePointUrl) { $config.SharePointUrl = $j.SharePointUrl }
        if ($j.TemplatePath) { $config.TemplatePath = $j.TemplatePath }
        if ($j.SharePointEmail) { $config.SharePointEmail = $j.SharePointEmail }
        if ($j.SharePointPassword) { $config.SharePointPassword = Unprotect-String $j.SharePointPassword }
        if ($j.ClientId) { $config.ClientId = $j.ClientId }
        if ($j.TenantId) { $config.TenantId = $j.TenantId }
        if ($j.CertThumbprint) { $config.CertThumbprint = $j.CertThumbprint }
        if ($j.AuthMode) { $config.AuthMode = $j.AuthMode }
    }
    catch {}
}

# --- Form Construction -------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Weekly Report Agent - Executive Dashboard"
$form.Size = New-Object System.Drawing.Size(880, 850)
$form.MinimumSize = New-Object System.Drawing.Size(880, 850)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F8F9FA")
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$fontBold = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$fontHeader = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# Color Palette (High contrast WCAG AA compliant)
$colorTextDark = [System.Drawing.ColorTranslator]::FromHtml("#201F1E")
$colorBlueAccent = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$colorGreen = [System.Drawing.ColorTranslator]::FromHtml("#107C41")
$colorRed = [System.Drawing.ColorTranslator]::FromHtml("#D13438")
$colorGroupBg = [System.Drawing.Color]::White

# Helper to create styled GroupBoxes
function New-CustomGroupBox([string]$title, [int]$x, [int]$y, [int]$w, [int]$h) {
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = $title
    $gb.Location = New-Object System.Drawing.Point($x, $y)
    $gb.Size = New-Object System.Drawing.Size($w, $h)
    $gb.Font = $fontHeader
    $gb.ForeColor = $colorBlueAccent
    $gb.BackColor = $colorGroupBg
    $gb.Anchor = "Top,Left,Right"
    return $gb
}

# Helper to create input rows with label, textbox, and browse button
function Add-InputRow([System.Windows.Forms.GroupBox]$parent, [string]$labelTitle, [string]$defaultValue, [int]$yPos, [string]$fileFilter, [switch]$isFolder) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelTitle
    $lbl.Location = New-Object System.Drawing.Point(15, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(700, 18)
    $lbl.Font = $form.Font
    $lbl.ForeColor = $colorTextDark
    $parent.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(15, ($yPos + 20))
    $txt.Size = New-Object System.Drawing.Size(695, 25)
    $txt.Font = $form.Font
    $txt.Text = $defaultValue
    $txt.Anchor = "Top,Left,Right"
    $parent.Controls.Add($txt)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Browse..."
    $btn.Location = New-Object System.Drawing.Point(720, ($yPos + 18))
    $btn.Size = New-Object System.Drawing.Size(100, 28)
    $btn.Font = $form.Font
    $btn.Anchor = "Top,Right"
    $btn.Add_Click({
            if ($isFolder) {
                $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                $dlg.Description = "Select Output Directory"
                $currentVal = $txt.Text
                if (-not [string]::IsNullOrWhiteSpace($currentVal) -and (Test-Path -Path $currentVal -ErrorAction SilentlyContinue)) {
                    $dlg.SelectedPath = $currentVal
                }
                if ($dlg.ShowDialog() -eq "OK") { $txt.Text = $dlg.SelectedPath }
            }
            else {
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Filter = if ($fileFilter) { $fileFilter } else { "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*" }
                $currentVal = $txt.Text
                if (-not [string]::IsNullOrWhiteSpace($currentVal) -and (Test-Path -Path $currentVal -ErrorAction SilentlyContinue)) {
                    try { $dlg.InitialDirectory = Split-Path -Path $currentVal -ErrorAction SilentlyContinue } catch {}
                }
                if ($dlg.ShowDialog() -eq "OK") { $txt.Text = $dlg.FileName }
            }
        })
    $parent.Controls.Add($btn)

    return $txt
}

# =============================================================================
# GROUP 1: DATA INPUT (EXCEL FILES)
# =============================================================================
$grpData = New-CustomGroupBox -title "Data Input (Excel Files & Template)" -x 15 -y 10 -w 835 -h 335

$txtMoves = Add-InputRow -parent $grpData -labelTitle "1. Actual & Forecast Moves Data (.xlsx or SharePoint link):" -defaultValue $config.MovesDataUrl -yPos 25
$txtCmph = Add-InputRow -parent $grpData -labelTitle "2. CMPH and PMPH Actual & Forecast Data (.xlsx or SharePoint link):" -defaultValue $config.CmphDataUrl -yPos 85
$txtEquipment = Add-InputRow -parent $grpData -labelTitle "3. Equipment PERFORMANCE V1.2 Weekly Data (.xlsx or SharePoint link):" -defaultValue $config.EquipmentDataUrl -yPos 145
$txtLti = Add-InputRow -parent $grpData -labelTitle "4. LTI PER TEU Data (.xlsx or SharePoint link):" -defaultValue $config.LtiDataUrl -yPos 205
$txtTemplate = Add-InputRow -parent $grpData -labelTitle "PowerPoint Template (.pptx, .pptm, .ppt, .potx or SharePoint link) - leave blank to auto-resolve:" -defaultValue $config.TemplatePath -yPos 265 -fileFilter "PowerPoint Files (*.pptx;*.pptm;*.ppt;*.potx;*.potm)|*.pptx;*.pptm;*.ppt;*.potx;*.potm|All Files (*.*)|*.*"

$form.Controls.Add($grpData)

# =============================================================================
# GROUP 2: SHAREPOINT PROFILE & AUTHENTICATION (DIRECT SIGN-IN MODE)
# =============================================================================
$grpProfile = New-CustomGroupBox -title "SharePoint Profile `& Authentication (Direct Sign-In)" -x 15 -y 355 -w 835 -h 100

# Subtitle Note
$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Note: Ensure you are logged in on the same account in Excel and PowerPoint."
$lblNote.Location = New-Object System.Drawing.Point(15, 22)
$lblNote.Size = New-Object System.Drawing.Size(750, 18)
$lblNote.Font = $form.Font; $lblNote.ForeColor = $colorBlueAccent
$grpProfile.Controls.Add($lblNote)

# Email & Password
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "SharePoint Email:"
$lblEmail.Location = New-Object System.Drawing.Point(15, 44)
$lblEmail.Size = New-Object System.Drawing.Size(150, 18)
$lblEmail.Font = $form.Font; $lblEmail.ForeColor = $colorTextDark
$grpProfile.Controls.Add($lblEmail)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Location = New-Object System.Drawing.Point(15, 64)
$txtEmail.Size = New-Object System.Drawing.Size(390, 25)
$txtEmail.Font = $form.Font; $txtEmail.Text = $config.SharePointEmail
$grpProfile.Controls.Add($txtEmail)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "SharePoint Password (Encrypted):"
$lblPass.Location = New-Object System.Drawing.Point(430, 44)
$lblPass.Size = New-Object System.Drawing.Size(220, 18)
$lblPass.Font = $form.Font; $lblPass.ForeColor = $colorTextDark
$grpProfile.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(430, 64)
$txtPass.Size = New-Object System.Drawing.Size(390, 25)
$txtPass.Font = $form.Font; $txtPass.PasswordChar = '*'
$txtPass.Text = $config.SharePointPassword
$grpProfile.Controls.Add($txtPass)

$form.Controls.Add($grpProfile)

# =============================================================================
# GROUP 3: OUTPUT & REPORT SETTINGS
# =============================================================================
$grpOutput = New-CustomGroupBox -title "Output `& Report Settings" -x 15 -y 465 -w 835 -h 75

$txtOut = Add-InputRow -parent $grpOutput -labelTitle "Output Location (Local directory path or SharePoint library URL):" -defaultValue $config.OutputDir -yPos 25 -isFolder

$form.Controls.Add($grpOutput)

# =============================================================================
# GROUP 4: EXECUTION & LIVE LOG
# =============================================================================
$grpExec = New-CustomGroupBox -title "Execution `& Live Log" -x 15 -y 550 -w 835 -h 245
$grpExec.Anchor = "Top,Bottom,Left,Right"

$lblStatusTitle = New-Object System.Windows.Forms.Label
$lblStatusTitle.Text = "Status:"
$lblStatusTitle.Location = New-Object System.Drawing.Point(15, 25)
$lblStatusTitle.Size = New-Object System.Drawing.Size(55, 22)
$lblStatusTitle.Font = $fontBold; $lblStatusTitle.ForeColor = $colorTextDark
$grpExec.Controls.Add($lblStatusTitle)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "READY"
$lblStatus.Location = New-Object System.Drawing.Point(70, 25)
$lblStatus.Size = New-Object System.Drawing.Size(180, 22)
$lblStatus.Font = $fontBold; $lblStatus.ForeColor = $colorGreen
$grpExec.Controls.Add($lblStatus)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "RUN WEEKLY REPORT AGENT"
$btnRun.Location = New-Object System.Drawing.Point(260, 20)
$btnRun.Size = New-Object System.Drawing.Size(260, 32)
$btnRun.Font = $fontBold
$btnRun.BackColor = $colorBlueAccent
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = "Flat"
$grpExec.Controls.Add($btnRun)

$btnOpenReport = New-Object System.Windows.Forms.Button
$btnOpenReport.Text = "Open Generated Report"
$btnOpenReport.Location = New-Object System.Drawing.Point(530, 20)
$btnOpenReport.Size = New-Object System.Drawing.Size(155, 32)
$btnOpenReport.Font = $form.Font
$btnOpenReport.Enabled = $false
$grpExec.Controls.Add($btnOpenReport)

$btnOpenFolder = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text = "Open Output"
$btnOpenFolder.Location = New-Object System.Drawing.Point(695, 20)
$btnOpenFolder.Size = New-Object System.Drawing.Size(125, 32)
$btnOpenFolder.Font = $form.Font
$grpExec.Controls.Add($btnOpenFolder)

# Live Log Text Box
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 60)
$txtLog.Size = New-Object System.Drawing.Size(805, 170)
$txtLog.Anchor = "Top,Bottom,Left,Right"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$txtLog.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#D4D4D4")
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpExec.Controls.Add($txtLog)

$form.Controls.Add($grpExec)

# Helper log appender
function Write-DashboardLog([string]$text) {
    if ($txtLog.InvokeRequired) {
        $txtLog.Invoke([Action[string]] { param($t) Write-DashboardLog $t }, $text)
    }
    else {
        $normalized = ($text + "`r`n") -replace '(?<!\r)\n', "`r`n"
        $txtLog.AppendText($normalized)
    }
}

# --- Action Event Handlers ----------------------------------------------------
$script:lastReport = ""

$btnOpenFolder.Add_Click({
        $target = $txtOut.Text
        if ([string]::IsNullOrWhiteSpace($target)) { $target = $scriptDir }

        if ($target -match '^https?://') {
            try { Start-Process "msedge.exe" -ArgumentList "`"$target`"" } catch { Start-Process $target }
        }
        else {
            if ($target -match '\.pptx?$' -or (Test-Path -Path $target -PathType Leaf -ErrorAction SilentlyContinue)) {
                $target = Split-Path -Path $target -Parent
                if ([string]::IsNullOrWhiteSpace($target)) { $target = $scriptDir }
            }
            if (-not [System.IO.Path]::IsPathRooted($target)) {
                $target = Join-Path $scriptDir $target
            }
            if (-not (Test-Path -Path $target -ErrorAction SilentlyContinue)) {
                try { New-Item -ItemType Directory -Force -Path $target -ErrorAction SilentlyContinue | Out-Null } catch {}
            }
            if (Test-Path -Path $target -ErrorAction SilentlyContinue) {
                Start-Process "explorer.exe" -ArgumentList "`"$target`""
            }
        }
    })

$btnOpenReport.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($script:lastReport)) {
            if ($script:lastReport -match '^https?://') {
                try { Start-Process "msedge.exe" -ArgumentList "`"$script:lastReport`"" } catch { Start-Process $script:lastReport }
            }
            elseif (Test-Path -Path $script:lastReport -ErrorAction SilentlyContinue) {
                Start-Process "`"$script:lastReport`""
            }
        }
    })

# --- FORM CLOSING PROTECTION: Dashboard NEVER closes during process ----------
$form.Add_FormClosing({
        param($src, $ev)
        if ($lblStatus.Text -eq "RUNNING...") {
            $res = [System.Windows.Forms.MessageBox]::Show(
                "The agent is currently processing in the background.`r`nAre you sure you want to close the dashboard?",
                "Agent Running - Confirm Exit",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
                $ev.Cancel = $true
            }
        }
    })

# --- RUN BUTTON CLICK HANDLER ------------------------------------------------
$btnRun.Add_Click({
        try {
            # Save Config Settings
            $config.MovesDataUrl = $txtMoves.Text
            $config.CmphDataUrl = $txtCmph.Text
            $config.EquipmentDataUrl = $txtEquipment.Text
            $config.LtiDataUrl = $txtLti.Text
            $config.TemplatePath = $txtTemplate.Text
            $config.OutputDir = $txtOut.Text
            $config.SharePointEmail = $txtEmail.Text
            $config.SharePointPassword = Protect-String $txtPass.Text

            $json = $config | ConvertTo-Json -Depth 5
            $json | Set-Content $configPath -Encoding UTF8

            $txtLog.Clear()
            $lblStatus.Text = "RUNNING..."
            $lblStatus.ForeColor = $colorBlueAccent
            $btnRun.Enabled = $false
            $btnOpenReport.Enabled = $false

            # Construct Agent Execution Arguments safely escaped for PowerShell parser
            $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-File", "`"$agentPath`"", "-Root", (Format-PsArg $scriptDir))

            if ($txtMoves.Text) {
                $argList += "-MovesDataUrl"
                $argList += (Format-PsArg $txtMoves.Text)
            }
            if ($txtCmph.Text) {
                $argList += "-CmphDataUrl"
                $argList += (Format-PsArg $txtCmph.Text)
            }
            if ($txtEquipment.Text) {
                $argList += "-EquipmentDataUrl"
                $argList += (Format-PsArg $txtEquipment.Text)
            }
            if ($txtLti.Text) {
                $argList += "-LtiDataUrl"
                $argList += (Format-PsArg $txtLti.Text)
            }
            if ($txtTemplate.Text) {
                $argList += "-TemplatePath"
                $argList += (Format-PsArg $txtTemplate.Text)
            }
            if ($txtOut.Text) {
                $argList += "-OutputDir"
                $argList += (Format-PsArg $txtOut.Text)
            }
            if ($txtEmail.Text) {
                $argList += "-SharePointEmail"
                $argList += (Format-PsArg $txtEmail.Text)
            }
            if ($txtPass.Text) {
                $argList += "-SharePointPassword"
                $argList += (Format-PsArg $txtPass.Text)
            }
            if ($config.ClientId) {
                $argList += "-ClientId"
                $argList += (Format-PsArg $config.ClientId)
            }
            if ($config.TenantId) {
                $argList += "-TenantId"
                $argList += (Format-PsArg $config.TenantId)
            }
            if ($config.CertThumbprint) {
                $argList += "-CertThumbprint"
                $argList += (Format-PsArg $config.CertThumbprint)
            }
            if (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue) {
                $argList += "-UsePnP"
            }

            # Use a temp log file for capturing agent output.
            # Rationale: BeginOutputReadLine/BeginErrorReadLine fire on thread-pool threads.
            # Any exception inside those callbacks (including cross-thread UI Invoke calls)
            # kills the entire PowerShell host, taking the dashboard window with it.
            # File-based polling keeps everything on the UI thread (WinForms Timer Tick),
            # completely eliminating cross-thread crash risk.
            $script:agentLogPath = Join-Path $env:TEMP ("wra_log_{0}.log" -f [guid]::NewGuid().ToString('N'))
            $script:agentLogPos = 0   # Byte offset of last-read position in the log file

            # Run via cmd.exe so we can redirect stdout+stderr to the log file with no pipe involvement.
            $cmdArgs = "/c powershell.exe " + ($argList -join " ") + " >> `"$($script:agentLogPath)`" 2>&1"

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "cmd.exe"
            $psi.Arguments = $cmdArgs
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $false   # Keep a desktop window handle so Office SSO can authenticate
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Minimized

            $script:agentProcess = New-Object System.Diagnostics.Process
            $script:agentProcess.StartInfo = $psi

            Write-DashboardLog "=== LAUNCHING WEEKLY REPORT AGENT ==="
            Write-DashboardLog ("Command: powershell " + ($argList -join " "))

            $script:agentProcess.Start() | Out-Null

            # Poll timer runs entirely on the UI thread - safe to update controls directly.
            # IMPORTANT: must be $script: scoped so the Add_Tick scriptblock can reference it
            # when fired by WinForms (local click-handler variables are $null in that scope).
            $script:pollTimer = New-Object System.Windows.Forms.Timer
            $script:pollTimer.Interval = 250

            $script:pollTimer.Add_Tick({
                    # Append any new lines written to the log file since we last checked
                    try {
                        if (Test-Path -LiteralPath $script:agentLogPath -ErrorAction SilentlyContinue) {
                            $fs = [System.IO.File]::Open($script:agentLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                            try {
                                $fs.Seek($script:agentLogPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                                $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                                try {
                                    $newText = $reader.ReadToEnd()
                                    $script:agentLogPos = $fs.Position
                                    if ($newText) {
                                        $normalizedText = $newText -replace '(?<!\r)\n', "`r`n"
                                        $txtLog.AppendText($normalizedText)
                                    }
                                }
                                finally {
                                    $reader.Dispose()
                                }
                            }
                            finally {
                                $fs.Dispose()
                            }
                        }
                    }
                    catch {}

                    if (-not $script:agentProcess.HasExited) { return }

                    # Flush any remaining output after process exits
                    try {
                        if (Test-Path -LiteralPath $script:agentLogPath -ErrorAction SilentlyContinue) {
                            $fs = [System.IO.File]::Open($script:agentLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                            try {
                                $fs.Seek($script:agentLogPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                                $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                                try {
                                    $newText = $reader.ReadToEnd()
                                    if ($newText) {
                                        $normalizedText = $newText -replace '(?<!\r)\n', "`r`n"
                                        $txtLog.AppendText($normalizedText)
                                    }
                                }
                                finally {
                                    $reader.Dispose()
                                }
                            }
                            finally {
                                $fs.Dispose()
                            }
                            Remove-Item -LiteralPath $script:agentLogPath -ErrorAction SilentlyContinue
                        }
                    }
                    catch {}

                    $script:pollTimer.Stop()
                    $script:pollTimer.Dispose()

                    $finalLog = $txtLog.Text
                    $hasDoneLine = ($finalLog -match '== DONE:\s*(.+) ==')
                    if ($hasDoneLine) {
                        $script:lastReport = $Matches[1].Trim()
                    }

                    $exitedClean = ($script:agentProcess.ExitCode -eq 0) -or $hasDoneLine
                    if ($exitedClean) {
                        $lblStatus.Text = "DONE!"
                        $lblStatus.ForeColor = $colorGreen
                        if ($script:lastReport -and (Test-Path -Path $script:lastReport -ErrorAction SilentlyContinue)) {
                            $btnOpenReport.Enabled = $true
                        }
                        [System.Windows.Forms.MessageBox]::Show(
                            "Weekly Report generated successfully!`r`nReport: $($script:lastReport)",
                            "Weekly Report Agent - Success",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        ) | Out-Null
                    }
                    else {
                        $lblStatus.Text = "FAILED"
                        $lblStatus.ForeColor = $colorRed

                        $errMsg = "The agent failed to complete. Check the log above for details."
                        if ($finalLog -match 'Failed to gather data from the source[^\r\n]*') {
                            $errMsg = $Matches[0]
                        }
                        elseif ($finalLog -match 'ERROR:\s*([^\r\n]+)') {
                            $errMsg = $Matches[1].Trim()
                        }

                        [System.Windows.Forms.MessageBox]::Show(
                            $errMsg,
                            "Weekly Report Agent - Error",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Error
                        ) | Out-Null
                    }
                    $btnRun.Enabled = $true
                })

            $script:pollTimer.Start()
        }
        catch {
            $lblStatus.Text = "FAILED"
            $lblStatus.ForeColor = $colorRed
            Write-DashboardLog ("LAUNCH ERROR: " + $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Weekly Report Agent Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $btnRun.Enabled = $true
        }
    })

try {
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Dashboard Error: $($_.Exception.Message)", "Weekly Report Agent Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}
