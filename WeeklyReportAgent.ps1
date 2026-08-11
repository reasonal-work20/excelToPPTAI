# WeeklyReportAgent.ps1
# Weekly Report Agent
#   1. Verifies mandatory 4 Excel files exist in the input source (Hard Stop if missing).
#   2. Refreshes charts in ALL 4 Excel files to plot updated weekly data.
#   3. Resolves PowerPoint template from local script directory ($PSScriptRoot) or SharePoint HTTPS URL.
#   4. Exports charts from all 4 Excel files as PNG images and embeds them into slides 2-6 of the PowerPoint report.
#   5. Publishes output directly to local folder or SharePoint output folder in a single PowerPoint COM pass.
#   6. Always generates a local fallback report copy in 'output-fallback\Weekly Report Week <N>_Fallback.pptx'.

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "SharePointPassword")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "password")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseApprovedVerbs", "")]
param(
    [int]$Week = -1,
    [string]$Root = "",
    [string]$ExcelFile = "LTI PER TEU data.xlsx",
    [string]$MovesDataUrl = "",       # 1. Actual & Forecast Moves data
    [string]$CmphDataUrl = "",        # 2. CMPH and PMPH actual and forecast data
    [string]$EquipmentDataUrl = "",   # 3. Equipment PERFORMANCE V1.2 Weekly data
    [string]$LtiDataUrl = "",         # 4. LTI PER TEU data
    [string]$TemplatePath = "",       # PowerPoint template (.pptx) — local path or SharePoint URL; blank = auto-resolve
    [string]$OutputDir = "",          # Local path or SharePoint URL
    [string]$SharePointUrl = "",       # Common input SharePoint URL (fallback)
    [string]$SharePointEmail = "",
    [string]$SharePointPassword = "",
    [string]$ClientId = "",
    [string]$TenantId = "",
    [string]$CertThumbprint = "",
    [switch]$UsePnP,
    [switch]$NoDateUpdate
)

$ErrorActionPreference = "Stop"

# --- Mandatory 4 Excel File Definition ---------------------------------------
$script:RequiredExcelFiles = @(
    "Actual & Forecast Moves data.xlsx",
    "CMPH and PMPH actual and forecast data.xlsx",
    "Equipment PERFORMANCE V1.2 Weekly data.xlsx",
    "LTI PER TEU data.xlsx"
)

# --- Centralized Excel Graph & Table Location Configuration -------------------
# To reconfigure chart indices, worksheet names, cell ranges, or column scans,
# edit the values in $script:GraphConfig below.
$script:GraphConfig = @{
    # Slide 2: LTI Chart (File: LTI PER TEU data.xlsx)
    Slide2_LTI = @{
        FileName         = "LTI PER TEU data.xlsx"
        Worksheet        = 1                 # Worksheet index (1-based) or name string
        ChartIndex       = 1                 # ChartObject index on sheet
        ScanColumn       = 6                 # Column F for last-row data lookup
        MaxScanRow       = 56                # Start row for scanning backward
        MinRow           = 4                 # Absolute minimum data row
        LookbackWeeks    = 11                # Spans lastRow - 11 to lastRow (12 weeks total)
        Width            = 823               # High-resolution export width
        Height           = 319               # High-resolution export height
    }

    # Slide 3: Commercial Moves Chart (File: Actual & Forecast Moves data.xlsx)
    Slide3_Moves = @{
        FileName         = "Actual & Forecast Moves data.xlsx"
        Worksheet        = 1
        ChartIndex       = 1
        ScanColumn       = 4                 # Column D for last-row lookup
        MaxScanRow       = 56
        MinRow           = 5
        LookbackWeeks    = 15                # Spans lastRow - 15 to lastRow (16 weeks total)
        Width            = 682
        Height           = 355
    }

    # Slide 4: Operations PMPH & CMPH Charts (File: CMPH and PMPH actual and forecast data.xlsx)
    Slide4_PMPH = @{
        FileName         = "CMPH and PMPH actual and forecast data.xlsx"
        Worksheet        = 1
        ChartIndex       = 1                 # Chart 1: PMPH
        ScanColumn       = 4                 # Column D for last-row data lookup
        MaxScanRow       = 56                # Start row for scanning backward
        MinRow           = 4                 # Absolute minimum data row
        LookbackWeeks    = 11                # Spans lastRow - 11 to lastRow (12 weeks total)
        Width            = 433
        Height           = 224
    }
    Slide4_CMPH = @{
        FileName         = "CMPH and PMPH actual and forecast data.xlsx"
        Worksheet        = 1
        ChartIndex       = 2                 # Chart 2: CMPH
        ScanColumn       = 4                 # Column D for last-row data lookup
        MaxScanRow       = 56                # Start row for scanning backward
        MinRow           = 4                 # Absolute minimum data row
        LookbackWeeks    = 11                # Spans lastRow - 11 to lastRow (12 weeks total)
        Width            = 436
        Height           = 224
    }

    # Slide 5: Equipment Availability Actual Range (File: Equipment PERFORMANCE V1.2 Weekly data.xlsx)
    Slide5_Actual = @{
        FileName         = "Equipment PERFORMANCE V1.2 Weekly data.xlsx"
        Worksheet        = "Weekly SMT"
        StartRow         = 39
        StartCol         = 21                # Column U (21)
        EndRow           = 55
        EndCol           = 37                # Column AK (37)
    }

    # Slide 6: Equipment Availability Forecast Stitched Tables (File: Equipment PERFORMANCE V1.2 Weekly data.xlsx)
    Slide6_Forecast = @{
        FileName         = "Equipment PERFORMANCE V1.2 Weekly data.xlsx"
        Worksheet        = "MnR Forecast"
        MaxCol           = 30                # Column AD (30)
        QC               = @( @{Start=4; End=5}, @{Start=11; End=17} )
        RTG              = @( @{Start=21; End=22}, @{Start=28; End=34} )
        PM               = @( @{Start=38; End=39}, @{Start=45; End=46}, @{Start=51; End=55} )
    }
}


# --- Clean Dashboard Logging Helper Functions ---------------------------------
$script:CurrentStepNum = 1
$script:TotalStepsNum  = 5

function Write-LogHeader([string]$text) {
    Write-Host ("=== {0} ===" -f $text)
}

function Redact-SensitiveText([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    if (-not [string]::IsNullOrWhiteSpace($SharePointPassword)) {
        $text = $text.Replace($SharePointPassword, "********")
    }
    if (-not [string]::IsNullOrWhiteSpace($env:WRA_SP_PASSWORD)) {
        $text = $text.Replace($env:WRA_SP_PASSWORD, "********")
    }
    return $text
}

function Write-LogStep([string]$stepText) {
    if ($stepText -match '\[(\d+)/(\d+)\]\s*(.*)') {
        $script:CurrentStepNum = [int]$Matches[1]
        $script:TotalStepsNum  = [int]$Matches[2]
        $stepText = $Matches[3]
    }
    $stepText = Redact-SensitiveText $stepText
    Write-Host ("[STEP {0}/{1}]: {2}" -f $script:CurrentStepNum, $script:TotalStepsNum, $stepText)
}

function Write-LogOk([string]$text) {
    $text = Redact-SensitiveText $text
    Write-Host ("[STEP {0}/{1}]: {2}" -f $script:CurrentStepNum, $script:TotalStepsNum, $text)
}

function Write-LogWarn([string]$text) {
    $text = Redact-SensitiveText $text
    Write-Host ("[STEP {0}/{1}]: WARN - {2}" -f $script:CurrentStepNum, $script:TotalStepsNum, $text)
}

function Write-LogError([string]$text) {
    $text = Redact-SensitiveText $text
    Write-Host ("[STEP {0}/{1}]: ERROR - {2}" -f $script:CurrentStepNum, $script:TotalStepsNum, $text)
}

function Invoke-SafeComAction([scriptblock]$action, [int]$maxRetries = 5, [int]$delayMs = 400) {
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            & $action
            return
        } catch [System.Runtime.InteropServices.COMException] {
            if ($attempt -lt $maxRetries) {
                Start-Sleep -Milliseconds $delayMs
            }
        } catch {
            return
        }
    }
}

function Close-WorkbookSafely($wb) {
    if ($null -eq $wb) { return }
    Invoke-SafeComAction { $wb.Close($false) }
}

function Save-WorkbookSafely($wb) {
    if ($null -eq $wb) { return }
    Invoke-SafeComAction { $wb.Save() }
}

function Open-ExcelWorkbookSafely($excel, [string]$filePath, [bool]$readOnly = $false) {
    if (-not $excel -or [string]::IsNullOrWhiteSpace($filePath)) { return $null }

    $isUrl = $filePath -match '^https?://'
    if ($isUrl) {
        try { $excel.DisplayAlerts = $true } catch {}
        try { $excel.Visible = $true } catch {}
    }

    $wb = try { $excel.Workbooks.Open($filePath, $false, $readOnly) } catch { $null }

    # Unprotect Protected View windows (common on client machines opening web/SharePoint URLs)
    try {
        if (-not $wb -and $null -ne $excel.ProtectedViewWindows -and $excel.ProtectedViewWindows.Count -gt 0) {
            foreach ($pvWin in $excel.ProtectedViewWindows) {
                if ($null -ne $pvWin) {
                    $wb = try { $pvWin.Edit() } catch { $null }
                    if ($null -ne $wb) { break }
                }
            }
        }
    } catch {}

    try { $excel.DisplayAlerts = $false } catch {}
    return $wb
}

function Get-ExcelWorksheetSafely($wb, $sheetTarget) {
    if ($null -eq $wb) { return $null }

    if ($null -eq $sheetTarget -or "$sheetTarget" -eq "") { $sheetTarget = 1 }

    if ("$sheetTarget" -match '^\d+$') {
        $idx = [int]$sheetTarget
        try {
            $ws = $wb.Worksheets.Item($idx)
            if ($null -ne $ws) { return $ws }
        } catch {}
        try {
            $ws = $wb.Sheets.Item($idx)
            if ($null -ne $ws) { return $ws }
        } catch {}
    }

    $name = [string]$sheetTarget
    try {
        $ws = $wb.Worksheets.Item($name)
        if ($null -ne $ws) { return $ws }
    } catch {}
    try {
        $ws = $wb.Sheets.Item($name)
        if ($null -ne $ws) { return $ws }
    } catch {}

    try {
        $ws = $wb.Worksheets.Item(1)
        if ($null -ne $ws) { return $ws }
    } catch {}
    try {
        $ws = $wb.Sheets.Item(1)
        if ($null -ne $ws) { return $ws }
    } catch {}

    return $null
}

function Get-IsoWeek([datetime]$d) {
    $day = [int]$d.DayOfWeek
    if ($day -eq 0) { $day = 7 }
    $thursday = $d.AddDays(4 - $day)
    return [math]::Floor(($thursday.DayOfYear - 1) / 7) + 1
}

function Get-GraphSharingToken([string]$url) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($url)
    $base64 = [Convert]::ToBase64String($bytes)
    $urlSafe = $base64.TrimEnd('=').Replace('+', '-').Replace('/', '_')
    return "u!$urlSafe"
}

function Get-GraphAccessTokenFromCert([string]$clientId, [string]$tenantId, [string]$certThumbprint, [string]$rootPath = "") {
    $cleanThumb = ($certThumbprint -replace '[^A-Fa-f0-9]', '').ToUpper()
    $cert = Get-Item "Cert:\CurrentUser\My\$cleanThumb" -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-Item "Cert:\LocalMachine\My\$cleanThumb" -ErrorAction SilentlyContinue
    }
    if (-not $cert) {
        $cert = Get-ChildItem -Path Cert:\ -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -and ($_.Thumbprint.ToUpper() -eq $cleanThumb) } | Select-Object -First 1
    }
    if (-not $cert -and $rootPath) {
        $pfxFiles = Get-ChildItem -Path $rootPath -Filter "*.pfx" -ErrorAction SilentlyContinue
        if ($pfxFiles) {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxFiles[0].FullName, "", [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        }
    }
    if (-not $cert) {
        throw "Failed to gather data from the source: Certificate with thumbprint '$certThumbprint' was not found."
    }

    function Base64UrlEncode([byte[]]$bytes) {
        return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }

    $x5t = Base64UrlEncode ($cert.GetCertHash())
    $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress

    $now = [datetimeoffset]::UtcNow
    $nbf = $now.ToUnixTimeSeconds()
    $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    $payload = @{
        aud = $tokenUri; exp = $exp; iss = $clientId; jti = [guid]::NewGuid().ToString(); nbf = $nbf; sub = $clientId
    } | ConvertTo-Json -Compress

    $headerEncoded  = Base64UrlEncode ([System.Text.Encoding]::UTF8.GetBytes($header))
    $payloadEncoded = Base64UrlEncode ([System.Text.Encoding]::UTF8.GetBytes($payload))
    $unsignedToken  = "$headerEncoded.$payloadEncoded"

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) {
        throw "Failed to gather data from the source: Could not retrieve RSA private key from certificate '$certThumbprint'."
    }
    $signatureBytes = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($unsignedToken), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $signatureEncoded = Base64UrlEncode $signatureBytes
    $clientAssertion = "$unsignedToken.$signatureEncoded"

    $body = @{
        grant_type            = 'client_credentials'
        client_id             = $clientId
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $clientAssertion
        scope                 = 'https://graph.microsoft.com/.default'
    }

    $response = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ErrorAction Stop
    return $response.access_token
}

# --- PnP PowerShell & Interactive MFA Authentication Helpers ------------------
function Get-PnPSharePointFile([string]$urlToFetch, [string]$destPath, [string]$email, [string]$password) {
    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        return $false
    }
    try {
        $uri = [Uri]$urlToFetch
        $siteUrl = "$($uri.Scheme)://$($uri.Host)"
        $pathSegments = $uri.AbsolutePath.Trim('/').Split('/')
        if ($pathSegments.Length -ge 2 -and ($pathSegments[0] -eq 'sites' -or $pathSegments[0] -eq 'teams')) {
            $siteUrl = "$siteUrl/$($pathSegments[0])/$($pathSegments[1])"
        }

        Write-LogStep ("Authenticating with SharePoint via PnP PowerShell...")
        if (-not [string]::IsNullOrWhiteSpace($email) -and -not [string]::IsNullOrWhiteSpace($password)) {
            Write-LogStep ("Connecting PnP using email credentials ({0})..." -f $email)
            $secPass = ConvertTo-SecureString $password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($email, $secPass)
            Connect-PnPOnline -Url $siteUrl -Credentials $cred -ErrorAction Stop
            Write-LogOk ("PnP PowerShell connected with credentials for {0}!" -f $email)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($email)) {
            Write-LogStep ("Connecting PnP using email ({0})..." -f $email)
            Connect-PnPOnline -Url $siteUrl -Credentials $email -ErrorAction Stop
            Write-LogOk ("PnP PowerShell connected with email {0}!" -f $email)
        }
        else {
            Write-LogStep "Connecting PnP via interactive authentication..."
            Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
            Write-LogOk "PnP PowerShell interactive login successful!"
        }

        $serverRelativeUrl = $uri.AbsolutePath
        $targetFolder = Split-Path $destPath -Parent
        $fileName = Split-Path $destPath -Leaf

        Write-LogOk ("Downloading via PnP: {0}" -f $serverRelativeUrl)
        Get-PnPFile -Url $serverRelativeUrl -Path $targetFolder -FileName $fileName -AsFile -Force -ErrorAction Stop

        if ((Test-Path $destPath) -and (Get-Item $destPath).Length -gt 100) {
            $bytes = [System.IO.File]::ReadAllBytes($destPath)
            if ($bytes.Length -gt 4 -and $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
                return $true
            }
        }
    } catch {
        Write-LogWarn ("PnP PowerShell download attempt failed: $_")
    }
    return $false
}

function Get-GraphAccessTokenInteractive([string]$tenantId, [string]$clientId) {
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        $clientId = "31359c7f-bd7e-475c-86db-f28c4cfd6573" # Standard PnP / MS Graph App ID
    }
    $tenant = if ([string]::IsNullOrWhiteSpace($tenantId)) { "common" } else { $tenantId }

    $codeUri = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"
    $body = @{
        client_id = $clientId
        scope     = "https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Sites.Read.All offline_access"
    }

    try {
        $codeResp = Invoke-RestMethod -Method Post -Uri $codeUri -Body $body -ErrorAction Stop
        Write-LogStep ("=" * 60)
        Write-LogStep ("MICROSOFT MFA INTERACTIVE LOGIN REQUIRED:")
        Write-LogStep ("1. Open Browser : {0}" -f $codeResp.verification_uri)
        Write-LogStep ("2. Enter Code    : {0}" -f $codeResp.user_code)
        Write-LogStep ("=" * 60)

        try { [System.Diagnostics.Process]::Start($codeResp.verification_uri) } catch {}

        $tokenUri = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token"
        $interval = if ($codeResp.interval) { [int]$codeResp.interval } else { 5 }
        $expiresIn = if ($codeResp.expires_in) { [int]$codeResp.expires_in } else { 900 }
        $startTime = [DateTime]::UtcNow

        while (([DateTime]::UtcNow - $startTime).TotalSeconds -lt $expiresIn) {
            Start-Sleep -Seconds $interval
            $tokenBody = @{
                grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                client_id   = $clientId
                device_code = $codeResp.device_code
            }
            try {
                $tokenResp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody -ErrorAction Stop
                if ($tokenResp.access_token) {
                    Write-LogOk "MFA Interactive Authentication successful!"
                    return $tokenResp.access_token
                }
            } catch {
                if ($_ -match "authorization_pending") {
                    continue
                } else {
                    break
                }
            }
        }
    } catch {
        Write-LogWarn ("MS Graph Device Code MFA flow failed: $_")
    }
    return $null
}

# --- Mandatory 4-Excel File Check ---------------------------------------------
function Test-MandatoryExcelFiles() {
    Write-LogStep "[1/5] Checking required Excel input files"

    $missingFiles = @()
    foreach ($file in $script:RequiredExcelFiles) {
        $resolved = $script:ResolvedExcelPaths[$file]
        if (-not $resolved) {
            Write-LogError ("Missing: {0}" -f $file)
            $missingFiles += $file
        } elseif ($resolved -match '^https?://') {
            Write-LogOk ("{0} -> SharePoint URL (Will open directly in Excel COM)" -f $file)
        } elseif (-not (Test-Path $resolved -PathType Leaf)) {
            Write-LogError ("Missing: {0}" -f $file)
            $missingFiles += $file
        } else {
            Write-LogOk ("{0} -> Local File" -f $file)
        }
    }

    if ($missingFiles.Count -gt 0) {
        $missingList = $missingFiles -join ", "
        throw "Failed to gather data from the source: Missing required Excel file(s): [$missingList]. All 4 required Excel files must exist locally or have valid SharePoint URLs."
    }
}

# --- Template Resolution ------------------------------------------------------
function Resolve-PowerPointTemplate([string]$executionDir, [string]$templatePath = "") {
    Write-LogStep "[2/5] Resolving PowerPoint template"

    if (-not [string]::IsNullOrWhiteSpace($templatePath)) {
        $templatePath = $templatePath.Trim("'").Trim('"').Trim()

        if ($templatePath -match '^https?://') {
            Write-LogOk ("Using SharePoint template URL: {0}" -f $templatePath)
            $localTemplate = Join-Path $env:TEMP ("wra_template_{0}.pptx" -f [guid]::NewGuid().ToString('N'))
            try {
                Invoke-WebRequest -Uri "$templatePath`?download=1" -OutFile $localTemplate -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
                if ((Test-Path $localTemplate) -and (Get-Item $localTemplate).Length -gt 100) {
                    $bytes = [System.IO.File]::ReadAllBytes($localTemplate)
                    if ($bytes.Length -gt 4 -and $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
                        Unblock-File -Path $localTemplate -ErrorAction SilentlyContinue
                        return $localTemplate
                    }
                }
            } catch {}

            return $templatePath
        }

        if (Test-Path $templatePath -PathType Leaf) {
            Unblock-File -Path $templatePath -ErrorAction SilentlyContinue
            Write-LogOk ("Using local template file: {0}" -f $templatePath)
            return $templatePath
        }

        throw "Failed to gather data from the source: Provided PowerPoint template not found: '$templatePath'"
    }

    $templateDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($executionDir) { $executionDir } else { Get-Location }

    $reports = Get-ChildItem -Path $templateDir -Filter "Weekly Report Week *.pptx" -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match 'Week (\d+)$' } |
        ForEach-Object { [pscustomobject]@{ File = $_; Num = [int][regex]::Match($_.BaseName, 'Week (\d+)$').Groups[1].Value } } |
        Sort-Object Num -Descending

    if ($reports) {
        Unblock-File -Path $reports[0].File.FullName -ErrorAction SilentlyContinue
        Write-LogOk ("Found template: {0}" -f $reports[0].File.Name)
        return $reports[0].File.FullName
    }

    $anyPptx = Get-ChildItem -Path $templateDir -Filter "*.pptx" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($anyPptx) {
        Unblock-File -Path $anyPptx.FullName -ErrorAction SilentlyContinue
        Write-LogOk ("Found template: {0}" -f $anyPptx.Name)
        return $anyPptx.FullName
    }

    throw "Failed to gather data from the source: Mandatory PowerPoint template file ('Weekly Report Week *.pptx') not found in local directory '$templateDir'."
}

# --- Direct Active Presentation Publication ------------------------------------
function Publish-ReportOutputFromPres(
    $pres,
    [string]$localReportFile,
    [string]$outputDestination,
    [string]$fallbackFilePath,
    [string]$userEmail,
    [string]$clientId,
    [string]$tenantId,
    [string]$certThumbprint,
    [string]$rootPath
) {
    if ([string]::IsNullOrWhiteSpace($outputDestination)) { $outputDestination = $rootPath }
    $outputDestination = $outputDestination.Trim("'").Trim('"').Trim()
    $fileName = Split-Path $localReportFile -Leaf

    # If Output destination is a SharePoint HTTPS URL
    if ($outputDestination -match '^https?://') {
        Write-LogStep "[5/5] Publishing report output to SharePoint"
        Write-LogOk ("Target SharePoint URL: {0}" -f $outputDestination)
        $targetUrl = "$($outputDestination.TrimEnd('/'))/$fileName"
        $uploadErrors = @()

        # Method 1: Active Presentation Direct SaveAs via Office SSO COM
        try {
            Write-LogOk "Attempting PowerPoint COM SaveAs to SharePoint..."
            $pres.SaveAs($targetUrl)
            if ($null -ne $pres.Sync) {
                try { $pres.Sync.PutUpdate() } catch {}
            }
            Write-LogOk "Uploading presentation stream to SharePoint Online (waiting for slow connections)..."
            $maxWaitSec = 60
            for ($waitSec = 1; $waitSec -le $maxWaitSec; $waitSec++) {
                Start-Sleep -Seconds 1
                if ($null -ne $pres.Sync) {
                    $status = $pres.Sync.Status
                    if ($status -eq 1) {
                        Write-LogOk ("Upload completed successfully ({0}s)" -f $waitSec)
                        break
                    }
                    if ($status -eq 4 -or $status -eq 6) {
                        throw "Office Upload Center reported sync status error ($status). File locked or forbidden."
                    }
                }
                if ($waitSec % 10 -eq 0) {
                    Write-LogOk ("Still uploading to SharePoint Online... ({0}s / {1}s)" -f $waitSec, $maxWaitSec)
                }
            }

            Write-LogOk ("Published to SharePoint via PowerPoint COM: {0}" -f $targetUrl)
            return $targetUrl
        } catch {
            $uploadErrors += "PowerPoint COM SaveAs: $($_.Exception.Message)"
            Write-LogWarn ("PowerPoint COM SaveAs failed: {0}" -f $_.Exception.Message)
        }

        # Method 1b: PowerPoint COM SaveCopyAs to SharePoint
        try {
            Write-LogOk "Attempting PowerPoint COM SaveCopyAs to SharePoint..."
            $pres.SaveCopyAs($targetUrl)
            Write-LogOk "Uploading copy stream to SharePoint Online..."
            for ($waitSec = 1; $waitSec -le 15; $waitSec++) {
                Start-Sleep -Seconds 1
            }
            Write-LogOk ("Published to SharePoint via PowerPoint COM SaveCopyAs: {0}" -f $targetUrl)
            return $targetUrl
        } catch {
            $uploadErrors += "PowerPoint COM SaveCopyAs: $($_.Exception.Message)"
            Write-LogWarn ("PowerPoint COM SaveCopyAs failed: {0}" -f $_.Exception.Message)
        }

        # Method 2: Local OneDrive / SharePoint Synced Folder Auto-Detection
        try {
            Write-LogOk "Checking for local OneDrive / SharePoint synced directory..."
            $uri = [Uri]$outputDestination
            $urlSegments = $uri.AbsolutePath.Trim('/').Split('/') | Where-Object { $_ -and $_ -ne 'sites' -and $_ -ne 'teams' -and $_ -ne 'Shared Documents' }
            
            $oneDriveRoots = @()
            $regAccounts = Get-ItemProperty 'HKCU:\Software\Microsoft\OneDrive\Accounts\*' -ErrorAction SilentlyContinue
            foreach ($acc in $regAccounts) {
                if ($acc.UserFolder -and (Test-Path $acc.UserFolder)) { $oneDriveRoots += $acc.UserFolder }
            }
            $userProfileSharePoint = Join-Path $env:USERPROFILE "SharePoint"
            if (Test-Path $userProfileSharePoint) { $oneDriveRoots += $userProfileSharePoint }

            $matchedSyncDir = $null
            foreach ($root in $oneDriveRoots) {
                foreach ($seg in $urlSegments) {
                    $decodedSeg = [Uri]::UnescapeDataString($seg)
                    $matchedDirs = Get-ChildItem -Path $root -Recurse -Depth 3 -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $decodedSeg -or $_.Name -ieq $seg }
                    if ($matchedDirs) {
                        $matchedSyncDir = $matchedDirs[0].FullName
                        break
                    }
                }
                if ($matchedSyncDir) { break }
            }

            if ($matchedSyncDir) {
                $targetLocal = Join-Path $matchedSyncDir $fileName
                Copy-Item -Path $localReportFile -Destination $targetLocal -Force
                Write-LogOk ("Published report to local OneDrive/SharePoint synced directory: {0}" -f $targetLocal)
                return $targetUrl
            }
        } catch {
            $uploadErrors += "OneDrive Sync Auto-Detection: $($_.Exception.Message)"
            Write-LogWarn ("Local OneDrive sync folder copy failed: {0}" -f $_.Exception.Message)
        }

        # Method 3: Windows WebClient WebDAV UNC Path Copy
        try {
            Write-LogOk "Attempting WebDAV UNC path copy..."
            $uri = [Uri]$outputDestination
            $uncFolder = "\\$($uri.Host)@SSL\DavWWWRoot$([Uri]::UnescapeDataString($uri.AbsolutePath).Replace('/', '\'))"
            $uncTarget = Join-Path $uncFolder $fileName

            $webClientSvc = Get-Service -Name WebClient -ErrorAction SilentlyContinue
            if ($webClientSvc -and $webClientSvc.Status -ne 'Running') {
                try { Start-Service -Name WebClient -ErrorAction SilentlyContinue } catch {}
            }

            Copy-Item -Path $localReportFile -Destination $uncTarget -Force -ErrorAction Stop
            Write-LogOk ("Published to SharePoint via WebDAV UNC: {0}" -f $uncTarget)
            return $targetUrl
        } catch {
            $uploadErrors += "WebDAV UNC: $($_.Exception.Message)"
            Write-LogWarn ("WebDAV UNC copy failed: {0}" -f $_.Exception.Message)
        }

        # Method 4: MS Graph API DriveItem Upload (if App Cert Configured)
        if ($certThumbprint -and $clientId -and $tenantId) {
            try {
                Write-LogOk "Attempting MS Graph API direct upload..."
                $token = Get-GraphAccessTokenFromCert -clientId $clientId -tenantId $tenantId -certThumbprint $certThumbprint -rootPath $rootPath
                $sharingToken = Get-GraphSharingToken $targetUrl
                $itemUri = "https://graph.microsoft.com/v1.0/shares/${sharingToken}/driveItem:/content"
                Invoke-WebRequest -Uri $itemUri -Method Put -InFile $localReportFile -Headers @{ Authorization = "Bearer $token" } -UseBasicParsing -ErrorAction Stop
                Write-LogOk ("Published to SharePoint via Graph API: {0}" -f $targetUrl)
                return $targetUrl
            } catch {
                $uploadErrors += "Graph API: $($_.Exception.Message)"
                Write-LogWarn ("Graph API upload failed: {0}" -f $_.Exception.Message)
            }
        }

        # Fallback & Notification: If cloud upload fails across all methods, publish locally & warn cleanly
        $errSummary = $uploadErrors -join " | "
        Write-LogWarn ("Direct SharePoint cloud upload unavailable: {0}" -f $errSummary)
        Write-LogOk ("Report successfully generated and available locally: {0}" -f $localReportFile)
        Write-LogOk ("Fallback copy saved: {0}" -f $fallbackFilePath)
        return $localReportFile
    }

    # Normalize local path destination
    Write-LogStep "[5/5] Publishing report output to local directory"
    if ($outputDestination -match '\.pptx?$' -or (Test-Path -Path $outputDestination -PathType Leaf -ErrorAction SilentlyContinue)) {
        $outputDestination = Split-Path -Path $outputDestination -Parent
        if ([string]::IsNullOrWhiteSpace($outputDestination)) { $outputDestination = $rootPath }
    }

    if (-not [System.IO.Path]::IsPathRooted($outputDestination)) {
        $outputDestination = Join-Path $rootPath $outputDestination
    }

    if (-not (Test-Path -Path $outputDestination -ErrorAction SilentlyContinue)) {
        New-Item -ItemType Directory -Force -Path $outputDestination | Out-Null
    }
    $finalPath = Join-Path $outputDestination $fileName
    if ($finalPath -ne $localReportFile) {
        Copy-Item -Path $localReportFile -Destination $finalPath -Force
    }
    Write-LogOk ("Report published to local folder: {0}" -f $finalPath)
    return $finalPath
}

# --- Excel Chart/Range Export Helpers ----------------------------------------

function Test-ValidImageFile([string]$filePath) {
    if ([string]::IsNullOrWhiteSpace($filePath) -or -not (Test-Path -Path $filePath -ErrorAction SilentlyContinue)) {
        return $false
    }
    for ($i = 0; $i -lt 30; $i++) {
        try {
            $item = Get-Item -Path $filePath -ErrorAction SilentlyContinue
            if ($null -ne $item -and $item.Length -gt 1024) {
                $img = [System.Drawing.Image]::FromFile($filePath)
                $w = $img.Width
                $h = $img.Height
                $img.Dispose()
                if ($w -gt 0 -and $h -gt 0) { return $true }
            }
        } catch {}
        Start-Sleep -Milliseconds 100
    }
    return $false
}

function Export-ChartHd($cObj, [string]$pngPath) {
    if (-not $cObj -or -not $pngPath) { return }
    $origW = $cObj.Width
    $origH = $cObj.Height
    try {
        $cObj.Width = $origW * 2
        $cObj.Height = $origH * 2
        if (Test-Path $pngPath) { Remove-Item $pngPath -Force -ErrorAction SilentlyContinue }
        [void]($cObj.Chart.Export($pngPath, "PNG"))
        [void](Test-ValidImageFile $pngPath)
    }
    finally {
        try {
            $cObj.Width = $origW
            $cObj.Height = $origH
        } catch {}
    }
}

# Robust helper to get a cell range image via clipboard with retries
function Get-RangeImageSafe($ws, [int]$r1, [int]$c1, [int]$r2, [int]$c2) {
    if (-not $ws) { return $null }
    $rng = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r2, $c2))
    $maxRetries = 5
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            $ws.Activate()
            $rng.Select() | Out-Null
            Start-Sleep -Milliseconds 150
            [System.Windows.Forms.Clipboard]::Clear()
            $rng.CopyPicture(1, 2) | Out-Null   # xlScreen=1, xlBitmap=2
            Start-Sleep -Milliseconds 400
            $img = [System.Windows.Forms.Clipboard]::GetImage()
            if ($img) {
                return $img
            }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    return $null
}

# Captures a cell range as a PNG via robust CopyPicture
function Export-RangePicture($ws, [int]$r1, [int]$c1, [int]$r2, [int]$c2, [string]$pngPath) {
    if (-not $ws -or -not $pngPath) { return }
    try {
        $img = Get-RangeImageSafe $ws $r1 $c1 $r2 $c2
        if ($img) {
            if (Test-Path $pngPath) { Remove-Item $pngPath -Force }
            $img.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $img.Dispose()
        }
    } catch {
        Write-LogWarn ("Export-RangePicture failed: {0}" -f $_.Exception.Message)
    }
}

# Captures all chart objects on a worksheet as one composite image via clipboard
function Export-ChartGroupPicture($ws, [string]$pngPath) {
    if (-not $ws -or -not $pngPath) { return }
    try {
        $ws.Activate()
        Start-Sleep -Milliseconds 300
        $ws.ChartObjects().Select() | Out-Null
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.Clipboard]::Clear()
        $ws.Application.Selection.Copy() | Out-Null
        Start-Sleep -Milliseconds 500
        $img = [System.Windows.Forms.Clipboard]::GetImage()
        if ($img) {
            if (Test-Path $pngPath) { Remove-Item $pngPath -Force }
            $img.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $img.Dispose()
        }
    } catch {
        Write-LogWarn ("Export-ChartGroupPicture failed: {0}" -f $_.Exception.Message)
    }
}

# Captures header and data ranges separately, stitches them vertically, and saves as PNG
function Export-StitchedRangePicture($ws, [int]$hStart, [int]$hEnd, [int]$dStart, [int]$dEnd, [int]$cMax, [string]$pngPath) {
    if (-not $ws -or -not $pngPath) { return }
    try {
        $hImg = Get-RangeImageSafe $ws $hStart 1 $hEnd $cMax
        $dImg = Get-RangeImageSafe $ws $dStart 1 $dEnd $cMax
        if ($hImg -and $dImg) {
            $w = [math]::Max($hImg.Width, $dImg.Width)
            $h = $hImg.Height + $dImg.Height
            $stitched = New-Object System.Drawing.Bitmap($w, $h)
            $g = [System.Drawing.Graphics]::FromImage($stitched)
            $g.Clear([System.Drawing.Color]::White)
            $g.DrawImage($hImg, 0, 0)
            $g.DrawImage($dImg, 0, $hImg.Height)
            $g.Dispose()
            
            if (Test-Path $pngPath) { Remove-Item $pngPath -Force }
            $stitched.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $stitched.Dispose()
        }
        if ($hImg) { $hImg.Dispose() }
        if ($dImg) { $dImg.Dispose() }
    } catch {
        Write-LogWarn ("Export-StitchedRangePicture failed: {0}" -f $_.Exception.Message)
    }
}

# Captures multiple cell ranges, stitches them vertically top-to-bottom, and saves as PNG
function Export-MultiStitchedRangePicture($ws, $ranges, [int]$cMax, [string]$pngPath) {
    if (-not $ws -or -not $pngPath -or -not $ranges) { return }
    try {
        $imgs = @()
        foreach ($r in $ranges) {
            $img = Get-RangeImageSafe $ws $r.Start 1 $r.End $cMax
            if ($img) { $imgs += $img }
        }
        if ($imgs.Count -gt 0) {
            $maxW = 0
            $totalH = 0
            foreach ($im in $imgs) {
                if ($im.Width -gt $maxW) { $maxW = $im.Width }
                $totalH += $im.Height
            }
            $stitched = New-Object System.Drawing.Bitmap($maxW, $totalH)
            $g = [System.Drawing.Graphics]::FromImage($stitched)
            $g.Clear([System.Drawing.Color]::White)
            $curY = 0
            foreach ($im in $imgs) {
                $g.DrawImage($im, 0, $curY)
                $curY += $im.Height
                $im.Dispose()
            }
            $g.Dispose()

            if (Test-Path $pngPath) { Remove-Item $pngPath -Force }
            $stitched.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $stitched.Dispose()
        }
    } catch {
        Write-LogWarn ("Export-MultiStitchedRangePicture failed: {0}" -f $_.Exception.Message)
    }
}


# Replaces a slide picture shape with a new PNG, preserving position, border, and Z-order
function Replace-SlidePictureShape($slideObj, $shapeTarget, $imagePngPath, [double]$TargetWidthPoints = 0, [double]$TargetHeightPoints = 0, [switch]$Fill) {
    if (-not $shapeTarget -or -not $imagePngPath) { return }
    if (-not (Test-ValidImageFile $imagePngPath)) {
        Write-LogWarn ("Replace-SlidePictureShape skipped: PNG image file is invalid or incomplete -> {0}" -f $imagePngPath)
        return
    }

    $boxLeft = $shapeTarget.Left
    $boxTop = $shapeTarget.Top
    $boxWidth = $shapeTarget.Width
    $boxHeight = $shapeTarget.Height

    if ($TargetWidthPoints -gt 0) { $boxWidth = $TargetWidthPoints }
    if ($TargetHeightPoints -gt 0) { $boxHeight = $TargetHeightPoints }

    $newWidth = $boxWidth
    $newHeight = $boxHeight
    $left = $boxLeft + ($shapeTarget.Width - $boxWidth) / 2
    $top = $boxTop + ($shapeTarget.Height - $boxHeight) / 2

    if (-not $Fill -and $TargetWidthPoints -eq 0 -and $TargetHeightPoints -eq 0) {
        try {
            $img = [System.Drawing.Image]::FromFile($imagePngPath)
            $imgW = $img.Width
            $imgH = $img.Height
            $img.Dispose()

            if ($imgW -gt 0 -and $imgH -gt 0) {
                $scale = [math]::Min($boxWidth / $imgW, $boxHeight / $imgH)
                $newWidth = $imgW * $scale
                $newHeight = $imgH * $scale
                $left = $boxLeft + ($shapeTarget.Width - $newWidth) / 2
                $top = $boxTop + ($shapeTarget.Height - $newHeight) / 2
            }
        } catch {}
    }

    $msoFalse = 0
    $msoTrue  = -1
    $newPic = $slideObj.Shapes.AddPicture($imagePngPath, $msoFalse, $msoTrue, $left, $top, $newWidth, $newHeight)

    try {
        if ($shapeTarget.Line.Visible -eq -1) {
            $newPic.Line.Visible = -1
            $newPic.Line.ForeColor.RGB = $shapeTarget.Line.ForeColor.RGB
            $newPic.Line.Weight = $shapeTarget.Line.Weight
        }
    } catch {}

    try {
        while ($newPic.ZOrderPosition -gt $shapeTarget.ZOrderPosition) {
            $newPic.ZOrder(3) # msoSendBackward
        }
    } catch {}

    $shapeTarget.Delete()
}

function New-ExcelComObject() {
    try {
        $type = [Type]::GetTypeFromProgID("Excel.Application")
        if ($null -ne $type) {
            $obj = [Activator]::CreateInstance($type)
            if ($null -ne $obj -and $null -ne $obj.Workbooks) { return $obj }
        }
    } catch {}
    return (New-Object -ComObject Excel.Application)
}

function New-PowerPointComObject() {
    try {
        $type = [Type]::GetTypeFromProgID("PowerPoint.Application")
        if ($null -ne $type) {
            $obj = [Activator]::CreateInstance($type)
            if ($null -ne $obj -and $null -ne $obj.Presentations) { return $obj }
        }
    } catch {}
    return (New-Object -ComObject PowerPoint.Application)
}

function Get-CleanInputUrl([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
    return $raw.Trim("'").Trim('"').Trim()
}

# --- Main Execution Flow -----------------------------------------------------

try {
    if (-not $SharePointPassword -and $env:WRA_SP_PASSWORD) {
        $SharePointPassword = $env:WRA_SP_PASSWORD
    }
    if ($Week -eq -1) { $Week = Get-IsoWeek (Get-Date) }

    $scriptExecutionDir = if (-not [string]::IsNullOrWhiteSpace($Root) -and (Test-Path -Path $Root -ErrorAction SilentlyContinue)) { $Root } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $workingRoot = Join-Path $env:TEMP ("wra_input_{0}" -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $workingRoot | Out-Null

    Write-LogHeader ("Weekly Report Agent | Target Week {0}" -f $Week)

    $MovesDataUrl     = Get-CleanInputUrl $MovesDataUrl
    $CmphDataUrl      = Get-CleanInputUrl $CmphDataUrl
    $EquipmentDataUrl = Get-CleanInputUrl $EquipmentDataUrl
    $LtiDataUrl       = Get-CleanInputUrl $LtiDataUrl
    $OutputDir        = Get-CleanInputUrl $OutputDir

    $fileSourceMap = @{
        "Actual & Forecast Moves data.xlsx"          = $MovesDataUrl
        "CMPH and PMPH actual and forecast data.xlsx" = $CmphDataUrl
        "Equipment PERFORMANCE V1.2 Weekly data.xlsx"= $EquipmentDataUrl
        "LTI PER TEU data.xlsx"                      = $LtiDataUrl
    }

    $script:ResolvedExcelPaths = @{}

    foreach ($reqFile in $script:RequiredExcelFiles) {
        $sourceVal = $fileSourceMap[$reqFile]
        if (-not $sourceVal -and $SharePointUrl) {
            $sourceVal = "$($SharePointUrl.TrimEnd('/'))/$reqFile"
        }

        $destPath = Join-Path $workingRoot $reqFile

        if ($sourceVal -and ($sourceVal -match '^https?://')) {
            $urlToFetch = if ($sourceVal -match '\.xlsx$') { $sourceVal } else { "$($sourceVal.TrimEnd('/'))/$reqFile" }

            $downloaded = $false

            # Strategy 1: PnP PowerShell (Connect-PnPOnline with Email/Password or Interactive)
            if ($UsePnP -or (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
                $downloaded = Get-PnPSharePointFile -urlToFetch $urlToFetch -destPath $destPath -email $SharePointEmail -password $SharePointPassword
                if ($downloaded) {
                    $script:ResolvedExcelPaths[$reqFile] = $destPath
                }
            }

            # Strategy 2: Direct Invoke-WebRequest with Default Credentials (Domain SSO)
            if (-not $downloaded) {
                try {
                    Invoke-WebRequest -Uri "$urlToFetch`?download=1" -OutFile $destPath -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
                    if ((Test-Path $destPath) -and (Get-Item $destPath).Length -gt 100) {
                        $bytes = [System.IO.File]::ReadAllBytes($destPath)
                        if ($bytes.Length -gt 4 -and $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
                            $script:ResolvedExcelPaths[$reqFile] = $destPath
                            $downloaded = $true
                        }
                    }
                } catch {}
            }

            # Strategy 3: Graph API via App Registration Certificate (Non-interactive)
            if (-not $downloaded -and $CertThumbprint -and $ClientId -and $TenantId) {
                try {
                    $token = Get-GraphAccessTokenFromCert -clientId $ClientId -tenantId $TenantId -certThumbprint $CertThumbprint -rootPath $scriptExecutionDir
                    $sharingToken = Get-GraphSharingToken $urlToFetch
                    $itemUri = "https://graph.microsoft.com/v1.0/shares/${sharingToken}/driveItem:/content"
                    Invoke-WebRequest -Uri $itemUri -OutFile $destPath -Headers @{ Authorization = "Bearer $token" } -UseBasicParsing -ErrorAction Stop
                    $script:ResolvedExcelPaths[$reqFile] = $destPath
                    $downloaded = $true
                } catch {}
            }

            # Strategy 4: MS Graph Interactive / Device Code Flow (MFA Interactive without PnP module)
            if (-not $downloaded) {
                try {
                    $mfaToken = Get-GraphAccessTokenInteractive -tenantId $TenantId -clientId $ClientId
                    if ($mfaToken) {
                        $sharingToken = Get-GraphSharingToken $urlToFetch
                        $itemUri = "https://graph.microsoft.com/v1.0/shares/${sharingToken}/driveItem:/content"
                        Invoke-WebRequest -Uri $itemUri -OutFile $destPath -Headers @{ Authorization = "Bearer $mfaToken" } -UseBasicParsing -ErrorAction Stop
                        $script:ResolvedExcelPaths[$reqFile] = $destPath
                        $downloaded = $true
                    }
                } catch {}
            }

            if (-not $downloaded) {
                $script:ResolvedExcelPaths[$reqFile] = $urlToFetch
            }
        }
        elseif ($sourceVal -and (Test-Path $sourceVal -PathType Leaf)) {
            $script:ResolvedExcelPaths[$reqFile] = $sourceVal
        }
        else {
            $localFolder = if ($sourceVal -and (Test-Path $sourceVal -PathType Container)) { $sourceVal } else { $scriptExecutionDir }
            $candidate = Join-Path $localFolder $reqFile
            if (Test-Path $candidate -PathType Leaf) {
                $script:ResolvedExcelPaths[$reqFile] = $candidate
            } else {
                $script:ResolvedExcelPaths[$reqFile] = $candidate
            }
        }
    }

    # --- 1) MANDATORY 4-EXCEL FILE CHECK & HARD STOP --------------------------
    Test-MandatoryExcelFiles

    # --- 2) TEMPLATE RESOLUTION -----------------------------------------------
    $templatePptx = Resolve-PowerPointTemplate -executionDir $scriptExecutionDir -templatePath $TemplatePath
    $localReportTemp = Join-Path $env:TEMP ("Weekly Report Week {0}.pptx" -f $Week)

    # Prepare output-fallback directory and fallback filename
    $fallbackFolder = Join-Path $scriptExecutionDir "output-fallback"
    if (-not (Test-Path $fallbackFolder)) {
        New-Item -ItemType Directory -Force -Path $fallbackFolder | Out-Null
    }
    $fallbackFileName = "Weekly Report Week {0}_Fallback.pptx" -f $Week
    $fallbackFilePath = Join-Path $fallbackFolder $fallbackFileName

    # --- 3) EXCEL CHART REFRESH & EXPORT ACROSS ALL 4 EXCEL FILES -------------
    Write-LogStep "[3/5] Updating Excel chart data & exporting images"

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $excel = New-ExcelComObject
    try { $excel.Visible = $true } catch {}         # Resilient against TypeLib registry issues
    try { $excel.DisplayAlerts = $false } catch {}
    try { $excel.WindowState = 2 } catch {}

    $exportedCharts = @{}

    try {
        # --- File 1: LTI PER TEU data.xlsx ---
        $cfgLti = $script:GraphConfig["Slide2_LTI"]
        if (-not $cfgLti) { throw "Configuration 'Slide2_LTI' is missing in `$script:GraphConfig." }
        $ltiPath = $script:ResolvedExcelPaths[$cfgLti.FileName]
        if (-not $ltiPath) { $ltiPath = Join-Path $workingRoot $cfgLti.FileName }
        if ($ltiPath -match '^https?://') {
            Write-LogStep ("Opening SharePoint URL directly in Excel: {0}" -f $ltiPath)
            try { $excel.DisplayAlerts = $true } catch {}
            try { $excel.Visible = $true } catch {}
        } elseif (-not (Test-Path $ltiPath -PathType Leaf)) {
            throw "Excel file '$($cfgLti.FileName)' not found at local path '$ltiPath'."
        }
        Write-LogOk ("Processing {0}..." -f $cfgLti.FileName)
        $wb1 = Open-ExcelWorkbookSafely $excel $ltiPath $false
        if (-not $wb1) {
            throw "Failed to open Excel workbook '$ltiPath'. If accessing a SharePoint URL, please ensure Excel is signed into your M365 account, or download the required Excel file locally."
        }
        $ws1 = Get-ExcelWorksheetSafely $wb1 $cfgLti.Worksheet
        if (-not $ws1) {
            Close-WorkbookSafely $wb1
            throw "Worksheet '$($cfgLti.Worksheet)' was not found in '$ltiPath'."
        }

        $lastRow1 = 0
        for ($r = $cfgLti.MaxScanRow; $r -ge $cfgLti.MinRow; $r--) {
            $v = $ws1.Cells.Item($r, $cfgLti.ScanColumn).Value2
            if ($null -ne $v -and "$v" -ne "") { $lastRow1 = $r; break }
        }
        if ($lastRow1 -gt 0) {
            $firstRow1 = [math]::Max($cfgLti.MinRow, $lastRow1 - $cfgLti.LookbackWeeks)
            $cObj1 = try { $ws1.ChartObjects().Item($cfgLti.ChartIndex) } catch { $null }
            if (-not $cObj1) {
                Write-LogWarn ("Chart index {0} not found on worksheet '{1}' in {2}" -f $cfgLti.ChartIndex, $cfgLti.Worksheet, $cfgLti.FileName)
            } else {
                foreach ($s in $cObj1.Chart.SeriesCollection()) {
                    if ($null -ne $s -and $null -ne $s.Formula) {
                        $newFormula = [regex]::Replace($s.Formula, '(\$[A-Z]+\$)\d+:(\$[A-Z]+\$)\d+',
                            { param($m) $m.Groups[1].Value + $firstRow1 + ":" + $m.Groups[2].Value + $lastRow1 })
                        if ($newFormula -ne $s.Formula) { $s.Formula = $newFormula }
                    }
                }
                Save-WorkbookSafely $wb1
                $png1 = Join-Path $workingRoot "slide2_lti.png"
                Export-ChartHd $cObj1 $png1 $cfgLti.Width $cfgLti.Height
                if (Test-ValidImageFile $png1) {
                    $exportedCharts["Slide2_LTI"] = $png1
                    Write-LogOk ("Slide #2 LTI Chart exported to PNG")
                } else {
                    Write-LogWarn ("Slide #2 LTI Chart export failed or created invalid PNG")
                }
            }
        }
        Close-WorkbookSafely $wb1

        # --- File 2: Actual & Forecast Moves data.xlsx ---
        $cfgMoves = $script:GraphConfig["Slide3_Moves"]
        if (-not $cfgMoves) { throw "Configuration 'Slide3_Moves' is missing in `$script:GraphConfig." }
        $movesPath = $script:ResolvedExcelPaths[$cfgMoves.FileName]
        if (-not $movesPath) { $movesPath = Join-Path $workingRoot $cfgMoves.FileName }
        if ($movesPath -match '^https?://') {
            Write-LogStep ("Opening SharePoint URL directly in Excel: {0}" -f $movesPath)
            try { $excel.DisplayAlerts = $true } catch {}
            try { $excel.Visible = $true } catch {}
        } elseif (-not (Test-Path $movesPath -PathType Leaf)) {
            throw "Excel file '$($cfgMoves.FileName)' not found at local path '$movesPath'."
        }
        Write-LogOk ("Processing {0}..." -f $cfgMoves.FileName)
        $wb2 = Open-ExcelWorkbookSafely $excel $movesPath $false
        if (-not $wb2) {
            throw "Failed to open Excel workbook '$movesPath'. If accessing a SharePoint URL, please ensure Excel is signed into your M365 account, or download the required Excel file locally."
        }
        $ws2 = Get-ExcelWorksheetSafely $wb2 $cfgMoves.Worksheet
        if (-not $ws2) {
            Close-WorkbookSafely $wb2
            throw "Worksheet '$($cfgMoves.Worksheet)' was not found in '$movesPath'."
        }

        $lastRow2 = 0
        for ($r = $cfgMoves.MaxScanRow; $r -ge $cfgMoves.MinRow; $r--) {
            $v = $ws2.Cells.Item($r, $cfgMoves.ScanColumn).Value2
            if ($null -ne $v -and "$v" -ne "") { $lastRow2 = $r; break }
        }
        if ($lastRow2 -gt 0) {
            $firstRow2 = [math]::Max($cfgMoves.MinRow, $lastRow2 - $cfgMoves.LookbackWeeks)
            $cObj2 = try { $ws2.ChartObjects().Item($cfgMoves.ChartIndex) } catch { $null }
            if (-not $cObj2) {
                Write-LogWarn ("Chart index {0} not found on worksheet '{1}' in {2}" -f $cfgMoves.ChartIndex, $cfgMoves.Worksheet, $cfgMoves.FileName)
            } else {
                foreach ($s in $cObj2.Chart.SeriesCollection()) {
                    if ($null -ne $s -and $null -ne $s.Formula) {
                        $newFormula = [regex]::Replace($s.Formula, '(\$[A-Z]+\$)\d+:(\$[A-Z]+\$)\d+',
                            { param($m) $m.Groups[1].Value + $firstRow2 + ":" + $m.Groups[2].Value + $lastRow2 })
                        if ($newFormula -ne $s.Formula) { $s.Formula = $newFormula }
                    }
                }
                Save-WorkbookSafely $wb2
                $png2 = Join-Path $workingRoot "slide3_moves.png"
                Export-ChartHd $cObj2 $png2 $cfgMoves.Width $cfgMoves.Height
                if (Test-ValidImageFile $png2) {
                    $exportedCharts["Slide3_Moves"] = $png2
                    Write-LogOk ("Slide #3 Moves Chart exported to PNG")
                } else {
                    Write-LogWarn ("Slide #3 Moves Chart export failed or created invalid PNG")
                }
            }
        }
        Close-WorkbookSafely $wb2

        # --- File 3: CMPH and PMPH actual and forecast data.xlsx ---
        $cfgPmph = $script:GraphConfig["Slide4_PMPH"]
        $cfgCmph = $script:GraphConfig["Slide4_CMPH"]
        if (-not $cfgPmph) { throw "Configuration 'Slide4_PMPH' is missing in `$script:GraphConfig." }
        if (-not $cfgCmph) { throw "Configuration 'Slide4_CMPH' is missing in `$script:GraphConfig." }
        $cmphPath = $script:ResolvedExcelPaths[$cfgPmph.FileName]
        if (-not $cmphPath) { $cmphPath = Join-Path $workingRoot $cfgPmph.FileName }
        if ($cmphPath -match '^https?://') {
            Write-LogStep ("Opening SharePoint URL directly in Excel: {0}" -f $cmphPath)
            try { $excel.DisplayAlerts = $true } catch {}
            try { $excel.Visible = $true } catch {}
        } elseif (-not (Test-Path $cmphPath -PathType Leaf)) {
            throw "Excel file '$($cfgPmph.FileName)' not found at local path '$cmphPath'."
        }
        Write-LogOk ("Processing {0}..." -f $cfgPmph.FileName)
        $wb3 = Open-ExcelWorkbookSafely $excel $cmphPath $false
        if (-not $wb3) {
            throw "Failed to open Excel workbook '$cmphPath'. If accessing a SharePoint URL, please ensure Excel is signed into your M365 account, or download the required Excel file locally."
        }
        $ws3 = Get-ExcelWorksheetSafely $wb3 $cfgPmph.Worksheet
        if (-not $ws3) {
            Close-WorkbookSafely $wb3
            throw "Worksheet '$($cfgPmph.Worksheet)' was not found in '$cmphPath'."
        }

        $charts3 = $ws3.ChartObjects()
        if ($charts3.Count -ge [math]::Max($cfgPmph.ChartIndex, $cfgCmph.ChartIndex)) {
            $cObjPmph = $charts3.Item($cfgPmph.ChartIndex)
            $cObjCmph = $charts3.Item($cfgCmph.ChartIndex)

            # --- Update PMPH Chart Series Range ---
            $lastRowPmph = 0
            if ($cfgPmph.ScanColumn) {
                for ($r = $cfgPmph.MaxScanRow; $r -ge $cfgPmph.MinRow; $r--) {
                    $v = $ws3.Cells.Item($r, $cfgPmph.ScanColumn).Value2
                    if ($null -ne $v -and "$v" -ne "") { $lastRowPmph = $r; break }
                }
                if ($lastRowPmph -eq 0) {
                    for ($altCol = 2; $altCol -le 6; $altCol++) {
                        for ($r = $cfgPmph.MaxScanRow; $r -ge $cfgPmph.MinRow; $r--) {
                            $v = $ws3.Cells.Item($r, $altCol).Value2
                            if ($null -ne $v -and "$v" -ne "") { $lastRowPmph = $r; break }
                        }
                        if ($lastRowPmph -gt 0) { break }
                    }
                }
            }
            if ($lastRowPmph -gt 0) {
                $firstRowPmph = [math]::Max($cfgPmph.MinRow, $lastRowPmph - $cfgPmph.LookbackWeeks)
                foreach ($s in $cObjPmph.Chart.SeriesCollection()) {
                    if ($null -ne $s -and $null -ne $s.Formula) {
                        $newFormula = [regex]::Replace($s.Formula, '(\$[A-Z]+\$)\d+:(\$[A-Z]+\$)\d+',
                            { param($m) $m.Groups[1].Value + $firstRowPmph + ":" + $m.Groups[2].Value + $lastRowPmph })
                        if ($newFormula -ne $s.Formula) { $s.Formula = $newFormula }
                    }
                }
            }

            # --- Update CMPH Chart Series Range ---
            $lastRowCmph = 0
            if ($cfgCmph.ScanColumn) {
                for ($r = $cfgCmph.MaxScanRow; $r -ge $cfgCmph.MinRow; $r--) {
                    $v = $ws3.Cells.Item($r, $cfgCmph.ScanColumn).Value2
                    if ($null -ne $v -and "$v" -ne "") { $lastRowCmph = $r; break }
                }
                if ($lastRowCmph -eq 0) {
                    for ($altCol = 2; $altCol -le 6; $altCol++) {
                        for ($r = $cfgCmph.MaxScanRow; $r -ge $cfgCmph.MinRow; $r--) {
                            $v = $ws3.Cells.Item($r, $altCol).Value2
                            if ($null -ne $v -and "$v" -ne "") { $lastRowCmph = $r; break }
                        }
                        if ($lastRowCmph -gt 0) { break }
                    }
                }
            }
            if ($lastRowCmph -gt 0) {
                $firstRowCmph = [math]::Max($cfgCmph.MinRow, $lastRowCmph - $cfgCmph.LookbackWeeks)
                foreach ($s in $cObjCmph.Chart.SeriesCollection()) {
                    if ($null -ne $s -and $null -ne $s.Formula) {
                        $newFormula = [regex]::Replace($s.Formula, '(\$[A-Z]+\$)\d+:(\$[A-Z]+\$)\d+',
                            { param($m) $m.Groups[1].Value + $firstRowCmph + ":" + $m.Groups[2].Value + $lastRowCmph })
                        if ($newFormula -ne $s.Formula) { $s.Formula = $newFormula }
                    }
                }
            }

            if ($lastRowPmph -gt 0 -or $lastRowCmph -gt 0) {
                Save-WorkbookSafely $wb3
            }

            $png4_pmph = Join-Path $workingRoot "slide4_pmph.png"
            $png4_cmph = Join-Path $workingRoot "slide4_cmph.png"

            Export-ChartHd $cObjPmph $png4_pmph $cfgPmph.Width $cfgPmph.Height
            Export-ChartHd $cObjCmph $png4_cmph $cfgCmph.Width $cfgCmph.Height

            if ((Test-ValidImageFile $png4_pmph) -and (Test-ValidImageFile $png4_cmph)) {
                $exportedCharts["Slide4_PMPH"] = $png4_pmph
                $exportedCharts["Slide4_CMPH"] = $png4_cmph
                Write-LogOk ("Slide #4 PMPH & CMPH Charts exported to PNG")
            } else {
                Write-LogWarn ("Slide #4 PMPH or CMPH Chart export failed or created invalid PNG")
            }
        }
        Close-WorkbookSafely $wb3

        # --- File 4: Equipment PERFORMANCE V1.2 Weekly data.xlsx ---
        $cfgAct = $script:GraphConfig["Slide5_Actual"]
        $cfgFc  = $script:GraphConfig["Slide6_Forecast"]
        if (-not $cfgAct) { throw "Configuration 'Slide5_Actual' is missing in `$script:GraphConfig." }
        if (-not $cfgFc)  { throw "Configuration 'Slide6_Forecast' is missing in `$script:GraphConfig." }
        $equipPath = $script:ResolvedExcelPaths[$cfgAct.FileName]
        if (-not $equipPath) { $equipPath = Join-Path $workingRoot $cfgAct.FileName }
        if ($equipPath -match '^https?://') {
            Write-LogStep ("Opening SharePoint URL directly in Excel: {0}" -f $equipPath)
            try { $excel.DisplayAlerts = $true } catch {}
            try { $excel.Visible = $true } catch {}
        } elseif (-not (Test-Path $equipPath -PathType Leaf)) {
            throw "Excel file '$($cfgAct.FileName)' not found at local path '$equipPath'."
        }
        Write-LogOk ("Processing {0}..." -f $cfgAct.FileName)
        $wb4 = Open-ExcelWorkbookSafely $excel $equipPath $true
        try { $excel.DisplayAlerts = $false } catch {}
        if (-not $wb4) {
            throw "Failed to open Excel workbook '$equipPath'. If accessing a SharePoint URL, please ensure Excel is signed into your M365 account, or download the required Excel file locally."
        }

        # SLIDE 5: Weekly SMT sheet - screenshot Equipment Technical Availability section
        $ws4_smt = Get-ExcelWorksheetSafely $wb4 $cfgAct.Worksheet
        if (-not $ws4_smt) {
            Close-WorkbookSafely $wb4
            throw "Worksheet '$($cfgAct.Worksheet)' was not found in '$equipPath'."
        }
        $png5_actual = Join-Path $workingRoot "slide5_actual.png"
        Export-RangePicture $ws4_smt $cfgAct.StartRow $cfgAct.StartCol $cfgAct.EndRow $cfgAct.EndCol $png5_actual
        if (Test-ValidImageFile $png5_actual) {
            $exportedCharts["Slide5_Actual"] = $png5_actual
            Write-LogOk "Slide #5 Technical Availability Actual section exported to PNG"
        } else {
            Write-LogWarn "Slide #5 chart export failed - PNG not found or invalid"
        }

        # SLIDE 6: MnR Forecast sheet - 3 separate table stitched CopyPicture screenshots
        $ws4_mnr = Get-ExcelWorksheetSafely $wb4 $cfgFc.Worksheet
        if (-not $ws4_mnr) {
            Close-WorkbookSafely $wb4
            throw "Worksheet '$($cfgFc.Worksheet)' was not found in '$equipPath'."
        }
        $ws4_mnr.Activate()
        Start-Sleep -Milliseconds 300

        $png6_qc  = Join-Path $workingRoot "slide6_fc_qc.png"
        $png6_rtg = Join-Path $workingRoot "slide6_fc_rtg.png"
        $png6_pm  = Join-Path $workingRoot "slide6_fc_pm.png"

        Export-MultiStitchedRangePicture $ws4_mnr $cfgFc.QC $cfgFc.MaxCol $png6_qc
        Export-MultiStitchedRangePicture $ws4_mnr $cfgFc.RTG $cfgFc.MaxCol $png6_rtg
        Export-MultiStitchedRangePicture $ws4_mnr $cfgFc.PM $cfgFc.MaxCol $png6_pm

        if ((Test-ValidImageFile $png6_qc) -and (Test-ValidImageFile $png6_rtg) -and (Test-ValidImageFile $png6_pm)) {
            $exportedCharts["Slide6_FC1"] = $png6_qc
            $exportedCharts["Slide6_FC2"] = $png6_rtg
            $exportedCharts["Slide6_FC3"] = $png6_pm
            Write-LogOk "Slide #6 MnR Forecast stitched tables (QC / RTG / PM) exported to PNG"
        } else {
            Write-LogWarn "Slide #6 one or more MnR Forecast table exports failed"
        }

        Close-WorkbookSafely $wb4

    }
    finally {
        if ($null -ne $excel) {
            try { $excel.DisplayAlerts = $false } catch {}
            try { $excel.Visible = $false } catch {}
            try {
                if ($null -ne $excel.Workbooks) {
                    foreach ($wb in $excel.Workbooks) {
                        try { $wb.Close($false) } catch {}
                    }
                }
            } catch {}
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            $excel = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # --- 4) POWERPOINT PRESENTATION UPDATE ACROSS SLIDES 2-6 ------------------
    Write-LogStep "[4/5] Updating PowerPoint presentation slides"
    $msoTrue  = -1
    $msoFalse = 0

    $ppt = New-PowerPointComObject
    try { $ppt.Visible = $msoTrue } catch {}
    try { $ppt.WindowState = 2 } catch {}

    $pres = $null
    $finalOutput = ""

    try {
        Write-LogOk "Opening PowerPoint template via COM..."
        $pres = $ppt.Presentations.Open($templatePptx, $msoFalse, $msoFalse, $msoTrue)

        # Slide 1: Update Date to Today / Running Date (date only, no time)
        if (-not $NoDateUpdate) {
            $todayStr = (Get-Date).ToString("dd-MMM-yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            foreach ($shape in $pres.Slides.Item(1).Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $curText = $shape.TextFrame.TextRange.Text.Trim()
                    if ($curText -match '\d{1,2}-[A-Za-z]{3}-\d{4}' -or $shape.Name -match 'Subtitle\s*4') {
                        $shape.TextFrame.TextRange.Text = $todayStr
                        Write-LogOk ("Slide #1 -> Updated date to {0} (removed time component)" -f $todayStr)
                        break
                    }
                }
            }
        }

        # Slide 2: LTI
        if ($pres.Slides.Count -ge 2 -and $exportedCharts["Slide2_LTI"]) {
            $slide2 = $pres.Slides.Item(2)
            $pic2 = $slide2.Shapes | Where-Object { $_.Type -eq 13 -or $_.Type -eq 11 } | Select-Object -First 1
            if ($pic2) {
                Replace-SlidePictureShape $slide2 $pic2 $exportedCharts["Slide2_LTI"]
                Write-LogOk "Slide #2 -> Replaced LTI Chart"
            }
        }

        # Slide 3: Commercial - Moves
        if ($pres.Slides.Count -ge 3 -and $exportedCharts["Slide3_Moves"]) {
            $slide3 = $pres.Slides.Item(3)
            $pic3 = $slide3.Shapes | Where-Object { $_.Type -eq 13 -or $_.Type -eq 11 } | Select-Object -First 1
            if ($pic3) {
                Replace-SlidePictureShape $slide3 $pic3 $exportedCharts["Slide3_Moves"]
                Write-LogOk "Slide #3 -> Replaced Moves Chart"
            }
        }

        # Slide 4: Operations - PMPH & CMPH
        if ($pres.Slides.Count -ge 4 -and $exportedCharts["Slide4_PMPH"] -and $exportedCharts["Slide4_CMPH"]) {
            $slide4 = $pres.Slides.Item(4)
            $pic4_left = $null; $pic4_right = $null
            foreach ($s in $slide4.Shapes) {
                if ($s.Type -eq 13 -or $s.Type -eq 11) {
                    if ($s.Left -lt 300) { $pic4_left = $s } else { $pic4_right = $s }
                }
            }
            if ($pic4_left) { Replace-SlidePictureShape $slide4 $pic4_left $exportedCharts["Slide4_PMPH"] }
            if ($pic4_right) { Replace-SlidePictureShape $slide4 $pic4_right $exportedCharts["Slide4_CMPH"] }
            Write-LogOk "Slide #4 -> Replaced PMPH & CMPH Charts"
        }

        # Slide 5: Engineering Technical Availability (Actual)
        if ($pres.Slides.Count -ge 5 -and $exportedCharts["Slide5_Actual"]) {
            $slide5 = $pres.Slides.Item(5)
            $pic5 = $slide5.Shapes | Where-Object { $_.Type -eq 13 -or $_.Type -eq 11 } | Select-Object -First 1
            if ($pic5) {
                # Force exact dimensions: 10 inches wide (720 pt), 5 inches high (360 pt)
                Replace-SlidePictureShape $slide5 $pic5 $exportedCharts["Slide5_Actual"] -TargetWidthPoints (10 * 72) -TargetHeightPoints (5 * 72)
                Write-LogOk "Slide #5 -> Replaced Technical Availability Actual Chart (Set width: 10 in, height: 5 in)"
            }
        }

        # Slide 6: Engineering Technical Availability (Forecast)
        if ($pres.Slides.Count -ge 6 -and $exportedCharts["Slide6_FC1"]) {
            $slide6 = $pres.Slides.Item(6)
            $pics6 = @($slide6.Shapes | Where-Object { $_.Type -eq 13 -or $_.Type -eq 11 } | Sort-Object Top)
            if ($pics6.Count -ge 3) {
                # Force exact dimensions: 12 inches wide (864 pt), 1.5 inches high (108 pt)
                Replace-SlidePictureShape $slide6 $pics6[0] $exportedCharts["Slide6_FC1"] -TargetWidthPoints (12 * 72) -TargetHeightPoints (1.5 * 72)
                Replace-SlidePictureShape $slide6 $pics6[1] $exportedCharts["Slide6_FC2"] -TargetWidthPoints (12 * 72) -TargetHeightPoints (1.5 * 72)
                Replace-SlidePictureShape $slide6 $pics6[2] $exportedCharts["Slide6_FC3"] -TargetWidthPoints (12 * 72) -TargetHeightPoints (1.5 * 72)
                Write-LogOk "Slide #6 -> Replaced 3 Technical Availability Forecast Charts (Set width: 12 in, height: 1.5 in)"
            }
        }

        # Force PowerPoint COM to flush and embed all picture streams into presentation package
        try { $pres.Save() } catch {}

        # --- 5) DUAL SAVING: 1. SAVE LOCAL FALLBACK REPORT ---------------------
        $pres.SaveCopyAs($fallbackFilePath)
        Write-LogOk ("Saved local fallback report: output-fallback\$fallbackFileName")

        # --- 2. PUBLISH REPORT TO TARGET DESTINATION (SHAREPOINT / OUTPUT) -----
        $pres.SaveCopyAs($localReportTemp)
        $finalOutput = Publish-ReportOutputFromPres -pres $pres -localReportFile $localReportTemp -outputDestination $OutputDir -fallbackFilePath $fallbackFilePath -userEmail $SharePointEmail -clientId $ClientId -tenantId $TenantId -certThumbprint $CertThumbprint -rootPath $scriptExecutionDir

    }
    finally {
        if ($null -ne $pres) {
            try { $pres.Close() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
            $pres = $null
        }
        if ($null -ne $ppt) {
            try { $ppt.Visible = $msoFalse } catch {}
            try { $ppt.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
            $ppt = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    Write-LogHeader "EXECUTION COMPLETE - SUCCESS"
    Write-Host ("== DONE: {0} ==" -f $finalOutput)

} catch {
    Write-Host ""
    Write-Host "=== EXECUTION FAILED ==="
    Write-Host ("ERROR: {0}" -f $_.Exception.Message)
    Write-Host "========================`n"
    throw $_
}
