# Weekly Report Agent

Automates the weekly executive PowerPoint report from Excel data in **5 fully automated, self-healing steps**:

1. **Mandatory 4-Excel Check** — Enforces resolution of all 4 required Excel data sources (from SharePoint HTTPS links or local disk paths).
2. **Dynamic Data Range Scan & Chart Refresh** — Scans rows dynamically for newly added weekly data and updates chart range formulas across all 4 Excel workbooks via Excel COM.
3. **Template Resolution** — Automatically resolves the highest-numbered `Weekly Report Week XX.pptx` template file.
4. **Report Generation & Picture Replacement** — Duplicates the template, embeds high-definition static PNG charts into placeholder shapes while preserving original bounds/Z-order, updates report date/period text, and flushes media streams into the presentation package.
5. **Multi-Tiered Publication** — Uploads the finished deck directly to SharePoint Online using PowerPoint COM `SaveAs` / `SaveCopyAs`, local OneDrive/SharePoint synced directory auto-detection, WebDAV UNC, or MS Graph API, with a guaranteed local fallback copy.

---

## Project Files

| File / Directory             | Purpose                                                                        |
| ---------------------------- | ------------------------------------------------------------------------------ |
| `main.bat`                   | **Start here.** Double-click to launch the executive GUI dashboard.            |
| `WeeklyReportDashboard.ps1`  | WinForms GUI dashboard — configures inputs, authenticates, and triggers runs.  |
| `WeeklyReportAgent.ps1`      | The core automation engine — handles Excel processing, PPTX building & upload. |
| `dashboard-config.json`      | Persisted dashboard input settings (auto-saved on run, encrypted via DPAPI).   |
| `Weekly Report Week XX.pptx` | The baseline weekly report deck, which serves as the template for next week.   |
| `output-fallback/`           | Directory where reports are saved as a fallback if SharePoint upload fails.    |

---

## Prerequisites

- **Windows 10 / 11** with PowerShell 5.1 or later.
- **Microsoft 365 desktop apps** installed — specifically Microsoft Excel and Microsoft PowerPoint (COM automation required).
- **Signed in to Microsoft 365** in the desktop apps with the account that has read/write access to your SharePoint site.
- No additional PowerShell modules, admin rights, or Azure AD app registrations are required for default Direct Sign-In mode.

---

## Quick Start

1. Place the 4 required Excel files in the project folder **or** have their SharePoint HTTPS links ready.
2. Ensure a `Weekly Report Week XX.pptx` file exists in the project folder — this is your baseline template.
3. Double-click **`main.bat`**.
4. Fill in the 4 Excel file paths/links and the SharePoint output URL in the dashboard.
5. Click **RUN WEEKLY REPORT AGENT**.

---

## Required Input Files

The agent enforces a **hard-stop 4-Excel check** before processing. All four files must be resolvable — either as a local `.xlsx` file path or a SharePoint HTTPS URL — or the agent halts immediately.

| #   | Required Filename                             | Target Slide / Purpose                            |
| --- | --------------------------------------------- | ------------------------------------------------- |
| 1   | `Actual & Forecast Moves data.xlsx`           | Slide 3 — Commercial Moves Chart                  |
| 2   | `CMPH and PMPH actual and forecast data.xlsx` | Slide 4 — Operations PMPH & CMPH Charts           |
| 3   | `Equipment PERFORMANCE V1.2 Weekly data.xlsx` | Slide 5 & Slide 6 — Technical Availability Tables |
| 4   | `LTI PER TEU data.xlsx`                       | Slide 2 — Safety LTI PER TEU Chart                |

> **Hard Stop:** If any of the 4 required files cannot be resolved, the agent logs an explicit error message and halts execution.

---

## Dashboard Layout & Features

The executive dashboard ([`WeeklyReportDashboard.ps1`](file:///c:/Users/User/exceltoPPTAI/WeeklyReportDashboard.ps1)) is split into 4 grouped sections:

### 1 — Data Input (Excel Files & Template)

Paste a **SharePoint HTTPS URL** (e.g. `https://tenant.sharepoint.com/sites/.../file.xlsx`) or a **local file path** for each of the 4 required Excel files. Use the **Browse…** button to pick a local file.

The fifth row is the **PowerPoint Template** field:

- **Leave blank** — auto-resolves the highest-numbered `Weekly Report Week XX.pptx` file in the script folder.
- **Local `.pptx` path** — browse to any local PowerPoint file to use as the template.
- **SharePoint HTTPS URL** — paste a direct link to a `.pptx` in SharePoint; the agent downloads it via Office SSO before processing.

### 2 — SharePoint Profile & Authentication (Direct Sign-In)

- **Direct Sign-In (Office SSO)**: Uses your active Microsoft 365 desktop session credentials.
- **SharePoint Email & Password**: Optional credentials saved encrypted on disk using Windows DPAPI.

### 3 — Output & Report Settings

- **Output Location**: Local folder path **or** a SharePoint library HTTPS URL (e.g. `https://tenant.sharepoint.com/sites/SiteName/Shared Documents/Output`).
- **Automatic Current Week Execution**: The agent automatically calculates and runs for the current calendar week.

### 4 — Execution & Live Log

- **Status badge**: Displays `READY` → `RUNNING...` → `DONE!` or `FAILED`.
- **Live log**: Real-time console output window polling child agent execution without UI thread freezing.
- **Action Buttons**: **RUN WEEKLY REPORT AGENT**, **Open Generated Report**, and **Open Output**.

---

## How the Agent Works (Step-by-Step Architecture)

```
main.bat
  └─► WeeklyReportDashboard.ps1 (WinForms GUI)
          └─► WeeklyReportAgent.ps1 (Powershell Automation Engine - STA Mode)
                  ├─► 1. Resolve & Validate 4 Excel Files + Template
                  ├─► 2. Excel COM: Dynamic Row Scan & Chart Range Updating
                  ├─► 3. GDI+ PNG Export & File Validation Polling
                  ├─► 4. PowerPoint COM: Shape Replacement & Date/Week Update
                  └─► 5. Multi-Tiered SharePoint / Local Publication Pipeline
```

### Step 1: Input Resolution & 4-Excel Validation

Clean quotes and resolves all 4 Excel URLs/paths and PowerPoint template. Resolves SharePoint URLs via Office SSO or Windows Integrated Auth. Enforces hard-stop check.

### Step 2: Dynamic Table Row Scanning & Chart Series Formula Refresh

For each workbook, the agent scans data columns backward from `MaxScanRow` down to `MinRow` to find the last row `$lastRow` containing non-empty weekly data. It calculates `$firstRow = [math]::Max($minRow, $lastRow - $lookback)` and regex-updates series formulas:
`=SERIES("Series", Sheet1!$B$4:$B$15, Sheet1!$C$4:$C$15, 1)` ➔ `=SERIES("Series", Sheet1!$B$4:$B$16, Sheet1!$C$4:$C$16, 1)`
When users expand Excel tables by adding new row data for new weeks, the graph view window automatically slides down to capture the updated range.

### Step 3: Image Export & Validation Polling (`Test-ValidImageFile`)

Charts and stitched cell ranges are exported into an isolated process working directory (`$workingRoot = %TEMP%\wra_input_<guid>`). Every PNG image is validated for existence, file size (> 1 KB), and GDI+ read capability with a 3-second polling loop before PowerPoint COM touches it.

### Step 4: PowerPoint Shape Replacement & Package Flushing

Opens the PowerPoint template via PowerPoint COM. Locates slide pictures, records bounds/Z-order, inserts the new PNG via `AddPicture($path, 0, -1, ...)` (`LinkToFile = 0`, `SaveWithDocument = -1`), deletes the old shape, updates the date string on Slide 1 (`dd-MMM-yyyy`), and calls `$pres.Save()` to flush all image streams (`ppt/media/imageX.png`) into the PPTX package archive.

### Step 5: Multi-Tiered Publication Pipeline

Attempts publication across 5 fallback channels:

1. **PowerPoint COM `SaveAs`**: Direct push over Office SSO. Includes 60-second upload stream timeout with 10-second status logging for slow internet connections.
2. **PowerPoint COM `SaveCopyAs`**: Secondary COM upload channel.
3. **OneDrive / SharePoint Synced Local Directory Auto-Detection**: Searches local OneDrive business folders (`C:\Users\User\OneDrive - SRKK Group 1`, `C:\Users\User\SharePoint`) for directories matching the target SharePoint folder name and copies the report directly into the synced folder for instant background cloud sync.
4. **WebDAV UNC Copy**: Copies to `\\tenant@SSL\DavWWWRoot\...`.
5. **MS Graph API**: Uploads via MS Graph REST API if App Cert is configured.
6. **Local & Fallback Report**: Saves fallback report to `output-fallback\Weekly Report Week XX_Fallback.pptx`.

---

## Configuring Excel Table Cell Scopes & Graph Locations

Excel chart object indices, worksheet names, cell range boundaries, column scan rules, and lookback windows are defined centrally in `$script:GraphConfig` near line 40 of [`WeeklyReportAgent.ps1`](file:///c:/Users/User/exceltoPPTAI/WeeklyReportAgent.ps1).

### Slide-to-Excel Central Mapping Table

| Slide #     | Target Topic / Slide         | Source Excel File                             | Sheet Name / Index | Target Object / Cell Scope Rules                                                               |
| :---------- | :--------------------------- | :-------------------------------------------- | :----------------- | :--------------------------------------------------------------------------------------------- |
| **Slide 2** | Safety LTI PER TEU YTD       | `LTI PER TEU data.xlsx`                       | `1` (Sheet1)       | Chart 1 (Scanned Col F / Col 6, Rows 4–56, 12-week window)                                     |
| **Slide 3** | Commercial Moves             | `Actual & Forecast Moves data.xlsx`           | `1` (Sheet1)       | Chart 1 (Scanned Col D / Col 4, Rows 5–56, 16-week window)                                     |
| **Slide 4** | Operations PMPH & CMPH       | `CMPH and PMPH actual and forecast data.xlsx` | `1` (Sheet1)       | Chart 1 (PMPH) & Chart 2 (CMPH) (Scanned Col D / Col 4, Rows 4–56, 12-week window)             |
| **Slide 5** | Tech Availability (Actual)   | `Equipment PERFORMANCE V1.2 Weekly data.xlsx` | `"Weekly SMT"`     | Cell Range `U39:AK55` (StartRow 39, StartCol 21, EndRow 55, EndCol 37)                         |
| **Slide 6** | Tech Availability (Forecast) | `Equipment PERFORMANCE V1.2 Weekly data.xlsx` | `"MnR Forecast"`   | Stitched Range Tables: QC (R4-5, R11-17), RTG (R21-22, R28-34), PM (R38-39, R45-55), MaxCol 30 |

---

### How to Edit Table Ranges, Scan Columns, & Lookback Windows in Code

Open [`WeeklyReportAgent.ps1`](file:///c:/Users/User/exceltoPPTAI/WeeklyReportAgent.ps1) and locate `$script:GraphConfig`:

```powershell
$script:GraphConfig = @{
    # Slide 2: LTI Chart
    Slide2_LTI = @{
        FileName         = "LTI PER TEU data.xlsx"
        Worksheet        = 1                 # Worksheet index (1-based) or sheet name string
        ChartIndex       = 1                 # ChartObject index on sheet
        ScanColumn       = 6                 # Column F for last-row data lookup
        MaxScanRow       = 56                # Start row for scanning backward
        MinRow           = 4                 # Absolute minimum data row
        LookbackWeeks    = 11                # Spans lastRow - 11 to lastRow (12 weeks total)
        Width            = 823               # PNG export width (pixels)
        Height           = 319               # PNG export height (pixels)
    }

    # Slide 4: Operations PMPH & CMPH Charts
    Slide4_PMPH = @{
        FileName         = "CMPH and PMPH actual and forecast data.xlsx"
        Worksheet        = 1
        ChartIndex       = 1                 # Chart 1: PMPH
        ScanColumn       = 4                 # Column D for last-row lookup
        MaxScanRow       = 56
        MinRow           = 4
        LookbackWeeks    = 11
        Width            = 433
        Height           = 224
    }

    # Slide 5: Equipment Availability Actual Range
    Slide5_Actual = @{
        FileName         = "Equipment PERFORMANCE V1.2 Weekly data.xlsx"
        Worksheet        = "Weekly SMT"
        StartRow         = 39                # Top row boundary
        StartCol         = 21                # Column U
        EndRow           = 55                # Bottom row boundary
        EndCol           = 37                # Column AK
    }

    # Slide 6: Equipment Availability Forecast Stitched Tables
    Slide6_Forecast = @{
        FileName         = "Equipment PERFORMANCE V1.2 Weekly data.xlsx"
        Worksheet        = "MnR Forecast"
        MaxCol           = 30                # Column AD
        QC               = @( @{Start=4; End=5}, @{Start=11; End=17} )
        RTG              = @( @{Start=21; End=22}, @{Start=28; End=34} )
        PM               = @( @{Start=38; End=39}, @{Start=45; End=46}, @{Start=51; End=55} )
    }
}
```

#### Step-by-Step Customization Guide:

1. **Changing Sheet Name or Index**: Update `Worksheet = "Sheet Name"` or `Worksheet = 1`.
2. **Changing Target Chart**: Update `ChartIndex = 2` if the chart object index on the worksheet changes.
3. **Changing Column Scan for Data Detection**: Change `ScanColumn = 4` (Col D), `ScanColumn = 6` (Col F), or `ScanColumn = 2` (Col B) depending on which column contains weekly entries.
4. **Changing Table Row Boundaries**:
   - For chart scanning: Adjust `MaxScanRow` (e.g. `100` if tables expand beyond week 53) and `MinRow` (header offset).
   - For Slide 5 cell range capture: Change `StartRow`, `EndRow`, `StartCol` (e.g. Col U = 21), or `EndCol` (e.g. Col AK = 37).
5. **Changing Lookback Window Size**: Change `LookbackWeeks = 11` (shows 12 weeks), `LookbackWeeks = 15` (shows 16 weeks), or set to `$null` to display all rows from `MinRow` to `lastRow`.

---

## Headless / Command-Line Usage

Run directly from PowerShell without the GUI dashboard:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ".\WeeklyReportAgent.ps1" `
    -MovesDataUrl    'https://tenant.sharepoint.com/sites/.../Actual & Forecast Moves data.xlsx' `
    -CmphDataUrl     'https://tenant.sharepoint.com/sites/.../CMPH and PMPH actual and forecast data.xlsx' `
    -EquipmentDataUrl 'https://tenant.sharepoint.com/sites/.../Equipment PERFORMANCE V1.2 Weekly data.xlsx' `
    -LtiDataUrl      'https://tenant.sharepoint.com/sites/.../LTI PER TEU data.xlsx' `
    -OutputDir       'https://tenant.sharepoint.com/sites/SiteName/Shared Documents/Output'
```

> **Important:** The `-STA` flag is mandatory. PowerShell must run in Single-Threaded Apartment mode for Office COM automation to operate correctly.

### Full Agent Parameter Reference

| Parameter           | Type     | Description                                                                     |
| :------------------ | :------- | :------------------------------------------------------------------------------ |
| `-Week`             | `int`    | Target week number. Defaults to current ISO week (`-1` = auto).                 |
| `-MovesDataUrl`     | `string` | Local path or SharePoint URL for `Actual & Forecast Moves data.xlsx`.           |
| `-CmphDataUrl`      | `string` | Local path or SharePoint URL for `CMPH and PMPH actual and forecast data.xlsx`. |
| `-EquipmentDataUrl` | `string` | Local path or SharePoint URL for `Equipment PERFORMANCE V1.2 Weekly data.xlsx`. |
| `-LtiDataUrl`       | `string` | Local path or SharePoint URL for `LTI PER TEU data.xlsx`.                       |
| `-TemplatePath`     | `string` | Local `.pptx` path or SharePoint URL for PowerPoint template.                   |
| `-OutputDir`        | `string` | Local directory path or SharePoint library URL for the output deck.             |
| `-SharePointEmail`  | `string` | Microsoft 365 email address.                                                    |
| `-NoDateUpdate`     | `switch` | Skip updating the title slide date.                                             |

---

## Troubleshooting & Best Practices

| Symptom                                                | Likely Cause                                 | Resolution                                                                                      |
| :----------------------------------------------------- | :------------------------------------------- | :---------------------------------------------------------------------------------------------- |
| `Missing required Excel file(s)`                       | File missing or inaccessible URL             | Verify local path or check SharePoint access in browser.                                        |
| `Mandatory PowerPoint template file not found`         | No `.pptx` template in script directory      | Place `Weekly Report Week XX.pptx` in the script directory or set `-TemplatePath`.              |
| `This picture cannot currently be displayed`           | Stale or un-flushed image stream             | Resolved via GDI+ validation polling (`Test-ValidImageFile`) & `$pres.Save()` package flushing. |
| `Direct SharePoint cloud upload unavailable`           | Tenant MFA or unauthenticated Office session | Script automatically publishes report locally to output folder and `output-fallback\`.          |
| `You cannot call a method on a null-valued expression` | Missing `-STA` flag                          | Always include `-STA` when launching PowerShell COM scripts.                                    |

---

## Configuration Persistence & Security

Dashboard settings are saved to `dashboard-config.json` every time you click **RUN**. Passwords are encrypted using **Windows Data Protection API (DPAPI)** bound to your Windows user profile. `dashboard-config.json` is listed in `.gitignore` to prevent credential exposure.
