<#
.SYNOPSIS
Run the Excel evidence session of docs\EXCEL_EVIDENCE.md for the checked-out
commit with one command and write a validated evidence bundle (#238).
.DESCRIPTION
For the commit checked out in this repository, with a clean working tree:
1. creates the bundle folder, by default irrbb-evidence-SHORTSHA next to the
   checkout, and starts session.log there;
2. records the environment as found: Windows, Excel version and build, Office
   bitness, the macro setting and "Trust access to the VBA project object
   model". It never changes any of them;
3. opens the macro-free template read-only and records whether Excel opened it
   without repair, its sheets, CodeNames, defined names, title fonts and data
   validation (each validated cell refuses an invalid value and accepts a
   valid one), and its appearance: tab colours, the brand strip and title,
   the dropdowns, the print titles and areas, and the page setup Excel writes
   when it saves a copy of the template into the bundle (template.log).
   Nothing is printed, exported or previewed, and no print object is read, so
   Excel never waits on a printer;
4. builds the development workbook with tools\devbuild\build_dev_workbook.py
   (build.log) and opens it in a visible Excel with its macros enabled for
   that Excel session only (AutomationSecurity "low"): the Trust Center is
   neither read for it nor changed. Should the macros still not run, it asks
   the operator to click Enable Content;
5. records the references, compiles the project (Debug > Compile VBAProject)
   and runs the harness (harness.log), its failure path: injected failures,
   a rerun, an injected write fault and a final run (failure.log), and the
   import: accepted, matching and differing control totals, a pasted invalid
   window date, a pasted invalid attestation, an account file another
   process holds open for writing (error 70) and a rerun (import.log). It then
   records every button, its label and the macro it runs, checks that each
   macro is a procedure of a standard module, and presses Import data and Run
   tests by running their macros as a click does (buttons.log); the
   file-choice buttons open a dialog, so only their wiring is checked. It
   then writes a large synthetic panel (about 47 MB) with
   tools\devbuild\make_large_panel.py and times its import (timing.log); the
   time is recorded, not judged against a target. Excel's
   calculation, events, screen updating and alerts are recorded before and
   after every run. The workbook is then closed without saving;
6. saves the template as IRRBB_Sync.xlsm, synchronizes it with
   tools\Sync-IRRBB-VBA.ps1 -Mode Development, compiles it and runs the
   harness (sync.log); exports it with tools\Export-IRRBB-VBA.ps1 and compares
   the export with tools\compare_vba_export.py (roundtrip.log); then runs the
   synchronizer again while another Excel holds the workbook open, which must
   fail and leave the workbook and the security settings unchanged
   (sync-failure.log);
7. writes session.json and runs tools\write_excel_evidence.py, which judges
   every stage from the logs, writes host.json and validates the bundle as
   tools\check_excel_evidence.py does.
Every log starts with CANDIDATE= and the SHA. The exit code is 0 only when the
bundle validates as PASS. The manual procedure of docs\EXCEL_EVIDENCE.md stays
available as a fallback.
.REQUIREMENTS
- Windows desktop Excel, and Python 3.11 or later on the PATH (or -Python).
- "Trust access to the VBA project object model" already on: the compile, the
  synchronizer and the export need it, and the script never turns it on.
- A clean checkout of the commit to certify.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-IRRBB.ps1
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-IRRBB.ps1 -BundleDir "C:\Evidence\irrbb-0.2.0" -Operator "Owner, release workstation"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Operator = "Repository owner, local Windows workstation",
    [Parameter(Mandatory = $false)]
    [string]$Python = "python",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw "Cannot determine the script path. Supply -RepoRoot explicitly."
    }
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$msoAutomationSecurityLow = 1
$msoAutomationSecurityForceDisable = 3
$compileControlId = 578
$xlOpenXMLWorkbookMacroEnabled = 52
$xlCellTypeAllValidation = -4174
$xlValidateList = 3
$xlListSeparator = 5
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$fixtures = Join-Path $RepoRoot "tests\fixtures\data_contract"
$templatePath = Join-Path $RepoRoot "src\workbook\IRRBB_Template.xlsx"

# ----------------------------------------------------------------------------
# Session log and evidence logs
# ----------------------------------------------------------------------------
function Get-Timestamp {
    return (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

function Write-Session {
    param([string]$Text)
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line
    [System.IO.File]::AppendAllText($script:sessionLog, $line + "`r`n", $script:utf8NoBom)
}

function New-EvidenceLog {
    # A log is never replaced: a failing record must not be overwritten.
    param([string]$Name)
    $path = Join-Path $script:BundleDir $Name
    if (Test-Path -LiteralPath $path) {
        throw "The log exists already: $path"
    }
    [System.IO.File]::WriteAllText($path, "CANDIDATE=$($script:sha)`r`n", $script:utf8NoBom)
    return $path
}

function Add-Evidence {
    param([string]$Path, [string]$Key, $Value)
    $text = ConvertTo-EvidenceText $Value
    $text = ($text -replace "[\r\n]+", " ")
    [System.IO.File]::AppendAllText($Path, "$Key=$text`r`n", $script:utf8NoBom)
}

function ConvertTo-EvidenceText {
    # Cell values as Value2 gives them: numbers in invariant culture, so 16
    # stays "16" and a date serial stays a number.
    param($Value)
    if ($null -eq $Value) {
        return ""
    }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        return ([double]$Value).ToString("R", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [int] -or $Value -is [long]) {
        return ([long]$Value).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
}

# ----------------------------------------------------------------------------
# Host security settings, read and never changed
# ----------------------------------------------------------------------------
function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-SecurityValue {
    # A value set by group policy takes precedence over the user's own setting.
    param([string]$OfficeVersion, [string]$Name)
    foreach ($root in @("HKCU:\Software\Policies\Microsoft\Office", "HKCU:\Software\Microsoft\Office")) {
        $value = Get-RegistryValue -Path "$root\$OfficeVersion\Excel\Security" -Name $Name
        if ($null -ne $value) {
            return $value
        }
    }
    return $null
}

function Get-MacroPolicy {
    param([string]$OfficeVersion)
    switch (Get-SecurityValue -OfficeVersion $OfficeVersion -Name "VBAWarnings") {
        1 { return "Enable VBA macros" }
        2 { return "Disable VBA macros with notification" }
        3 { return "Disable VBA macros except digitally signed macros" }
        4 { return "Disable VBA macros without notification" }
        default { return "Disable VBA macros with notification (default; no setting found)" }
    }
}

function Get-VbaProjectAccess {
    param([string]$OfficeVersion)
    if ((Get-SecurityValue -OfficeVersion $OfficeVersion -Name "AccessVBOM") -eq 1) {
        return "On"
    }
    return "Off"
}

# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
function New-Excel {
    param([bool]$Visible, [int]$Security)
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $Visible
    $app.AutomationSecurity = $Security
    return $app
}

function Close-Excel {
    param($App, $Book)
    if ($null -ne $Book) {
        try { $Book.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Book)
    }
    if ($null -ne $App) {
        try { $App.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Get-ExcelState {
    param($App)
    return "{0}|{1}|{2}|{3}" -f $App.Calculation, $App.EnableEvents, $App.ScreenUpdating, $App.DisplayAlerts
}

function ConvertTo-Rgb {
    # Excel gives a colour as a BGR integer; the template states it as RRGGBB.
    param($Color)
    if (-not ($Color -is [double] -or $Color -is [int] -or $Color -is [long])) { return [string]$Color }
    $value = [long]$Color
    return "{0:X2}{1:X2}{2:X2}" -f ($value -band 0xFF), (($value -shr 8) -band 0xFF), (($value -shr 16) -band 0xFF)
}

function Read-ZipXml {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { throw "$Name is missing from the saved copy" }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-XmlAttribute {
    param($Node, [string]$Name)
    if ($null -eq $Node) { return "" }
    return [string]$Node.GetAttribute($Name)
}

function Read-SavedPageLayout {
    # Each sheet's page layout as Excel saved it: read from the saved package,
    # never through Excel's print objects, which query the printer. One entry
    # per sheet: name|orientation|paperSize|fitToWidth|fitToHeight|fitToPage.
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $rows = New-Object System.Collections.Generic.List[string]
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbook = Read-ZipXml $zip "xl/workbook.xml"
        $targets = @{}
        foreach ($rel in (Read-ZipXml $zip "xl/_rels/workbook.xml.rels").GetElementsByTagName("Relationship")) {
            $targets[$rel.GetAttribute("Id")] = $rel.GetAttribute("Target")
        }
        foreach ($sheet in $workbook.GetElementsByTagName("sheet")) {
            $target = [string]$targets[$sheet.GetAttribute("id", $relNs)]
            $part = if ($target.StartsWith("/")) { $target.Substring(1) } else { "xl/" + $target }
            $xml = Read-ZipXml $zip $part
            $setup = $xml.GetElementsByTagName("pageSetup") | Select-Object -First 1
            $fit = $xml.GetElementsByTagName("pageSetUpPr") | Select-Object -First 1
            $rows.Add(("{0}|{1}|{2}|{3}|{4}|{5}" -f $sheet.GetAttribute("name"),
                       (Get-XmlAttribute $setup "orientation"), (Get-XmlAttribute $setup "paperSize"),
                       (Get-XmlAttribute $setup "fitToWidth"), (Get-XmlAttribute $setup "fitToHeight"),
                       (Get-XmlAttribute $fit "fitToPage")))
        }
    }
    finally {
        $zip.Dispose()
    }
    return ,$rows
}

function Write-Buttons {
    # Every button of the development workbook and the macro it runs; then
    # Import data and Run tests pressed by running their actions as recorded,
    # as a click does. An action that names no workbook, or the file format's
    # "this workbook" ([0]), runs in this workbook.
    param($App, $Book, [string]$Qualified, [string]$Log)
    $actions = @{}
    foreach ($sheet in $Book.Worksheets) {
        foreach ($shape in $sheet.Shapes) {
            $label = [string]$shape.TextFrame2.TextRange.Text
            $actions[$label] = [string]$shape.OnAction
            Add-Evidence $Log "BUTTON" ("{0}|{1}|{2}" -f $sheet.Name, $label, $shape.OnAction)
        }
    }
    $macros = @($actions.Values | ForEach-Object { ($_ -split '!')[-1] } | Sort-Object -Unique)
    foreach ($macro in $macros) {
        $where = ""
        foreach ($component in $Book.VBProject.VBComponents) {
            try {
                [void]$component.CodeModule.ProcStartLine($macro, 0)
                $where = "{0}|{1}" -f $component.Name, $component.Type
                break
            }
            catch {}
        }
        Add-Evidence $Log "MACRO" ("{0}|{1}" -f $macro, $where)
    }
    foreach ($press in @(@("Import data", "DataSummary_Status"), @("Run tests", "Checks_Status"))) {
        Add-Evidence $Log "PRESS" $press[0]
        Add-Evidence $Log "STATE_BEFORE" (Get-ExcelState $App)
        try {
            $action = [string]$actions[$press[0]]
            if ($action -notmatch '!' -or $action.StartsWith("[0]!")) { $action = $Qualified + ($action -split '!')[-1] }
            [void]$App.Run($action)
        }
        catch {
            Add-Evidence $Log "ERROR" $_.Exception.Message
        }
        Add-Evidence $Log "STATE_AFTER" (Get-ExcelState $App)
        $status = (Get-Area -Book $Book -Name $press[1]).Value2
        Add-Evidence $Log "STATUS" $status
        Write-Session ("  button {0}: {1}" -f $press[0], $status)
    }
}

function Get-Area {
    # A Range is enumerable: returned bare, PowerShell would unroll it into its
    # cells, and .Value2 would then be a flat list instead of the area's array.
    param($Book, [string]$Name)
    return ,($Book.Names.Item($Name).RefersToRange)
}

function Get-Rows {
    # The rows of a named area whose first cell is not empty, each joined with |.
    param($Book, [string]$Name)
    $area = Get-Area -Book $Book -Name $Name
    $data = $area.Value2
    $rows = New-Object System.Collections.Generic.List[string]
    if ($null -eq $data) {
        return ,$rows
    }
    if (-not ($data -is [System.Array])) {
        $rows.Add((ConvertTo-EvidenceText $data))
        return ,$rows
    }
    # Value2 is a 1-based two-dimensional array; GetValue takes its true indices,
    # where PowerShell's own indexer would count from 0.
    $firstColumn = $data.GetLowerBound(1)
    for ($r = $data.GetLowerBound(0); $r -le $data.GetUpperBound(0); $r++) {
        $first = ConvertTo-EvidenceText $data.GetValue($r, $firstColumn)
        if ($first -eq "") {
            continue
        }
        $cells = for ($c = $firstColumn; $c -le $data.GetUpperBound(1); $c++) {
            ConvertTo-EvidenceText $data.GetValue($r, $c)
        }
        $rows.Add(($cells -join "|"))
    }
    # The comma keeps an empty list a list: unrolled, it would arrive as $null.
    return ,$rows
}

function Get-Column {
    # Every cell of a one-column named area, empty ones included, in order.
    param($Book, [string]$Name)
    $data = (Get-Area -Book $Book -Name $Name).Value2
    $out = New-Object System.Collections.Generic.List[string]
    if (-not ($data -is [System.Array])) {
        $out.Add((ConvertTo-EvidenceText $data))
        return ,$out
    }
    for ($r = $data.GetLowerBound(0); $r -le $data.GetUpperBound(0); $r++) {
        $out.Add((ConvertTo-EvidenceText $data.GetValue($r, $data.GetLowerBound(1))))
    }
    return ,$out
}

function Wait-Macros {
    # The call succeeds only when the workbook's macros may run, so the operator
    # is asked to enable content instead of any macro setting being changed.
    param($App, [string]$Qualified)
    while ($true) {
        try {
            [void]$App.Run($Qualified + "TEST_Harness.TEST_Summary")
            return
        }
        catch {
            Write-Host ""
            Write-Host "The workbook's macros are disabled. In the Excel window, click Enable Content in"
            Write-Host "the message bar (or File > Info > Enable Content), then press Enter here."
            $answer = Read-Host "Press Enter to try again, or type Q to stop"
            if ($answer -match '^\s*q') {
                throw "Stopped by the operator: the macros were not enabled."
            }
        }
    }
}

function Invoke-Compile {
    # Debug > Compile VBAProject acts on the VBE's active project and is greyed
    # out once that project is compiled. Returns PASS or FAIL.
    param($App, $Book)
    $vbe = $App.VBE
    $vbe.MainWindow.Visible = $true
    try {
        $vbe.ActiveVBProject = $Book.VBProject
    }
    catch {
        Write-Session ("The workbook's project could not be made the active one: " + $_.Exception.Message)
    }
    $control = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $control) {
        Write-Session "Debug > Compile VBAProject was not found in the VBE"
        return "FAIL"
    }
    if ($control.Enabled) {
        $control.Execute()
    }
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        Start-Sleep -Milliseconds 500
        if (-not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled) {
            $vbe.MainWindow.Visible = $false
            return "PASS"
        }
    }
    Write-Host ""
    Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
    Write-Host "Debug > Compile VBAProject yourself. Type Y if it completes with no error and"
    Write-Host "the command is then greyed out; type N if it reports a compile error."
    $answer = Read-Host "Compiled with no error? (Y/N)"
    $vbe.MainWindow.Visible = $false
    if ($answer -match '^\s*y') {
        Write-Session "Compiled with no error, as the operator saw in the VBE"
        return "PASS"
    }
    return "FAIL"
}

function Write-HarnessRun {
    # One harness entry point and what Checks shows after it.
    param($App, $Book, [string]$Qualified, [string]$Log, [string]$Entry)
    Add-Evidence $Log "RUN" $Entry
    Add-Evidence $Log "STATE_BEFORE" (Get-ExcelState $App)
    try {
        [void]$App.Run($Qualified + $Entry)
    }
    catch {
        Add-Evidence $Log "ERROR" $_.Exception.Message
    }
    Add-Evidence $Log "STATE_AFTER" (Get-ExcelState $App)
    Add-Evidence $Log "STATUS" (Get-Area -Book $Book -Name "Checks_Status").Value2
    foreach ($value in (Get-Column -Book $Book -Name "Checks_Info")) {
        Add-Evidence $Log "INFO" $value
    }
    foreach ($row in (Get-Rows -Book $Book -Name "Checks_Cases")) {
        Add-Evidence $Log "CASE" $row
    }
    Add-Evidence $Log "ASSERTION_ROWS" (Get-Rows -Book $Book -Name "Checks_Assertions").Count
    Write-Session ("  {0}: {1}" -f $Entry, (Get-Area -Book $Book -Name "Checks_Status").Value2)
}

function Set-Setting {
    param($Book, [string]$Name, $Value)
    $area = Get-Area -Book $Book -Name $Name
    if ($Value -is [datetime]) {
        $area.Value = $Value
    }
    else {
        $area.Value2 = $Value
    }
}

function Write-ImportScenario {
    # Settings as a user would leave them (or paste them), the Import data
    # button's macro, and what DataSummary shows after it.
    param($App, $Book, [string]$Qualified, [string]$Log, [string]$Name, [System.Collections.Specialized.OrderedDictionary]$Settings)
    Add-Evidence $Log "SCENARIO" $Name
    foreach ($key in $Settings.Keys) {
        Set-Setting -Book $Book -Name $key -Value $Settings[$key]
        Add-Evidence $Log "SET" ("{0}|{1}" -f $key, (ConvertTo-EvidenceText $Settings[$key]))
    }
    Add-Evidence $Log "STATE_BEFORE" (Get-ExcelState $App)
    try {
        [void]$App.Run($Qualified + "WB_Import.WB_ImportData")
    }
    catch {
        Add-Evidence $Log "ERROR" $_.Exception.Message
    }
    Add-Evidence $Log "STATE_AFTER" (Get-ExcelState $App)
    Add-Evidence $Log "STATUS" (Get-Area -Book $Book -Name "DataSummary_Status").Value2
    foreach ($value in (Get-Column -Book $Book -Name "DataSummary_Info")) {
        Add-Evidence $Log "INFO" $value
    }
    Add-Evidence $Log "TOTAL_ROWS" (Get-Rows -Book $Book -Name "DataSummary_Totals").Count
    Add-Evidence $Log "OUTCOME_ROWS" (Get-Rows -Book $Book -Name "DataSummary_Outcomes").Count
    foreach ($row in (Get-Rows -Book $Book -Name "DataSummary_Findings")) {
        Add-Evidence $Log "FINDING" $row
    }
    foreach ($row in (Get-Rows -Book $Book -Name "DataSummary_Lineage")) {
        Add-Evidence $Log "LINEAGE" $row
    }
    Write-Session ("  import {0}: {1}" -f $Name, (Get-Area -Book $Book -Name "DataSummary_Status").Value2)
}

function Invoke-Logged {
    # Run a command, write each output line and its exit code to the log.
    param([string]$Log, [string]$Label, [scriptblock]$Command)
    Add-Evidence $Log "COMMAND" $Label
    # Windows PowerShell turns a native command's stderr line into an error
    # record, which "Stop" would make terminating: the output is only recorded.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & $Command 2>&1
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previous
    }
    foreach ($line in $output) {
        Add-Evidence $Log "OUTPUT" ([string]$line)
    }
    Add-Evidence $Log "EXIT" $code
    Write-Session ("  {0}: exit code {1}" -f $Label, $code)
    return $code
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

# ----------------------------------------------------------------------------
# Candidate and bundle
# ----------------------------------------------------------------------------
$sha = (& git -C $RepoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit of $RepoRoot."
}
$dirty = & git -C $RepoRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0 -or -not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    throw "The working tree has changes. Certify a clean checkout: git status must show none."
}
if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ("irrbb-evidence-" + $sha.Substring(0, 7))
}
if (Test-Path -LiteralPath $BundleDir) {
    if (@(Get-ChildItem -LiteralPath $BundleDir -Force).Count -gt 0) {
        throw "The bundle folder is not empty: $BundleDir. A bundle is never overwritten."
    }
}
else {
    New-Item -ItemType Directory -Path $BundleDir | Out-Null
}
$BundleDir = (Resolve-Path -LiteralPath $BundleDir).Path
if ($BundleDir.StartsWith($RepoRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The bundle folder must be outside the checkout: $BundleDir"
}
$script:sessionLog = Join-Path $BundleDir "session.log"
$startedAt = Get-Timestamp
Write-Session "K-IRRBB Excel evidence session for candidate $sha"
Write-Session "Started at $startedAt; bundle $BundleDir"

# ----------------------------------------------------------------------------
# Environment, as found
# ----------------------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$osName = ($osInfo.Caption -replace '^Microsoft\s+', '').Trim()
$osText = "{0} (version {1}, build {2})" -f $osName, $osInfo.Version, $osInfo.BuildNumber
$osArchitecture = switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { "x64" }
    "ARM64" { "ARM64" }
    "x86" { "x86" }
    default { $env:PROCESSOR_ARCHITECTURE }
}
$clickToRun = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$platform = Get-RegistryValue -Path $clickToRun -Name "Platform"
$officeBitness = switch ($platform) {
    "x64" { "64-bit" }
    "x86" { "32-bit" }
    default { "unknown" }
}
$versionToReport = Get-RegistryValue -Path $clickToRun -Name "VersionToReport"
Write-Session "Windows: $osText; architecture $osArchitecture"

$pythonVersion = & $Python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Python was not found as '$Python'. Install Python 3.11 or later, or pass -Python."
}
Write-Session "Python: $pythonVersion"

$session = [ordered]@{
    operator           = $Operator
    started_at         = $startedAt
    finished_at        = $null
    excel_version      = "unknown"
    excel_build        = "unknown"
    office_bitness     = $officeBitness
    os                 = $osText
    os_architecture    = $osArchitecture
    references         = @()
    macro_policy       = "unknown"
    vba_project_access = "unknown"
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    workbook_closed    = $false
}

$excel = New-Excel -Visible $false -Security $msoAutomationSecurityForceDisable
$officeVersion = [string]$excel.Version
$session.excel_version = $officeVersion
if ($null -ne $versionToReport) {
    $session.excel_build = "Version $versionToReport (build $($excel.Build))"
}
else {
    $session.excel_build = "Version $officeVersion (build $($excel.Build))"
}
Close-Excel -App $excel -Book $null
$session.macro_policy = Get-MacroPolicy -OfficeVersion $officeVersion
$session.vba_project_access = Get-VbaProjectAccess -OfficeVersion $officeVersion
Write-Session "Excel: $($session.excel_build), $officeBitness"
Write-Session "Macro setting: $($session.macro_policy); VBA project access: $($session.vba_project_access)"
if ($session.vba_project_access -ne "On") {
    Write-Session ("Stopped: 'Trust access to the VBA project object model' is off. The compile, the " +
                   "synchronizer and the export need it; turn it on yourself in Excel > Options > " +
                   "Trust Center if your policy allows. This script never changes it.")
    exit 1
}

# ----------------------------------------------------------------------------
# Stage: the macro-free template (template.log)
# ----------------------------------------------------------------------------
Write-Session "Template: opening src\workbook\IRRBB_Template.xlsx read-only"
$log = New-EvidenceLog "template.log"
Add-Evidence $log "FILE" "src/workbook/IRRBB_Template.xlsx"
$excel = $null
$book = $null
try {
    $excel = New-Excel -Visible $false -Security $msoAutomationSecurityForceDisable
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($templatePath, 0, $true)
    Add-Evidence $log "OPENED" "True"
    Add-Evidence $log "CAPTION" $book.Windows.Item(1).Caption
    try { $null = $book.VBProject.Name } catch {}
    # A list's values are shown joined by the list separator of the Windows locale.
    Add-Evidence $log "LIST_SEPARATOR" $excel.International($xlListSeparator)
    foreach ($sheet in $book.Worksheets) {
        Add-Evidence $log "SHEET" ("{0}|{1}" -f $sheet.Name, $sheet.CodeName)
        Add-Evidence $log "FONT" ("{0}|{1}" -f $sheet.Name, $sheet.Range("A2").Font.Name)
        Add-Evidence $log "TAB" ("{0}|{1}" -f $sheet.Name, (ConvertTo-Rgb $sheet.Tab.Color))
        foreach ($ref in @("A1", "A2")) {
            $cell = $sheet.Range($ref)
            Add-Evidence $log "STRIP" ("{0}!{1}|{2}|{3}|{4}" -f $sheet.Name, $ref, $cell.Value2,
                                       (ConvertTo-Rgb $cell.Interior.Color), (ConvertTo-Rgb $cell.Font.Color))
        }
        $validated = $null
        try { $validated = $sheet.Cells.SpecialCells($xlCellTypeAllValidation) } catch {}
        if ($null -ne $validated) {
            foreach ($cell in $validated.Cells) {
                if ($cell.Validation.Type -eq $xlValidateList) {
                    Add-Evidence $log "DROPDOWN" ("{0}!{1}|{2}|{3}" -f $sheet.Name, $cell.Address($false, $false),
                                                  $cell.Validation.InCellDropdown, $cell.Validation.Formula1)
                }
            }
        }
    }
    foreach ($name in $book.Names) {
        if ($name.Name -notmatch '!') {
            Add-Evidence $log "NAME" ("{0}|{1}" -f $name.Name, $name.RefersTo)
        }
        elseif ($name.Name -match '!Print_(Titles|Area)$') {
            Add-Evidence $log "PRINT" ("{0}|{1}" -f $name.Name, $name.RefersTo)
        }
    }
    $probes = @(
        @("Config", "B8", "x"), @("Config", "B8", (Get-Date -Year 2025 -Month 1 -Day 31).Date),
        @("Config", "B11", "Maybe"), @("Config", "B11", "Yes"),
        @("Segments", "C6", "Maybe"), @("Segments", "C6", "No")
    )
    foreach ($probe in $probes) {
        $cell = $book.Worksheets.Item($probe[0]).Range($probe[1])
        if ($probe[2] -is [datetime]) { $cell.Value = $probe[2] } else { $cell.Value2 = $probe[2] }
        $outcome = if ($cell.Validation.Value) { "accepted" } else { "refused" }
        Add-Evidence $log "VALIDATION" ("{0}!{1}|{2}|{3}" -f $probe[0], $probe[1], (ConvertTo-EvidenceText $probe[2]), $outcome)
    }
    # The page layout as Excel saves it, read back from a copy in the bundle.
    $copyPath = Join-Path $BundleDir "template-copy.xlsx"
    $book.SaveCopyAs($copyPath)
    Add-Evidence $log "COPY" "template-copy.xlsx"
    foreach ($entry in (Read-SavedPageLayout -Path $copyPath)) {
        Add-Evidence $log "PAGE" $entry
    }
}
catch {
    if ($null -eq $book) { Add-Evidence $log "OPENED" "False" }
    Add-Evidence $log "ERROR" $_.Exception.Message
    Write-Session ("Template: " + $_.Exception.Message)
}
finally {
    Close-Excel -App $excel -Book $book
}

# ----------------------------------------------------------------------------
# Stages 1 to 5: build, compile, harness, failure path, import (development workbook)
# ----------------------------------------------------------------------------
$devPath = Join-Path $BundleDir "IRRBB_Dev.xlsm"
$log = New-EvidenceLog "build.log"
Write-Session "Build: tools\devbuild\build_dev_workbook.py"
$buildCode = Invoke-Logged -Log $log -Label "python tools/devbuild/build_dev_workbook.py IRRBB_Dev.xlsm" -Command {
    & $Python (Join-Path $RepoRoot "tools\devbuild\build_dev_workbook.py") $devPath
}
$excel = $null
$book = $null
if ($buildCode -eq 0 -and (Test-Path -LiteralPath $devPath)) {
    try {
        $excel = New-Excel -Visible $true -Security $msoAutomationSecurityLow
        $book = $excel.Workbooks.Open($devPath)
        $qualified = "'" + $book.Name + "'!"
        Wait-Macros -App $excel -Qualified $qualified
        Write-Session "Macros enabled"
        $session.references = @(foreach ($reference in $book.VBProject.References) { [string]$reference.Description })
        Write-Session ("References: " + ($session.references -join "; "))
        $compiled = Invoke-Compile -App $excel -Book $book
        $session.compile = [ordered]@{ status = $compiled; detail = "Debug > Compile VBAProject on IRRBB_Dev.xlsm" }
        Write-Session "Compile: $compiled"
        if ($compiled -eq "PASS") {
            $log = New-EvidenceLog "harness.log"
            Write-HarnessRun -App $excel -Book $book -Qualified $qualified -Log $log -Entry "TEST_Harness.RunTests"

            $log = New-EvidenceLog "failure.log"
            foreach ($entry in @("TEST_Harness.TEST_RunTestsWithInjectedFailures", "TEST_Harness.RunTests",
                                 "TEST_Harness.TEST_RunTestsWithWriteFault", "TEST_Harness.RunTests")) {
                Write-HarnessRun -App $excel -Book $book -Qualified $qualified -Log $log -Entry $entry
            }

            $log = New-EvidenceLog "import.log"
            $base = [ordered]@{
                Config_AccountsFile      = (Join-Path $fixtures "accounts.csv")
                Config_RatesFile         = (Join-Path $fixtures "market_rates.csv")
                Config_WindowStart       = (Get-Date -Year 2025 -Month 1 -Day 31).Date
                Config_WindowEnd         = (Get-Date -Year 2025 -Month 6 -Day 30).Date
                Config_RateSeries        = "EUR_1M"
                Config_ClosureAttested   = "Yes"
                Config_ControlTotalsFile = ""
            }
            $include = Get-Area -Book $book -Name "Segments_Include"
            for ($i = 1; $i -le $include.Rows.Count; $i++) { $include.Cells.Item($i, 1).Value2 = "Yes" }
            Write-ImportScenario $excel $book $qualified $log "accepted" $base
            Write-ImportScenario $excel $book $qualified $log "control-match" ([ordered]@{
                Config_ControlTotalsFile = (Join-Path $fixtures "control_totals.csv") })
            Write-ImportScenario $excel $book $qualified $log "control-differ" ([ordered]@{
                Config_ControlTotalsFile = (Join-Path $fixtures "control_totals_balance.csv") })
            # A value pasted over a validated cell: Excel's validation does not see it, VBA must.
            Write-ImportScenario $excel $book $qualified $log "pasted-window" ([ordered]@{
                Config_ControlTotalsFile = ""; Config_WindowStart = "2025-02-15" })
            Write-ImportScenario $excel $book $qualified $log "pasted-attestation" ([ordered]@{
                Config_WindowStart = (Get-Date -Year 2025 -Month 1 -Day 31).Date; Config_ClosureAttested = "Maybe" })
            # An account file another program holds open for writing: the import must
            # refuse it when it opens it (error 70), and the rerun must then succeed.
            $lockedPath = Join-Path $BundleDir "locked_accounts.csv"
            Copy-Item -LiteralPath (Join-Path $fixtures "accounts.csv") -Destination $lockedPath
            $holder = [System.IO.File]::Open($lockedPath, [System.IO.FileMode]::Open,
                                             [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
            try {
                Write-ImportScenario $excel $book $qualified $log "locked-file" ([ordered]@{
                    Config_ClosureAttested = "Yes"; Config_AccountsFile = $lockedPath })
            }
            finally {
                $holder.Dispose()
            }
            Write-ImportScenario $excel $book $qualified $log "rerun" ([ordered]@{
                Config_AccountsFile = (Join-Path $fixtures "accounts.csv") })

            $log = New-EvidenceLog "buttons.log"
            Write-Buttons -App $excel -Book $book -Qualified $qualified -Log $log

            # The import of a large synthetic panel, timed (#71). The time is a
            # measurement, not a pass mark.
            $timingLog = New-EvidenceLog "timing.log"
            $panelDir = Join-Path $BundleDir "large_panel"
            Write-Session "Timing: a large synthetic panel (about 47 MB); the import can take several minutes"
            $panelCode = Invoke-Logged -Log $timingLog -Label "python tools/devbuild/make_large_panel.py" -Command {
                & $Python (Join-Path $RepoRoot "tools\devbuild\make_large_panel.py") --out $panelDir
            }
            if ($panelCode -eq 0) {
                $panelPath = Join-Path $panelDir "accounts_large.csv"
                Add-Evidence $timingLog "FILE_BYTES" (Get-Item -LiteralPath $panelPath).Length
                $clock = [System.Diagnostics.Stopwatch]::StartNew()
                Write-ImportScenario $excel $book $qualified $timingLog "large-panel" ([ordered]@{
                    Config_AccountsFile = $panelPath
                    Config_RatesFile    = (Join-Path $panelDir "market_rates_large.csv")
                    Config_WindowStart  = (Get-Date -Year 2024 -Month 1 -Day 31).Date
                    Config_WindowEnd    = (Get-Date -Year 2025 -Month 4 -Day 30).Date
                })
                $clock.Stop()
                $seconds = [math]::Round($clock.Elapsed.TotalSeconds, 1)
                Add-Evidence $timingLog "SECONDS" $seconds.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                Write-Session ("  large panel: {0:N1} MB imported in {1} s" -f ((Get-Item -LiteralPath $panelPath).Length / 1MB),
                               $seconds.ToString([System.Globalization.CultureInfo]::InvariantCulture))
            }
        }
    }
    catch {
        Write-Session ("Stopped in the development workbook: " + $_.Exception.Message)
    }
    finally {
        if ($null -ne $book) {
            try {
                $book.Close($false)
                $session.workbook_closed = $true
                Write-Session "Development workbook closed without saving"
            }
            catch {
                Write-Session ("The development workbook could not be closed: " + $_.Exception.Message)
            }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
            $book = $null
        }
        Close-Excel -App $excel -Book $null
    }
}
else {
    Write-Session "The development workbook was not built; stages 2 to 5 not run"
}

# ----------------------------------------------------------------------------
# Stage 6 and the synchronizer: synchronize, round trip, deliberate failure
# ----------------------------------------------------------------------------
$syncPath = Join-Path $BundleDir "IRRBB_Sync.xlsm"
$log = New-EvidenceLog "sync.log"
Write-Session "Synchronizer: the template saved as IRRBB_Sync.xlsm, then Sync-IRRBB-VBA.ps1 -Mode Development"
$excel = $null
$book = $null
try {
    $excel = New-Excel -Visible $false -Security $msoAutomationSecurityForceDisable
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($templatePath, 0, $true)
    $book.SaveAs($syncPath, $xlOpenXMLWorkbookMacroEnabled)
}
catch {
    Add-Evidence $log "ERROR" ("Saving the template as .xlsm: " + $_.Exception.Message)
}
finally {
    Close-Excel -App $excel -Book $book
}
$syncScript = Join-Path $RepoRoot "tools\Sync-IRRBB-VBA.ps1"
$syncCode = Invoke-Logged -Log $log -Label "Sync-IRRBB-VBA.ps1 -Mode Development -IncludeSheets -NoBackup" -Command {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -WorkbookPath $syncPath `
        -Mode Development -IncludeSheets -NoBackup -RepoRoot $RepoRoot
}
$excel = $null
$book = $null
if ($syncCode -eq 0) {
    try {
        $excel = New-Excel -Visible $true -Security $msoAutomationSecurityLow
        $book = $excel.Workbooks.Open($syncPath)
        $qualified = "'" + $book.Name + "'!"
        Wait-Macros -App $excel -Qualified $qualified
        # The synchronizer imports code only: the commit is entered as INSTALLATION.md says.
        Set-Setting -Book $book -Name "Readme_SourceCommit" -Value $sha
        $book.Save()
        $compiled = Invoke-Compile -App $excel -Book $book
        Add-Evidence $log "COMPILE" $compiled
        if ($compiled -eq "PASS") {
            Write-HarnessRun -App $excel -Book $book -Qualified $qualified -Log $log -Entry "TEST_Harness.RunTests"
        }
    }
    catch {
        Add-Evidence $log "ERROR" $_.Exception.Message
        Write-Session ("Stopped in the synchronized workbook: " + $_.Exception.Message)
    }
    finally {
        Close-Excel -App $excel -Book $book
    }
}

$log = New-EvidenceLog "roundtrip.log"
$exportDir = Join-Path $BundleDir "export"
Write-Session "Round trip: Export-IRRBB-VBA.ps1 and compare_vba_export.py"
$exportCode = Invoke-Logged -Log $log -Label "Export-IRRBB-VBA.ps1 IRRBB_Sync.xlsm export" -Command {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "tools\Export-IRRBB-VBA.ps1") `
        -WorkbookPath $syncPath -OutDir $exportDir
}
if ($exportCode -eq 0) {
    [void](Invoke-Logged -Log $log -Label "compare_vba_export.py --mode development" -Command {
        & $Python (Join-Path $RepoRoot "tools\compare_vba_export.py") --root $RepoRoot --export $exportDir `
            --mode development --candidate-sha $sha --workbook $syncPath
    })
}

$log = New-EvidenceLog "sync-failure.log"
Write-Session "Deliberate failure: the synchronizer on IRRBB_Sync.xlsm while another Excel holds it open"
if (Test-Path -LiteralPath $syncPath) {
    Add-Evidence $log "HASH_BEFORE" (Get-FileSha256 $syncPath)
    Add-Evidence $log "MACRO_POLICY_BEFORE" (Get-MacroPolicy -OfficeVersion $officeVersion)
    Add-Evidence $log "VBA_PROJECT_ACCESS_BEFORE" (Get-VbaProjectAccess -OfficeVersion $officeVersion)
    $excel = $null
    $book = $null
    try {
        $excel = New-Excel -Visible $false -Security $msoAutomationSecurityForceDisable
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Open($syncPath)
        [void](Invoke-Logged -Log $log -Label "Sync-IRRBB-VBA.ps1 -Mode Development -IncludeSheets -NoBackup" -Command {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -WorkbookPath $syncPath `
                -Mode Development -IncludeSheets -NoBackup -RepoRoot $RepoRoot
        })
    }
    catch {
        Add-Evidence $log "ERROR" $_.Exception.Message
    }
    finally {
        Close-Excel -App $excel -Book $book
    }
    Add-Evidence $log "HASH_AFTER" (Get-FileSha256 $syncPath)
    Add-Evidence $log "MACRO_POLICY_AFTER" (Get-MacroPolicy -OfficeVersion $officeVersion)
    Add-Evidence $log "VBA_PROJECT_ACCESS_AFTER" (Get-VbaProjectAccess -OfficeVersion $officeVersion)
}
else {
    Add-Evidence $log "ERROR" "IRRBB_Sync.xlsm does not exist"
}

# ----------------------------------------------------------------------------
# Record and validate
# ----------------------------------------------------------------------------
$session.finished_at = Get-Timestamp
Write-Session "Finished at $($session.finished_at)"
[System.IO.File]::WriteAllText((Join-Path $BundleDir "session.json"),
    ($session | ConvertTo-Json -Depth 4), $utf8NoBom)
Write-Host ""
& $Python (Join-Path $RepoRoot "tools\write_excel_evidence.py") --root $RepoRoot --candidate-sha $sha --bundle $BundleDir
$code = $LASTEXITCODE
Write-Host ""
Write-Host "Bundle: $BundleDir"
if ($code -eq 0) {
    Write-Host "RESULT=PASS: the bundle validates for candidate $sha."
    exit 0
}
Write-Host "RESULT=FAIL: see the stages above and session.log. Keep the bundle as it is."
exit $code
