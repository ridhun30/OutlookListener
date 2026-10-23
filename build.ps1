<#
build.ps1

Compiles two standalone exes using csc.exe (ships with .NET Framework - no
separate install needed):
  - OutlookListener.exe   - the background listener (console app)
  - RulesDashboard.exe    - GUI for adding/editing/deleting rules
Both share RulesModel.cs (rule data + JSON storage). Only OutlookListener.exe
also compiles RulesEngine.cs, since that's the only piece that touches live
Outlook COM objects.
#>

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$csc = Get-ChildItem -Path "$env:WINDIR\Microsoft.NET\Framework64" -Recurse -Filter "csc.exe" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $csc) {
    $csc = Get-ChildItem -Path "$env:WINDIR\Microsoft.NET\Framework" -Recurse -Filter "csc.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $csc) {
    throw "csc.exe not found under $env:WINDIR\Microsoft.NET - cannot compile."
}

Write-Output "Using compiler: $csc"

# OUTLOOK_INTEROP lets CI (no Office installed) point at the interop DLL from NuGet.
$outlookInterop = $env:OUTLOOK_INTEROP
if (-not $outlookInterop) {
    $outlookInterop = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook" -Recurse -Filter "Microsoft.Office.Interop.Outlook.dll" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $outlookInterop) {
    throw "Microsoft.Office.Interop.Outlook.dll not found in the GAC - is Outlook (with primary interop assemblies) installed?"
}

$netDir = Split-Path -Parent $csc
$sysWebExtensions = Join-Path $netDir "System.Web.Extensions.dll"
$sysWindowsForms = Join-Path $netDir "System.Windows.Forms.dll"
$sysDrawing = Join-Path $netDir "System.Drawing.dll"
$sysDll = Join-Path $netDir "System.dll"
$sysSecurity = Join-Path $netDir "System.Security.dll"
$sysCore = Join-Path $netDir "System.Core.dll"

if (-not (Test-Path $sysWebExtensions)) {
    throw "System.Web.Extensions.dll not found under $netDir - needed for JSON (de)serialization."
}

# Stop any running instances first - csc can't overwrite a locked exe.
Get-Process -Name "OutlookListener","RulesDashboard" -ErrorAction SilentlyContinue | Stop-Process -Force

$appIcon = Join-Path $root "OutlookListener.ico"
$iconArg = if (Test-Path $appIcon) { "/win32icon:$appIcon" } else { "" }

# TKE icon assets embedded into RulesDashboard.exe as resources (button icons) -
# sourced from Resources\Icon Library per the icons-only-from-there rule; copied
# locally into Icons\ so the built exe is self-contained at runtime.
$iconsDir = Join-Path $root "Icons"
$buttonIcons = @("create-folder-white.png", "blog-black.png", "bin-black.png", "life-cycle-black.png", "log-black.png", "document-black.png")
$resourceArgs = @()
foreach ($iconFile in $buttonIcons) {
    $iconPath = Join-Path $iconsDir $iconFile
    if (-not (Test-Path $iconPath)) {
        throw "Expected icon asset not found: $iconPath"
    }
    $resourceArgs += "/resource:$iconPath,$iconFile"
}

# --- OutlookListener.exe (listener) ---
$listenerExe = Join-Path $root "OutlookListener.exe"
$listenerRefs = @($sysDll, $sysCore, $sysWindowsForms, $sysDrawing, $sysWebExtensions, $sysSecurity, $outlookInterop) -join ","
$listenerSrc = @(
    (Join-Path $root "OutlookListener.cs"),
    (Join-Path $root "RulesModel.cs"),
    (Join-Path $root "RulesEngine.cs"),
    (Join-Path $root "WorkTrackSettings.cs"),
    (Join-Path $root "AskDetector.cs"),
    (Join-Path $root "WorkTrackConnector.cs"),
    (Join-Path $root "WorkTrackSync.cs"),
    (Join-Path $root "WorkTrackRules.cs"),
    (Join-Path $root "WorkTrackReminders.cs"),
    (Join-Path $root "TkeTheme.cs")
)

& $csc /nologo /target:exe /out:"$listenerExe" $iconArg /reference:$listenerRefs $listenerSrc
if ($LASTEXITCODE -ne 0) {
    throw "OutlookListener.exe compilation failed (exit code $LASTEXITCODE)."
}
Write-Output "Built: $listenerExe"

# --- RulesDashboard.exe (GUI) ---
$dashboardExe = Join-Path $root "RulesDashboard.exe"
$dashboardRefs = @($sysDll, $sysWindowsForms, $sysDrawing, $sysWebExtensions, $sysSecurity) -join ","
$dashboardSrc = @(
    (Join-Path $root "RulesDashboard.cs"),
    (Join-Path $root "RuleEditorForm.cs"),
    (Join-Path $root "WorkTrackForm.cs"),
    (Join-Path $root "TkeTheme.cs"),
    (Join-Path $root "RulesModel.cs"),
    (Join-Path $root "WorkTrackSettings.cs"),
    (Join-Path $root "AskDetector.cs"),
    (Join-Path $root "WorkTrackConnector.cs")
)

& $csc /nologo /target:winexe /out:"$dashboardExe" $iconArg $resourceArgs /reference:$dashboardRefs $dashboardSrc
if ($LASTEXITCODE -ne 0) {
    throw "RulesDashboard.exe compilation failed (exit code $LASTEXITCODE)."
}
Write-Output "Built: $dashboardExe"
