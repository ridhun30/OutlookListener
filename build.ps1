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

$outlookInterop = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook" -Recurse -Filter "Microsoft.Office.Interop.Outlook.dll" -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $outlookInterop) {
    throw "Microsoft.Office.Interop.Outlook.dll not found in the GAC - is Outlook (with primary interop assemblies) installed?"
}

$netDir = Split-Path -Parent $csc
$sysWebExtensions = Join-Path $netDir "System.Web.Extensions.dll"
$sysWindowsForms = Join-Path $netDir "System.Windows.Forms.dll"
$sysDrawing = Join-Path $netDir "System.Drawing.dll"
$sysDll = Join-Path $netDir "System.dll"

if (-not (Test-Path $sysWebExtensions)) {
    throw "System.Web.Extensions.dll not found under $netDir - needed for JSON (de)serialization."
}

# Stop any running instances first - csc can't overwrite a locked exe.
Get-Process -Name "OutlookListener","RulesDashboard" -ErrorAction SilentlyContinue | Stop-Process -Force

# --- OutlookListener.exe (listener) ---
$listenerExe = Join-Path $root "OutlookListener.exe"
$listenerRefs = @($sysDll, $sysWindowsForms, $sysWebExtensions, $outlookInterop) -join ","
$listenerSrc = @(
    (Join-Path $root "OutlookListener.cs"),
    (Join-Path $root "RulesModel.cs"),
    (Join-Path $root "RulesEngine.cs")
)

& $csc /nologo /target:exe /out:"$listenerExe" /reference:$listenerRefs $listenerSrc
if ($LASTEXITCODE -ne 0) {
    throw "OutlookListener.exe compilation failed (exit code $LASTEXITCODE)."
}
Write-Output "Built: $listenerExe"

# --- RulesDashboard.exe (GUI) ---
$dashboardExe = Join-Path $root "RulesDashboard.exe"
$dashboardRefs = @($sysDll, $sysWindowsForms, $sysDrawing, $sysWebExtensions) -join ","
$dashboardSrc = @(
    (Join-Path $root "RulesDashboard.cs"),
    (Join-Path $root "RuleEditorForm.cs"),
    (Join-Path $root "TkeTheme.cs"),
    (Join-Path $root "RulesModel.cs")
)

& $csc /nologo /target:winexe /out:"$dashboardExe" /reference:$dashboardRefs $dashboardSrc
if ($LASTEXITCODE -ne 0) {
    throw "RulesDashboard.exe compilation failed (exit code $LASTEXITCODE)."
}
Write-Output "Built: $dashboardExe"
