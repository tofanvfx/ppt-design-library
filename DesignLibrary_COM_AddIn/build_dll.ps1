# build_dll.ps1 - Compiles a C# COM Add-in and registers it with PowerPoint
$ErrorActionPreference = "Stop"
Write-Host "`n=== Design Library COM Add-in Build ===" -ForegroundColor Cyan

# Paths
$srcDir = Join-Path $PSScriptRoot "src"
$outDll = Join-Path $PSScriptRoot "DesignLibraryAddIn.dll"
$cscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$regasmPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

if (-not (Test-Path $cscPath)) {
    throw "C# Compiler (csc.exe) not found at $cscPath"
}

# Find Interop Assemblies in GAC
Write-Host "[1] Locating Office PIAs in GAC..." -ForegroundColor Yellow
$gacPaths = @(
    "C:\Windows\assembly\GAC_MSIL",
    "C:\Windows\assembly\GAC"
)

function Find-PIA($name) {
    foreach ($base in $gacPaths) {
        if (Test-Path $base) {
            $files = Get-ChildItem -Path $base -Recurse -Filter "$name.dll" -ErrorAction SilentlyContinue
            if ($files.Count -gt 0) {
                # Pick the latest version
                $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                return $latest.FullName
            }
        }
    }
    return $null
}

$extensibilityDll = Find-PIA "Extensibility"
$officeDll = Find-PIA "office"
$pptDll = Find-PIA "Microsoft.Office.Interop.PowerPoint"

if (-not $extensibilityDll) { throw "Extensibility.dll not found in GAC" }
if (-not $officeDll) { throw "office.dll not found in GAC" }
if (-not $pptDll) { throw "Microsoft.Office.Interop.PowerPoint.dll not found in GAC" }

Write-Host "    Found Extensibility: $extensibilityDll"
Write-Host "    Found Office: $officeDll"
Write-Host "    Found PowerPoint: $pptDll"

# Build command args
Write-Host "`n[2] Compiling C# source code..." -ForegroundColor Yellow
$csFiles = Get-ChildItem -Path $srcDir -Filter "*.cs" | Select-Object -ExpandProperty FullName
$csFilesList = $csFiles -join "`" `""
$references = "/reference:`"$extensibilityDll`",`"$officeDll`",`"$pptDll`",System.Windows.Forms.dll,System.Drawing.dll,Microsoft.VisualBasic.dll,System.Core.dll"

# Make sure Ribbon.xml is copied as well
$ribbonSrc = Join-Path $srcDir "Ribbon.xml"
$ribbonDst = Join-Path $PSScriptRoot "Ribbon.xml"
if (Test-Path $ribbonSrc) { Copy-Item $ribbonSrc $ribbonDst -Force }

# Execute csc.exe
$compilerArgs = "/target:library /platform:anycpu /out:`"$outDll`" $references `"$csFilesList`""
Write-Host "    Running csc.exe..."
$process = Start-Process -FilePath $cscPath -ArgumentList $compilerArgs -NoNewWindow -Wait -PassThru

if ($process.ExitCode -ne 0) {
    throw "Compilation failed with exit code $($process.ExitCode). Please check syntax errors."
}

if (-not (Test-Path $outDll)) {
    throw "Compilation finished but $outDll was not produced."
}
$dllSize = (Get-Item $outDll).Length
Write-Host "    SUCCESS: Compiled DesignLibrary.dll ($dllSize bytes)" -ForegroundColor Green

# Register COM Add-in (Per-User to avoid Admin requirement)
Write-Host "`n[3] Generating COM Add-in registry keys..." -ForegroundColor Yellow
$regFile = Join-Path $PSScriptRoot "DesignLibrary.reg"
if (Test-Path $regFile) { Remove-Item $regFile -Force }

Write-Host "    Running regasm.exe /regfile..."
$oldErrorAction = $ErrorActionPreference
$ErrorActionPreference = "Continue"
Start-Process -FilePath $regasmPath -ArgumentList "`"$outDll`" /codebase /regfile:`"$regFile`"" -NoNewWindow -Wait -PassThru | Out-Null
$ErrorActionPreference = $oldErrorAction

if (Test-Path $regFile) {
    Write-Host "    Modifying .reg file for Current User..."
    $text = [System.IO.File]::ReadAllText($regFile)
    $text = $text.Replace("HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes")
    [System.IO.File]::WriteAllText($regFile, $text, [System.Text.Encoding]::Unicode)

    Write-Host "    Importing registry keys..."
    Start-Process -FilePath "reg.exe" -ArgumentList "import `"$regFile`"" -NoNewWindow -Wait -PassThru | Out-Null
    Write-Host "    COM Registration successful." -ForegroundColor Green
}
else {
    throw "Failed to generate .reg file."
}

# Add Registry Keys for PowerPoint
Write-Host "`n[4] Writing PowerPoint Add-in Registry Keys..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\DesignLibraryAddIn.AddIn"

if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

Set-ItemProperty -Path $regPath -Name "Description" -Value "PowerPoint Design Library Add-in"
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Design Library"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 # 3 = Load at startup

Write-Host "    Registry keys written successfully." -ForegroundColor Green

Write-Host "`n=== SETUP COMPLETE ===" -ForegroundColor Green
Write-Host "Restart PowerPoint to see the 'Design Library' tab in the Ribbon." -ForegroundColor White
