$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $out | Out-Null

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "Visual Studio Build Tools were not found. Install 'Build Tools for Visual Studio' with the C++ desktop workload, then run this script again."
}

$install = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $install) {
    throw "MSVC C++ tools were not found. Install the 'Desktop development with C++' workload."
}

$vcvars = Join-Path $install "VC\Auxiliary\Build\vcvars64.bat"
if (-not (Test-Path $vcvars)) {
    throw "Could not find vcvars64.bat at $vcvars"
}

$cmd = @"
call "$vcvars"
cd /d "$root"
rc /nologo /fo "$out\resource.res" resource.rc
cl /nologo /std:c++17 /EHsc /W4 /DUNICODE /D_UNICODE /MT /O2 AudioOptimizer.cpp "$out\resource.res" /link /SUBSYSTEM:WINDOWS /OUT:"$out\AudioOptimizer.exe" user32.lib gdi32.lib comctl32.lib ole32.lib propsys.lib shell32.lib advapi32.lib powrprof.lib
"@
$cmd | Set-Content -Encoding ASCII -Path (Join-Path $out "build.cmd")

$buildCmd = Join-Path $out "build.cmd"
cmd.exe /c "`"$buildCmd`""
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path (Join-Path $out "AudioOptimizer.exe"))) {
    throw "Build completed without producing dist\AudioOptimizer.exe"
}

Write-Host "Built $out\AudioOptimizer.exe"
