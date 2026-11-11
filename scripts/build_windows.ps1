param([switch]$SkipInstaller)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw '先按 BUILD_GUIDE.md 创建 Python 3.12 x64 虚拟环境。'
}
Set-Location -LiteralPath $projectRoot

# PyInstaller and pip write progress to stderr; with ErrorActionPreference=Stop
# PowerShell would turn that into a terminating NativeCommandError.
$ErrorActionPreference = 'Continue'

& $pythonExe -m pip check
if ($LASTEXITCODE -ne 0) { throw '依赖冲突，构建终止。' }

$buildId = Get-Date -Format 'yyyyMMdd-HHmmss'
$bundleRoot = Join-Path $projectRoot "build\bundle-$buildId"
$pythonBase = (& $pythonExe -c 'import sys; print(sys.base_prefix)').Trim()
$savedPath = $env:PATH
# Native dependency analysis must not pick identically named DLLs from other
# applications (e.g. Poppler's versioned ICU instead of the Windows one).
$env:PATH = "$pythonBase;$(Split-Path -Parent $pythonExe);$env:SystemRoot\System32;$env:SystemRoot"
& $pythonExe -m PyInstaller --noconfirm --onedir --windowed --name MultiTableSearch `
    --distpath $bundleRoot --workpath "build\pyinstaller-$buildId" --specpath build `
    --paths $projectRoot --hidden-import win32timezone --hidden-import pythoncom `
    --hidden-import pywintypes --hidden-import win32com.client `
    --exclude-module pytest --exclude-module pytestqt `
    --exclude-module tests --exclude-module scripts `
    app\launcher.py
$buildExit = $LASTEXITCODE
$env:PATH = $savedPath
if ($buildExit -ne 0) { throw 'PyInstaller 构建失败。' }

$bundle = Join-Path $bundleRoot 'MultiTableSearch'
$forbidden = @(Get-ChildItem -LiteralPath $bundle -File -Recurse | Where-Object {
    $_.Extension -match '^\.(xlsx|xlsm|xls|xlsb|db|sqlite|csv|tsv)$'
})
if ($forbidden.Count -gt 0) { throw '构建目录包含数据文件，拒绝生成安装包。' }
Set-Content -LiteralPath (Join-Path $projectRoot 'build\last-bundle-path.txt') -Value $bundle -Encoding utf8
if ($SkipInstaller) { Write-Output $bundle; exit 0 }

$compilerCandidates = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe'
)
$compiler = $compilerCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) { throw '缺少 Inno Setup 6，请安装后重试。' }

& $compiler "/DBundleDir=$bundle" scripts\create_installer.iss
if ($LASTEXITCODE -ne 0) { throw 'Inno Setup 构建失败。' }
Get-FileHash -LiteralPath 'dist\MultiTableSearch_Setup.exe' -Algorithm SHA256
