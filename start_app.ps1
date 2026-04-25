$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$appPath = Join-Path $projectRoot "app.py"

if (-not (Test-Path $venvPython)) {
    Write-Error "未找到虚拟环境 Python：$venvPython"
}

if (-not (Test-Path $appPath)) {
    Write-Error "未找到 app.py：$appPath"
}

Set-Location $projectRoot
& $venvPython -m streamlit run $appPath
