# Genera venv\ con dependencias (ejecutar en Windows con Internet, una vez).
# Uso: powershell -ExecutionPolicy Bypass -File .\build.ps1
$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
$Req = Join-Path $Root "requirements-portable.txt"
if (-not (Test-Path $Req)) {
    $Req = Join-Path $Root "requirements.txt"
}

function Find-Python312 {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            $exe = & py -3.12 -c "import sys; print(sys.executable)" 2>$null
            if ($exe -and (Test-Path $exe)) { return $exe.Trim() }
        } catch {}
        try {
            $exe = & py -3 -c "import sys; print(sys.executable)" 2>$null
            if ($exe -and (Test-Path $exe)) { return $exe.Trim() }
        } catch {}
    }
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) {
        $ver = & $cmd.Source -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
        if ([version]$ver -ge [version]"3.12") { return $cmd.Source }
    }
    return $null
}

$Py = Find-Python312
if (-not $Py) {
    Write-Host "[ERROR] Se necesita Python 3.12+ en PATH (python o py -3.12)." -ForegroundColor Red
    Write-Host "Instale desde python.org o use el zip completo dist\HorasHombre (Python embebido)."
    exit 1
}

Write-Host "Python: $Py"
$Venv = Join-Path $Root "venv"
if (Test-Path $Venv) {
    Write-Host "Eliminando venv anterior..."
    Remove-Item -Recurse -Force $Venv
}

& $Py -m venv $Venv
$Pip = Join-Path $Venv "Scripts\pip.exe"
& $Pip install --upgrade pip
& $Pip install -r $Req

$Data = Join-Path $Root "data"
if (-not (Test-Path $Data)) { New-Item -ItemType Directory -Path $Data | Out-Null }

Write-Host ""
Write-Host "Listo. Ejecute HorasHombre.bat" -ForegroundColor Green
Write-Host "Para llevar a otra PC sin Internet: copie TODA esta carpeta (incluido venv\)."
Write-Host "Si la otra PC es distinta, conviene ejecutar build.ps1 alli tambien."