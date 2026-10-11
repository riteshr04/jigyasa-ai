$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$python = Join-Path $projectRoot ".venv312\Scripts\python.exe"
if (-not (Test-Path $python)) {
    py -3.12 -m venv (Join-Path $projectRoot ".venv312")
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the Python 3.12 environment. Install Python 3.12 and run this script again."
    }
}

& $python -c "import fastapi, uvicorn, pypdf, pptx, docx, dotenv" 2>$null
if ($LASTEXITCODE -ne 0) {
    & $python -m pip install -r (Join-Path $projectRoot "requirements-backend.txt")
    if ($LASTEXITCODE -ne 0) {
        throw "Could not install Jigyasa AI backend dependencies."
    }
}

Write-Host "Jigyasa AI is starting at http://127.0.0.1:8123"
Write-Host "Keep this window open while using the site. Press Ctrl+C to stop."
& $python -m uvicorn backend.main:app --host 127.0.0.1 --port 8123
