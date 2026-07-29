$ErrorActionPreference = 'Stop'
$appDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = 'C:\Users\Hp\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$env:OPENAI_API_KEY = [Environment]::GetEnvironmentVariable('OPENAI_API_KEY', 'User')

if (-not (Test-Path $python)) {
    Write-Host 'The bundled Codex Python runtime was not found. Open this app from Codex, or install Python with openpyxl, pypdf, and pdfplumber.' -ForegroundColor Red
    Read-Host 'Press Enter to close'
    exit 1
}

$env:BANK_PARSER_PORT = '8081'
Start-Process 'http://localhost:8081'
Set-Location $appDirectory
& $python run_current.py
