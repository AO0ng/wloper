# run.ps1 - 供 Windows 任务计划程序调用的入口脚本
# 用法: powershell.exe -ExecutionPolicy Bypass -File "C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work\run.ps1"

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logDir = Join-Path $scriptDir "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logFile = Join-Path $logDir "run-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').log"

function Write-Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line
}

try {
    Write-Log "=== 水文周报抓取任务开始 ==="
    Set-Location $scriptDir

    if (-not (Test-Path "node_modules")) {
        Write-Log "正在安装依赖..."
        cmd /c "npm install" 2>&1 | ForEach-Object { Write-Log $_ }
    }

    Write-Log "启动抓取脚本..."
    $output = cmd /c "node scraper.js" 2>&1
    $output | ForEach-Object { Write-Log $_ }

    Write-Log "=== 任务结束 ==="
} catch {
    Write-Log "错误: $_"
    exit 1
}
