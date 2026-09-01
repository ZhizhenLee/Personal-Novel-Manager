# 打包脚本：将 novel_manager.py 打包为单文件 通天录.exe
# 使用方法：在 PowerShell 中运行  ./build_exe.ps1
$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

Write-Host "==> 安装打包依赖（仅需一次）..." -ForegroundColor Cyan
python -m pip install --quiet pyinstaller

Write-Host "==> 开始打包 通天录.exe ..." -ForegroundColor Cyan
python -m PyInstaller `
    --onefile `
    --windowed `
    --noconfirm `
    --clean `
    --name "通天录" `
    --icon "app.ico" `
    --distpath "dist" `
    --workpath "build" `
    --specpath "." `
    novel_manager.py

if (Test-Path "dist\通天录.exe") {
    Write-Host ""
    Write-Host "打包成功！" -ForegroundColor Green
    Write-Host "输出: dist\通天录.exe" -ForegroundColor Green
    Write-Host "双下运行即可，无需安装，单文件无多余文件。" -ForegroundColor Green
} else {
    Write-Host "打包失败，未找到输出文件。" -ForegroundColor Red
    exit 1
}
