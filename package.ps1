# Excel合并工具打包脚本
# 使用方法: .\package.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Excel表格合并工具 - 自动打包脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查虚拟环境是否存在
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "错误: 未找到虚拟环境！" -ForegroundColor Red
    Write-Host "请先运行: python311 -m venv .venv" -ForegroundColor Yellow
    exit 1
}

Write-Host "[1/4] 激活虚拟环境..." -ForegroundColor Green
& .\.venv\Scripts\Activate.ps1

Write-Host "[2/4] 检查依赖..." -ForegroundColor Green
$python = ".\.venv\Scripts\python.exe"
& $python -m pip install --quiet openpyxl PyQt6 pyinstaller xlrd

Write-Host "[3/4] 清理旧的构建文件..." -ForegroundColor Green
if (Test-Path "package") {
    Remove-Item -Recurse -Force package
}
if (Test-Path "dist") {
    Remove-Item -Recurse -Force dist
}

Write-Host "[4/4] 开始打包..." -ForegroundColor Green
& $python -m PyInstaller --workpath package excel_merger.spec

Write-Host ""
if ($LASTEXITCODE -eq 0) {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "打包成功！" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "可执行文件位置: dist\Excel表格合并工具.exe" -ForegroundColor Cyan
    
    # 显示文件大小
    $exeFile = Get-Item "dist\Excel表格合并工具.exe"
    $sizeMB = [math]::Round($exeFile.Length / 1MB, 2)
    Write-Host "文件大小: $sizeMB MB" -ForegroundColor Cyan
} else {
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "打包失败！请检查错误信息" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    exit 1
}
