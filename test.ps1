# Excel合并工具测试脚本
# 使用方法: .\test.ps1

Write-Host "启动Excel表格合并工具..." -ForegroundColor Cyan
Write-Host ""

# 检查虚拟环境是否存在
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "错误: 未找到虚拟环境！" -ForegroundColor Red
    Write-Host "请先运行: python311 -m venv .venv" -ForegroundColor Yellow
    exit 1
}

# 激活虚拟环境并运行主程序
$python = ".\.venv\Scripts\python.exe"
& $python excel_merger_gui.py
