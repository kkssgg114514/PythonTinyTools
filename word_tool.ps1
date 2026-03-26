param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ToolArgs
)

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$script = Join-Path $PSScriptRoot "word_tool.py"

if (-not (Test-Path $python)) {
    Write-Error "未找到虚拟环境 Python: $python"
    exit 1
}

if (-not (Test-Path $script)) {
    Write-Error "未找到脚本: $script"
    exit 1
}

& $python $script @ToolArgs
exit $LASTEXITCODE
