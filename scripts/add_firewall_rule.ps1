# 允許區網連入 Hymn Search 遙控 API（TCP 8765，程式預設 port）
#
# 執行方式（擇一）：
#   A) 右鍵 add_firewall_rule.cmd → 以系統管理員身分執行（推薦）
#   B) 管理員 PowerShell：
#        powershell -ExecutionPolicy Bypass -File .\add_firewall_rule.ps1
#   C) 管理員 CMD 一行：
#        netsh advfirewall firewall add rule name="Hymn Search Remote" dir=in action=allow protocol=TCP localport=8765 profile=private,domain,public enable=yes

param(
    [int]$Port = 8765
)

$ruleName = "Hymn Search Remote"

$existing = netsh advfirewall firewall show rule name="$ruleName" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "規則已存在：$ruleName"
    netsh advfirewall firewall show rule name="$ruleName"
    exit 0
}

netsh advfirewall firewall add rule `
    name="$ruleName" `
    dir=in `
    action=allow `
    protocol=TCP `
    localport=$Port `
    profile=private,domain,public `
    enable=yes

if ($LASTEXITCODE -eq 0) {
    Write-Host "已新增防火牆規則：$ruleName (TCP $Port, 私人/網域網路)"
    Write-Host "防火牆可保持開啟；其他電腦請用 http://<你部機IP>:$Port/api/health 測試"
} else {
    Write-Host "失敗：請確認以系統管理員身分執行。" -ForegroundColor Red
    exit 1
}
