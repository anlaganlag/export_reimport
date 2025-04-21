# 清除Windows凭据管理器中的GitHub凭据
# 此脚本将帮助您自动清除Windows凭据管理器中保存的GitHub凭据

Write-Host "正在寻找并清除Windows凭据管理器中的GitHub凭据..." -ForegroundColor Cyan

# 使用cmdkey命令列出所有凭据
$credentials = cmdkey /list | Select-String -Pattern "github|git"

if ($credentials) {
    Write-Host "发现以下GitHub相关的凭据:" -ForegroundColor Yellow
    $credentials | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    
    $confirm = Read-Host "`n是否删除这些凭据? (y/n)"
    
    if ($confirm -eq "y") {
        # 为每个找到的GitHub凭据执行删除
        $credentials | ForEach-Object {
            $credLine = $_ -replace '[\r\n]', ''
            
            # 从命令行输出中提取Target值
            if ($credLine -match 'Target: (.+)') {
                $target = $matches[1]
                Write-Host "正在删除凭据: $target" -ForegroundColor Cyan
                cmdkey /delete:$target
            }
        }
        
        Write-Host "`n凭据已清除。现在您可以执行以下步骤:" -ForegroundColor Green
        Write-Host "1. 运行 '.\fix_github_account.ps1' 设置Git配置" -ForegroundColor Cyan
        Write-Host "2. 执行一次git push，使用您想默认的账号密码" -ForegroundColor Cyan
    }
    else {
        Write-Host "操作已取消。" -ForegroundColor Red
    }
}
else {
    Write-Host "未找到GitHub相关的凭据。" -ForegroundColor Green
    Write-Host "您可以继续运行 '.\fix_github_account.ps1' 来设置Git配置。" -ForegroundColor Cyan
}

# 如果使用控制面板方式删除凭据的指导
Write-Host "`n如果自动方式不起作用，您可以手动删除凭据:" -ForegroundColor Yellow
Write-Host "1. 打开控制面板 -> 用户账户 -> 凭据管理器" -ForegroundColor Yellow
Write-Host "2. 在Windows凭据下，查找并删除所有与GitHub相关的凭据" -ForegroundColor Yellow
Write-Host "3. 然后运行 '.\fix_github_account.ps1' 并重新进行git push" -ForegroundColor Yellow 