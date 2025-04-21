# 修复GitHub账号选择问题
# 此脚本将帮助您解决每次git push都需要选择账号的问题

# 1. 首先显示当前配置
Write-Host "当前Git配置:" -ForegroundColor Cyan
git config --list | Select-String -Pattern "credential|user"

# 2. 禁用Windows凭据管理器
Write-Host "`n正在配置Git以使用特定账号..." -ForegroundColor Green

# 设置全局配置为使用store而非manager
git config --global credential.helper store
git config --global --unset credential.helper manager
git config --global --unset credential.manager

# 3. 为当前仓库设置特定配置
$repo_url = git remote get-url origin
Write-Host "`n为仓库 $repo_url 设置固定账号" -ForegroundColor Green

# 4. 清除Windows凭据管理器中的GitHub凭据
Write-Host "`n您可能需要清除Windows凭据管理器中的GitHub凭据:" -ForegroundColor Yellow
Write-Host "1. 打开控制面板 -> 用户账户 -> 凭据管理器" -ForegroundColor Yellow
Write-Host "2. 在Windows凭据下，查找并删除所有与GitHub相关的凭据" -ForegroundColor Yellow
Write-Host "3. 然后重新进行一次git push，并输入您要默认使用的账号和密码" -ForegroundColor Yellow

# 5. 设置默认账号
$username = Read-Host "`n请输入您希望默认使用的GitHub用户名"
$email = Read-Host "请输入对应的GitHub邮箱"

git config --global user.name $username
git config --global user.email $email

# 6. 针对特定仓库的URL设置凭据
$credential_section = "[credential `"$repo_url`"]"
$helper_line = "	helper = store"

# 检查全局配置文件并添加特定配置
$configPath = "$env:USERPROFILE\.gitconfig"
$configContent = Get-Content $configPath -Raw -ErrorAction SilentlyContinue

if ($configContent -notmatch [regex]::Escape($credential_section)) {
    Write-Host "`n正在添加特定仓库凭据配置..." -ForegroundColor Green
    Add-Content -Path $configPath -Value "`n$credential_section`n$helper_line"
}

Write-Host "`n配置已更新。执行以下步骤完成设置:" -ForegroundColor Green
Write-Host "1. 删除Windows凭据管理器中的GitHub凭据" -ForegroundColor Cyan
Write-Host "2. 执行一次git push操作" -ForegroundColor Cyan
Write-Host "3. 输入您的用户名($username)和密码" -ForegroundColor Cyan
Write-Host "`n完成这些步骤后，Git应该会记住您的选择，不再显示账号选择对话框" -ForegroundColor Green 