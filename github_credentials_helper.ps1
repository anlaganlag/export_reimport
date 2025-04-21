# GitHub账号设置助手
# 此脚本帮助您设置Git以记住GitHub凭据，避免每次push时都需要选择账号

# 设置Git凭据存储
Write-Host "正在配置Git凭据存储..." -ForegroundColor Green
git config --global credential.helper store

# 设置用户名
$username = Read-Host "请输入您的GitHub用户名 (例如: anlaganlag)"
git config --global user.name $username
Write-Host "用户名已设置为: $username" -ForegroundColor Green

# 设置邮箱
$email = Read-Host "请输入您的GitHub邮箱"
git config --global user.email $email
Write-Host "邮箱已设置为: $email" -ForegroundColor Green

# 指导用户完成一次推送以保存凭据
Write-Host "`n现在进行一次git push操作，输入您的GitHub密码或个人访问令牌" -ForegroundColor Yellow
Write-Host "这将是最后一次需要输入密码，之后Git将记住您的凭据" -ForegroundColor Yellow
Write-Host "`n执行以下命令:" -ForegroundColor Cyan
Write-Host "git push" -ForegroundColor Cyan

Write-Host "`n设置完成！下次push时将不再需要选择账号" -ForegroundColor Green 