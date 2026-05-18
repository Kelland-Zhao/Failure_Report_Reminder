param(
    [string]$message = "update"
)

git add .
git commit -m $message
Write-Host "推送完成 / Push completed: $message"
