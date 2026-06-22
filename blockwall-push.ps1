# blockwall-push.ps1
# Run from C:\Users\VW\Desktop\blockwall
# One-shot: init + commit + push to GitHub

$ErrorActionPreference = 'Stop'
Set-Location "C:\Users\VW\Desktop\blockwall"

# 1) git init if needed
if (-not (Test-Path .git)) {
    git init
    git branch -M main
}

# 2) .gitignore — add the things that should NOT be in the repo
$gi = @'
venv/
__pycache__/
*.log
test.html
Invoke-Ssh.ps1
run_hermes*.sh
hermes_*.sh
hermes_prompt.txt
deploy_step*.sh
cleanup_host.sh
verify_app.sh
blockwall.env
app.log
.gitignore.tmp
start.bat
*.xlsx
'@
Add-Content -Path .gitignore -Value $gi -Encoding utf8

# 3) stage + commit
git add .
git status

$msg = "blockwall for render deploy"
git commit -m $msg

# 4) add remote if not present
$remote = git remote get-url origin 2>$null
if (-not $remote) {
    git remote add origin git@github.com:richodin/blockwall.git
    Write-Host "Added remote: git@github.com:richodin/blockwall.git" -ForegroundColor Cyan
} else {
    Write-Host "Remote already set: $remote" -ForegroundColor Yellow
}

# 5) push
git push -u origin main
