# Upload D:\TraeNewPaper11 to YUWENJIE10086/myPaper2612 (main).
# This script never force-pushes, never removes user files, and requires a confirmation.
$ErrorActionPreference = 'Stop'
$source = 'D:\TraeNewPaper11'
$remote = 'https://github.com/YUWENJIE10086/myPaper2612.git'
$expectedRemote = 'github.com/yuwenjie10086/mypaper2612'

function RunGit {
    param([string[]]$GitArguments)
    & git -C $source @GitArguments
    if ($LASTEXITCODE -ne 0) { throw "git $($GitArguments -join ' ') failed. Exit code: $LASTEXITCODE" }
}
function GetGitValue {
    param([string[]]$GitArgs)
    $value = & git -C $source @GitArgs 2>$null
    if ($LASTEXITCODE -eq 0) { return ($value | Out-String).Trim() }
    return ''
}
try {
    Write-Host ''
    Write-Host 'Local folder -> GitHub repository' -ForegroundColor Cyan
    Write-Host "Source: $source"
    Write-Host "Remote: $remote"
    Write-Host 'Branch: main'
    Write-Host ''
    if (-not (Test-Path -LiteralPath $source -PathType Container)) {
        throw "Cannot find $source. Confirm the D: drive and folder exist on this computer."
    }
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'Git for Windows is not installed or is not on PATH. Install from https://git-scm.com/downloads/win and rerun.'
    }

    # Before any remote upload, check obvious secrets and GitHub file size limit.
    $root = (Resolve-Path -LiteralPath $source).Path.TrimEnd('\')
    $files = @(Get-ChildItem -LiteralPath $source -File -Force -Recurse | Where-Object {
        -not $_.FullName.StartsWith(($root + '\.git\'), [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($files.Count -eq 0) { throw 'Source folder has no files to upload.' }
    Write-Host "Found $($files.Count) files (including hidden files)."
    $oversize = @($files | Where-Object { $_.Length -ge 100000000 })
    if ($oversize.Count -gt 0) {
        $oversize | ForEach-Object { Write-Host "Too large: $($_.FullName) ($($_.Length) bytes)" -ForegroundColor Red }
        throw 'GitHub blocks files of 100 MB or larger. Use Git LFS or remove these files first.'
    }
    $sensitive = @($files | Where-Object {
        $_.Name -match '^(\.env(\..*)?|id_rsa|id_ed25519|credentials(\..*)?|\.npmrc|\.pypirc)$' -or
        $_.Extension -match '^\.(pem|p12|pfx|key)$' -or
        $_.FullName -match '[\\/](\.ssh|\.aws)[\\/]'
    })
    if ($sensitive.Count -gt 0) {
        $sensitive | ForEach-Object { Write-Host "Potential secret: $($_.FullName)" -ForegroundColor Yellow }
        throw 'Potentially sensitive files detected. Review and exclude them explicitly before uploading to this PUBLIC repository.'
    }

    $gitFolder = Join-Path $source '.git'
    if (Test-Path -LiteralPath $gitFolder) {
        $repoRoot = GetGitValue -GitArgs @('rev-parse','--show-toplevel')
        if ($repoRoot -and ([IO.Path]::GetFullPath($repoRoot).TrimEnd('\') -ne [IO.Path]::GetFullPath($source).TrimEnd('\'))) {
            throw 'Git repository root does not match selected folder; stopping to protect other projects.'
        }
        $oldHead = GetGitValue -GitArgs @('rev-parse','--verify','HEAD')
        if ($oldHead) {
            throw 'This folder already has Git commit history. Stopping so old commits are not accidentally published. Ask for a history-safe upload method.'
        }
    } else {
        RunGit -GitArguments @('init', '-b', 'main')
    }

    $origin = GetGitValue -GitArgs @('remote','get-url','origin')
    if (-not $origin) {
        RunGit -GitArguments @('remote', 'add', 'origin', $remote)
    } else {
        $originNorm = $origin.ToLowerInvariant() -replace '\.git$','' -replace '^https?://','' -replace '^git@github.com:','github.com/'
        if ($originNorm -ne $expectedRemote) { throw "Existing origin points elsewhere: $origin. Refusing to change it." }
    }

    # An empty remote branch is expected. Never overwrite an existing branch.
    $existingRemoteMain = GetGitValue -GitArgs @('ls-remote','--heads','origin','main')
    if ($existingRemoteMain) { throw 'Remote main already has commits. Stopping to avoid overwriting existing changes.' }

    $name = GetGitValue -GitArgs @('config','user.name')
    if (-not $name) { RunGit -GitArguments @('config', 'user.name', 'YUWENJIE10086') }
    $email = GetGitValue -GitArgs @('config','user.email')
    if (-not $email) { RunGit -GitArguments @('config', 'user.email', '44671256+YUWENJIE10086@users.noreply.github.com') }

    RunGit -GitArguments @('add', '-A')
    $stagedText = GetGitValue -GitArgs @('diff','--cached','--name-only')
    $staged = @($stagedText -split "`r?`n" | Where-Object { $_.Trim() })
    if ($staged.Count -eq 0) { throw 'No staged files. Check existing .gitignore rules.' }
    Write-Host ''
    Write-Host "Staged $($staged.Count) files. First 30 paths:" -ForegroundColor Cyan
    $staged | Select-Object -First 30 | ForEach-Object { Write-Host "  $_" }
    if ($staged.Count -gt 30) { Write-Host "  ... and $($staged.Count - 30) more." }
    Write-Host ''
    Write-Host 'IMPORTANT: The GitHub repository is PUBLIC. Files will be visible to everyone.' -ForegroundColor Yellow
    $answer = Read-Host 'Enter UPLOAD to commit and push these files to main'
    if ($answer -cne 'UPLOAD') { throw 'Cancelled by user. Nothing was uploaded.' }

    RunGit -GitArguments @('commit', '-m', 'Initial upload: TraeNewPaper11 (2026-11-16)')
    RunGit -GitArguments @('push', '-u', 'origin', 'HEAD:main')
    Write-Host ''
    Write-Host 'SUCCESS: Upload completed.' -ForegroundColor Green
    Write-Host 'https://github.com/YUWENJIE10086/myPaper2612'
    RunGit -GitArguments @('rev-parse', 'HEAD')
} catch {
    Write-Host ''
    Write-Host ("STOPPED: " + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'No force push or deletion was performed. Existing user files remain untouched.'
    exit 1
}
