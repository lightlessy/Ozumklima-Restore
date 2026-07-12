param(
    [string]$Root = (Resolve-Path (Join-Path $PSScriptRoot ".."))
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
$win1254 = [System.Text.Encoding]::GetEncoding(1254)
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = Join-Path $Root ".utf8-backup-$timestamp"

$targets = @(
    (Join-Path $Root "ust.asp"),
    (Join-Path $Root "alt.asp"),
    (Join-Path $Root "uygulama-ornegi.asp"),
    (Join-Path $Root "uygulama-ornekleri.asp")
)

$wpanel = Join-Path $Root "wpanel"
if (Test-Path $wpanel) {
    $targets += Get-ChildItem -Path $wpanel -Recurse -File -Filter "*.asp" | ForEach-Object { $_.FullName }
}

$targets = $targets | Where-Object { Test-Path $_ } | Sort-Object -Unique

function Read-TextSafely([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }

    try {
        return $utf8Strict.GetString($bytes)
    }
    catch {
        return $win1254.GetString($bytes)
    }
}

function Ensure-PagePreamble([string]$Text, [string]$RelativePath) {
    $normalized = $Text.TrimStart([char]0xFEFF)
    $isInclude = $RelativePath -in @("ust.asp", "alt.asp", "wpanel/menu.asp")
    $looksLikePage = $normalized -match '(?is)<!DOCTYPE|<html\b|Response\.Redirect|Request\.(Form|QueryString)'

    if (-not $isInclude -and $looksLikePage) {
        if ($normalized -notmatch '^\s*<%@\s*Language\s*=\s*["'']VBScript["'']\s+CodePage\s*=\s*["'']?65001["'']?\s*%>') {
            $normalized = "<%@ Language=\"VBScript\" CodePage=\"65001\" %>`r`n" + $normalized
        }

        if ($normalized -notmatch 'Response\.CodePage\s*=\s*65001') {
            $directiveEnd = $normalized.IndexOf("%>") + 2
            $runtime = "`r`n<%`r`nResponse.Buffer = True`r`nResponse.CodePage = 65001`r`nResponse.Charset = \"utf-8\"`r`n%>"
            $normalized = $normalized.Insert($directiveEnd, $runtime)
        }
    }

    return $normalized
}

New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
$changed = @()

foreach ($path in $targets) {
    $relative = $path.Substring($Root.Length).TrimStart('\','/')
    $backupPath = Join-Path $backupRoot $relative
    $backupDir = Split-Path $backupPath -Parent
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Copy-Item $path $backupPath -Force

    $text = Read-TextSafely $path
    $before = $text

    $text = [regex]::Replace(
        $text,
        '(?is)<meta\s+http-equiv\s*=\s*["'']Content-Type["'']\s+content\s*=\s*["'']text/html;\s*charset\s*=\s*windows-1254["'']\s*/?>',
        '<meta charset="utf-8">'
    )
    $text = [regex]::Replace(
        $text,
        '(?is)<meta\s+charset\s*=\s*["'']windows-1254["'']\s*/?>',
        '<meta charset="utf-8">'
    )
    $text = Ensure-PagePreamble $text $relative

    [System.IO.File]::WriteAllText($path, $text, $utf8NoBom)
    if ($text -ne $before) { $changed += $relative }
}

Write-Host "Backup: $backupRoot"
Write-Host "Processed: $($targets.Count) files"
Write-Host "Text changed: $($changed.Count) files"
$changed | ForEach-Object { Write-Host "  $_" }
Write-Host "Run next: git diff --check; git diff --stat; git status"
