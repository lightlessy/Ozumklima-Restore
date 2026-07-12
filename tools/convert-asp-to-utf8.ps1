param(
    [string]$Root = (Resolve-Path (Join-Path $PSScriptRoot ".."))
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
$legacy1254 = [System.Text.Encoding]::GetEncoding(1254)
$extensions = @("*.asp", "*.asa", "*.txt", "*.xml")

$files = foreach ($pattern in $extensions) {
    Get-ChildItem -Path $Root -Recurse -File -Filter $pattern | Where-Object {
        $_.FullName -notlike "*\.git\*"
    }
}

$processed = 0
$replaced = 0
$legacyDecoded = 0

foreach ($file in ($files | Sort-Object FullName -Unique)) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $wasLegacy = $false

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    else {
        try {
            # Valid UTF-8 without BOM: preserve it as UTF-8.
            $text = $utf8Strict.GetString($bytes)
        }
        catch [System.Text.DecoderFallbackException] {
            # Invalid UTF-8: treat the legacy Classic ASP source as Windows-1254.
            $text = $legacy1254.GetString($bytes)
            $wasLegacy = $true
            $legacyDecoded++
        }
    }

    $original = $text

    $text = [regex]::Replace(
        $text,
        '<meta\s+http-equiv=["'']Content-Type["'']\s+content=["'']text/html;\s*charset=windows-1254["'']\s*/?>',
        '<meta charset="utf-8">',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $text = [regex]::Replace(
        $text,
        'charset\s*=\s*["'']?windows-1254["'']?',
        'charset="utf-8"',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $text = [regex]::Replace($text, 'Response\.CodePage\s*=\s*1254', 'Response.CodePage = 65001', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $text = [regex]::Replace($text, 'Session\.CodePage\s*=\s*1254', 'Session.CodePage = 65001', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $text = [regex]::Replace($text, 'Response\.Charset\s*=\s*["'']windows-1254["'']', 'Response.Charset = "utf-8"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    [System.IO.File]::WriteAllText($file.FullName, $text, $utf8NoBom)
    $processed++

    if ($text -ne $original -or $wasLegacy) {
        $relative = $file.FullName.Substring($Root.Length).TrimStart('\','/')
        $source = if ($wasLegacy) { "windows-1254" } else { "utf-8" }
        Write-Host "Converted: $relative (source: $source)"
        $replaced++
    }
}

Write-Host "UTF-8 conversion completed. Files processed: $processed"
Write-Host "Legacy Windows-1254 files decoded: $legacyDecoded"
Write-Host "Files changed or re-encoded: $replaced"
Write-Host "Review with: git diff --check; git diff --stat; git diff"
