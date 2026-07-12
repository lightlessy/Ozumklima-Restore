param(
    [string]$Root = (Resolve-Path (Join-Path $PSScriptRoot ".."))
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$extensions = @("*.asp", "*.asa", "*.txt", "*.xml")
$skipDirs = @(".git")

$files = foreach ($pattern in $extensions) {
    Get-ChildItem -Path $Root -Recurse -File -Filter $pattern | Where-Object {
        $relative = $_.FullName.Substring($Root.Length).TrimStart('\','/')
        -not ($skipDirs | Where-Object { $relative -like "$_*" })
    }
}

$changed = 0
foreach ($file in ($files | Sort-Object FullName -Unique)) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    # Decode legacy Turkish files as Windows-1254 unless a UTF-8 BOM is present.
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    else {
        $legacy = [System.Text.Encoding]::GetEncoding(1254)
        $text = $legacy.GetString($bytes)
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
    $text = $text -replace 'Response\.CodePage\s*=\s*1254', 'Response.CodePage = 65001'
    $text = $text -replace 'Session\.CodePage\s*=\s*1254', 'Session.CodePage = 65001'
    $text = $text -replace 'Response\.Charset\s*=\s*["'']windows-1254["'']', 'Response.Charset = "utf-8"'

    if ($text -ne $original -or -not [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($text)).Equals($text)) {
        [System.IO.File]::WriteAllText($file.FullName, $text, $utf8NoBom)
        Write-Host "Converted: $($file.FullName.Substring($Root.Length).TrimStart('\','/'))"
        $changed++
    }
    else {
        # Still normalize the file bytes to UTF-8 without BOM.
        [System.IO.File]::WriteAllText($file.FullName, $text, $utf8NoBom)
    }
}

Write-Host "UTF-8 conversion completed. Files processed: $($files.Count). Files with textual replacements: $changed"
Write-Host "Review with: git diff --check; git diff"
