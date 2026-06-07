$ErrorActionPreference = 'Continue'
$rdpCandidates = @(
    'C:\工程管理AIプロジェクト_JAVA\Default.pm-ai-signed.rdp',
    'C:\Users\0585\OneDrive\ドキュメント\Default.rdp'
)
Write-Host '=== RDP files ==='
foreach ($p in $rdpCandidates) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host "MISSING: $p"
        continue
    }
    $bytes = [System.IO.File]::ReadAllBytes($p)
    $utf16 = [System.Text.Encoding]::Unicode.GetString($bytes)
    $hasSig = $utf16 -match '(?i)signature:s:'
    Write-Host "FILE: $p"
    Write-Host "  size: $($bytes.Length) bytes"
    Write-Host "  signature:s: $hasSig"
}
Write-Host ''
Write-Host '=== Registry ==='
foreach ($hive in @('HKLM', 'HKCU')) {
    $path = "$hive`:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
    $p = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    Write-Host "$hive :"
    if ($null -eq $p) { Write-Host '  not configured'; continue }
    Write-Host "  AllowSignedFiles=$($p.AllowSignedFiles)"
    $thumb = [string]$p.TrustedCertThumbprints
    if ($thumb) {
        $parts = $thumb -split '[;,]'
        Write-Host "  thumbprints count=$($parts.Count) first=$($parts[0])"
    } else { Write-Host '  TrustedCertThumbprints empty' }
}
Write-Host ''
Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "CERT: $($_.Subject) SHA1=$($_.Thumbprint)"
}
