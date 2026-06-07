$rdp = Get-ChildItem -Path 'C:\' -Filter 'Default.pm-ai-signed.rdp' -Recurse -ErrorAction SilentlyContinue -Depth 4 | Select-Object -First 1
if (-not $rdp) { Write-Error 'Default.pm-ai-signed.rdp not found'; exit 1 }
Write-Host "RDP=$($rdp.FullName)"
$rdpsign = "$env:Windir\System32\rdpsign.exe"
& $rdpsign -v $rdp.FullName 2>&1 | ForEach-Object { Write-Host $_ }
Write-Host "rdpsign exit=$LASTEXITCODE"
try {
  $sig = Get-AuthenticodeSignature -LiteralPath $rdp.FullName
  Write-Host "Authenticode Status=$($sig.Status)"
  if ($sig.SignerCertificate) {
    Write-Host "Signer=$($sig.SignerCertificate.Subject)"
    Write-Host "Thumbprint=$($sig.SignerCertificate.Thumbprint)"
  }
} catch { Write-Host "Authenticode error: $_" }
$reg = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' -ErrorAction SilentlyContinue
Write-Host "HKLM Trusted=$($reg.TrustedCertThumbprints)"
