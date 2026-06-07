Add-Type -AssemblyName System.Security
$rdp = Get-ChildItem -Path 'C:\' -Filter 'Default.pm-ai-signed.rdp' -Recurse -ErrorAction SilentlyContinue -Depth 4 | Select-Object -First 1
$text = [System.IO.File]::ReadAllText($rdp.FullName, [System.Text.Encoding]::Unicode)
$line = ($text -split "`r?`n" | Where-Object { $_ -match '(?i)^signature:s:' } | Select-Object -First 1)
$b64 = ($line -replace '(?i)^signature:s:', '' -replace '\s','')
$bytes = [Convert]::FromBase64String($b64)
$cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
$cms.Decode($bytes)
Write-Host "Signing cert thumbprints:"
$thumbs = @()
foreach ($c in $cms.Certificates) {
  Write-Host "  $($c.Thumbprint) | $($c.Subject)"
  $thumbs += $c.Thumbprint.ToUpper()
}
$reg = [string](Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' -EA SilentlyContinue).TrustedCertThumbprints
Write-Host "HKLM trusted list: $reg"
foreach ($t in $thumbs) {
  if ($reg.ToUpper().Contains($t)) { Write-Host "OK: $t is registered" }
  else { Write-Host "NG: $t is NOT registered" }
}
