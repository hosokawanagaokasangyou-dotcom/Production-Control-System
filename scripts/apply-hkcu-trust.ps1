$thumb = 'B39638421424858590EFB220131450E145D2B6D3'
$path = 'HKCU:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
try {
  New-Item -Path $path -Force | Out-Null
  New-ItemProperty -Path $path -Name AllowSignedFiles -PropertyType DWord -Value 1 -Force | Out-Null
  New-ItemProperty -Path $path -Name TrustedCertThumbprints -PropertyType String -Value $thumb -Force | Out-Null
  Write-Host 'HKCU OK'
  Get-ItemProperty $path | Select-Object AllowSignedFiles, TrustedCertThumbprints | Format-List
} catch {
  Write-Host "HKCU FAILED: $($_.Exception.Message)"
  exit 1
}
