package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

/**
 * mstsc の RDP セキュリティ警告ダイアログを UI Automation（PowerShell）で操作する。
 *
 * <p>2026 年更新以降、リダイレクトチェックはダイアログ上で既定 OFF のため、RPA 相当で ON にして「接続」を押す。
 */
public final class RdpSecurityDialogAutomator {

    /** 0 / false / off で無効。未設定時は有効。 */
    public static final String KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG =
            "PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG";

    public static final int DEFAULT_TIMEOUT_SEC = 45;

    private RdpSecurityDialogAutomator() {}

    public static boolean isAutoConfirmEnabled(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG);
        }
        if (raw == null || raw.isBlank()) {
            return true;
        }
        String v = raw.trim().toLowerCase(Locale.ROOT);
        return !("0".equals(v) || "false".equals(v) || "off".equals(v) || "no".equals(v));
    }

    /**
     * 監視スクリプトを起動したうえで {@code mstsc.exe} を起動する（非同期・呼び出し元は待たない）。
     */
    public static Path launchWithAutomatedConfirm(Path mstscExe, Path rdpProfile, Map<String, String> ui)
            throws IOException {
        Objects.requireNonNull(mstscExe, "mstscExe");
        Objects.requireNonNull(rdpProfile, "rdpProfile");
        if (!RemoteDesktopLauncher.isSupportedPlatform()) {
            throw new IOException("RDP セキュリティダイアログ自動操作は Windows のみ対応です。");
        }
        Path scriptDir = resolveScriptDir();
        Path scriptFile = scriptDir.resolve("rdp-security-auto.ps1");
        Path pidMarkerFile = scriptDir.resolve("rdp-mstsc-" + System.currentTimeMillis() + ".pid");
        writeUtf8PowerShell(
                scriptFile,
                buildAutomationScript(mstscExe, rdpProfile, DEFAULT_TIMEOUT_SEC, pidMarkerFile));
        ProcessBuilder pb =
                new ProcessBuilder(
                        "powershell.exe",
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-WindowStyle",
                        "Hidden",
                        "-File",
                        scriptFile.toAbsolutePath().normalize().toString());
        pb.redirectErrorStream(true);
        pb.start();
        return pidMarkerFile;
    }

    static String buildAutomationScript(
            Path mstscExe, Path rdpProfile, int timeoutSec, Path pidMarkerFile) {
        String mstsc = escapePowerShellSingleQuoted(mstscExe.toAbsolutePath().normalize().toString());
        String rdp = escapePowerShellSingleQuoted(rdpProfile.toAbsolutePath().normalize().toString());
        String pidFile =
                escapePowerShellSingleQuoted(pidMarkerFile.toAbsolutePath().normalize().toString());
        int timeout = Math.max(5, timeoutSec);
        return """
                $ErrorActionPreference = 'Continue'
                Add-Type -AssemblyName UIAutomationClient
                Add-Type -AssemblyName UIAutomationTypes

                function Find-SecurityDialog {
                  param([int]$TimeoutSec)
                  $deadline = (Get-Date).AddSeconds($TimeoutSec)
                  $windowType = [System.Windows.Automation.ControlType]::Window
                  $typeCond = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $windowType)
                  while ((Get-Date) -lt $deadline) {
                    $root = [System.Windows.Automation.AutomationElement]::RootElement
                    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $typeCond)
                    foreach ($w in $windows) {
                      try {
                        $name = [string]$w.Current.Name
                        if ($name -match 'セキュリティ警告' -or $name -match 'security warning') {
                          return $w
                        }
                      } catch { }
                    }
                    Start-Sleep -Milliseconds 250
                  }
                  return $null
                }

                function Enable-AllCheckboxes {
                  param($Window)
                  $cbType = [System.Windows.Automation.ControlType]::CheckBox
                  $cbCond = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $cbType)
                  $pattern = [System.Windows.Automation.TogglePattern]::Pattern
                  $checked = 0
                  foreach ($cb in $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cbCond)) {
                    try {
                      $tp = $cb.GetCurrentPattern($pattern)
                      if ($tp.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) {
                        $tp.Toggle()
                        $checked++
                      }
                    } catch { }
                  }
                  return $checked
                }

                function Invoke-ConnectButton {
                  param($Window)
                  $btnType = [System.Windows.Automation.ControlType]::Button
                  $btnCond = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $btnType)
                  $pattern = [System.Windows.Automation.InvokePattern]::Pattern
                  foreach ($btn in $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)) {
                    try {
                      $name = [string]$btn.Current.Name
                      if ($name -match '^接続' -or $name -match '^Connect') {
                        $btn.GetCurrentPattern($pattern).Invoke()
                        return $true
                      }
                    } catch { }
                  }
                  return $false
                }

                $mstsc = '%s'
                $rdp = '%s'
                $pidFile = '%s'
                $timeout = %d
                if (-not (Test-Path -LiteralPath $mstsc)) { exit 3 }
                if (-not (Test-Path -LiteralPath $rdp)) { exit 4 }
                try {
                  $proc = Start-Process -FilePath $mstsc -ArgumentList @($rdp) -WindowStyle Normal -PassThru
                  if ($null -ne $proc) {
                    Set-Content -LiteralPath $pidFile -Value ([string]$proc.Id) -Encoding ascii -NoNewline
                  }
                } catch {
                  exit 5
                }
                $dialog = Find-SecurityDialog -TimeoutSec $timeout
                if ($null -eq $dialog) { exit 0 }
                Enable-AllCheckboxes -Window $dialog | Out-Null
                Start-Sleep -Milliseconds 300
                if (Invoke-ConnectButton -Window $dialog) { exit 0 }
                exit 1
                """
                .formatted(mstsc, rdp, pidFile, timeout);
    }

    static String escapePowerShellSingleQuoted(String value) {
        return value.replace("'", "''");
    }

    private static Path resolveScriptDir() throws IOException {
        String tmp = System.getenv("TEMP");
        Path dir =
                tmp != null && !tmp.isBlank()
                        ? Path.of(tmp.trim(), "PM-AI-rdp-sign")
                        : Path.of(System.getProperty("java.io.tmpdir", "."), "PM-AI-rdp-sign");
        Files.createDirectories(dir);
        return dir.toAbsolutePath().normalize();
    }

    /** PowerShell 5.1 が UTF-8 として解釈できるよう BOM 付き UTF-8 で書き込む。 */
    private static void writeUtf8PowerShell(Path file, String content) throws IOException {
        byte[] body = content.getBytes(StandardCharsets.UTF_8);
        byte[] bom = new byte[] {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        byte[] out = new byte[bom.length + body.length];
        System.arraycopy(bom, 0, out, 0, bom.length);
        System.arraycopy(body, 0, out, bom.length, body.length);
        Files.write(file, out);
    }
}
