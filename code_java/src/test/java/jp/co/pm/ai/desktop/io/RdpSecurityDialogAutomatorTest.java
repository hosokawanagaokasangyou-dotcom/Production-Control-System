package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;

class RdpSecurityDialogAutomatorTest {

    @Test
    void isAutoConfirmEnabled_defaultsTrueWhenUnset() {
        assertTrue(RdpSecurityDialogAutomator.isAutoConfirmEnabled(Map.of()));
        assertTrue(RdpSecurityDialogAutomator.isAutoConfirmEnabled(null));
    }

    @Test
    void isAutoConfirmEnabled_respectsDisableValues() {
        assertFalse(
                RdpSecurityDialogAutomator.isAutoConfirmEnabled(
                        Map.of(RdpSecurityDialogAutomator.KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG, "0")));
        assertFalse(
                RdpSecurityDialogAutomator.isAutoConfirmEnabled(
                        Map.of(RdpSecurityDialogAutomator.KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG, "false")));
        assertFalse(
                RdpSecurityDialogAutomator.isAutoConfirmEnabled(
                        Map.of(RdpSecurityDialogAutomator.KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG, "off")));
    }

    @Test
    void buildAutomationScript_containsMstscAndRdpPaths() {
        String script =
                RdpSecurityDialogAutomator.buildAutomationScript(
                        Path.of("C:\\Windows\\System32\\mstsc.exe"),
                        Path.of("C:\\repo\\Default.pm-ai-signed.rdp"),
                        30,
                        Path.of("C:\\Temp\\PM-AI-rdp-sign\\rdp-mstsc-test.pid"));
        assertTrue(script.contains("mstsc.exe"));
        assertTrue(script.contains("Default.pm-ai-signed.rdp"));
        assertTrue(script.contains("rdp-mstsc-test.pid"));
        assertTrue(script.contains("Find-SecurityDialog") || script.contains("Find-SecurityWarningWindow"));
        assertTrue(script.contains("Enable-AllCheckboxes") || script.contains("Check-AllCheckBoxes"));
        assertTrue(script.contains("Invoke-ConnectButton") || script.contains("Click-ConnectButton"));
        assertTrue(script.contains("Start-Process"));
        assertTrue(script.contains("-PathType Leaf"), "ディレクトリを .rdp として渡さない");
        assertTrue(script.contains("ArgumentList $rdpArg"), "mstsc へ引用符付き .rdp を渡す");
    }

    @Test
    void escapePowerShellSingleQuoted_doublesSingleQuotes() {
        assertTrue(
                RdpSecurityDialogAutomator.escapePowerShellSingleQuoted("O'Brien").contains("O''Brien"));
    }
}
