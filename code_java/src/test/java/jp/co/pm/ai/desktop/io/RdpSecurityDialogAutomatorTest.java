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
        assertTrue(script.contains("Find-SecurityDialog"));
        assertTrue(script.contains("Enable-AllCheckboxes"));
        assertTrue(script.contains("Invoke-ConnectButton"));
        assertTrue(script.contains("Start-Process"));
    }

    @Test
    void escapePowerShellSingleQuoted_doublesSingleQuotes() {
        assertTrue(
                RdpSecurityDialogAutomator.escapePowerShellSingleQuoted("O'Brien").contains("O''Brien"));
    }
}
