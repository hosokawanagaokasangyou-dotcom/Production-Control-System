package jp.co.pm.ai.desktop.debug;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.Test;

class AgentDebugLogMirrorPathTest {

    @Test
    void buildWslUncPathString_typicalDrivePath() {
        assertEquals(
                "\\\\wsl$\\Ubuntu\\mnt\\c\\repo\\.cursor\\debug-478e5a.log",
                AgentDebugLog.buildWslUncPathString(
                        "C:\\repo\\.cursor\\debug-478e5a.log", "Ubuntu"));
    }

    @Test
    void buildWslUncPathString_wslLocalhostPrefix() {
        assertEquals(
                "\\\\wsl.localhost\\Ubuntu\\mnt\\c\\repo\\.cursor\\x.log",
                AgentDebugLog.buildWslUncPathString(
                        "C:\\repo\\.cursor\\x.log", "Ubuntu", "\\\\wsl.localhost\\"));
    }

    @Test
    void buildWslUncPathString_forwardSlashesNormalized() {
        assertEquals(
                "\\\\wsl$\\Ubuntu\\mnt\\c\\repo\\.cursor\\x.log",
                AgentDebugLog.buildWslUncPathString("C:/repo/.cursor/x.log", "Ubuntu"));
    }

    @Test
    void buildWslUncPathString_nullWhenNoDistro() {
        assertNull(AgentDebugLog.buildWslUncPathString("C:\\a.log", ""));
        assertNull(AgentDebugLog.buildWslUncPathString("C:\\a.log", null));
    }

    @Test
    void buildWslUncPathString_nullWhenNotDriveLetterPath() {
        assertNull(AgentDebugLog.buildWslUncPathString("\\\\wsl$\\Ubuntu\\mnt\\c\\a.log", "Ubuntu"));
    }
}
