package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class RdpMstscSessionMonitorTest {

    @Test
    void profileMatchKey_normalizesSeparators() {
        assertEquals(
                "C:\\data\\factory.rdp",
                RdpMstscSessionMonitor.profileMatchKey(Path.of("C:/data/factory.rdp")));
    }

    @Test
    void processRefersToProfile_matchesArgumentPath() {
        ProcessHandle current = ProcessHandle.current();
        assertFalse(
                RdpMstscSessionMonitor.processRefersToProfile(
                        current, Path.of("C:\\missing\\other.rdp")));
    }
}
