package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import org.junit.jupiter.api.Test;

class RdpMstscCloserTest {

    @Test
    void closeForProfile_nonWindows_returnsNotFound() {
        if (RemoteDesktopLauncher.isSupportedPlatform()) {
            return;
        }
        RdpMstscCloser.CloseResult result =
                RdpMstscCloser.closeForProfile(java.nio.file.Path.of("C:\\test.rdp"));
        assertFalse(result.closed());
        assertEquals(RdpMstscCloser.CloseMethod.NOT_FOUND, result.method());
    }
}
