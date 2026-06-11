package jp.co.pm.ai.desktop.io.win32;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

class MstscWindowEmbedderTest {

    @Test
    void unsupportedOnNonWindows() {
        MstscWindowEmbedder embedder = new MstscWindowEmbedder();
        if (RemoteDesktopLauncher.isSupportedPlatform()) {
            assertTrue(embedder.isSupported());
        } else {
            assertFalse(embedder.isSupported());
            assertFalse(embedder.isAttached());
        }
    }
}
