package jp.co.pm.ai.desktop.io.win32;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class MstscWindowLocatorTest {

    @Test
    void clientSurfaceClass() {
        assertTrue(MstscWindowLocator.isClientSurfaceClass("TscShellContainerClass"));
        assertFalse(MstscWindowLocator.isClientSurfaceClass("#32770"));
    }

    @Test
    void rdpTitleHeuristic() {
        assertTrue(
                MstscWindowLocator.looksLikeRdpSessionTitle(
                        "Default - 192.168.0.182 - リモート デスクトップ接続"));
        assertFalse(MstscWindowLocator.looksLikeRdpSessionTitle("Windows セキュリティ"));
    }

    @Test
    void scoreCandidate_prefersClientSurface() {
        int client =
                MstscWindowLocator.scoreCandidate(
                        "host - Remote Desktop Connection", "TscShellContainerClass", true);
        int dialog = MstscWindowLocator.scoreCandidate("資格情報", "#32770", false);
        assertTrue(client > dialog);
    }
}
