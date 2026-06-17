package jp.co.pm.ai.desktop.io.win32;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.RemoteDesktopLauncherAppIdentity;

class MstscWindowLocatorTest {

    @Test
    void clientSurfaceClass() {
        assertTrue(MstscWindowLocator.isClientSurfaceClass("TscShellContainerClass"));
        assertTrue(MstscWindowLocator.isClientSurfaceClass("IM Client Area"));
        assertFalse(MstscWindowLocator.isClientSurfaceClass("#32770"));
    }

    @Test
    void rdpTitleHeuristic() {
        assertTrue(
                MstscWindowLocator.looksLikeRdpSessionTitle(
                        "Default - 192.168.0.182 - リモート デスクトップ接続"));
        assertTrue(
                MstscWindowLocator.looksLikeRdpSessionTitle(
                        "host - Remote Desktop Connection"));
        assertFalse(MstscWindowLocator.looksLikeRdpSessionTitle("Windows セキュリティ"));
    }

    @Test
    void launcherTitle_isNotRdpSession() {
        assertFalse(
                MstscWindowLocator.looksLikeRdpSessionTitle(
                        RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE));
        assertTrue(
                MstscWindowLocator.isLauncherWindowTitle(
                        RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE));
    }

    @Test
    void excludedCaptureWindow_launcherAndSelfProcess() {
        assertTrue(
                MstscWindowLocator.isExcludedCaptureWindow(
                        RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE, "Glass Window", 99999));
        int self = (int) ProcessHandle.current().pid();
        assertTrue(
                MstscWindowLocator.isExcludedCaptureWindow(
                        "host - リモート デスクトップ接続", "#32770", self));
    }

    @Test
    void scoreCandidate_prefersClientSurface() {
        int client =
                MstscWindowLocator.scoreCandidate(
                        "host - Remote Desktop Connection", "TscShellContainerClass", true);
        int dialog = MstscWindowLocator.scoreCandidate("資格情報", "#32770", false);
        assertTrue(client > dialog);
    }

    @Test
    void scoreCandidate_fullscreenEmptyTitle_clientSurfaceStillMatches() {
        int emptyTitleSurface =
                MstscWindowLocator.scoreCandidate("", "TscShellContainerClass", true);
        int imClientArea =
                MstscWindowLocator.scoreCandidate("", "IM Client Area", true);
        assertTrue(emptyTitleSurface > 0);
        assertTrue(imClientArea > 0);
        assertTrue(emptyTitleSurface >= imClientArea);
    }

    @Test
    void resolveRootHwnd_invalidHandle_returnsZeroOrSelf() {
        assertEquals(0L, MstscWindowLocator.resolveRootHwnd(0L));
        assertEquals(0L, MstscWindowLocator.resolveRootHwnd(-1L));
    }

    @Test
    void scoreCandidate_launcherTitle_isZero() {
        assertEquals(
                0,
                MstscWindowLocator.scoreCandidate(
                        RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE, "Glass Window", false));
        int self = (int) ProcessHandle.current().pid();
        assertEquals(
                0,
                MstscWindowLocator.scoreCandidate(
                        RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE,
                        "Glass Window",
                        false,
                        self));
    }

    @Test
    void scoreCandidate_looseRemoteDesktopTitleWithoutConnection_isZero() {
        assertEquals(
                0,
                MstscWindowLocator.scoreCandidate("リモートデスクトップRPAランチャー", "#32770", false));
    }
}
