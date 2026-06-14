package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class StartupSplashBrandingTest {

    @Test
    void pmdShowsFactorySiteAndKeepsDispatchBranding() {
        StartupSplashBranding branding = StartupSplashBranding.PMD;

        assertTrue(branding.showFactorySite());
        assertEquals("工程管理 AI 配台", branding.title());
        assertEquals("", branding.rootStyleClass());
    }

    @Test
    void remoteDesktopLauncherUsesProductTitleWithoutFactorySite() {
        StartupSplashBranding branding = StartupSplashBranding.REMOTE_DESKTOP_LAUNCHER;

        assertFalse(branding.showFactorySite());
        assertEquals(
                RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE, branding.title());
        assertEquals("splash-app-rdp-launcher", branding.rootStyleClass());
        assertEquals(
                "部署別オペレータ · RDP 接続 · Aladdin RPA", branding.subtitleJa());
        assertEquals(
                "DEPARTMENT · REMOTE DESKTOP · RPA", branding.subtitleEn());
        assertEquals(
                "リモートデスクトップ RPA ランチャーを起動しています…",
                branding.statusText());
        assertEquals("", branding.backgroundResource());
    }

    @Test
    void remoteDesktopLauncherBrandingDoesNotMentionFactoryNames() {
        StartupSplashBranding branding = StartupSplashBranding.REMOTE_DESKTOP_LAUNCHER;

        assertFalse(branding.title().contains("湖南工場"));
        assertFalse(branding.subtitleJa().contains("湖南工場"));
        assertFalse(branding.subtitleEn().contains("湖南工場"));
        assertFalse(branding.statusText().contains("湖南工場"));
        assertFalse(branding.title().contains("国分工場"));
        assertFalse(branding.subtitleJa().contains("国分工場"));
    }
}
