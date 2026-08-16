package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.util.List;

import javafx.application.Platform;
import javafx.scene.image.Image;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class AppWindowIconSupportTest {

    @BeforeAll
    static void initFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void desktopIconsAreBundled() {
        List<Image> icons = AppWindowIconSupport.loadIcons(AppWindowIconSupport.Variant.DESKTOP);
        assertFalse(icons.isEmpty(), "desktop icons");
        assertEquals(6, icons.size());
        assertNotNull(icons.get(icons.size() - 1).getUrl());
    }

    @Test
    void rdpLauncherIconsAreBundled() {
        List<Image> icons = AppWindowIconSupport.loadIcons(AppWindowIconSupport.Variant.RDP_LAUNCHER);
        assertFalse(icons.isEmpty(), "rdp launcher icons");
        assertEquals(6, icons.size());
        assertNotNull(icons.get(0).getUrl());
    }
}
