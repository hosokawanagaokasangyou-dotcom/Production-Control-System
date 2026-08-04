package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.scene.image.Image;

class AppWindowIconSupportTest {

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
