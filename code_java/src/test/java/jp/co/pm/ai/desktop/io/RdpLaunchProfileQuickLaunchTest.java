package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.util.List;

import org.junit.jupiter.api.Test;

class RdpLaunchProfileQuickLaunchTest {

    @Test
    void catalogOrderProfileNumbers_takesFirstNInOrder() {
        List<Integer> catalog = List.of(1, 2, 3, 4, 5, 6, 7, 8);
        assertEquals(List.of(1, 2, 3), RdpLaunchProfileQuickLaunch.catalogOrderProfileNumbers(catalog, 3));
        assertEquals(
                List.of(1, 2, 3, 4, 5, 6, 7, 8),
                RdpLaunchProfileQuickLaunch.catalogOrderProfileNumbers(catalog, 8));
    }

    @Test
    void catalogOrderProfileNumbers_whenFewerThanLimit_returnsAll() {
        List<Integer> catalog = List.of(1, 2);
        assertEquals(List.of(1, 2), RdpLaunchProfileQuickLaunch.quickLaunchProfileNumbers(catalog));
    }

    @Test
    void slotProfileNumbers_padsWithNull() {
        List<Integer> slots = RdpLaunchProfileQuickLaunch.slotProfileNumbers(List.of(1, 2));
        assertEquals(8, slots.size());
        assertEquals(1, slots.get(0));
        assertEquals(2, slots.get(1));
        assertNull(slots.get(2));
        assertNull(slots.get(7));
    }

    @Test
    void buttonLabel_truncatesLongText() {
        String full = "2: アラジン 工程別加工計画問い合わせとマスタ更新の長い名称";
        String shortLabel = RdpLaunchProfileQuickLaunch.buttonLabel(full, 20);
        assertEquals(20, shortLabel.length());
        assertEquals("2: アラジン 工程別加工計画問い合わ…", shortLabel);
    }

    @Test
    void buttonLabel_keepsShortText() {
        assertEquals(
                "1: アラジン起動のみ",
                RdpLaunchProfileQuickLaunch.buttonLabel("1: アラジン起動のみ"));
    }
}
