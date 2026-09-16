package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.MainShellTabId;

class MainShellKouchinTabPlacementTest {

    @Test
    void defaultFlatPlacesKouchinImmediatelyBeforeProcessingTrend() {
        List<String> keys = MainShellTabLayoutDefaults.DEFAULT_FLAT_TAB_KEY_ORDER;
        int k = keys.indexOf(MainShellTabId.KOUCHIN.key());
        int p = keys.indexOf(MainShellTabId.PROCESSING_TREND.key());
        assertTrue(k >= 0 && p >= 0);
        assertEquals(p - 1, k);
    }

    @Test
    void groupedLayoutPlacesKouchinImmediatelyBeforeProcessingTrend() {
        List<String> leaves = new java.util.ArrayList<>();
        for (MainShellTabLayoutNode n : MainShellTabLayoutDefaults.groupedLayout()) {
            collect(n, leaves);
        }
        int k = leaves.indexOf(MainShellTabId.KOUCHIN.key());
        int p = leaves.indexOf(MainShellTabId.PROCESSING_TREND.key());
        assertEquals(p - 1, k, leaves.toString());
    }

    @Test
    void insertMissingKeyBeforeProcessingTrend() {
        List<String> in = List.of("run", "processingTrend", "env");
        assertEquals(List.of("run", "kouchin", "processingTrend", "env"),
                MainShellKouchinTabPlacement.insertMissingKey(in));
    }

    @Test
    void insertMissingKeyAppendsWhenNoProcessingTrend() {
        List<String> in = List.of("run", "env");
        assertEquals(List.of("run", "env", "kouchin"), MainShellKouchinTabPlacement.insertMissingKey(in));
    }

    @Test
    void insertMissingLeafInsideGroup() {
        List<MainShellTabLayoutNode> top = List.of(
                MainShellTabLayoutNode.groupNode(
                        "g",
                        "#000",
                        List.of(
                                MainShellTabLayoutNode.tabNode("run", ""),
                                MainShellTabLayoutNode.tabNode("processingTrend", ""))));
        List<MainShellTabLayoutNode> out = MainShellKouchinTabPlacement.insertMissingLeaf(top);
        List<String> leaves = new java.util.ArrayList<>();
        for (MainShellTabLayoutNode n : out) {
            collect(n, leaves);
        }
        assertEquals(List.of("run", "kouchin", "processingTrend"), leaves);
    }

    private static void collect(MainShellTabLayoutNode n, List<String> out) {
        if (n.isTab()) {
            out.add(n.id());
        } else {
            for (MainShellTabLayoutNode c : n.children()) {
                collect(c, out);
            }
        }
    }
}
