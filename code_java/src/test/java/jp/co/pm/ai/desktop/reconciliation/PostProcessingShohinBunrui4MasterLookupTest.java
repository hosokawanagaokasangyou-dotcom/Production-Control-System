package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingShohinBunrui4MasterLookupTest {

    @Test
    void normalizeCode_stripsDecimal() {
        assertEquals("999999", PostProcessingShohinBunrui4MasterLookup.normalizeCode("999999.0"));
        assertEquals("510100", PostProcessingShohinBunrui4MasterLookup.normalizeCode("510100"));
    }

    @Test
    void productColumnDetection() {
        assertTrue(
                PostProcessingShohinBunrui4MasterLookup.isBunrui4ProductColumn("商品分類4コード"));
    }

    @Test
    @EnabledIf("aladdinBunrui4MasterPresent")
    void snapshot_loadsFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingShohinBunrui4MasterLookup.snapshot(
                        java.util.Map.of(
                                "PM_AI_ALADDIN_MASTER_DIR",
                                root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertEquals("対象外", snap.codeToName().get("999999"));
        assertEquals("後加工-売上", snap.codeToName().get("510100"));
    }

    static boolean aladdinBunrui4MasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        return Files.isRegularFile(repo.resolve("アラジンマスタ/後加工商品分類4マスタ.xlsx"));
    }
}
