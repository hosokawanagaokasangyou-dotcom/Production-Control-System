package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingShuruiMasterLookupTest {

    @Test
    void normalizeCode_stripsDecimalAndLabel() {
        assertEquals("551", PostProcessingShuruiMasterLookup.normalizeCode("551.0"));
        assertEquals("999999", PostProcessingShuruiMasterLookup.normalizeCode("999999"));
        assertEquals("411", PostProcessingShuruiMasterLookup.normalizeCode("411:フィルム"));
        assertEquals("551", PostProcessingShuruiMasterLookup.normalizeCode("551 ペフ"));
    }

    @Test
    void productColumnDetection() {
        assertTrue(PostProcessingShuruiMasterLookup.isShuruiProductColumn("商品分類1コード"));
    }

    @Test
    @EnabledIf("aladdinShuruiMasterPresent")
    void snapshot_loadsFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingShuruiMasterLookup.snapshot(
                        java.util.Map.of(
                                "PM_AI_ALADDIN_MASTER_DIR",
                                root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertEquals("ペフ", snap.codeToName().get("551"));
        assertEquals("フィルム", snap.codeToName().get("411"));
    }

    static boolean aladdinShuruiMasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        return Files.isRegularFile(repo.resolve("アラジンマスタ/後加工種類マスタ.xlsx"));
    }
}
