package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingYotoMasterLookupTest {

    @Test
    void normalizeCode_stripsDecimalAndLabel() {
        assertEquals("55109", PostProcessingYotoMasterLookup.normalizeCode("55109.0"));
        assertEquals("50101", PostProcessingYotoMasterLookup.normalizeCode("50101 トーレペフ"));
    }

    @Test
    void productColumnDetection() {
        assertTrue(PostProcessingYotoMasterLookup.isYotoProductColumn("商品分類2コード"));
    }

    @Test
    @EnabledIf("aladdinYotoMasterPresent")
    void snapshot_loadsFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingYotoMasterLookup.snapshot(
                        java.util.Map.of(
                                "PM_AI_ALADDIN_MASTER_DIR",
                                root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertEquals("ﾍﾟﾌ　その他", snap.codeToName().get("55109"));
        assertEquals("トーレペフ　W(自材)", snap.codeToName().get("50101"));
    }

    static boolean aladdinYotoMasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        return Files.isRegularFile(repo.resolve("アラジンマスタ/後加工用途マスタ.xlsx"));
    }
}
