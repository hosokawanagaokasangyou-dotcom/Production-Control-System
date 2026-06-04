package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingKouteiNaiyoMasterLookupTest {

    @Test
    void normalizeCode_padsFourDigits() {
        assertEquals("0110", PostProcessingKouteiNaiyoMasterLookup.normalizeNaiyoCode("110"));
        assertEquals("0001", PostProcessingKouteiNaiyoMasterLookup.normalizeKouteiCode("1"));
        assertEquals("", PostProcessingKouteiNaiyoMasterLookup.normalizeCode("", 4));
    }

    @Test
    void stepColumnDetection() {
        assertTrue(PostProcessingKouteiNaiyoMasterLookup.isKouteiCodeColumn("工程コード3"));
        assertTrue(PostProcessingKouteiNaiyoMasterLookup.isNaiyoCodeColumn("加工内容コード8"));
        assertEquals(5, PostProcessingKouteiNaiyoMasterLookup.stepIndex("加工内容コード5").orElse(-1));
    }

    @Test
    @EnabledIf("aladdinMasterPresent")
    void snapshot_loadsMastersFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingKouteiNaiyoMasterLookup.snapshot(
                        java.util.Map.of("PM_AI_ALADDIN_MASTER_DIR", root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertNotNull(snap.kouteiCodeToName().get("0001"));
        PostProcessingKouteiNaiyoMasterLookup.NaiyoEntry e =
                snap.naiyoCodeToEntry().get("0110");
        assertNotNull(e);
        assertEquals("0001", e.kouteiCode());
    }

    static boolean aladdinMasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        return Files.isRegularFile(repo.resolve("アラジンマスタ/後加工工程マスタ.xlsx"))
                && Files.isRegularFile(repo.resolve("アラジンマスタ/後加工加工内容マスタ.xlsx"));
    }
}
