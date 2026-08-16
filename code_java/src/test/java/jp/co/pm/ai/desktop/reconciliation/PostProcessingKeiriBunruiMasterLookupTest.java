package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingKeiriBunruiMasterLookupTest {

    @Test
    void normalizeCode_padsTwoDigits() {
        assertEquals("01", PostProcessingKeiriBunruiMasterLookup.normalizeCode("1"));
        assertEquals("06", PostProcessingKeiriBunruiMasterLookup.normalizeCode("6.0"));
        assertEquals("11", PostProcessingKeiriBunruiMasterLookup.normalizeCode("11"));
        assertEquals("06", PostProcessingKeiriBunruiMasterLookup.normalizeCode("06 社内製品"));
    }

    @Test
    void productColumnDetection() {
        assertTrue(
                PostProcessingKeiriBunruiMasterLookup.isKeiriBunruiProductColumn("単価分類コード"));
    }

    @Test
    @EnabledIf("aladdinKeiriMasterPresent")
    void snapshot_loadsFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingKeiriBunruiMasterLookup.snapshot(
                        java.util.Map.of(
                                "PM_AI_ALADDIN_MASTER_DIR",
                                root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertEquals("社内製品", snap.codeToName().get("06"));
        assertEquals("社外製品", snap.codeToName().get("01"));
    }

    static boolean aladdinKeiriMasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        Path master = repo.resolve("アラジンマスタ/後加工経理分類マスタ.xlsx");
        if (!Files.isRegularFile(master)) {
            return false;
        }
        try (var in = Files.newInputStream(master)) {
            return in.read() >= 0;
        } catch (IOException ex) {
            return false;
        }
    }
}
