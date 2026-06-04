package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class PostProcessingZaikoBunruiMasterLookupTest {

    @Test
    void normalizeLookupCode_padsSixDigits() {
        assertEquals("000001", PostProcessingZaikoBunruiMasterLookup.normalizeLookupCode("1"));
        assertEquals("000005", PostProcessingZaikoBunruiMasterLookup.normalizeLookupCode("5.0"));
    }

    @Test
    void toProductColumnValue_stripsLeadingZeros() {
        assertEquals("2", PostProcessingZaikoBunruiMasterLookup.toProductColumnValue("000002"));
        assertEquals("1", PostProcessingZaikoBunruiMasterLookup.resolveCodeFromComboInput(null, "000001"));
    }

    @Test
    void productColumnDetection() {
        assertTrue(
                PostProcessingZaikoBunruiMasterLookup.isZaikoBunruiProductColumn("商品分類6コード"));
    }

    @Test
    @EnabledIf("aladdinZaikoMasterPresent")
    void snapshot_loadsFromRepo() throws Exception {
        Path root = Path.of(System.getProperty("user.dir")).getParent();
        if (root == null || !Files.isDirectory(root.resolve("アラジンマスタ"))) {
            root = Path.of(".").toAbsolutePath().normalize();
        }
        var snap =
                PostProcessingZaikoBunruiMasterLookup.snapshot(
                        java.util.Map.of(
                                "PM_AI_ALADDIN_MASTER_DIR",
                                root.resolve("アラジンマスタ").toString()));
        assertTrue(snap.loaded());
        assertEquals("2製品", snap.codeToName().get("000001"));
        assertEquals("1商品", snap.codeToName().get("000002"));
    }

    static boolean aladdinZaikoMasterPresent() {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        return Files.isRegularFile(repo.resolve("アラジンマスタ/後加工在庫分類マスタ.xlsx"));
    }
}
