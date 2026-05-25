package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

/** ローカル原本が存在するときのみ E6-2 のセル読取を検証する。 */
class RequestFormOriginalExtractorE6ProbeTest {

    private static final Path REAL_BOOK =
            Path.of(
                    "/mnt/c/Users/0585/OneDrive/ドキュメント/加工依頼書（湖南）/依頼書",
                    "E-6月加工依頼書（湖南E）2026.xlsm");

    @Test
    void e6_2_readsKakochinFromAe20AndTokkiFromXColumn() throws Exception {
        File file = REAL_BOOK.toFile();
        assumeTrue(file.isFile(), "real workbook not available: " + file);

        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = WorkbookFactory.create(fis)) {
            Sheet sheet = wb.getSheet("E6-2");
            assumeTrue(sheet != null, "E6-2 sheet missing");

            Map<String, String> raw =
                    RequestFormOriginalExtractor.buildRawMapFromSheet(file, "E6-2", sheet);

            assertEquals("E6-2", raw.get("依頼Ｎｏ").replace(" ", ""));
            assertEquals("49", raw.get("加工賃").replace(".00", "").strip());
            assertTrue(
                    raw.getOrDefault("特記事項1", "").contains("Q面"),
                    "tokki1=" + raw.get("特記事項1"));
            assertTrue(
                    !raw.getOrDefault("特記事項1", "").contains("加工"),
                    "tokki1 must not be processing label: " + raw.get("特記事項1"));
        }
    }
}
