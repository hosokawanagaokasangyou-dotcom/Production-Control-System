package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RequestFormOriginalAttacherTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("月ラベル付き依頼NOを分け、一致行は添付対象にしない")
    void splitsIraiAndSkipsMatch() {
        assertEquals(List.of("V9-9", "C8-1"), RequestFormOriginalAttacher.splitIrai("(9月)V9-9, C8-1"));
        assertTrue(RequestFormOriginalAttacher.anomalyA(Judge.MISMATCH));
        assertTrue(RequestFormOriginalAttacher.anomalyA(Judge.ONLY_1));
        assertTrue(!RequestFormOriginalAttacher.anomalyA(Judge.MATCH));
        assertTrue(!RequestFormOriginalAttacher.anomalyA(Judge.PREV_ADJUST));
    }

    @Test
    @DisplayName("異常依頼の原本シートを結果ブックへコピーし、シートとファイルの両方へリンクする")
    void attachesSheetAndLinksFile() throws Exception {
        Path original = tempDir.resolve("2026加工依頼書.xlsm");
        try (XSSFWorkbook book = new XSSFWorkbook()) {
            Sheet index = book.createSheet("目次");
            Row header = index.createRow(0);
            header.createCell(0).setCellValue("加工依頼NO");
            index.createRow(1).createCell(0).setCellValue("V9-9");
            Sheet form = book.createSheet("V9-9");
            form.createRow(0).createCell(0).setCellValue("依頼書本文");
            try (var out = Files.newOutputStream(original)) {
                book.write(out);
            }
        }

        Map<String, Object> info = new HashMap<>();
        info.put("対象月ラベル", "2026年9月度");
        info.put("実行日時", "2026-10-01");
        info.put("実行者", "test");
        info.put("許容差", 0.5);
        info.put("入庫場所", "A010");
        info.put("②名称", "長岡明細");
        info.put("②金額列", "AA");
        VerifyResult result = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(new RecordA("192265M", "V9-9", 100.0, 80.0, 20.0, 20.0, Judge.MISMATCH, "差")),
                List.of(),
                info,
                List.of(),
                null,
                null,
                null);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, tempDir.toString());

        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(result, null, null, ui)) {
            Sheet copied = wb.getSheet("原本_V9-9");
            assertNotNull(copied, "添付シートが無い");
            assertEquals("依頼書本文", copied.getRow(0).getCell(0).getStringCellValue());

            Sheet a = wb.getSheet("検証A_契約NO(①vs②)");
            Row data = a.getRow(1);
            Cell attached = data.getCell(8);
            Cell fileLink = data.getCell(9);
            assertEquals("V9-9", attached.getStringCellValue());
            assertEquals(HyperlinkType.DOCUMENT, attached.getHyperlink().getType());
            assertTrue(attached.getHyperlink().getAddress().contains("原本_V9-9"), attached.getHyperlink().getAddress());
            assertEquals(HyperlinkType.FILE, fileLink.getHyperlink().getType());
            String fileAddress = java.net.URLDecoder.decode(fileLink.getHyperlink().getAddress(), java.nio.charset.StandardCharsets.UTF_8);
            assertTrue(fileAddress.contains("2026加工依頼書.xlsm"), fileAddress);

            Sheet index = wb.getSheet(RequestFormOriginalAttacher.INDEX_SHEET);
            assertNotNull(index);
            assertEquals("V9-9", index.getRow(1).getCell(0).getStringCellValue());
        }
    }
}
