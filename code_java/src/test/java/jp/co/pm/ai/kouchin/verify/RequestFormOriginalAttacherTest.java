package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
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
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFee;

class RequestFormOriginalAttacherTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("Gemini応答のJSONから加工賃を読む")
    void readsFeeFromModelJson() {
        RequestFormOriginalFee.Result fee = RequestFormOriginalFee.parseModelJson(
                "```json\n{\"amountYen\":207900,\"meters\":6300,\"yenPerMeter\":33,\"reason\":\"6300×33\"}\n```");
        assertEquals(207900.0, fee.amountYen(), 0.001);
        assertEquals(6300.0, fee.meters(), 0.001);
        assertEquals(33.0, fee.yenPerMeter(), 0.001);
        assertEquals("6300×33", fee.reason());
        assertTrue(fee.contracts().isEmpty());
        RequestFormOriginalFee.Result split = RequestFormOriginalFee.parseModelJson(
                "{\"lines\":[{\"keiyaku\":\"192591B\",\"amountYen\":8800,\"meters\":200,\"reason\":\"44×200\"},"
                        + "{\"keiyaku\":\"192592\",\"amountYen\":1500,\"meters\":100,\"reason\":\"15×100\"}]}");
        assertEquals(10300.0, split.amountYen(), 0.001);
        assertEquals(2, split.contracts().size());
        assertEquals("192592", split.contracts().get(1).keiyaku());
        assertNull(RequestFormOriginalFee.parseModelJson("{\"amountYen\":null,\"reason\":\"不明\"}"));
    }

    @Test
    @DisplayName("一致だけの結果は依頼書の加工賃を計算しない")
    void matchRowsAreNotPrepared() {
        Map<String, Object> info = new HashMap<>();
        info.put("許容差", 0.5);
        VerifyResult result = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(new RecordA("192265M", "V9-9", 100.0, 100.0, 0.0, null, Judge.MATCH, "")),
                List.of(),
                info,
                List.of(),
                null,
                null,
                null);
        RequestFormOriginalAttacher.Plan plan = RequestFormOriginalAttacher.prepare(result, Map.of(
                AppPaths.KEY_PM_AI_SKIP_GEMINI_API, "1",
                AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, tempDir.toString(),
                AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, tempDir.resolve("no-juchu.xlsx").toString()));
        assertTrue(plan.items().isEmpty(), plan.items().toString());
        assertTrue(plan.warnings().isEmpty(), plan.warnings().toString());
    }

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
            form.createRow(9).createCell(30).setCellValue(100);
            form.createRow(12).createCell(15).setCellValue(18);
            form.createRow(13).createCell(15).setCellValue(15);
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
        Map<String, String> ui = Map.of(
                AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, tempDir.toString(),
                AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, tempDir.resolve("no-juchu.xlsx").toString(),
                AppPaths.KEY_PM_AI_SKIP_GEMINI_API, "1");

        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(result, null, null, ui)) {
            Sheet copied = wb.getSheet("原本_V9-9");
            assertNotNull(copied, "添付シートが無い");
            assertEquals("依頼書本文", copied.getRow(0).getCell(0).getStringCellValue());

            Sheet a = wb.getSheet("検証A_契約NO(①vs②)");
            Row data = a.getRow(1);
            assertEquals("", data.getCell(8).getStringCellValue());
            Cell attached = data.getCell(9);
            Cell fileLink = data.getCell(10);
            assertEquals("V9-9", attached.getStringCellValue());
            assertEquals(HyperlinkType.DOCUMENT, attached.getHyperlink().getType());
            assertTrue(attached.getHyperlink().getAddress().contains("原本_V9-9"), attached.getHyperlink().getAddress());
            assertEquals(HyperlinkType.FILE, fileLink.getHyperlink().getType());
            String fileAddress = java.net.URLDecoder.decode(fileLink.getHyperlink().getAddress(), java.nio.charset.StandardCharsets.UTF_8);
            assertTrue(fileAddress.contains("2026加工依頼書.xlsm"), fileAddress);

            Sheet index = wb.getSheet(RequestFormOriginalAttacher.INDEX_SHEET);
            assertNotNull(index);
            Row iraiRow = null;
            for (Row row : index) {
                Cell c = row.getCell(0);
                if (c != null && "V9-9".equals(c.getStringCellValue())) {
                    iraiRow = row;
                    break;
                }
            }
            assertNotNull(iraiRow);
            assertEquals("", iraiRow.getCell(4).getStringCellValue());
        }
    }
}
