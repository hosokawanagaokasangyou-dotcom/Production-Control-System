package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class AmountCellRefTest {

    @Test
    @DisplayName("金額列はJとAA")
    void columnLetters() {
        assertEquals("J", AmountCellRef.column(9));
        assertEquals("AA", AmountCellRef.column(26));
        assertEquals("'①東レCSV原本'!J12", AmountCellRef.address(SourceRawSheets.CSV_SHEET, 9, 12));
        assertEquals("'②東レまとめ(当月)'!AA413",
                AmountCellRef.address(SourceRawSheets.copiedSheetName(FactoryProfile.MATOME_SHEET), 26, 413));
    }

    @Test
    @DisplayName("検証Aの①②金額は元セルへリンクする")
    void sheetAAmountsLinkToSourceCells() throws Exception {
        Map<String, Object> info = new HashMap<>();
        info.put("対象月ラベル", "2026年9月度");
        info.put("実行日時", "2026-10-02");
        info.put("実行者", "test");
        info.put("許容差", 0.5);
        info.put("入庫場所", "A010");
        info.put("②名称", "長岡明細");
        info.put("②金額列", "AA");
        info.put("①総額", 1L);
        info.put("②総額", 1L);
        info.put("③総額", 0L);
        info.put("①金額セル", Map.of("192265M", List.of("'①東レCSV原本'!J20", "'①東レCSV原本'!J21")));
        info.put("②金額セル", Map.of("192265M", List.of("'②東レまとめ(当月)'!AA30")));
        VerifyResult result = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(new RecordA("192265M", "V9-9", 307808.0, 480930.0, -173122.0, -173122.0, Judge.MISMATCH, "")),
                List.of(),
                info,
                List.of(),
                null,
                null,
                null);
        try (XSSFWorkbook book = ResultExcelExporter.buildWorkbook(result, null, null)) {
            XSSFSheet sheet = book.getSheet("検証A_契約NO(①vs②)");
            var toray = sheet.getRow(1).getCell(2);
            var nagaoka = sheet.getRow(1).getCell(3);
            assertEquals(CellType.NUMERIC, toray.getCellType());
            assertEquals(307808.0, toray.getNumericCellValue(), 0.001);
            assertEquals("'①東レCSV原本'!J20", toray.getHyperlink().getAddress());
            assertEquals("'②東レまとめ(当月)'!AA30", nagaoka.getHyperlink().getAddress());
            assertNull(sheet.getRow(1).getCell(4).getHyperlink());
        }
    }
}
