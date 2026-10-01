package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class CombinedVerifyWorkbookTest {

    @Test
    @DisplayName("片方だけの検証結果では統合しない")
    void combinedOnlyWhenBothExist() {
        assertFalse(CombinedVerifyWorkbook.bothVerifiable(null, null));
        assertFalse(CombinedVerifyWorkbook.bothVerifiable(stub(), null));
        assertFalse(CombinedVerifyWorkbook.bothVerifiable(null, stub()));
        assertTrue(CombinedVerifyWorkbook.bothVerifiable(stub(), stub()));
    }

    @Test
    @DisplayName("統合ブックは工場プレフィックスを付け、シート内リンクも合わせる")
    void prefixesSheetsAndRewritesLinks() throws Exception {
        try (XSSFWorkbook kokubu = book("サマリ", "'検証A'!A1");
                XSSFWorkbook konan = book("サマリ", "'検証A'!B2");
                XSSFWorkbook combined = CombinedVerifyWorkbook.merge(kokubu, konan)) {
            assertNotNull(combined.getSheet("統合"));
            assertNotNull(combined.getSheet("国分_サマリ"));
            assertNotNull(combined.getSheet("湖南_サマリ"));
            XSSFSheet kokubuSummary = combined.getSheet("国分_サマリ");
            Hyperlink link = kokubuSummary.getRow(0).getCell(0).getHyperlink();
            assertEquals(HyperlinkType.DOCUMENT, link.getType());
            assertEquals("'国分_検証A'!A1", link.getAddress());
            assertEquals("'湖南_検証A'!B2",
                    combined.getSheet("湖南_サマリ").getRow(0).getCell(0).getHyperlink().getAddress());
        }
    }

    @Test
    @DisplayName("引用符付きシート名のリンクも工場名を付ける")
    void rewritesQuotedSheetName() {
        assertEquals("'国分_原本_V9-9'!A1",
                CombinedVerifyWorkbook.rewriteDocumentAddress(
                        "'原本_V9-9'!A1",
                        java.util.Map.of("原本_V9-9", "国分_原本_V9-9")));
    }

    private static VerifyResult stub() {
        return new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                null,
                java.util.List.of(),
                java.util.List.of(),
                java.util.Map.of(),
                java.util.List.of(),
                null,
                null,
                null);
    }

    private static XSSFWorkbook book(String summary, String linkAddress) {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(summary);
        Hyperlink link = wb.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress(linkAddress);
        sheet.createRow(0).createCell(0).setHyperlink(link);
        sheet.getRow(0).getCell(0).setCellValue("開く");
        wb.createSheet("検証A");
        return wb;
    }
}
