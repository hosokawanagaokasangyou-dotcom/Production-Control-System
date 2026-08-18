package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

class RequestFormOriginalIndexSheetCatalogPaneTest {

    @Test
    void openSourceFile_reportsMissingPath() {
        List<String> statuses = new ArrayList<>();
        RequestFormOriginalIndexSheetCatalogPane.DisplayRow row = row("", "C7-1", "", "", "");
        RequestFormOriginalIndexSheetCatalogPane.openSourceFile(row, statuses::add, null);
        assertEquals("原本ファイルのパスがありません。", statuses.getFirst());
    }

    @Test
    void formatCountLabel_unfilteredAndFiltered() {
        assertEquals("登録 581 件", RequestFormOriginalIndexSheetCatalogPane.formatCountLabel(581, 581, 0));
        assertEquals(
                "表示 12 / 登録 581 件",
                RequestFormOriginalIndexSheetCatalogPane.formatCountLabel(12, 581, 0));
        assertEquals(
                "登録 581 件（警告 3 件）",
                RequestFormOriginalIndexSheetCatalogPane.formatCountLabel(581, 581, 3));
        assertEquals(
                "表示 12 / 登録 581 件（警告 3 件）",
                RequestFormOriginalIndexSheetCatalogPane.formatCountLabel(12, 581, 3));
    }

    @Test
    void matchesFilter_blankMatchesAll() {
        RequestFormOriginalIndexSheetCatalogPane.DisplayRow row =
                row("book.xlsm", "C7-27", "7/23", "7/17", "107857T");
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, ""));
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "   "));
    }

    @Test
    void matchesFilter_hitsBusinessKeysIncludingDates() {
        RequestFormOriginalIndexSheetCatalogPane.DisplayRow row =
                row("C-7月（2024年）加工依頼書.xlsm", "C7-27", "7/23", "7/17", "107857T");
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "c7-27"));
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "7/23"));
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "7/17"));
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "107857t"));
        assertTrue(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "7月"));
        assertFalse(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "C8-99"));
        assertFalse(RequestFormOriginalIndexSheetCatalogPane.matchesFilter(row, "8/1"));
    }

    private static RequestFormOriginalIndexSheetCatalogPane.DisplayRow row(
            String fileName, String iraiNo, String deliveryDate, String inputDate, String contractNo) {
        return RequestFormOriginalIndexSheetCatalogPane.DisplayRow.from(
                new RequestFormOriginalIndexSheetCatalog.Row(
                        fileName,
                        "",
                        iraiNo,
                        "",
                        "",
                        inputDate,
                        deliveryDate,
                        "",
                        "",
                        contractNo,
                        ""));
    }
}
