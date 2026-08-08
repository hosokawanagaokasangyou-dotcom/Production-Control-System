package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

class RequestFormOriginalIndexSheetCatalogPaneTest {

    @Test
    void openSourceFile_reportsMissingPath() {
        List<String> statuses = new ArrayList<>();
        RequestFormOriginalIndexSheetCatalogPane.DisplayRow row =
                RequestFormOriginalIndexSheetCatalogPane.DisplayRow.from(
                        new RequestFormOriginalIndexSheetCatalog.Row(
                                "book.xlsm",
                                "",
                                "C7-1",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                ""));
        RequestFormOriginalIndexSheetCatalogPane.openSourceFile(row, statuses::add, null);
        assertEquals("原本ファイルのパスがありません。", statuses.getFirst());
    }
}
