package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo.TabularSheet;

class SummaryAiDispatchExportColumnSupportTest {

    @Test
    void applySheetLayout_reordersNonDateColumnsOnly() {
        TabularSheet src =
                new TabularSheet(
                        List.of("B", "2026年5月17日(土)", "A", "2026年5月18日(日)"),
                        List.of(List.of("b1", "d1", "a1", "d2")));
        TabularSheet out =
                SummaryAiDispatchExportColumnSupport.applySheetLayout(
                        src,
                        SummaryAiDispatchExportPrefs.SheetKey.MAIN_COMPARE,
                        List.of("A", "B"));
        assertEquals(List.of("A", "B", "2026年5月17日(土)", "2026年5月18日(日)"), out.headers());
        assertEquals(List.of("a1", "b1", "d1", "d2"), out.rows().get(0));
    }

    @Test
    void isDateColumnHeader_dispatchStaticAndSlashDates() {
        assertTrue(
                SummaryAiDispatchExportColumnSupport.isDateColumnHeader(
                        SummaryAiDispatchExportPrefs.SheetKey.DISPATCH, "受注日"));
        assertTrue(
                SummaryAiDispatchExportColumnSupport.isDateColumnHeader(
                        SummaryAiDispatchExportPrefs.SheetKey.DISPATCH, "2026/05/17"));
        assertTrue(
                !SummaryAiDispatchExportColumnSupport.isDateColumnHeader(
                        SummaryAiDispatchExportPrefs.SheetKey.DISPATCH, "依頼NO"));
    }
}
