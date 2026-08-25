package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ResultDispatchTableTabController.AladdinEntryExportOutcome;
import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;

class AladdinIdentityCheckLocalFlowTest {

    @Test
    void afterExport_usesGeneratedLatestExcelWhenExportSucceeded() {
        Path latest = Path.of("code", "アラジン入力用_配台計画.xlsx");
        AladdinEntryExportOutcome outcome =
                new AladdinEntryExportOutcome(
                        new DispatchAladdinEntryWorkbookExporter.ExportResult(
                                latest, Path.of("gen.xlsx")),
                        List.of(),
                        null);

        AladdinIdentityCheckLocalFlow.NextStep next =
                AladdinIdentityCheckLocalFlow.afterExport(outcome);

        assertTrue(next.canCheck());
        assertEquals(latest, next.excelPath());
        assertNull(next.errorMessage());
    }

    @Test
    void afterExport_blocksIdentityCheckWhenExportFailed() {
        AladdinEntryExportOutcome outcome =
                new AladdinEntryExportOutcome(
                        null, List.of(), new IllegalStateException("結果_配台表.json が見つかりません"));

        AladdinIdentityCheckLocalFlow.NextStep next =
                AladdinIdentityCheckLocalFlow.afterExport(outcome);

        assertFalse(next.canCheck());
        assertNull(next.excelPath());
        assertTrue(next.errorMessage().contains("アラジン入力用Excelの生成に失敗"));
        assertTrue(next.errorMessage().contains("結果_配台表.json が見つかりません"));
    }

    @Test
    void afterExport_blocksIdentityCheckWhenLatestPathMissing() {
        AladdinEntryExportOutcome outcome =
                new AladdinEntryExportOutcome(
                        new DispatchAladdinEntryWorkbookExporter.ExportResult(null, Path.of("gen.xlsx")),
                        List.of(),
                        null);

        AladdinIdentityCheckLocalFlow.NextStep next =
                AladdinIdentityCheckLocalFlow.afterExport(outcome);

        assertFalse(next.canCheck());
        assertNull(next.excelPath());
        assertTrue(next.errorMessage().contains("生成した配台計画 Excel が見つかりません"));
    }
}
