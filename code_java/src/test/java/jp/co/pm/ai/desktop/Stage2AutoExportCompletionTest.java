package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ResultDispatchTableTabController.AladdinEntryExportOutcome;
import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;

class Stage2AutoExportCompletionTest {

    @Test
    void successMessageEmphasizesGeneratedExcelAndKeepsStage2Success() {
        AladdinEntryExportOutcome outcome =
                new AladdinEntryExportOutcome(
                        new DispatchAladdinEntryWorkbookExporter.ExportResult(
                                Path.of("latest.xlsx"), Path.of("generation.xlsx")),
                        List.of(),
                        null);

        assertTrue(MainShellController.stage2CompletionHeader(outcome).contains("生成しました"));
        assertTrue(MainShellController.stage2CompletionContent(outcome).contains("正常終了"));
        assertTrue(MainShellController.stage2CompletionContent(outcome).contains("latest.xlsx"));
    }

    @Test
    void failureMessageEmphasizesExcelFailureButKeepsStage2Success() {
        AladdinEntryExportOutcome outcome =
                new AladdinEntryExportOutcome(
                        null, List.of(), new IllegalStateException("共有先へ書き込めません"));

        assertTrue(MainShellController.stage2CompletionHeader(outcome).contains("失敗しました"));
        assertTrue(MainShellController.stage2CompletionContent(outcome).contains("正常終了"));
        assertTrue(MainShellController.stage2CompletionContent(outcome).contains("共有先へ書き込めません"));
        assertFalse(MainShellController.stage2CompletionContent(outcome).contains("段階2 の処理に失敗"));
    }
}
