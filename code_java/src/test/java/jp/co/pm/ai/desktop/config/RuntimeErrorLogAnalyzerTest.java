package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RuntimeErrorLogAnalyzerTest {

    @Test
    void scan_readsDailyDiagnosticLinesAndMetaError(@TempDir Path temp) throws Exception {
        Map<String, String> ui = archiveUi(temp);
        Path daily =
                RemoteSupportLogArchive.appendDailyUiLog(
                        ui,
                        "古家",
                        "[kouchin] ドロップされたファイルがありません（Outlookの添付はファイルとしてドロップしてください）\n"
                                + "段階2を開始します\n"
                                + "[kouchin] 検証C 月次実績① 読取不可\n",
                        java.time.LocalDate.of(2026, 9, 18));
        assertTrue(Files.isRegularFile(daily));

        RemoteSupportLogArchive.archiveAfterStage(
                ui,
                "古家",
                "kouchin",
                1,
                "ドロップ空",
                "[kouchin] ドロップされたファイルがありません\n",
                java.time.LocalDateTime.of(2026, 9, 18, 7, 35, 0));

        List<RuntimeErrorLogAnalyzer.Row> rows = RuntimeErrorLogAnalyzer.scan(ui);
        assertTrue(rows.size() >= 2, rows.toString());
        assertTrue(
                rows.stream().anyMatch(r -> r.excerpt().contains("ドロップされたファイルがありません")),
                rows.toString());
        assertTrue(
                rows.stream().anyMatch(r -> r.excerpt().contains("読取不可")),
                rows.toString());
        assertTrue(rows.stream().anyMatch(r -> "error".equals(r.severity())), rows.toString());
        assertTrue(rows.stream().allMatch(r -> "古家".equals(r.operator())), rows.toString());
        assertEquals(
                List.of("古家"),
                RuntimeErrorLogAnalyzer.listOperators(ui));
    }

    @Test
    void severityOf_classifiesErrorAndWarn() {
        assertEquals(
                "error",
                RuntimeErrorLogAnalyzer.severityOf(
                        "[kouchin] ドロップされたファイルがありません"));
        assertEquals("error", RuntimeErrorLogAnalyzer.severityOf("java.io.IOException: x"));
        assertEquals("error", RuntimeErrorLogAnalyzer.severityOf("[kouchin] 検証C 読取不可"));
        assertEquals("warn", RuntimeErrorLogAnalyzer.severityOf("[kouchin] 取り込み完了: 1件"));
    }

    private static Map<String, String> archiveUi(Path temp) throws Exception {
        Path summary = temp.resolve("サマリ_AI配台.xlsx");
        Files.writeString(summary, "x", StandardCharsets.UTF_8);
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summary.toString());
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, temp.toString());
        ui.put(AppPaths.KEY_PM_AI_FACTORY_SITE, FactorySite.KONAN.name());
        ui.put(AppPaths.KEY_PM_AI_REMOTE_LOG, "1");
        return ui;
    }
}
