package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.Stage2AladdinTodayExcludeNextDayDispatchDialog;
import jp.co.pm.ai.desktop.ui.Stage2InProgressNextDayDispatchDialog;
import jp.co.pm.ai.planning.stage2.Stage2PlanRunDateResolver;

class PlanInputNextDayDialogRowsTest {

    private static final String MACHINE = "スリット機1\u3000湖南";
    private static final List<String> HEADERS =
            List.of(
                    "依頼NO",
                    "工程名",
                    "機械名",
                    "換算数量",
                    "実加工数",
                    "未加工",
                    "配台使用残数量",
                    "(原反)ロール単位長さ");

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void inProgressAndAladdinExcludeRowSetsDoNotOverlap(@TempDir Path temp) throws Exception {
        Path outputDir = temp.resolve("output");
        Files.createDirectories(outputDir);
        Path extractionWorkbook = temp.resolve("data.xlsx");
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = wb.createSheet("加工計画DATA");
            sheet.createRow(0).createCell(0).setCellValue("データ抽出時間");
            sheet.createRow(1).createCell(0).setCellValue("2026/06/09 08:00:00");
            try (var out = java.nio.file.Files.newOutputStream(extractionWorkbook)) {
                wb.write(out);
            }
        }

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                        extractionWorkbook.toString(),
                        AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                        outputDir.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        outputDir.toString());
        String planDateKey = Stage2PlanRunDateResolver.planDateColumnKey(ui);

        Path shapedJson = outputDir.resolve(AppPaths.SHAPED_ALADDIN_PLAN_JSON_BASENAME);
        JsonTableIo.saveArrayTable(
                shapedJson,
                List.of("機械名", "依頼NO", "工程名", planDateKey),
                List.of(
                        List.of(MACHINE, "T-IN", "スリット", "0"),
                        List.of(MACHINE, "T-AL", "スリット", "6090"),
                        List.of(MACHINE, "T-NONE", "スリット", "0")));
        Files.writeString(
                outputDir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME), "{\"rows\":[]}\n");

        PlanInputTabController ctrl = new PlanInputTabController();
        injectGrid(
                ctrl,
                HEADERS,
                List.of(
                        row("T-IN", "スリット", MACHINE, "13530", "2870", "10660", "10660", "3045"),
                        row("T-AL", "スリット", MACHINE, "8000", "0", "8000", "8000", "3045"),
                        row("T-NONE", "スリット", MACHINE, "5000", "0", "5000", "5000", "3045")));

        List<Stage2InProgressNextDayDispatchDialog.Row> inProgress =
                ctrl.collectInProgressRowsForNextDayDialog(ui);
        List<Stage2AladdinTodayExcludeNextDayDispatchDialog.Row> aladdin =
                ctrl.collectAladdinTodayExcludeRowsForNextDayDialog(ui);

        Set<String> inProgressIds =
                inProgress.stream().map(Stage2InProgressNextDayDispatchDialog.Row::taskId).collect(Collectors.toSet());
        Set<String> aladdinIds =
                aladdin.stream()
                        .map(Stage2AladdinTodayExcludeNextDayDispatchDialog.Row::taskId)
                        .collect(Collectors.toSet());

        assertEquals(Set.of("T-IN"), inProgressIds);
        assertEquals(Set.of("T-AL"), aladdinIds);
        assertTrue(inProgressIds.stream().noneMatch(aladdinIds::contains));
        assertEquals(0.0, inProgress.get(0).aladdinTodayM(), 1e-9);
    }

    private static List<String> row(
            String taskId,
            String process,
            String machine,
            String converted,
            String actual,
            String unprocessed,
            String remaining,
            String rollUnit) {
        return List.of(taskId, process, machine, converted, actual, unprocessed, remaining, rollUnit);
    }

    private static void injectGrid(
            PlanInputTabController ctrl, List<String> headers, List<List<String>> dataRows)
            throws Exception {
        Field headersField = PlanInputTabController.class.getDeclaredField("headersRef");
        headersField.setAccessible(true);
        @SuppressWarnings("unchecked")
        List<String> headersRef = (List<String>) headersField.get(ctrl);
        headersRef.clear();
        headersRef.addAll(headers);

        Field rowsField = PlanInputTabController.class.getDeclaredField("rows");
        rowsField.setAccessible(true);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        for (List<String> data : dataRows) {
            rows.add(FXCollections.observableArrayList(data));
        }
        rowsField.set(ctrl, rows);
    }
}
