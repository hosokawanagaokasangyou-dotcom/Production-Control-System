package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class PlanInputEditedCellMarksTest {

    private static List<String> headers() {
        return List.of("依頼NO", "工程名", "機械名", "原反投入日");
    }

    private static ObservableList<ObservableList<String>> rows(String date) {
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("JR260602", "スリット", "SL1", date));
        return rows;
    }

    @Test
    void recompute_marksChangedCellAndUnmarksOnRevert() {
        List<String> headers = headers();
        ObservableList<ObservableList<String>> rows = rows("2026/6/14");
        Map<String, String> baseline =
                PlanInputEditedCellMarks.captureBaseline(headers, rows);
        Set<String> persisted = new LinkedHashSet<>();
        Set<String> edited = new LinkedHashSet<>();

        rows.get(0).set(3, "2026/6/13");
        PlanInputEditedCellMarks.recompute(headers, rows, baseline, persisted, edited);
        String key =
                PlanInputEditedCellMarks.markKey(
                        PlanInputEditedCellMarks.rowKey(headers, rows.get(0)), "原反投入日");
        assertTrue(edited.contains(key));

        rows.get(0).set(3, "2026/6/14");
        PlanInputEditedCellMarks.recompute(headers, rows, baseline, persisted, edited);
        assertFalse(edited.contains(key), "基準値に戻したら未保存マークは外れる");
    }

    @Test
    void recompute_keepsPersistedMarkOnRevert() {
        List<String> headers = headers();
        ObservableList<ObservableList<String>> rows = rows("2026/6/14");
        Map<String, String> baseline =
                PlanInputEditedCellMarks.captureBaseline(headers, rows);
        String key =
                PlanInputEditedCellMarks.markKey(
                        PlanInputEditedCellMarks.rowKey(headers, rows.get(0)), "原反投入日");
        Set<String> persisted = new LinkedHashSet<>(List.of(key));
        Set<String> edited = new LinkedHashSet<>(List.of(key));

        PlanInputEditedCellMarks.recompute(headers, rows, baseline, persisted, edited);
        assertTrue(edited.contains(key), "保存済みマークは基準値一致でも保持する");
    }

    @Test
    void saveLoadRoundTrip(@TempDir Path dir) {
        Path planInput = dir.resolve("plan_input_tasks.xlsx");
        Set<String> marks = new LinkedHashSet<>(List.of("JR260602\u0001スリット\u0001SL1\u0001原反投入日"));
        PlanInputEditedCellMarks.save(planInput, marks);
        assertEquals(marks, PlanInputEditedCellMarks.load(planInput));

        PlanInputEditedCellMarks.save(planInput, new LinkedHashSet<>());
        assertTrue(PlanInputEditedCellMarks.load(planInput).isEmpty(), "空マークは sidecar 削除");
    }

    @Test
    void namespaceSeparatesSidecars(@TempDir Path dir) {
        Path planInput = dir.resolve("plan_input_tasks.xlsx");
        PlanInputEditedCellMarks.save(planInput, new LinkedHashSet<>(List.of("a")), "");
        PlanInputEditedCellMarks.save(planInput, new LinkedHashSet<>(List.of("b")), "stage3");
        assertEquals(Set.of("a"), PlanInputEditedCellMarks.load(planInput, ""));
        assertEquals(Set.of("b"), PlanInputEditedCellMarks.load(planInput, "stage3"));
    }

    @Test
    void filterToPresentRows_dropsMissingRowKeys() {
        List<String> headers = headers();
        ObservableList<ObservableList<String>> rows = rows("2026/6/14");
        Set<String> marks =
                new LinkedHashSet<>(
                        List.of(
                                "JR260602\u0001スリット\u0001SL1\u0001原反投入日",
                                "GONE\u0001工程\u0001機\u0001原反投入日"));
        Set<String> kept = PlanInputEditedCellMarks.filterToPresentRows(headers, rows, marks);
        assertEquals(Set.of("JR260602\u0001スリット\u0001SL1\u0001原反投入日"), kept);
    }
}
