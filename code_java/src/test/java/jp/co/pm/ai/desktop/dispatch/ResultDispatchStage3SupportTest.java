package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ResultDispatchStage3SupportTest {

    @Test
    void applyStage3DisplayQuantities_copiesActualToPlanColumn() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            row.put(c, "");
        }
        row.put(ResultDispatchSchema.COL_DISPATCH_QTY, "4400");
        row.put(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL, "3000");
        rows.add(row);

        ResultDispatchStage3Support.applyStage3DisplayQuantities(cols, rows);

        assertEquals("3000", rows.get(0).get(ResultDispatchSchema.COL_DISPATCH_QTY));
    }

    @Test
    void removeRedundantActualColumn_dropsColumnFromMaps() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            row.put(c, "1");
        }
        rows.add(row);

        ResultDispatchStage3Support.removeRedundantActualColumnFromMaps(cols, rows);

        assertFalse(cols.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL));
        assertFalse(rows.get(0).containsKey(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL));
    }

    @Test
    void hasStage3ActualColumn_falseWhenAbsent() {
        assertFalse(
                ResultDispatchStage3Support.hasStage3ActualColumn(
                        ResultDispatchSchema.canonicalColumnOrder()));
    }

    @Test
    void hasStage3ActualColumn_trueWhenPresent() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        assertTrue(ResultDispatchStage3Support.hasStage3ActualColumn(cols));
    }

    @Test
    void detectStage21TrialFromSidecar() throws Exception {
        java.nio.file.Path dir = java.nio.file.Files.createTempDirectory("stage21-badge");
        java.nio.file.Path dispatchJson = dir.resolve("結果_配台表.json");
        java.nio.file.Files.writeString(dispatchJson, "{}", java.nio.charset.StandardCharsets.UTF_8);
        java.nio.file.Path stage21Json = dir.resolve("stage21").resolve("結果_配台表.json");
        java.nio.file.Files.createDirectories(stage21Json.getParent());
        java.nio.file.Files.writeString(stage21Json, "{}", java.nio.charset.StandardCharsets.UTF_8);
        Stage21TrialSnapshotStore.writeWithMeta(
                dispatchJson,
                java.util.Map.of(),
                stage21Json,
                dir.resolve("stage21").resolve("overtime_simulation_overrides.json"),
                new Stage21TrialSnapshotStore.OverrideSummary(1, 0, 0));
        assertTrue(ResultDispatchStage3Support.detectStage21TrialFromDispatchJsonPath(dispatchJson));
        assertEquals(
                ResultDispatchStage3Support.PlanningStage.STAGE2,
                ResultDispatchStage3Support.detectPlanningStage(dispatchJson));

        Stage21TrialSnapshotStore.writePromotedMeta(
                dispatchJson,
                dir.resolve("overtime_simulation_overrides.json"),
                new Stage21TrialSnapshotStore.OverrideSummary(1, 0, 0));
        assertEquals(
                ResultDispatchStage3Support.PlanningStage.STAGE21,
                ResultDispatchStage3Support.detectPlanningStage(dispatchJson));
    }

    @Test
    void stage32Sidecar_resolvesStage3PlanningVariantForBadge(@TempDir Path dir) throws Exception {
        Path dispatchJson = dir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{\"columns\":[],\"rows\":[]}", StandardCharsets.UTF_8);
        Stage3PlanningMetaStore.writeVariant(
                dispatchJson, Stage3PlanningMetaStore.Variant.STAGE3_2);

        assertEquals(
                ResultDispatchStage3Support.PlanningStage.STAGE3,
                ResultDispatchStage3Support.detectPlanningStage(dispatchJson));
        assertEquals(
                ResultDispatchStage3Support.Stage3PlanningVariant.STAGE3_2,
                Stage3PlanningMetaStore.readPlanningVariant(dispatchJson));
        assertEquals(
                "段階3.2",
                Stage3PlanningMetaStore.readPlanningVariant(dispatchJson).badgeText());
    }
}
