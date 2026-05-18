package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.scene.control.Label;

/**
 * 結果_配台表 JSON の段階3試行有無と、配台結果タブ／納期管理比較の表示方針。
 */
public final class ResultDispatchStage3Support {

    private static final ObjectMapper JSON = new ObjectMapper();

    public static final String BADGE_STAGE2 = "\u6bb5\u968e2";

    public static final String BADGE_STAGE3 = "\u6bb5\u968e3";

    private static final double EPS = 1e-6;

    private ResultDispatchStage3Support() {}

    public static boolean hasStage3ActualColumn(List<String> columns) {
        return columns != null
                && columns.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
    }

    public static boolean detectStage3FromDispatchJsonPath(Path jsonPath) {
        if (jsonPath == null || !Files.isRegularFile(jsonPath)) {
            return false;
        }
        try {
            String raw = Files.readString(jsonPath, StandardCharsets.UTF_8);
            JsonNode root = JSON.readTree(raw);
            JsonNode columnsNode = root.get("columns");
            if (columnsNode == null || !columnsNode.isArray()) {
                return false;
            }
            for (JsonNode c : columnsNode) {
                if (c != null
                        && c.isTextual()
                        && ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL.equals(c.asText())) {
                    return true;
                }
            }
        } catch (Exception ignored) {
            return false;
        }
        return false;
    }

    /**
     * 段階3試行後: 配台結果タブの主数量はタイムライン実績（実配台数量）とする。
     * {@link ResultDispatchInteractiveConsolidator} の後に呼ぶ。
     */
    public static void applyStage3DisplayQuantities(
            List<String> columns, List<Map<String, String>> rows) {
        if (!hasStage3ActualColumn(columns) || rows == null) {
            return;
        }
        String planCol = ResultDispatchSchema.COL_DISPATCH_QTY;
        String actualCol = ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL;
        for (Map<String, String> row : rows) {
            double actual = ResultDispatchNormalizer.parseDouble(row.get(actualCol));
            if (actual > EPS) {
                row.put(planCol, ResultDispatchNormalizer.formatQty(actual));
            }
        }
    }

    /** 段階3表示時は実配台数量列を表から外す（当日配台数量に統合済み）。 */
    public static void removeRedundantActualColumn(
            List<String> columns, List<List<String>> rowLines) {
        if (!hasStage3ActualColumn(columns)) {
            return;
        }
        int idx = columns.indexOf(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        if (idx < 0) {
            return;
        }
        columns.remove(idx);
        if (rowLines == null) {
            return;
        }
        for (List<String> line : rowLines) {
            if (line != null && idx < line.size()) {
                line.remove(idx);
            }
        }
    }

    public static void removeRedundantActualColumnFromMaps(
            List<String> columns, List<Map<String, String>> rows) {
        if (!hasStage3ActualColumn(columns)) {
            return;
        }
        int idx = columns.indexOf(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        if (idx < 0) {
            return;
        }
        columns.remove(idx);
        if (rows == null) {
            return;
        }
        for (Map<String, String> row : rows) {
            if (row != null) {
                row.remove(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
            }
        }
    }

    public static void applyPlanningStageBadge(Label badge, boolean stage3) {
        if (badge == null) {
            return;
        }
        badge.setText(stage3 ? BADGE_STAGE3 : BADGE_STAGE2);
        badge.getStyleClass().removeAll(
                "pm-planning-stage-badge-stage2", "pm-planning-stage-badge-stage3");
        badge.getStyleClass().add(
                stage3 ? "pm-planning-stage-badge-stage3" : "pm-planning-stage-badge-stage2");
        if (!badge.getStyleClass().contains("pm-planning-stage-badge")) {
            badge.getStyleClass().add(0, "pm-planning-stage-badge");
        }
        badge.setVisible(true);
        badge.setManaged(true);
    }
}
