package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.scene.control.Label;

/**
 * 結果_配台表 JSON の段階2 / 段階2.1 判定と配台結果タブの段階バッジ表示。
 * 旧段階3試行由来の {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 列の読取補助も含む。
 */
public final class ResultDispatchPlanningStageSupport {

    private static final ObjectMapper JSON = new ObjectMapper();

    public static final String BADGE_STAGE2 = "段階2";

    public static final String BADGE_STAGE21 = "段階2.1";

    public enum PlanningStage {
        STAGE2,
        STAGE21
    }

    private static final double EPS = 1e-6;

    private ResultDispatchPlanningStageSupport() {}

    public static boolean hasActualDispatchQtyColumn(List<String> columns) {
        return columns != null
                && columns.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
    }

    /**
     * 配台計画手動修正タブ用: 段階2の孤立目標行を除き、配台結果タブと同じ行構成にする。
     * {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 列は維持する。
     */
    public static ResultDispatchDocument prepareForDispatchInteractiveWideGrid(
            ResultDispatchDocument doc) {
        if (doc == null) {
            return ResultDispatchDocument.empty();
        }
        if (!hasActualDispatchQtyColumn(doc.columns())) {
            return doc;
        }
        List<String> cols = new ArrayList<>(doc.columns());
        List<Map<String, String>> rows = new ArrayList<>();
        for (Map<String, String> r : doc.rows()) {
            rows.add(new LinkedHashMap<>(r));
        }
        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, rows);
        ResultDispatchDocument out = new ResultDispatchDocument(cols, rows);
        out.setFormatVersion(doc.formatVersion());
        out.setSheetName(doc.sheetName());
        out.setExcelTableName(doc.excelTableName());
        return out;
    }

    public static boolean detectActualQtyColumnFromDispatchJsonPath(Path jsonPath) {
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

    public static boolean detectStage21TrialFromDispatchJsonPath(Path jsonPath) {
        return Stage21TrialSnapshotStore.tryLoadMeta(jsonPath).hasTrialApplied();
    }

    public static PlanningStage detectPlanningStage(Path jsonPath) {
        Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                Stage21TrialSnapshotStore.tryLoadMeta(jsonPath);
        if (meta.hasPromotedToMain()) {
            return PlanningStage.STAGE21;
        }
        return PlanningStage.STAGE2;
    }

    /** 実配台数量列がある旧 JSON: 主数量をタイムライン実績へ寄せる。 */
    public static void applyActualQtyDisplayQuantities(
            List<String> columns, List<Map<String, String>> rows) {
        if (!hasActualDispatchQtyColumn(columns) || rows == null) {
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

    public static void removeRedundantActualColumnFromMaps(
            List<String> columns, List<Map<String, String>> rows) {
        if (!hasActualDispatchQtyColumn(columns)) {
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

    public static void applyPlanningStageBadge(Label badge, PlanningStage stage) {
        if (badge == null) {
            return;
        }
        PlanningStage s = stage != null ? stage : PlanningStage.STAGE2;
        badge.setText(s == PlanningStage.STAGE21 ? BADGE_STAGE21 : BADGE_STAGE2);
        badge.getStyleClass()
                .removeAll(
                        "pm-planning-stage-badge-stage2",
                        "pm-planning-stage-badge-stage3",
                        "pm-planning-stage-badge-stage21",
                        "pm-planning-stage-badge-stage35");
        badge.getStyleClass()
                .add(
                        s == PlanningStage.STAGE21
                                ? "pm-planning-stage-badge-stage21"
                                : "pm-planning-stage-badge-stage2");
        if (!badge.getStyleClass().contains("pm-planning-stage-badge")) {
            badge.getStyleClass().add(0, "pm-planning-stage-badge");
        }
        badge.setVisible(true);
        badge.setManaged(true);
    }

    public static void applyPlanningStageBadgeFromDispatchJson(Label badge, Path jsonPath) {
        if (badge == null) {
            return;
        }
        applyPlanningStageBadge(badge, detectPlanningStage(jsonPath));
    }

    /** 旧段階3試行 JSON・セル表記の読取互換（表示専用）。 */
    public enum Stage3PlanningVariant {
        NONE,
        LEGACY,
        STAGE3_0,
        STAGE3_1,
        STAGE3_2;

        public String actualQtyLabel() {
            return switch (this) {
                case STAGE3_0, STAGE3_1, STAGE3_2, LEGACY -> DispatchQtyCellLineLabels.MANUAL_RESULT;
                case NONE -> "";
            };
        }

        public String revisedQtyLabel() {
            return switch (this) {
                case STAGE3_0, STAGE3_1, STAGE3_2 -> DispatchQtyCellLineLabels.MANUAL_RESULT;
                case LEGACY -> DispatchQtyCellLineLabels.DISPATCH_PLAN_REVISED;
                case NONE -> "";
            };
        }

        public String badgeText() {
            return switch (this) {
                case STAGE3_0 -> "段階3.0";
                case STAGE3_1 -> "段階3.1";
                case STAGE3_2 -> "段階3.2";
                case LEGACY -> "段階3";
                case NONE -> BADGE_STAGE2;
            };
        }
    }

    public static boolean isStage3AfterQtyLine(String line) {
        return DispatchQtyCellLineLabels.isManualResultLine(line);
    }

    public static boolean isStage3RevisedQtyLine(String line) {
        return DispatchQtyCellLineLabels.isDispatchPlanRevisedLine(line);
    }

    public static void applyPlanningStageBadge(
            Label badge, PlanningStage stage, Stage3PlanningVariant ignored) {
        applyPlanningStageBadge(badge, stage);
    }
}
