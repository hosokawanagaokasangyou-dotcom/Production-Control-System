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
 * 結果_配台表 JSON の段階2/2.1/3.0～3.2 試行有無と、配台結果タブ／納期管理比較の表示方針。
 */
public final class ResultDispatchStage3Support {

    private static final ObjectMapper JSON = new ObjectMapper();

    public static final String BADGE_STAGE2 = "\u6bb5\u968e2";

    public static final String BADGE_STAGE3 = "\u6bb5\u968e3";

    public static final String BADGE_STAGE21 = "\u6bb5\u968e2.1";

    /** 配台手動修正の日付セル数量行: 段階3.0/3.1/3.2 または旧配台試行。 */
    public enum Stage3PlanningVariant {
        NONE,
        /** 旧 {@code dispatch_interactive_trial} 等（sidecar 無し・実配台数量列あり）。 */
        LEGACY,
        STAGE3_0,
        STAGE3_1,
        STAGE3_2;

        public static Stage3PlanningVariant fromMetaVariant(Stage3PlanningMetaStore.Variant variant) {
            if (variant == null) {
                return LEGACY;
            }
            return switch (variant) {
                case STAGE3_0 -> STAGE3_0;
                case STAGE3_1 -> STAGE3_1;
                case STAGE3_2 -> STAGE3_2;
            };
        }

        public String actualQtyLabel() {
            return switch (this) {
                case STAGE3_0 -> "(段階3.0後)";
                case STAGE3_1 -> "(段階3.1後)";
                case STAGE3_2 -> "(段階3.2後)";
                case LEGACY -> "(段階3後)";
                case NONE -> "";
            };
        }

        /** 手修正行。3.0～3.2 は実績行と同じ {@code (段階3.x後)} 表記（旧 {@code 改} は {@link #isStage3RevisedQtyLine} で読取）。 */
        public String revisedQtyLabel() {
            return switch (this) {
                case STAGE3_0, STAGE3_1, STAGE3_2 -> actualQtyLabel();
                case LEGACY -> "(段階3改)";
                case NONE -> "";
            };
        }

        public String badgeText() {
            return switch (this) {
                case STAGE3_0 -> "段階3.0";
                case STAGE3_1 -> "段階3.1";
                case STAGE3_2 -> "段階3.2";
                case LEGACY -> BADGE_STAGE3;
                case NONE -> BADGE_STAGE2;
            };
        }
    }

    /** 旧ラベル {@code (段階3後)} および 3.0～3.2 実績行。 */
    public static boolean isStage3AfterQtyLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        if (line.startsWith("(段階3後)")) {
            return true;
        }
        return line.startsWith("(段階3.0後)")
                || line.startsWith("(段階3.1後)")
                || line.startsWith("(段階3.2後)");
    }

    /** 旧ラベル {@code (段階3改)} および 3.0～3.2 手修正行。 */
    public static boolean isStage3RevisedQtyLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        if (line.startsWith("(段階3改)")) {
            return true;
        }
        return line.startsWith("(段階3.0改)")
                || line.startsWith("(段階3.1改)")
                || line.startsWith("(段階3.2改)");
    }

    public enum PlanningStage {
        STAGE2,
        STAGE3,
        STAGE21
    }

    private static final double EPS = 1e-6;

    private ResultDispatchStage3Support() {}

    public static boolean hasStage3ActualColumn(List<String> columns) {
        return columns != null
                && columns.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
    }

    /**
     * 配台計画手動修正タブ用: 段階2の孤立目標行（タイムライン無しの旧 配台日 行）を除き、配台結果タブと同じ行構成にする。
     * {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 列は維持（ワイド表の (段階3.x後) 表示用）。
     */
    public static ResultDispatchDocument prepareForDispatchInteractiveWideGrid(
            ResultDispatchDocument doc) {
        if (doc == null) {
            return ResultDispatchDocument.empty();
        }
        if (!hasStage3ActualColumn(doc.columns())) {
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

    public static boolean detectStage21TrialFromDispatchJsonPath(Path jsonPath) {
        return Stage21TrialSnapshotStore.tryLoadMeta(jsonPath).hasTrialApplied();
    }

    /** @deprecated 段階3.5 廃止。{@link #detectStage21TrialFromDispatchJsonPath} を使用。 */
    @Deprecated
    public static boolean detectStage35FromDispatchJsonPath(Path jsonPath) {
        return detectStage21TrialFromDispatchJsonPath(jsonPath);
    }

    public static PlanningStage detectPlanningStage(Path jsonPath) {
        if (detectStage3FromDispatchJsonPath(jsonPath)) {
            return PlanningStage.STAGE3;
        }
        if (Stage3PlanningMetaStore.hasPipelinePlanningVariant(jsonPath)) {
            return PlanningStage.STAGE3;
        }
        Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                Stage21TrialSnapshotStore.tryLoadMeta(jsonPath);
        if (meta.hasPromotedToMain()) {
            return PlanningStage.STAGE21;
        }
        return PlanningStage.STAGE2;
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
        applyPlanningStageBadge(
                badge, stage3 ? PlanningStage.STAGE3 : PlanningStage.STAGE2);
    }

    public static void applyPlanningStageBadge(Label badge, PlanningStage stage) {
        applyPlanningStageBadge(badge, stage, Stage3PlanningVariant.NONE);
    }

    /** {@code 結果_配台表.json} と sidecar から段階バッジを解決して適用する。 */
    public static void applyPlanningStageBadgeFromDispatchJson(Label badge, Path jsonPath) {
        if (badge == null) {
            return;
        }
        PlanningStage stage = detectPlanningStage(jsonPath);
        Stage3PlanningVariant variant = Stage3PlanningMetaStore.readPlanningVariant(jsonPath);
        applyPlanningStageBadge(
                badge,
                stage,
                stage == PlanningStage.STAGE3 ? variant : Stage3PlanningVariant.NONE);
    }

    public static void applyPlanningStageBadge(
            Label badge, PlanningStage stage, Stage3PlanningVariant stage3Variant) {
        if (badge == null) {
            return;
        }
        PlanningStage s = stage != null ? stage : PlanningStage.STAGE2;
        Stage3PlanningVariant v =
                stage3Variant != null ? stage3Variant : Stage3PlanningVariant.NONE;
        badge.setText(
                switch (s) {
                    case STAGE21 -> BADGE_STAGE21;
                    case STAGE3 ->
                            v == Stage3PlanningVariant.NONE ? BADGE_STAGE3 : v.badgeText();
                    case STAGE2 -> BADGE_STAGE2;
                });
        badge.getStyleClass()
                .removeAll(
                        "pm-planning-stage-badge-stage2",
                        "pm-planning-stage-badge-stage3",
                        "pm-planning-stage-badge-stage21",
                        "pm-planning-stage-badge-stage35");
        badge.getStyleClass()
                .add(
                        switch (s) {
                            case STAGE21 -> "pm-planning-stage-badge-stage21";
                            case STAGE3 -> "pm-planning-stage-badge-stage3";
                            case STAGE2 -> "pm-planning-stage-badge-stage2";
                        });
        if (!badge.getStyleClass().contains("pm-planning-stage-badge")) {
            badge.getStyleClass().add(0, "pm-planning-stage-badge");
        }
        badge.setVisible(true);
        badge.setManaged(true);
    }
}
