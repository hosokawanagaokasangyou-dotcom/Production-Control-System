package jp.co.pm.ai.desktop.dispatch;

/** 配台計画手動修正・納期管理ビューの日付セル内数量行ラベル（表示名と旧 JSON 互換）。 */
public final class DispatchQtyCellLineLabels {

    /** 編集目標（当日配台数量）。旧称: (段階3前)。 */
    public static final String DISPATCH_PLAN = "(配台計画)";

    /** 配台試行後の実績行（読取専用）。旧称: (段階3後) 等。 */
    public static final String MANUAL_RESULT = "(手動改定)";

    /** 試行後に手動で当日配台数量を変更したセル。旧称: (段階3改) 等。 */
    public static final String DISPATCH_PLAN_REVISED = "(配台計画改)";

    private static final String LEGACY_PLAN = "(段階3前)";
    private static final String LEGACY_RESULT = "(段階3後)";
    private static final String LEGACY_REVISED = "(段階3改)";

    private DispatchQtyCellLineLabels() {}

    public static boolean isDispatchPlanLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        return line.startsWith(DISPATCH_PLAN) || line.startsWith(LEGACY_PLAN);
    }

    public static boolean isManualResultLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        return line.startsWith(MANUAL_RESULT)
                || line.startsWith(LEGACY_RESULT)
                || line.startsWith("(段階3.0後)")
                || line.startsWith("(段階3.1後)")
                || line.startsWith("(段階3.2後)");
    }

    public static boolean isDispatchPlanRevisedLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        return line.startsWith(DISPATCH_PLAN_REVISED)
                || line.startsWith(LEGACY_REVISED)
                || line.startsWith("(段階3.0改)")
                || line.startsWith("(段階3.1改)")
                || line.startsWith("(段階3.2改)");
    }
}
