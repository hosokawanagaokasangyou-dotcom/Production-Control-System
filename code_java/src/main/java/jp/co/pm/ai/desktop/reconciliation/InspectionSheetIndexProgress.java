package jp.co.pm.ai.desktop.reconciliation;

/** 検査表索引生成の進捗表示。 */
public final class InspectionSheetIndexProgress {

    public static final String PHASE_WALK = "フォルダ走査中";
    public static final String PHASE_INDEX = "索引更新中";

    private InspectionSheetIndexProgress() {}

    public static String format(String phase, int processed, int total) {
        String p = phase != null && !phase.isBlank() ? phase : PHASE_INDEX;
        if (total > 0) {
            return "検査表索引 " + p + " " + processed + " / " + total;
        }
        return "検査表索引 " + p + " " + Math.max(0, processed) + " 件";
    }

    public static double fraction(int processed, int total) {
        if (total <= 0) {
            return Double.NaN;
        }
        return Math.max(0.0, Math.min(1.0, (double) processed / (double) total));
    }
}
