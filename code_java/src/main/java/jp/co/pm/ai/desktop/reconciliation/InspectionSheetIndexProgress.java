package jp.co.pm.ai.desktop.reconciliation;

/** 検査表索引生成の進捗表示。 */
public final class InspectionSheetIndexProgress {

    public static final String PHASE_WALK = "フォルダ走査中";
    public static final String PHASE_INDEX = "索引更新中";

    /** ステータスバー更新の最短間隔。FX スレッドを他タブ操作に譲る。 */
    public static final long UI_THROTTLE_NS = 250_000_000L;

    private InspectionSheetIndexProgress() {}

    public static boolean shouldPublishUi(long lastPublishedNs, long nowNs, boolean force) {
        if (force) {
            return true;
        }
        return nowNs - lastPublishedNs >= UI_THROTTLE_NS;
    }

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
