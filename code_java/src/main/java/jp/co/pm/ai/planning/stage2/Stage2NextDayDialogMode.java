package jp.co.pm.ai.planning.stage2;

/** 段階2直前の翌日設定ダイアログ種別（配台計画_タスク入力タブのラジオ）。 */
public enum Stage2NextDayDialogMode {
    /** ① 加工途中タスクの翌日配台量のみ。 */
    IN_PROGRESS,
    /** ② アラジン当日配台の翌日除外量のみ。 */
    ALADDIN_TODAY_EXCLUDE,
    /** ③ ①→② を連続表示（既定）。 */
    BOTH,
    /** ダイアログを表示しない。 */
    NONE;

    public boolean runsInProgressDialog() {
        return this == IN_PROGRESS || this == BOTH;
    }

    public boolean runsAladdinExcludeDialog() {
        return this == ALADDIN_TODAY_EXCLUDE || this == BOTH;
    }

    public static Stage2NextDayDialogMode parse(String raw) {
        if (raw == null || raw.isBlank()) {
            return defaultMode();
        }
        String s = raw.strip();
        for (Stage2NextDayDialogMode m : values()) {
            if (m.name().equalsIgnoreCase(s)) {
                return m;
            }
        }
        return defaultMode();
    }

    public static Stage2NextDayDialogMode defaultMode() {
        return BOTH;
    }
}
