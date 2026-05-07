package jp.co.pm.ai.desktop.ui;

/**
 * 設備ガントタイムライン1セルの文言から、帯の種類（通常／休憩／始業）を判別する。
 *
 * <p>加工ギャップ分割のラベル「休憩前／休憩後」は語に「休憩」を含むが、暦の休憩帯ではない。
 * {@code contains("休憩")} 単独では誤って {@link #BREAK} になるため先に除外する。
 */
public enum GanttScheduleSlotBarKind {
    DEFAULT,
    BREAK,
    STARTUP;

    public static GanttScheduleSlotBarKind fromTimelineCell(String t) {
        if (t == null || t.isEmpty()) {
            return DEFAULT;
        }
        if (t.contains("休憩前") || t.contains("休憩後")) {
            return DEFAULT;
        }
        if (t.contains("休憩") || t.contains("（休憩）")) {
            return BREAK;
        }
        if (t.contains("日次始業準備")) {
            return STARTUP;
        }
        return DEFAULT;
    }
}
