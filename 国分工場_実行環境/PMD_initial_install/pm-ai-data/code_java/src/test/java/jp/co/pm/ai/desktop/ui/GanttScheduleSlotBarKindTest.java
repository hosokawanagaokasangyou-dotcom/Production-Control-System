package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class GanttScheduleSlotBarKindTest {

    @Test
    void gapSplit_machining_phase_labels_areDefault_notBreak() {
        assertEquals(
                GanttScheduleSlotBarKind.DEFAULT,
                GanttScheduleSlotBarKind.fromTimelineCell("Y5-3 休憩前 1000m"));
        assertEquals(
                GanttScheduleSlotBarKind.DEFAULT,
                GanttScheduleSlotBarKind.fromTimelineCell("Y5-3 休憩後 1200m"));
    }

    @Test
    void calendar_rest_paren_label_isBreak() {
        assertEquals(
                GanttScheduleSlotBarKind.BREAK,
                GanttScheduleSlotBarKind.fromTimelineCell("（休憩）"));
    }

    @Test
    void daily_startup_isStartup() {
        assertEquals(
                GanttScheduleSlotBarKind.STARTUP,
                GanttScheduleSlotBarKind.fromTimelineCell("日次始業準備"));
    }
}
