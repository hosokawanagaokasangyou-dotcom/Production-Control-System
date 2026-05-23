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

    @Test
    void request_switch_prep_isDedicatedKind_notDefault() {
        assertEquals(
                GanttScheduleSlotBarKind.REQUEST_SWITCH_PREP,
                GanttScheduleSlotBarKind.fromTimelineCell("依頼切替準備"));
    }

    @Test
    void break_resume_prep_isDedicatedKind_notCalendarBreak() {
        assertEquals(
                GanttScheduleSlotBarKind.BREAK_RESUME_PREP,
                GanttScheduleSlotBarKind.fromTimelineCell("休憩再開準備"));
    }

    @Test
    void post_machining_cleanup_isDedicatedKind() {
        assertEquals(
                GanttScheduleSlotBarKind.POST_MACHINING_CLEANUP,
                GanttScheduleSlotBarKind.fromTimelineCell("後始末"));
    }

    @Test
    void request_interval_buffer_isDedicatedKind() {
        assertEquals(
                GanttScheduleSlotBarKind.REQUEST_INTERVAL_BUFFER,
                GanttScheduleSlotBarKind.fromTimelineCell("依頼間余裕"));
    }
}
