package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import javafx.application.Platform;
import javafx.scene.control.Label;
import javafx.scene.layout.Priority;
import javafx.scene.layout.RowConstraints;
import javafx.scene.layout.StackPane;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class AttendanceGridCellSizingTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void applyMemberNameLabel_locksHeightAndDoesNotWrap() {
        Label label = new Label("春樹 真由美");
        int px = AttendanceGridCellSizing.DEFAULT_CELL_PX;
        AttendanceGridCellSizing.applyMemberNameLabel(label, px);
        int h = AttendanceGridCellSizing.memberCellHeight(px);
        int w = AttendanceGridCellSizing.memberNameColumnWidth(px);
        assertEquals(h, label.getMinHeight(), 0.01);
        assertEquals(h, label.getPrefHeight(), 0.01);
        assertEquals(h, label.getMaxHeight(), 0.01);
        assertEquals(w, label.getMaxWidth(), 0.01);
        assertFalse(label.isWrapText());
    }

    @Test
    void applyMemberRoleLabel_usesRoleColumnWidth() {
        Label label = new Label("後加工");
        int px = AttendanceGridCellSizing.DEFAULT_CELL_PX;
        AttendanceGridCellSizing.applyMemberRoleLabel(label, px);
        int h = AttendanceGridCellSizing.memberCellHeight(px);
        int w = AttendanceGridCellSizing.memberPrimaryRoleColumnWidth(px);
        assertEquals(h, label.getMaxHeight(), 0.01);
        assertEquals(w, label.getPrefWidth(), 0.01);
        assertFalse(label.isWrapText());
    }

    @Test
    void memberRowConstraints_lockHeight() {
        int px = AttendanceGridCellSizing.DEFAULT_CELL_PX;
        int h = AttendanceGridCellSizing.memberCellHeight(px);
        RowConstraints rc = AttendanceGridCellSizing.memberRowConstraints(px);
        assertEquals(h, rc.getMinHeight(), 0.01);
        assertEquals(h, rc.getPrefHeight(), 0.01);
        assertEquals(h, rc.getMaxHeight(), 0.01);
        assertEquals(Priority.NEVER, rc.getVgrow());
    }

    @Test
    void applyMemberCellWrap_locksSizeToDayCell() {
        StackPane wrap = new StackPane();
        int px = AttendanceGridCellSizing.DEFAULT_CELL_PX;
        AttendanceGridCellSizing.applyMemberCellWrap(wrap, px);
        assertEquals(AttendanceGridCellSizing.memberDayColumnWidth(px), wrap.getPrefWidth(), 0.01);
        assertEquals(AttendanceGridCellSizing.memberCellHeight(px), wrap.getPrefHeight(), 0.01);
        assertEquals(AttendanceGridCellSizing.memberCellHeight(px), wrap.getMaxHeight(), 0.01);
    }
}
