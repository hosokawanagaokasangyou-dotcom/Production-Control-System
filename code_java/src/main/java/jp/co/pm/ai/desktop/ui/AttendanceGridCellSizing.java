package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.Priority;
import javafx.scene.layout.RowConstraints;
import javafx.scene.layout.StackPane;

/** 会社カレンダー・メンバー勤怠グリッドのセル寸法。 */
public final class AttendanceGridCellSizing {

    public static final int MIN_PX = 28;
    public static final int MAX_PX = 56;
    /** 旧既定 30px より少し大きめ。 */
    public static final int DEFAULT_CELL_PX = 38;
    /** 機械カレンダー設備列の幅（行のセル高さとは独立）。 */
    public static final int MACHINE_CALENDAR_COLUMN_MIN_PX = 32;
    public static final int MACHINE_CALENDAR_COLUMN_MAX_PX = 120;
    public static final int DEFAULT_MACHINE_CALENDAR_COLUMN_PX = 64;
    /** 機械カレンダー設備列の横方向隙間（GridPane hgap）。 */
    public static final int MACHINE_CALENDAR_COLUMN_GAP_MIN_PX = 0;
    public static final int MACHINE_CALENDAR_COLUMN_GAP_MAX_PX = 24;
    public static final int DEFAULT_MACHINE_CALENDAR_COLUMN_GAP_PX = 4;

    private AttendanceGridCellSizing() {}

    public static int clamp(int px) {
        return Math.max(MIN_PX, Math.min(MAX_PX, px));
    }

    public static int clampMachineCalendarColumnWidth(int px) {
        return Math.max(
                MACHINE_CALENDAR_COLUMN_MIN_PX,
                Math.min(MACHINE_CALENDAR_COLUMN_MAX_PX, px));
    }

    public static int clampMachineCalendarColumnGap(int px) {
        return Math.max(
                MACHINE_CALENDAR_COLUMN_GAP_MIN_PX,
                Math.min(MACHINE_CALENDAR_COLUMN_GAP_MAX_PX, px));
    }

    public static int memberCellWidth(int cellPx) {
        return clamp(cellPx);
    }

    public static int memberCellHeight(int cellPx) {
        return Math.max(24, clamp(cellPx) - 4);
    }

    public static int memberNameColumnWidth(int cellPx) {
        return Math.max(72, clamp(cellPx) + 48);
    }

    public static int memberPrimaryRoleColumnWidth(int cellPx) {
        return Math.max(52, clamp(cellPx) + 28);
    }

    public static int machineEquipmentColumnWidth(int cellPx) {
        return Math.max(72, clamp(cellPx) + 34);
    }

    /** 機械カレンダー編集グリッドの列幅（セル高さとは独立）。 */
    public static int machineCalendarDataColumnWidth(int columnWidthPx) {
        return clampMachineCalendarColumnWidth(columnWidthPx);
    }

    public static void applyMachineCalendarDataCell(Button cell, int columnWidthPx, int rowCellPx) {
        int w = machineCalendarDataColumnWidth(columnWidthPx);
        int h = memberCellHeight(rowCellPx);
        cell.setMinSize(w, h);
        cell.setPrefSize(w, h);
        cell.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        cell.setStyle(buttonFontStyle(rowCellPx));
    }

    /** 機械カレンダー時刻列（HH:mm 表示用の最小幅）。 */
    public static int machineCalendarTimeColumnWidth(int cellPx) {
        return Math.max(44, memberCellWidth(cellPx) + 8);
    }

    public static void applyMachineCalendarCell(Button cell, int cellPx) {
        applyMachineCalendarDataCell(cell, cellPx, cellPx);
    }

    public static void applyMachineCalendarTimeLabel(Label label, int cellPx) {
        int w = machineCalendarTimeColumnWidth(cellPx);
        int h = memberCellHeight(cellPx);
        label.setMinSize(w, h);
        label.setPrefSize(w, h);
        label.setMaxWidth(w);
        label.setMaxHeight(Double.MAX_VALUE);
        label.setAlignment(Pos.CENTER);
        label.setStyle(headerFontStyle(cellPx));
    }

    public static int timeSlotColumnWidth(int cellPx) {
        return Math.max(56, clamp(cellPx) + 18);
    }

    public static void applyTimeSlotLabel(Label label, int cellPx) {
        int w = timeSlotColumnWidth(cellPx);
        int h = memberCellHeight(cellPx);
        label.setMinSize(w, h);
        label.setPrefSize(w, h);
        label.setMaxWidth(w);
        label.setAlignment(Pos.CENTER);
        label.setStyle(headerFontStyle(cellPx));
    }

    public static int memberDayColumnWidth(int cellPx) {
        return clamp(cellPx) + 2;
    }

    public static String buttonFontStyle(int cellPx) {
        int font = Math.max(8, Math.min(15, (int) Math.round(clamp(cellPx) * 0.32)));
        return "-fx-font-size: " + font + "px;";
    }

    public static String headerFontStyle(int cellPx) {
        int font = Math.max(8, Math.min(13, (int) Math.round(clamp(cellPx) * 0.28)));
        return "-fx-font-size: " + font + "px;";
    }

    public static void applySquareCell(Button cell, int cellPx) {
        int px = clamp(cellPx);
        cell.setMinSize(px, px);
        cell.setPrefSize(px, px);
        cell.setMaxSize(px, px);
        cell.setStyle(buttonFontStyle(px));
    }

    public static void applyMemberCell(Button cell, int cellPx) {
        int w = memberCellWidth(cellPx);
        int h = memberCellHeight(cellPx);
        cell.setMinSize(w, h);
        cell.setPrefSize(w, h);
        cell.setMaxSize(w, h);
        cell.setStyle(buttonFontStyle(cellPx));
    }

    public static void applyHeaderLabel(Label label, int cellPx) {
        label.setStyle(headerFontStyle(cellPx));
    }

    /** 左右グリッドで同一行高にするための行制約。 */
    public static RowConstraints memberRowConstraints(int cellPx) {
        int h = memberCellHeight(cellPx);
        RowConstraints rc = new RowConstraints();
        rc.setMinHeight(h);
        rc.setPrefHeight(h);
        rc.setMaxHeight(h);
        rc.setValignment(VPos.CENTER);
        rc.setVgrow(Priority.NEVER);
        return rc;
    }

    public static void applyMemberCellWrap(StackPane wrap, int cellPx) {
        int w = memberDayColumnWidth(cellPx);
        int h = memberCellHeight(cellPx);
        wrap.setMinSize(w, h);
        wrap.setPrefSize(w, h);
        wrap.setMaxSize(w, h);
    }

    /** メンバー名列の見出しセル（行の勤怠セル高さに揃える）。 */
    public static void applyMemberNameLabel(Label label, int cellPx) {
        applyFixedRowLabel(label, memberNameColumnWidth(cellPx), cellPx);
    }

    /** 主担当列の見出しセル（氏名列と同じ行高、列幅は主担当用）。 */
    public static void applyMemberRoleLabel(Label label, int cellPx) {
        applyFixedRowLabel(label, memberPrimaryRoleColumnWidth(cellPx), cellPx);
    }

    private static void applyFixedRowLabel(Label label, int widthPx, int cellPx) {
        int h = memberCellHeight(cellPx);
        label.setMinSize(widthPx, h);
        label.setPrefSize(widthPx, h);
        label.setMaxSize(widthPx, h);
        label.setAlignment(Pos.CENTER);
        label.setWrapText(false);
        label.setStyle(headerFontStyle(cellPx));
    }
}
