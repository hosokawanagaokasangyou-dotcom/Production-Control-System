package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;

/** 会社カレンダー・メンバー勤怠グリッドのセル寸法。 */
public final class AttendanceGridCellSizing {

    public static final int MIN_PX = 28;
    public static final int MAX_PX = 56;
    /** 旧既定 30px より少し大きめ。 */
    public static final int DEFAULT_CELL_PX = 38;

    private AttendanceGridCellSizing() {}

    public static int clamp(int px) {
        return Math.max(MIN_PX, Math.min(MAX_PX, px));
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

    /** メンバー名列の見出しセル（行の勤怠セル高さに揃える）。 */
    public static void applyMemberNameLabel(Label label, int cellPx) {
        int h = memberCellHeight(cellPx);
        int w = memberNameColumnWidth(cellPx);
        label.setMinSize(w, h);
        label.setPrefSize(w, h);
        label.setMaxWidth(w);
        label.setAlignment(Pos.CENTER);
        label.setWrapText(true);
        label.setStyle(headerFontStyle(cellPx));
    }
}
