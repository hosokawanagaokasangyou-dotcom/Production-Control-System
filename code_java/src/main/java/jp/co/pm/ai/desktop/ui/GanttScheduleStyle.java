package jp.co.pm.ai.desktop.ui;

import java.util.List;

import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * Excel 条件付き書式風のタイムライン表示（設備割つき・勤務カレンダー用）。
 *
 * <p>「結果_設備ガント」参照の {@link GanttSheetKind#EQUIPMENT_TIMELINE} では、濃納見出し行、
 * 灰色の時刻軸、青系の割当ブロック、進度列の黄系エリアを再現する。
 */
public final class GanttScheduleStyle {

    /** Excel 見出し行（Dark Blue テーマ） */
    private static final String XL_SECTION_ROW_STYLE =
            "-fx-background-color: #1f4e79; -fx-control-inner-background: #1f4e79;"
                    + " -fx-text-fill: #ffffff; -fx-font-weight: bold; -fx-font-size: 11px;";

    /** 時刻列（行ラベル）— Excel 無彩填色に近い灰 */
    private static final String XL_TIME_AXIS_STYLE =
            "-fx-background-color: #d9d9d9; -fx-control-inner-background: #d9d9d9;"
                    + " -fx-text-fill: #000000; -fx-font-size: 11px;"
                    + " -fx-border-color: #bfbfbf; -fx-border-width: 0 1 0 0;";

    private static final String XL_GRID_EMPTY_LIGHT =
            "-fx-background-color: #ffffff; -fx-control-inner-background: #ffffff;"
                    + " -fx-border-color: #d9d9d9; -fx-border-width: 0.5;";

    private static final String XL_GRID_EMPTY_BAND =
            "-fx-background-color: #f2f2f2; -fx-control-inner-background: #f2f2f2;"
                    + " -fx-border-color: #d9d9d9; -fx-border-width: 0.5;";

    /** 割当テキスト（Accent 青 + 白文字 + 柴線） */
    private static final String XL_TASK_BAR_STYLE_PREFIX =
            "-fx-background-color: #5b9bd5; -fx-control-inner-background: #5b9bd5;"
                    + " -fx-text-fill: #ffffff; -fx-font-size: 10px; -fx-font-weight: bold;"
                    + " -fx-border-color: #2e5597; -fx-border-width: 0.5;";

    /** 進度列・有値 */
    private static final String XL_PROGRESS_FILLED_STYLE =
            "-fx-background-color: #fff2cc; -fx-control-inner-background: #fff2cc;"
                    + " -fx-text-fill: #333333; -fx-font-size: 10px;"
                    + " -fx-border-color: #d6b656; -fx-border-width: 0.5;";

    /** 進度列・空セル（薄い黄ソフト） */
    private static final String XL_PROGRESS_EMPTY_STYLE =
            "-fx-background-color: #fffbf0; -fx-control-inner-background: #fffbf0;"
                    + " -fx-border-color: #f0e1b7; -fx-border-width: 0.5;";

    /** 準備帯: 薄地 + 補色の斜線ハッチ + 枠（設備ガント・メンバー共通） */
    private static final String XL_PREP_STARTUP_HATCH_STYLE =
            hatchTimelineCellStyle("#fff7d6", "#b45309", "#78350f");

    private static final String XL_PREP_REQUEST_SWITCH_HATCH_STYLE =
            hatchTimelineCellStyle("#fce7f3", "#a21caf", "#701a75");

    private static final String XL_PREP_BREAK_RESUME_HATCH_STYLE =
            hatchTimelineCellStyle("#d1fae5", "#047857", "#064e3b");

    private static final String SECTION_ROW_STYLE =
            "-fx-background-color: #4472c4; -fx-control-inner-background: #4472c4;"
                    + " -fx-text-fill: #ffffff; -fx-font-weight: bold;";

    private static final String ROW_HEADER_STYLE =
            "-fx-background-color: #f1f5f9; -fx-control-inner-background: #f1f5f9;"
                    + " -fx-text-fill: #0f172a;";

    private static final String EMPTY_STYLE =
            "-fx-background-color: #ffffff; -fx-control-inner-background: #ffffff;";

    private GanttScheduleStyle() {}

    /**
     * JSON シート名と列定義から、Excel 風設備ガントを適用するか判定する。
     */
    public static GanttSheetKind resolveKind(String sheetName, List<String> columns) {
        if (columns != null && !columns.isEmpty() && "日時帯".equals(columns.get(0))) {
            return GanttSheetKind.EQUIPMENT_TIMELINE;
        }
        if (sheetName != null) {
            if (sheetName.contains("設備")
                    && (sheetName.contains("ガント")
                            || sheetName.contains("時間割"))) {
                return GanttSheetKind.EQUIPMENT_TIMELINE;
            }
        }
        return GanttSheetKind.DEFAULT;
    }

    /**
     * @param sectionRow 行頭（日時帯等）に日付セクション行（■を含む）のとき全列に適用
     * @param dataRowIndex データ行番号（0起点）。柵柵まち用
     */
    public static void applyTimelineCell(
            SpreadsheetCell cell,
            int colIndex,
            String columnTitle,
            String raw,
            boolean sectionRow,
            int dataRowIndex,
            GanttSheetKind kind) {
        if (cell == null) {
            return;
        }
        if (kind == GanttSheetKind.EQUIPMENT_TIMELINE) {
            applyEquipmentTimeline(cell, colIndex, columnTitle, raw, sectionRow, dataRowIndex);
            return;
        }
        applyDefaultTimeline(cell, colIndex, columnTitle, raw, sectionRow);
    }

    private static void applyEquipmentTimeline(
            SpreadsheetCell cell,
            int colIndex,
            String columnTitle,
            String raw,
            boolean sectionRow,
            int dataRowIndex) {
        if (sectionRow) {
            cell.setStyle(XL_SECTION_ROW_STYLE);
            return;
        }
        if (colIndex == 0) {
            cell.setStyle(XL_TIME_AXIS_STYLE);
            return;
        }
        String t = raw != null ? raw.strip() : "";
        boolean progressCol = columnTitle != null && columnTitle.endsWith("進度");
        if (t.isEmpty()) {
            if (progressCol) {
                cell.setStyle(XL_PROGRESS_EMPTY_STYLE);
            } else {
                cell.setStyle(stripeEmptyGrid(dataRowIndex));
            }
            return;
        }
        if (progressCol) {
            cell.setStyle(XL_PROGRESS_FILLED_STYLE);
            return;
        }
        String prepHatch = prepTimelineHatchStyle(t);
        if (prepHatch != null) {
            cell.setStyle(prepHatch);
            return;
        }
        cell.setStyle(XL_TASK_BAR_STYLE_PREFIX);
    }

    private static String hatchTimelineCellStyle(String bg, String line, String text) {
        return "-fx-background-color: repeating-linear-gradient(135deg, "
                + bg
                + " 0px, "
                + bg
                + " 5px, "
                + line
                + " 5px, "
                + line
                + " 6px), "
                + bg
                + "; -fx-control-inner-background: "
                + bg
                + "; -fx-text-fill: "
                + text
                + "; -fx-font-size: 10px; -fx-font-weight: bold;"
                + " -fx-border-color: "
                + line
                + "; -fx-border-width: 1;";
    }

    private static String prepTimelineHatchStyle(String t) {
        return switch (t) {
            case "日次始業準備" -> XL_PREP_STARTUP_HATCH_STYLE;
            case "依頼切替準備" -> XL_PREP_REQUEST_SWITCH_HATCH_STYLE;
            case "休憩再開準備" -> XL_PREP_BREAK_RESUME_HATCH_STYLE;
            default -> null;
        };
    }

    private static String stripeEmptyGrid(int dataRowIndex) {
        return (dataRowIndex & 1) == 0 ? XL_GRID_EMPTY_LIGHT : XL_GRID_EMPTY_BAND;
    }

    private static void applyDefaultTimeline(
            SpreadsheetCell cell, int colIndex, String columnTitle, String raw, boolean sectionRow) {
        if (sectionRow) {
            cell.setStyle(SECTION_ROW_STYLE);
            return;
        }
        if (colIndex == 0) {
            cell.setStyle(ROW_HEADER_STYLE);
            return;
        }
        String t = raw != null ? raw.strip() : "";
        if (t.isEmpty()) {
            cell.setStyle(EMPTY_STYLE);
            return;
        }
        if (columnTitle != null && columnTitle.endsWith("進度")) {
            String pastel = hashPastelBackground(t + ":progress");
            cell.setStyle(
                    "-fx-background-color: "
                            + pastel
                            + "; -fx-control-inner-background: "
                            + pastel
                            + "; -fx-text-fill: #475569; -fx-font-size: 11px;");
            return;
        }
        String prepHatch = prepTimelineHatchStyle(t);
        if (prepHatch != null) {
            cell.setStyle(prepHatch);
            return;
        }
        String semantic = semanticMemberOrTaskColor(t);
        if (semantic != null) {
            cell.setStyle(
                    "-fx-background-color: "
                            + semantic
                            + "; -fx-control-inner-background: "
                            + semantic
                            + "; -fx-text-fill: #0f172a;");
            return;
        }
        String bg = hashPastelBackground(t);
        cell.setStyle(
                "-fx-background-color: "
                        + bg
                        + "; -fx-control-inner-background: "
                        + bg
                        + "; -fx-text-fill: #0f172a;");
    }

    /** 休憩・勤務状態など固定色（メンバーシート用）。 */
    private static String semanticMemberOrTaskColor(String t) {
        return switch (t) {
            case "休" -> "#e2e8f0";
            case "勤務外" -> "#f1f5f9";
            case "年休" -> "#fef9c3";
            default -> null;
        };
    }

    private static String hashPastelBackground(String key) {
        int h = mixHash(key.hashCode());
        double hue = (h & 0xffff) % 360;
        double sat = 0.42 + (h >>> 3 & 7) * 0.02;
        double light = 0.88 + (h >>> 7 & 3) * 0.02;
        return hslToCss(hue, sat, light);
    }

    private static int mixHash(int x) {
        x ^= x >>> 16;
        x *= 0x85ebca6b;
        x ^= x >>> 13;
        x *= 0xc2b2ae35;
        x ^= x >>> 16;
        return x;
    }

    private static String hslToCss(double h, double s, double l) {
        double c = (1 - Math.abs(2 * l - 1)) * s;
        double x = c * (1 - Math.abs((h / 60) % 2 - 1));
        double m = l - c / 2;
        double rp;
        double gp;
        double bp;
        if (h < 60) {
            rp = c;
            gp = x;
            bp = 0;
        } else if (h < 120) {
            rp = x;
            gp = c;
            bp = 0;
        } else if (h < 180) {
            rp = 0;
            gp = c;
            bp = x;
        } else if (h < 240) {
            rp = 0;
            gp = x;
            bp = c;
        } else if (h < 300) {
            rp = x;
            gp = 0;
            bp = c;
        } else {
            rp = c;
            gp = 0;
            bp = x;
        }
        int r = (int) Math.round((rp + m) * 255);
        int g = (int) Math.round((gp + m) * 255);
        int b = (int) Math.round((bp + m) * 255);
        r = Math.clamp(r, 0, 255);
        g = Math.clamp(g, 0, 255);
        b = Math.clamp(b, 0, 255);
        return String.format("#%02x%02x%02x", r, g, b);
    }
}
