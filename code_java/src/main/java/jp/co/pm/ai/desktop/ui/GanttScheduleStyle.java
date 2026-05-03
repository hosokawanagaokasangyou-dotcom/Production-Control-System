package jp.co.pm.ai.desktop.ui;

import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * Excel \u6761\u4ef6\u4ed8\u304d\u66f8\u5f0f\u98a8\u306e\u30bf\u30a4\u30e0\u30e9\u30a4\u30f3\u8868\u793a\uff08\u8a2d\u5099\u5272\u3064\u304d\u30fb\u52e4\u52d9\u30ab\u30ec\u30f3\u30c0\u30fc\u7528\uff09\u3002
 */
public final class GanttScheduleStyle {

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
     * @param sectionRow \u884c\u982d\uff08\u65e5\u6642\u5e2f\u7b49\uff09\u306b\u65e5\u4ed8\u30bb\u30af\u30b7\u30e7\u30f3\u884c\uff08\u25a0\u3092\u542b\u3080\uff09\u306e\u3068\u304d\u5168\u5217\u306b\u9002\u7528
     */
    public static void applyTimelineCell(
            SpreadsheetCell cell, int colIndex, String columnTitle, String raw, boolean sectionRow) {
        if (cell == null) {
            return;
        }
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
        if (columnTitle != null && columnTitle.endsWith("\u9032\u5ea6")) {
            String pastel = hashPastelBackground(t + ":progress");
            cell.setStyle(
                    "-fx-background-color: "
                            + pastel
                            + "; -fx-control-inner-background: "
                            + pastel
                            + "; -fx-text-fill: #475569; -fx-font-size: 11px;");
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

    /** \u4f11\u61a9\u30fb\u52e4\u52d9\u72b6\u614b\u306a\u3069\u56fa\u5b9a\u8272\uff08\u30e1\u30f3\u30d0\u30fc\u30b7\u30fc\u30c8\u7528\uff09\u3002 */
    private static String semanticMemberOrTaskColor(String t) {
        return switch (t) {
            case "\u4f11" -> "#e2e8f0";
            case "\u52e4\u52d9\u5916" -> "#f1f5f9";
            case "\u5e74\u4f11" -> "#fef9c3";
            case "\u65e5\u6b21\u59cb\u696d\u6e96\u5099" -> "#fed7aa";
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
