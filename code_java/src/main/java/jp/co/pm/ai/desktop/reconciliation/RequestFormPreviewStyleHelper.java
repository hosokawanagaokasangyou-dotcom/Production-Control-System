package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFTextRun;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;

final class RequestFormPreviewStyleHelper {

    private static final String DEFAULT_BG = "#FFFFFF";
    private static final String DEFAULT_FG = "#000000";

    private RequestFormPreviewStyleHelper() {}

    static RequestFormPreviewCellStyle cellStyle(
            Workbook wb, CellStyle style, Map<Integer, Font> fontCache) {
        if (style == null) {
            return RequestFormPreviewCellStyle.defaults();
        }
        Font font = fontCache.computeIfAbsent(style.getFontIndexAsInt(), wb::getFontAt);
        RequestFormPreviewCellStyle base =
                fromFont(wb, font, fillHex(style), fontHex(font));
        return new RequestFormPreviewCellStyle(
                base.fontSizePx(),
                base.foreground(),
                base.background(),
                base.bold(),
                base.italic(),
                base.underline(),
                base.doubleUnderline(),
                base.strike(),
                base.doubleStrike(),
                base.fontFamily(),
                style.getVerticalAlignment() != null
                        ? style.getVerticalAlignment()
                        : VerticalAlignment.CENTER,
                style.getWrapText(),
                borderCss(style));
    }

    static RequestFormPreviewCellStyle fromFont(
            Workbook wb, Font font, String background, String foreground) {
        if (font == null) {
            return RequestFormPreviewCellStyle.defaults();
        }
        double sizePx = fontSizePx(font);
        byte underline = font.getUnderline();
        boolean underlineSingle =
                underline == Font.U_SINGLE || underline == Font.U_SINGLE_ACCOUNTING;
        boolean underlineDouble =
                underline == Font.U_DOUBLE || underline == Font.U_DOUBLE_ACCOUNTING;
        boolean strike = font.getStrikeout();
        boolean doubleStrike = strike && isDoubleStrikeFont(font);
        return new RequestFormPreviewCellStyle(
                sizePx,
                foreground != null ? foreground : DEFAULT_FG,
                background != null ? background : DEFAULT_BG,
                font.getBold(),
                font.getItalic(),
                underlineSingle || underlineDouble,
                underlineDouble,
                strike && !doubleStrike,
                doubleStrike,
                fontFamily(font),
                VerticalAlignment.CENTER,
                false,
                "");
    }

    static RequestFormPreviewCellStyle fromTextRun(XSSFTextRun run, String background) {
        if (run == null) {
            return RequestFormPreviewCellStyle.defaults();
        }
        double sizePx = run.getFontSize() > 0 ? run.getFontSize() * 96.0 / 72.0 : 11.0 * 96.0 / 72.0;
        java.awt.Color awt = run.getFontColor();
        String fg =
                awt != null
                        ? String.format(
                                Locale.ROOT, "#%02X%02X%02X", awt.getRed(), awt.getGreen(), awt.getBlue())
                        : DEFAULT_FG;
        boolean strike = run.isStrikethrough();
        boolean doubleStrike = strike && isDoubleStrikeTextRun(run);
        return new RequestFormPreviewCellStyle(
                sizePx,
                fg,
                background != null ? background : DEFAULT_BG,
                run.isBold(),
                run.isItalic(),
                run.isUnderline(),
                false,
                strike && !doubleStrike,
                doubleStrike,
                run.getFontFamily(),
                VerticalAlignment.CENTER,
                false,
                "");
    }

    static List<RequestFormPreviewTextRun> richTextRuns(
            Workbook wb, Cell cell, CellStyle style, Map<Integer, Font> fontCache) {
        if (cell == null || !hasRichTextStringValue(cell)) {
            return List.of();
        }
        RichTextString rich;
        try {
            rich = cell.getRichStringCellValue();
        } catch (IllegalStateException ex) {
            return List.of();
        }
        if (!(rich instanceof XSSFRichTextString xssfRich) || xssfRich.numFormattingRuns() <= 1) {
            return List.of();
        }
        String background = fillHex(style);
        List<RequestFormPreviewTextRun> runs = new ArrayList<>();
        for (int i = 0; i < xssfRich.numFormattingRuns(); i++) {
            int start = xssfRich.getIndexOfFormattingRun(i);
            int length = xssfRich.getLengthOfFormattingRun(i);
            if (length <= 0) {
                continue;
            }
            String text = xssfRich.getString().substring(start, start + length);
            XSSFFont runFont = xssfRich.getFontOfFormattingRun(i);
            runs.add(
                    new RequestFormPreviewTextRun(
                            text, fromFont(wb, runFont, background, fontHex(runFont))));
        }
        return runs;
    }

    static double fontSizePx(Font font) {
        double points = font.getFontHeightInPoints();
        if (points <= 0) {
            points = 11.0;
        }
        return points * 96.0 / 72.0;
    }

    static String fillHex(CellStyle style) {
        if (style == null || style.getFillPattern() == FillPatternType.NO_FILL) {
            return null;
        }
        return colorToHex(style.getFillForegroundColorColor());
    }

    static String fontHex(Font font) {
        if (font instanceof XSSFFont xssfFont) {
            String hex = colorToHex(xssfFont.getXSSFColor());
            if (hex != null) {
                return hex;
            }
        }
        return DEFAULT_FG;
    }

    static String colorToHex(org.apache.poi.ss.usermodel.Color color) {
        if (!(color instanceof XSSFColor xssfColor)) {
            return null;
        }
        byte[] rgb = xssfColor.getRGBWithTint();
        if (rgb == null || rgb.length < 3) {
            rgb = xssfColor.getRGB();
        }
        if (rgb == null || rgb.length < 3) {
            return null;
        }
        return String.format(
                Locale.ROOT, "#%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    static String colorToHex(java.awt.Color color) {
        if (color == null) {
            return null;
        }
        return String.format(
                Locale.ROOT, "#%02X%02X%02X", color.getRed(), color.getGreen(), color.getBlue());
    }

    private static String fontFamily(Font font) {
        String name = font.getFontName();
        return name != null && !name.isBlank() ? name : null;
    }

    private static boolean isDoubleStrikeFont(Font font) {
        if (!(font instanceof XSSFFont xssfFont)) {
            return false;
        }
        CTFont ct = xssfFont.getCTFont();
        if (ct == null) {
            return false;
        }
        String xml = ct.xmlText();
        return xml != null && xml.contains("dblStrike");
    }

    private static boolean isDoubleStrikeTextRun(XSSFTextRun run) {
        return false;
    }

    private static boolean hasRichTextStringValue(Cell cell) {
        CellType type = cell.getCellType();
        if (type == CellType.STRING) {
            return true;
        }
        if (type == CellType.FORMULA) {
            try {
                return cell.getCachedFormulaResultType() == CellType.STRING;
            } catch (IllegalStateException ex) {
                return false;
            }
        }
        return false;
    }

    static String borderCss(CellStyle style) {
        if (style == null) {
            return "";
        }
        double top = borderWidthPx(style.getBorderTop());
        double right = borderWidthPx(style.getBorderRight());
        double bottom = borderWidthPx(style.getBorderBottom());
        double left = borderWidthPx(style.getBorderLeft());
        if (top <= 0 && right <= 0 && bottom <= 0 && left <= 0) {
            return "";
        }
        String topColor = sideBorderColorHex(style, Side.TOP);
        String rightColor = sideBorderColorHex(style, Side.RIGHT);
        String bottomColor = sideBorderColorHex(style, Side.BOTTOM);
        String leftColor = sideBorderColorHex(style, Side.LEFT);
        return String.format(
                Locale.ROOT,
                "-fx-border-width: %.1f %.1f %.1f %.1f; -fx-border-color: %s %s %s %s;",
                top,
                right,
                bottom,
                left,
                topColor,
                rightColor,
                bottomColor,
                leftColor);
    }

    private static double borderWidthPx(BorderStyle border) {
        if (border == null || border == BorderStyle.NONE) {
            return 0.0;
        }
        return switch (border) {
            case THICK, MEDIUM, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, MEDIUM_DASHED ->
                    2.0;
            default -> 1.0;
        };
    }

    private enum Side {
        TOP,
        RIGHT,
        BOTTOM,
        LEFT
    }

    private static String sideBorderColorHex(CellStyle style, Side side) {
        if (!(style instanceof XSSFCellStyle xssfStyle)) {
            return "#000000";
        }
        XSSFColor color =
                switch (side) {
                    case TOP -> xssfStyle.getTopBorderXSSFColor();
                    case RIGHT -> xssfStyle.getRightBorderXSSFColor();
                    case BOTTOM -> xssfStyle.getBottomBorderXSSFColor();
                    case LEFT -> xssfStyle.getLeftBorderXSSFColor();
                };
        String hex = colorToHex(color);
        return hex != null ? hex : "#000000";
    }
}
