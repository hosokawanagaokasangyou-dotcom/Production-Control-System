package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 備考や場所に書かれた「413行目」「東レV.C!C83」を、結果ブック内のコピーシートへのリンクにする。
 */
public final class ResultSheetAnchors {

    private static final Pattern NAGAOKA = Pattern.compile(
            "②?長岡明細\\s*(\\d+)行目(?:\\s*([A-Z]{1,2})列)?");
    private static final Pattern QUOTED = Pattern.compile(
            "「([^」]+)」\\s*(\\d+)行目(?:\\s*([A-Z]{1,2})列)?");
    private static final Pattern BANG = Pattern.compile(
            "(東レ[^\\s!↔/、。（）()]{1,24})!\\$?([A-Z]{1,2})?\\$?(\\d+)");
    private static final Pattern CSV = Pattern.compile(
            "①\\s*CSV\\s*(\\d+)(?:\\s*[〜～~]\\s*\\d+)?行目");
    private static final Pattern FILE_ROW = Pattern.compile(
            "②\\S+\\s+(\\d+)行目(?:\\s*([A-Z]{1,2})列)?");
    private static final Pattern NEAR_COL = Pattern.compile("([A-Z]{1,2})列");

    private ResultSheetAnchors() {}

    /** 文中の位置。{@code start} 以上 {@code end} 未満。 */
    public record Hit(int start, int end, String sheet, String cell) {

        public String address() {
            return "'" + sheet.replace("'", "''") + "'!" + cell;
        }
    }

    public static List<Hit> find(String text) {
        if (text == null || text.isBlank()) {
            return List.of();
        }
        List<Hit> hits = new ArrayList<>();
        collect(text, NAGAOKA, hits, m -> hit(
                m, SourceRawSheets.copiedSheetName(FactoryProfile.MATOME_SHEET),
                column(m, 2, text), m.group(1)));
        collect(text, QUOTED, hits, m -> hit(
                m, SourceRawSheets.copiedSheetName(m.group(1)),
                column(m, 3, text), m.group(2)));
        collect(text, BANG, hits, m -> {
            String col = m.group(2);
            return hit(m, SourceRawSheets.copiedSheetName(m.group(1)),
                    col == null || col.isBlank() ? "A" : col, m.group(3));
        });
        collect(text, CSV, hits, m -> hit(
                m, SourceRawSheets.CSV_SHEET, "A", m.group(1)));
        collect(text, FILE_ROW, hits, m -> hit(
                m, SourceRawSheets.copiedSheetName(FactoryProfile.MATOME_SHEET),
                column(m, 2, text), m.group(1)));
        hits.sort(Comparator.comparingInt(Hit::start).thenComparing((a, b) -> Integer.compare(b.end(), a.end())));
        List<Hit> kept = new ArrayList<>();
        int covered = -1;
        for (Hit hit : hits) {
            if (hit.start() < covered) {
                continue;
            }
            if (!valid(hit)) {
                continue;
            }
            kept.add(hit);
            covered = hit.end();
        }
        return List.copyOf(kept);
    }

    /** 位置が1件以上あれば、先頭の位置へブック内リンクを付ける。文言の位置は下線にする。 */
    public static void link(Cell cell) {
        if (!(cell instanceof XSSFCell xssf) || xssf.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {
            return;
        }
        String text = xssf.getStringCellValue();
        List<Hit> hits = find(text);
        if (hits.isEmpty()) {
            return;
        }
        applyLinkFont(xssf, text, hits);
        Hyperlink link = xssf.getSheet().getWorkbook().getCreationHelper()
                .createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress(hits.get(0).address());
        if (link instanceof org.apache.poi.xssf.usermodel.XSSFHyperlink xssfLink) {
            xssfLink.setTooltip(tooltip(hits));
        }
        xssf.setHyperlink(link);
    }

    private static void applyLinkFont(XSSFCell cell, String text, List<Hit> hits) {
        XSSFWorkbook book = cell.getSheet().getWorkbook();
        XSSFFont base = cell.getCellStyle() == null ? book.createFont() : cell.getCellStyle().getFont();
        XSSFFont linkFont = linkFont(book, base);
        XSSFRichTextString rich = new XSSFRichTextString(text);
        rich.applyFont(base);
        for (Hit hit : hits) {
            rich.applyFont(hit.start(), hit.end(), linkFont);
        }
        cell.setCellValue(rich);
    }

    private static XSSFFont linkFont(XSSFWorkbook book, XSSFFont base) {
        int fonts = book.getNumberOfFontsAsInt();
        for (int i = 0; i < fonts; i++) {
            XSSFFont existing = book.getFontAt(i);
            if (existing.getUnderline() == Font.U_SINGLE
                    && base.getFontName().equals(existing.getFontName())
                    && existing.getFontHeight() == base.getFontHeight()) {
                return existing;
            }
        }
        XSSFFont font = book.createFont();
        font.setFontName(base.getFontName());
        font.setFontHeight(base.getFontHeight());
        font.setBold(base.getBold());
        font.setUnderline(Font.U_SINGLE);
        font.setColor(new org.apache.poi.xssf.usermodel.XSSFColor(
                new byte[] {5, 99, (byte) 193}, null));
        return font;
    }

    private static String tooltip(List<Hit> hits) {
        StringBuilder sb = new StringBuilder();
        for (Hit hit : hits) {
            if (sb.length() > 0) {
                sb.append(" / ");
            }
            sb.append(hit.address());
            if (sb.length() > 240) {
                break;
            }
        }
        return sb.toString();
    }

    private interface HitMapper {
        Hit map(Matcher matcher);
    }

    private static void collect(String text, Pattern pattern, List<Hit> into, HitMapper mapper) {
        Matcher matcher = pattern.matcher(text);
        while (matcher.find()) {
            Hit hit = mapper.map(matcher);
            if (hit != null) {
                into.add(hit);
            }
        }
    }

    private static Hit hit(Matcher matcher, String sheet, String column, String rowText) {
        int row;
        try {
            row = Integer.parseInt(rowText);
        } catch (NumberFormatException ex) {
            return null;
        }
        if (row < 1 || row > 1_048_576 || sheet == null || sheet.isBlank()) {
            return null;
        }
        String col = column == null || column.isBlank() ? "A" : column;
        return new Hit(matcher.start(), matcher.end(), sheet, col + row);
    }

    private static String column(Matcher matcher, int group, String text) {
        String explicit = matcher.group(group);
        if (explicit != null && !explicit.isBlank()) {
            return explicit;
        }
        int from = Math.min(text.length(), matcher.end());
        int to = Math.min(text.length(), from + 80);
        Matcher near = NEAR_COL.matcher(text.substring(from, to));
        if (near.find()) {
            return near.group(1);
        }
        return "A";
    }

    private static boolean valid(Hit hit) {
        return hit.end() > hit.start() && hit.cell != null && !hit.cell.isBlank();
    }

    static RichTextString visible(Cell cell) {
        return cell.getRichStringCellValue();
    }
}
