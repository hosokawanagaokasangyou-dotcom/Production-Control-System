package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 原本コピーなどの付帯シート先頭に、検証シートへ戻るリンクを置く。
 * 行を1行ずらすので、付帯シートを指す既存リンクの行番号も1つ下げる。
 */
public final class AttachedSheetNav {

    public record Target(String label, String sheet) {}

    private static final Pattern ADDRESS = Pattern.compile("^'([^']+)'!([A-Z]+)(\\d+)$");

    private AttachedSheetNav() {}

    public static void apply(XSSFWorkbook book, List<Target> targets) {
        if (book == null || targets == null || targets.isEmpty()) {
            return;
        }
        List<Target> present = new ArrayList<>();
        Set<String> primary = new HashSet<>();
        for (Target target : targets) {
            if (target == null || target.sheet() == null) {
                continue;
            }
            primary.add(target.sheet());
            if (book.getSheet(target.sheet()) != null) {
                present.add(target);
            }
        }
        if (present.isEmpty()) {
            return;
        }
        primary.add("サマリ");
        primary.add("レポートの見方");
        primary.add(NagaokaExcelFixAdvice.SHEET);
        primary.add("報告メール下書き");
        List<String> shifted = new ArrayList<>();
        XSSFCellStyle linkStyle = linkStyle(book);
        for (int i = 0; i < book.getNumberOfSheets(); i++) {
            XSSFSheet sheet = book.getSheetAt(i);
            if (primary.contains(sheet.getSheetName())) {
                continue;
            }
            insertNav(sheet, present, linkStyle);
            shifted.add(sheet.getSheetName());
        }
        bumpLinks(book, new HashSet<>(shifted));
    }

    /** 付帯シート先頭行を挿入した分、リンク先の行番号を1つ増やす。 */
    static String bump(String address) {
        if (address == null) {
            return null;
        }
        Matcher matcher = ADDRESS.matcher(address);
        if (!matcher.matches()) {
            return address;
        }
        int row = Integer.parseInt(matcher.group(3));
        return "'" + matcher.group(1).replace("'", "''") + "'!" + matcher.group(2) + (row + 1);
    }

    private static void insertNav(XSSFSheet sheet, List<Target> targets, XSSFCellStyle linkStyle) {
        if (sheet.getPhysicalNumberOfRows() > 0) {
            sheet.shiftRows(0, sheet.getLastRowNum(), 1);
        }
        Row row = sheet.createRow(0);
        XSSFWorkbook book = sheet.getWorkbook();
        for (int i = 0; i < targets.size(); i++) {
            Target target = targets.get(i);
            Cell cell = row.createCell(i);
            cell.setCellValue(target.label());
            cell.setCellStyle(linkStyle);
            Hyperlink link = book.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
            link.setAddress("'" + target.sheet().replace("'", "''") + "'!A1");
            cell.setHyperlink(link);
        }
        sheet.createFreezePane(0, 1);
    }

    private static void bumpLinks(XSSFWorkbook book, Set<String> shifted) {
        for (int i = 0; i < book.getNumberOfSheets(); i++) {
            XSSFSheet sheet = book.getSheetAt(i);
            for (Row row : sheet) {
                if (row == null) {
                    continue;
                }
                for (Cell cell : row) {
                    Hyperlink link = cell.getHyperlink();
                    if (link == null || link.getType() != HyperlinkType.DOCUMENT) {
                        continue;
                    }
                    String address = link.getAddress();
                    Matcher matcher = ADDRESS.matcher(address == null ? "" : address);
                    if (!matcher.matches() || !shifted.contains(matcher.group(1))) {
                        continue;
                    }
                    link.setAddress(bump(address));
                    if (link instanceof XSSFHyperlink xssfLink && xssfLink.getTooltip() != null) {
                        xssfLink.setTooltip(bumpAll(xssfLink.getTooltip(), shifted));
                    }
                }
            }
        }
    }

    private static String bumpAll(String text, Set<String> shifted) {
        Matcher matcher = Pattern.compile("'([^']+)'!([A-Z]+)(\\d+)").matcher(text);
        StringBuilder sb = new StringBuilder();
        while (matcher.find()) {
            String sheet = matcher.group(1);
            String replacement = matcher.group();
            if (shifted.contains(sheet)) {
                int row = Integer.parseInt(matcher.group(3)) + 1;
                replacement = "'" + sheet + "'!" + matcher.group(2) + row;
            }
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    private static XSSFCellStyle linkStyle(XSSFWorkbook book) {
        XSSFCellStyle style = book.createCellStyle();
        XSSFFont font = book.createFont();
        font.setUnderline(Font.U_SINGLE);
        font.setColor(new XSSFColor(new byte[] {5, 99, (byte) 193}, null));
        style.setFont(font);
        return style;
    }
}
