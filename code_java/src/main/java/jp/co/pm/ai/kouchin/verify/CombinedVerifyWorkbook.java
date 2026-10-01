package jp.co.pm.ai.kouchin.verify;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 国分・湖南の検証ブックを1つの統合ブックへコピーする。
 * シート名は {@code 国分_} {@code 湖南_} を付ける。ブック内リンク先のシート名も合わせる。
 */
public final class CombinedVerifyWorkbook {

    public static final String FILE_PREFIX = "検証結果_統合_";
    public static final String KOKUBU_PREFIX = "国分_";
    public static final String KONAN_PREFIX = "湖南_";

    private CombinedVerifyWorkbook() {}

    /** 両方の検証結果があるときだけ統合ブックを作る。 */
    public static boolean bothVerifiable(VerifyResult kokubu, VerifyResult konan) {
        return kokubu != null && konan != null;
    }

    public static XSSFWorkbook merge(XSSFWorkbook kokubu, XSSFWorkbook konan) {
        XSSFWorkbook dest = new XSSFWorkbook();
        writeCover(dest);
        copyAll(kokubu, dest, KOKUBU_PREFIX);
        copyAll(konan, dest, KONAN_PREFIX);
        return dest;
    }

    private static void writeCover(XSSFWorkbook dest) {
        XSSFSheet ws = dest.createSheet("統合");
        Row row = ws.createRow(0);
        row.createCell(0).setCellValue(
                "国分工場と湖南工場の検証結果を1つのブックにまとめたものです。"
                        + "シート名の国分_ / 湖南_ が工場です。"
                        + "個別の検証結果Excelも別に出力しています。"
                        + "このブックは両方の工場が検証できたときだけ作ります。");
        ws.setColumnWidth(0, 120 * 256);
        row.setHeightInPoints(48f);
    }

    private static void copyAll(XSSFWorkbook src, XSSFWorkbook dest, String prefix) {
        if (src == null) {
            return;
        }
        Map<String, String> rename = new LinkedHashMap<>();
        for (int i = 0; i < src.getNumberOfSheets(); i++) {
            String oldName = src.getSheetName(i);
            rename.put(oldName, uniqueName(dest, prefix + oldName));
        }
        for (int i = 0; i < src.getNumberOfSheets(); i++) {
            XSSFSheet from = src.getSheetAt(i);
            XSSFSheet to = dest.createSheet(rename.get(from.getSheetName()));
            copySheet(from, to, dest, rename);
        }
    }

    static String rewriteDocumentAddress(String address, Map<String, String> rename) {
        if (address == null || address.isBlank() || rename == null || rename.isEmpty()) {
            return address;
        }
        String sheet;
        String rest;
        if (address.startsWith("'")) {
            StringBuilder name = new StringBuilder();
            int i = 1;
            while (i < address.length()) {
                char ch = address.charAt(i);
                if (ch == '\'' && i + 1 < address.length() && address.charAt(i + 1) == '\'') {
                    name.append('\'');
                    i += 2;
                    continue;
                }
                if (ch == '\'') {
                    i++;
                    break;
                }
                name.append(ch);
                i++;
            }
            sheet = name.toString();
            rest = address.substring(Math.min(i, address.length()));
        } else {
            int bang = address.indexOf('!');
            if (bang < 0) {
                return address;
            }
            sheet = address.substring(0, bang);
            rest = address.substring(bang);
        }
        String renamed = rename.get(sheet);
        if (renamed == null) {
            return address;
        }
        if (!rest.startsWith("!")) {
            rest = "!" + rest;
        }
        return "'" + renamed.replace("'", "''") + "'" + rest;
    }

    private static void copySheet(XSSFSheet from, XSSFSheet to, XSSFWorkbook dest, Map<String, String> rename) {
        if (from.getTabColor() != null) {
            to.setTabColor(from.getTabColor());
        }
        int maxCol = 0;
        Map<Short, CellStyle> styles = new HashMap<>();
        for (int r = 0; r <= from.getLastRowNum(); r++) {
            Row srcRow = from.getRow(r);
            if (srcRow == null) {
                continue;
            }
            Row destRow = to.createRow(r);
            destRow.setHeight(srcRow.getHeight());
            short last = srcRow.getLastCellNum();
            maxCol = Math.max(maxCol, last);
            for (int c = 0; c < last; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) {
                    continue;
                }
                Cell destCell = destRow.createCell(c);
                copyValue(srcCell, destCell);
                CellStyle copied = copyStyle(srcCell.getCellStyle(), dest, styles);
                if (copied != null) {
                    destCell.setCellStyle(copied);
                }
                copyLink(srcCell, destCell, dest, rename);
            }
        }
        for (int c = 0; c < maxCol; c++) {
            to.setColumnWidth(c, from.getColumnWidth(c));
        }
        for (CellRangeAddress region : from.getMergedRegions()) {
            to.addMergedRegion(region.copy());
        }
    }

    private static void copyValue(Cell src, Cell dest) {
        CellType type = src.getCellType();
        if (type == CellType.FORMULA) {
            try {
                type = src.getCachedFormulaResultType();
            } catch (RuntimeException ex) {
                dest.setCellValue("");
                return;
            }
        }
        switch (type) {
            case NUMERIC -> dest.setCellValue(src.getNumericCellValue());
            case BOOLEAN -> dest.setCellValue(src.getBooleanCellValue());
            case STRING -> dest.setCellValue(src.getStringCellValue());
            case BLANK -> dest.setBlank();
            default -> dest.setCellValue(src.toString());
        }
    }

    private static CellStyle copyStyle(CellStyle src, XSSFWorkbook dest, Map<Short, CellStyle> cache) {
        if (src == null) {
            return null;
        }
        return cache.computeIfAbsent((short) src.getIndex(), idx -> {
            XSSFCellStyle created = dest.createCellStyle();
            created.cloneStyleFrom(src);
            return created;
        });
    }

    private static void copyLink(Cell src, Cell dest, XSSFWorkbook destWb, Map<String, String> rename) {
        Hyperlink srcLink = src.getHyperlink();
        if (srcLink == null || srcLink.getAddress() == null) {
            return;
        }
        Hyperlink link = destWb.getCreationHelper().createHyperlink(srcLink.getType());
        String address = srcLink.getAddress();
        if (srcLink.getType() == HyperlinkType.DOCUMENT) {
            address = rewriteDocumentAddress(address, rename);
        }
        link.setAddress(address);
        dest.setHyperlink(link);
    }

    private static String uniqueName(XSSFWorkbook wb, String raw) {
        String base = raw == null || raw.isBlank() ? "sheet" : raw;
        if (base.length() > 31) {
            base = base.substring(0, 31);
        }
        String name = base;
        int i = 2;
        while (wb.getSheet(name) != null) {
            String suffix = "_" + i++;
            int keep = Math.max(1, 31 - suffix.length());
            name = base.substring(0, Math.min(base.length(), keep)) + suffix;
        }
        return name;
    }
}
