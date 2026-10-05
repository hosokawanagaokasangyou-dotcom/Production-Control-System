package jp.co.pm.ai.desktop.io;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.UnaryOperator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;

/**
 * 同一化チェック履歴に保存した配台計画 Excel の上段（現アラ計）を、チェックに使った加工計画の値で書き直す。
 *
 * <p>差異セル（シス計あり かつ 加工計画と不一致）は {@link #DIFF_FILL_RGB} で塗る。判定条件は
 * {@link AladdinEntryDispatchPlanIdentityCheck#compare} と同一で、塗ったセル数は差異件数と一致する。
 * 下段（シス計）は変更しない。
 */
public final class IdentityCheckHistoryWorkbookAnnotator {

    /** 差異セルの背景（赤系 #FFC7CE）。 */
    static final byte[] DIFF_FILL_RGB = new byte[] {(byte) 0xFF, (byte) 0xC7, (byte) 0xCE};

    private static final byte[] WEEKEND_FILL_RGB = new byte[] {(byte) 0xF2, (byte) 0xF2, (byte) 0xF2};

    private static final double EPS = DispatchAladdinEntrySheetBuilder.QTY_MATCH_EPS;

    private IdentityCheckHistoryWorkbookAnnotator() {}

    private enum CellKind {
        DIFF,
        WEEKDAY,
        WEEKEND
    }

    private record LineFonts(Font aladdin, Font system) {}

    /**
     * @param sheetToMachine シート名 → 加工計画の機械名
     * @return 差異として塗ったセル数
     */
    public static int rewriteAladdinLine(
            Path xlsx,
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup,
            LocalDate referenceDate,
            UnaryOperator<String> sheetToMachine)
            throws IOException {
        if (xlsx == null || !Files.isRegularFile(xlsx)) {
            throw new IOException("配台計画 Excel がありません: " + xlsx);
        }
        Map<String, Map<String, Map<String, Map<String, Double>>>> plan =
                lookup != null ? lookup : Map.of();
        LocalDate ref = referenceDate != null ? referenceDate : LocalDate.now();
        UnaryOperator<String> toMachine = sheetToMachine != null ? sheetToMachine : s -> s;
        int diffCells = 0;
        byte[] bytes = Files.readAllBytes(xlsx);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Map<String, CellStyle> styleCache = new HashMap<>();
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                diffCells += rewriteSheet(wb, wb.getSheetAt(s), plan, ref, toMachine, styleCache);
            }
            Path tmp = xlsx.resolveSibling(xlsx.getFileName() + ".tmp");
            try {
                try (OutputStream out = Files.newOutputStream(tmp)) {
                    wb.write(out);
                }
                Files.move(tmp, xlsx, StandardCopyOption.REPLACE_EXISTING);
            } finally {
                Files.deleteIfExists(tmp);
            }
        }
        return diffCells;
    }

    private static int rewriteSheet(
            XSSFWorkbook wb,
            Sheet sh,
            Map<String, Map<String, Map<String, Map<String, Double>>>> plan,
            LocalDate ref,
            UnaryOperator<String> toMachine,
            Map<String, CellStyle> styleCache) {
        if (sh == null) {
            return 0;
        }
        String sheetName = sh.getSheetName();
        if (sheetName == null || sheetName.isBlank() || "データなし".equals(sheetName)) {
            return 0;
        }
        Row header = sh.getRow(0);
        if (header == null) {
            return 0;
        }
        int tidCol =
                AladdinEntryDispatchPlanWorkbookReader.findHeaderCol(
                        header, AladdinEntryDispatchPlanWorkbookReader.COL_TID);
        int procCol =
                AladdinEntryDispatchPlanWorkbookReader.findHeaderCol(
                        header, AladdinEntryDispatchPlanWorkbookReader.COL_PROCESS);
        if (tidCol < 0) {
            return 0;
        }
        Map<Integer, LocalDate> dateCols =
                AladdinEntryDispatchPlanWorkbookReader.dateColumns(header, ref);
        if (dateCols.isEmpty()) {
            return 0;
        }
        String machine = toMachine.apply(sheetName.strip());

        Row totalRow = null;
        LineFonts dataFonts = null;
        LineFonts totalFonts = null;
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            boolean isTotal = isTotalRow(row, tidCol);
            if (isTotal) {
                totalRow = row;
            }
            for (int c : dateCols.keySet()) {
                LineFonts f = lineFonts(row.getCell(c));
                if (f == null) {
                    continue;
                }
                if (isTotal && totalFonts == null) {
                    totalFonts = f;
                } else if (!isTotal && dataFonts == null) {
                    dataFonts = f;
                }
            }
        }

        Map<Integer, double[]> totals = new LinkedHashMap<>();
        int diffCells = 0;
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null || isTotalRow(row, tidCol)) {
                continue;
            }
            String tid = ExcelCellReadSupport.cellToDisplayString(row.getCell(tidCol)).strip();
            if (tid.isEmpty()) {
                continue;
            }
            String process =
                    procCol >= 0
                            ? ExcelCellReadSupport.cellToDisplayString(row.getCell(procCol)).strip()
                            : "";
            String tidKey = JuchuTransferValueNormalizer.normalizeKey(tid);
            for (Map.Entry<Integer, LocalDate> e : dateCols.entrySet()) {
                int c = e.getKey();
                LocalDate d = e.getValue();
                Cell cell = row.getCell(c);
                String oldText = cell != null ? ExcelCellReadSupport.cellToDisplayString(cell) : "";
                double system = AladdinEntryDispatchPlanWorkbookReader.parseSystemQty(oldText);
                double planQty =
                        AladdinShapedPlanQtyLookup.lookup(plan, machine, tidKey, isoSlash(d), process);
                double[] sum = totals.computeIfAbsent(c, k -> new double[2]);
                sum[0] += planQty;
                sum[1] += system;

                DispatchAladdinEntrySheetBuilder.EntryCell ec =
                        new DispatchAladdinEntrySheetBuilder.EntryCell(planQty, system);
                if (ec.isEmpty() && oldText.isBlank()) {
                    continue;
                }
                if (cell == null) {
                    cell = row.createCell(c);
                }
                boolean diff =
                        Math.abs(system) > EPS && Math.abs(planQty - system) > EPS;
                if (diff) {
                    diffCells++;
                }
                CellKind kind = diff ? CellKind.DIFF : isWeekend(d) ? CellKind.WEEKEND : CellKind.WEEKDAY;
                LineFonts own = lineFonts(cell);
                cell.setCellStyle(derivedStyle(wb, cell.getCellStyle(), kind, styleCache));
                writeTwoLineText(cell, ec, own != null ? own : dataFonts);
            }
        }

        if (totalRow != null) {
            for (Map.Entry<Integer, double[]> e : totals.entrySet()) {
                double[] sum = e.getValue();
                DispatchAladdinEntrySheetBuilder.EntryCell ec =
                        new DispatchAladdinEntrySheetBuilder.EntryCell(sum[0], sum[1]);
                Cell cell = totalRow.getCell(e.getKey());
                if (cell == null) {
                    if (ec.isEmpty()) {
                        continue;
                    }
                    cell = totalRow.createCell(e.getKey());
                }
                writeTwoLineText(cell, ec, totalFonts);
            }
        }
        return diffCells;
    }

    private static boolean isTotalRow(Row row, int tidCol) {
        return DispatchAladdinEntryWorkbookExporter.DAILY_PROCESSING_TOTAL_LABEL.equals(
                ExcelCellReadSupport.cellToDisplayString(row.getCell(tidCol)).strip());
    }

    private static void writeTwoLineText(
            Cell cell, DispatchAladdinEntrySheetBuilder.EntryCell ec, LineFonts fonts) {
        String text = ec.cellText();
        if (text.isEmpty()) {
            cell.setCellValue("");
            return;
        }
        XSSFRichTextString rich =
                fonts != null
                        ? DispatchAladdinEntryWorkbookExporter.buildDateCellRichText(
                                text, fonts.aladdin(), fonts.system())
                        : null;
        if (rich != null) {
            cell.setCellValue(rich);
        } else {
            cell.setCellValue(text);
        }
    }

    private static LineFonts lineFonts(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.STRING) {
            return null;
        }
        if (!(cell.getRichStringCellValue() instanceof XSSFRichTextString rt)) {
            return null;
        }
        String text = rt.getString();
        int nl = text != null ? text.indexOf('\n') : -1;
        if (nl <= 0 || nl + 1 >= text.length()) {
            return null;
        }
        Font upper = rt.getFontAtIndex(0);
        Font lower = rt.getFontAtIndex(nl + 1);
        if (upper == null || lower == null) {
            return null;
        }
        return new LineFonts(upper, lower);
    }

    private static CellStyle derivedStyle(
            XSSFWorkbook wb, CellStyle base, CellKind kind, Map<String, CellStyle> cache) {
        String key = (base != null ? base.getIndex() : -1) + ":" + kind;
        return cache.computeIfAbsent(
                key,
                k -> {
                    XSSFCellStyle s = wb.createCellStyle();
                    if (base != null) {
                        s.cloneStyleFrom(base);
                    }
                    switch (kind) {
                        case DIFF -> fill(s, DIFF_FILL_RGB);
                        case WEEKEND -> fill(s, WEEKEND_FILL_RGB);
                        case WEEKDAY -> s.setFillPattern(FillPatternType.NO_FILL);
                    }
                    return s;
                });
    }

    private static void fill(XSSFCellStyle style, byte[] rgb) {
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(new XSSFColor(rgb, null));
    }

    private static boolean isWeekend(LocalDate d) {
        return d.getDayOfWeek() == DayOfWeek.SATURDAY || d.getDayOfWeek() == DayOfWeek.SUNDAY;
    }

    private static String isoSlash(LocalDate d) {
        return String.format("%04d/%02d/%02d", d.getYear(), d.getMonthValue(), d.getDayOfMonth());
    }
}
