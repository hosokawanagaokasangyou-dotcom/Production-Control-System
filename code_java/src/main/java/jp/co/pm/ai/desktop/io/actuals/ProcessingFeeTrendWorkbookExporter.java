package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.RequestPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

/**
 * 加工賃トレンド画面の日別・依頼NO別・AO実績相違を Excel 出力する。
 *
 * <p>加工量トレンド出力と同系の書式（BIZ UDPゴシック、見出し色、土日色、カンマ区切り、列幅）。
 */
public final class ProcessingFeeTrendWorkbookExporter {

    private static final DateTimeFormatter FILE_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final String[] WEEKDAY_JA = {"月", "火", "水", "木", "金", "土", "日"};

    private ProcessingFeeTrendWorkbookExporter() {}

    public static String suggestFileName(LocalDate from, LocalDate to, LocalDateTime now) {
        LocalDateTime t = now != null ? now : LocalDateTime.now();
        String a = from != null ? from.toString() : "from";
        String b = to != null ? to.toString() : "to";
        return "加工賃トレンド_" + a + "_" + b + "_" + t.format(FILE_TS) + ".xlsx";
    }

    public static void write(Result result, Path target) throws IOException {
        if (target == null) {
            throw new IllegalArgumentException("target");
        }
        Result r = result != null ? result : emptyResult();
        if (target.getParent() != null) {
            Files.createDirectories(target.getParent());
        }
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles s = new Styles(wb);
            writeDays(wb.createSheet("日別"), s, r);
            writeRequests(
                    wb.createSheet("依頼NO別"),
                    s,
                    ProcessingFeeTrendAggregator.withLeadingTotalRow(r.requests()));
            writeMismatches(
                    wb.createSheet("AO実績相違"),
                    s,
                    ProcessingFeeTrendAggregator.aoActualMismatchesOnOrBeforeToday(
                            r.requests(), r.to(), r.today()));
            try (OutputStream out = Files.newOutputStream(target)) {
                wb.write(out);
            }
        }
    }

    private static Result emptyResult() {
        return new Result(
                List.of(), List.of(), 0, 0, LocalDate.now(), LocalDate.now(), LocalDate.now(), 0, 0, 0);
    }

    private static void writeDays(Sheet sh, Styles s, Result result) {
        configureSheet(sh);
        sh.createFreezePane(0, 1);
        String[] headers = {"日付", "曜日", "実績", "未了", "実績累計", "未了累計", "見込累計"};
        Row head = sh.createRow(0);
        head.setHeightInPoints(22);
        writeHeader(head, s, headers);

        int row = 1;
        LocalDate today = result.today();
        List<DayPoint> days = result.days();
        double sumActual = 0;
        double sumRemain = 0;
        if (days != null) {
            for (DayPoint d : days) {
                if (d == null || d.date() == null) {
                    continue;
                }
                Row r = sh.createRow(row++);
                r.setHeightInPoints(18);
                CellStyle dateStyle = s.dateStyle(d.date(), today);
                CellStyle textStyle = s.textStyle(d.date(), today);
                CellStyle yenStyle = s.yenStyle(d.date(), today);
                int c = 0;
                setText(r, c++, d.date().format(DATE_FMT), dateStyle);
                setText(r, c++, weekdayJa(d.date()), textStyle);
                setNumber(r, c++, d.actualYen(), yenStyle);
                setNumber(r, c++, d.planYen(), yenStyle);
                setNumber(r, c++, d.actualCumYen(), yenStyle);
                setNumber(r, c++, d.planCumYen(), yenStyle);
                setNumber(r, c++, d.projectedCumYen(), yenStyle);
                sumActual += d.actualYen();
                sumRemain += d.planYen();
            }
        }
        int lastDataRow = row - 1;
        Row total = sh.createRow(row++);
        total.setHeightInPoints(20);
        setText(total, 0, "合計", s.totalText);
        setText(total, 1, "", s.totalText);
        setNumber(total, 2, sumActual, s.totalYen);
        setNumber(total, 3, sumRemain, s.totalYen);
        setText(total, 4, "", s.totalText);
        setText(total, 5, "", s.totalText);
        setText(total, 6, "", s.totalText);
        finishTable(sh, headers.length, lastDataRow + 1);
        sh.setColumnWidth(0, 15 * 256);
        sh.setColumnWidth(1, 8 * 256);
        for (int i = 2; i < headers.length; i++) {
            sh.setColumnWidth(i, 16 * 256);
        }
        if (lastDataRow >= 1 && sh instanceof XSSFSheet xssf) {
            ProcessingTrendWorkbookCharts.addLineChart(
                    xssf,
                    "加工賃トレンド（累計）",
                    headers.length + 1,
                    0,
                    headers.length + 13,
                    18,
                    0,
                    new int[] {4, 6},
                    new String[] {"実績累計", "見込累計"},
                    1,
                    lastDataRow,
                    "円");
        }
    }

    private static void writeRequests(Sheet sh, Styles s, List<RequestPoint> requests) {
        configureSheet(sh);
        sh.createFreezePane(0, 1);
        String[] headers = {"依頼NO", "AO(受注額)", "円/m", "実績 (m)", "未了 (m)", "実績", "未了"};
        Row head = sh.createRow(0);
        head.setHeightInPoints(22);
        writeHeader(head, s, headers);

        int row = 1;
        if (requests != null) {
            for (RequestPoint p : requests) {
                if (p == null) {
                    continue;
                }
                boolean total = p.isTotalRow();
                Row r = sh.createRow(row++);
                r.setHeightInPoints(18);
                CellStyle text = total ? s.totalText : s.text;
                CellStyle yen = total ? s.totalYen : s.yen;
                CellStyle meters = total ? s.totalMeters : s.meters;
                int c = 0;
                setText(r, c++, p.requestNo(), text);
                setNumber(r, c++, p.aoYen(), yen);
                if (p.rateMissing()) {
                    setText(r, c++, "—", text);
                } else {
                    setNumber(r, c++, p.rateYenPerM(), yen);
                }
                setNumber(r, c++, p.actualMeters(), meters);
                setNumber(r, c++, p.planMeters(), meters);
                setNumber(r, c++, p.actualYen(), yen);
                setNumber(r, c++, p.planYen(), yen);
            }
        }
        finishTable(sh, headers.length, row);
        sh.setColumnWidth(0, 16 * 256);
        for (int i = 1; i < headers.length; i++) {
            sh.setColumnWidth(i, 16 * 256);
        }
    }

    private static void writeMismatches(Sheet sh, Styles s, List<RequestPoint> requests) {
        configureSheet(sh);
        sh.createFreezePane(0, 1);
        String[] headers = {"依頼NO", "AO(受注額)", "実績", "差額", "実績 (m)", "未了 (m)", "未了"};
        Row head = sh.createRow(0);
        head.setHeightInPoints(22);
        writeHeader(head, s, headers);

        int row = 1;
        if (requests != null) {
            for (RequestPoint p : requests) {
                if (p == null) {
                    continue;
                }
                Row r = sh.createRow(row++);
                r.setHeightInPoints(18);
                double diff = p.aoYen() - p.actualYen();
                int c = 0;
                setText(r, c++, p.requestNo(), s.text);
                setNumber(r, c++, p.aoYen(), s.yen);
                setNumber(r, c++, p.actualYen(), s.yen);
                setNumber(r, c++, diff, diff >= 0 ? s.diffPos : s.diffNeg);
                setNumber(r, c++, p.actualMeters(), s.meters);
                setNumber(r, c++, p.planMeters(), s.meters);
                setNumber(r, c++, p.planYen(), s.yen);
            }
        }
        finishTable(sh, headers.length, row);
        sh.setColumnWidth(0, 16 * 256);
        for (int i = 1; i < headers.length; i++) {
            sh.setColumnWidth(i, 16 * 256);
        }
    }

    private static void writeHeader(Row head, Styles s, String[] headers) {
        for (int i = 0; i < headers.length; i++) {
            Cell c = head.createCell(i);
            c.setCellValue(headers[i]);
            c.setCellStyle(s.header);
        }
    }

    private static void finishTable(Sheet sh, int colCount, int rowCount) {
        int lastRow = Math.max(0, rowCount - 1);
        int lastCol = Math.max(0, colCount - 1);
        sh.setAutoFilter(new CellRangeAddress(0, lastRow, 0, lastCol));
        sh.setRepeatingRows(new CellRangeAddress(0, 0, 0, lastCol));
    }

    private static void configureSheet(Sheet sh) {
        PrintSetup ps = sh.getPrintSetup();
        ps.setLandscape(true);
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        sh.setAutobreaks(true);
        sh.setDisplayGridlines(false);
    }

    private static void setText(Row r, int col, String value, CellStyle style) {
        Cell c = r.createCell(col);
        c.setCellValue(value != null ? value : "");
        c.setCellStyle(style);
    }

    private static void setNumber(Row r, int col, double value, CellStyle style) {
        Cell c = r.createCell(col);
        c.setCellValue(value);
        c.setCellStyle(style);
    }

    private static String weekdayJa(LocalDate d) {
        return WEEKDAY_JA[d.getDayOfWeek().getValue() - 1];
    }

    private static final class Styles {
        private static final byte[] HEADER = new byte[] {(byte) 0xD9, (byte) 0xE1, (byte) 0xF2};
        private static final byte[] SAT = new byte[] {(byte) 0xEB, (byte) 0xF2, (byte) 0xFA};
        private static final byte[] SUN = new byte[] {(byte) 0xFD, (byte) 0xEC, (byte) 0xEC};
        private static final byte[] TODAY = new byte[] {(byte) 0xFE, (byte) 0xF3, (byte) 0xC7};
        private static final byte[] TOTAL = new byte[] {(byte) 0xE2, (byte) 0xE8, (byte) 0xF0};
        private static final byte[] POS = new byte[] {(byte) 0x15, (byte) 0x80, (byte) 0x3D};
        private static final byte[] NEG = new byte[] {(byte) 0xB9, (byte) 0x1C, (byte) 0x1C};

        final CellStyle header;
        final CellStyle text;
        final CellStyle yen;
        final CellStyle meters;
        final CellStyle totalText;
        final CellStyle totalYen;
        final CellStyle totalMeters;
        final CellStyle diffPos;
        final CellStyle diffNeg;
        private final CellStyle date;
        private final CellStyle dateSat;
        private final CellStyle dateSun;
        private final CellStyle dateToday;
        private final CellStyle textSat;
        private final CellStyle textSun;
        private final CellStyle textToday;
        private final CellStyle yenSat;
        private final CellStyle yenSun;
        private final CellStyle yenToday;

        Styles(XSSFWorkbook wb) {
            String fontName = DispatchAladdinEntryWorkbookExporter.DEFAULT_WORKBOOK_FONT_FAMILY;
            DataFormat df = wb.createDataFormat();
            short yenFmt = df.getFormat("#,##0");
            short mFmt = df.getFormat("#,##0.0");

            XSSFFont base = wb.createFont();
            base.setFontName(fontName);
            base.setFontHeightInPoints((short) 10);

            XSSFFont bold = wb.createFont();
            bold.setFontName(fontName);
            bold.setFontHeightInPoints((short) 10);
            bold.setBold(true);

            XSSFFont pos = wb.createFont();
            pos.setFontName(fontName);
            pos.setFontHeightInPoints((short) 10);
            pos.setBold(true);
            pos.setColor(new XSSFColor(POS, null));

            XSSFFont neg = wb.createFont();
            neg.setFontName(fontName);
            neg.setFontHeightInPoints((short) 10);
            neg.setBold(true);
            neg.setColor(new XSSFColor(NEG, null));

            header = boxed(wb, bold);
            header.setFillForegroundColor(new XSSFColor(HEADER, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            header.setAlignment(HorizontalAlignment.CENTER);

            text = boxed(wb, base);
            yen = boxed(wb, base);
            yen.setAlignment(HorizontalAlignment.RIGHT);
            yen.setDataFormat(yenFmt);
            meters = boxed(wb, base);
            meters.setAlignment(HorizontalAlignment.RIGHT);
            meters.setDataFormat(mFmt);

            totalText = boxed(wb, bold);
            totalText.setFillForegroundColor(new XSSFColor(TOTAL, null));
            totalText.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            totalYen = boxed(wb, bold);
            totalYen.setAlignment(HorizontalAlignment.RIGHT);
            totalYen.setDataFormat(yenFmt);
            totalYen.setFillForegroundColor(new XSSFColor(TOTAL, null));
            totalYen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            totalMeters = boxed(wb, bold);
            totalMeters.setAlignment(HorizontalAlignment.RIGHT);
            totalMeters.setDataFormat(mFmt);
            totalMeters.setFillForegroundColor(new XSSFColor(TOTAL, null));
            totalMeters.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            diffPos = boxed(wb, pos);
            diffPos.setAlignment(HorizontalAlignment.RIGHT);
            diffPos.setDataFormat(yenFmt);
            diffNeg = boxed(wb, neg);
            diffNeg.setAlignment(HorizontalAlignment.RIGHT);
            diffNeg.setDataFormat(yenFmt);

            date = boxed(wb, base);
            date.setAlignment(HorizontalAlignment.CENTER);
            dateSat = fill(boxed(wb, base), SAT);
            dateSat.setAlignment(HorizontalAlignment.CENTER);
            dateSun = fill(boxed(wb, base), SUN);
            dateSun.setAlignment(HorizontalAlignment.CENTER);
            dateToday = fill(boxed(wb, bold), TODAY);
            dateToday.setAlignment(HorizontalAlignment.CENTER);

            textSat = fill(boxed(wb, base), SAT);
            textSat.setAlignment(HorizontalAlignment.CENTER);
            textSun = fill(boxed(wb, base), SUN);
            textSun.setAlignment(HorizontalAlignment.CENTER);
            textToday = fill(boxed(wb, bold), TODAY);

            yenSat = fill(boxed(wb, base), SAT);
            yenSat.setAlignment(HorizontalAlignment.RIGHT);
            yenSat.setDataFormat(yenFmt);
            yenSun = fill(boxed(wb, base), SUN);
            yenSun.setAlignment(HorizontalAlignment.RIGHT);
            yenSun.setDataFormat(yenFmt);
            yenToday = fill(boxed(wb, bold), TODAY);
            yenToday.setAlignment(HorizontalAlignment.RIGHT);
            yenToday.setDataFormat(yenFmt);
        }

        CellStyle dateStyle(LocalDate d, LocalDate today) {
            if (today != null && d.equals(today)) {
                return dateToday;
            }
            if (d.getDayOfWeek() == DayOfWeek.SUNDAY) {
                return dateSun;
            }
            if (d.getDayOfWeek() == DayOfWeek.SATURDAY) {
                return dateSat;
            }
            return date;
        }

        CellStyle textStyle(LocalDate d, LocalDate today) {
            if (today != null && d.equals(today)) {
                return textToday;
            }
            if (d.getDayOfWeek() == DayOfWeek.SUNDAY) {
                return textSun;
            }
            if (d.getDayOfWeek() == DayOfWeek.SATURDAY) {
                return textSat;
            }
            return text;
        }

        CellStyle yenStyle(LocalDate d, LocalDate today) {
            if (today != null && d.equals(today)) {
                return yenToday;
            }
            if (d.getDayOfWeek() == DayOfWeek.SUNDAY) {
                return yenSun;
            }
            if (d.getDayOfWeek() == DayOfWeek.SATURDAY) {
                return yenSat;
            }
            return yen;
        }

        private static XSSFCellStyle boxed(XSSFWorkbook wb, XSSFFont font) {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setFont(font);
            cs.setBorderTop(BorderStyle.THIN);
            cs.setBorderBottom(BorderStyle.THIN);
            cs.setBorderLeft(BorderStyle.THIN);
            cs.setBorderRight(BorderStyle.THIN);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            return cs;
        }

        private static XSSFCellStyle fill(XSSFCellStyle cs, byte[] rgb) {
            cs.setFillForegroundColor(new XSSFColor(rgb, null));
            cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return cs;
        }
    }
}
