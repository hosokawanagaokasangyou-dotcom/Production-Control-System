package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.RequestPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;

/** 加工賃トレンド画面の日別・依頼NO別を簡易 Excel 出力する。 */
public final class ProcessingFeeTrendWorkbookExporter {

    private static final DateTimeFormatter FILE_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

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
        Files.createDirectories(target.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            writeDays(wb.createSheet("日別"), r.days());
            writeRequests(
                    wb.createSheet("依頼NO別"),
                    ProcessingFeeTrendAggregator.withLeadingTotalRow(r.requests()));
            try (OutputStream out = Files.newOutputStream(target)) {
                wb.write(out);
            }
        }
    }

    private static Result emptyResult() {
        return new Result(
                List.of(), List.of(), 0, 0, LocalDate.now(), LocalDate.now(), LocalDate.now(), 0, 0, 0);
    }

    private static void writeDays(Sheet sh, List<DayPoint> days) {
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("日付");
        h.createCell(1).setCellValue("実績円");
        h.createCell(2).setCellValue("未了円");
        h.createCell(3).setCellValue("実績累計円");
        h.createCell(4).setCellValue("未了累計円");
        h.createCell(5).setCellValue("見込累計円");
        int row = 1;
        if (days != null) {
            for (DayPoint d : days) {
                if (d == null) {
                    continue;
                }
                Row r = sh.createRow(row++);
                r.createCell(0).setCellValue(d.date().toString());
                r.createCell(1).setCellValue(d.actualYen());
                r.createCell(2).setCellValue(d.planYen());
                r.createCell(3).setCellValue(d.actualCumYen());
                r.createCell(4).setCellValue(d.planCumYen());
                r.createCell(5).setCellValue(d.projectedCumYen());
            }
        }
    }

    private static void writeRequests(Sheet sh, List<RequestPoint> requests) {
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("依頼NO");
        h.createCell(1).setCellValue("AO(受注額)");
        h.createCell(2).setCellValue("円/m");
        h.createCell(3).setCellValue("実績(m)");
        h.createCell(4).setCellValue("未了(m)");
        h.createCell(5).setCellValue("実績円");
        h.createCell(6).setCellValue("未了円");
        int row = 1;
        if (requests != null) {
            for (RequestPoint p : requests) {
                if (p == null) {
                    continue;
                }
                Row r = sh.createRow(row++);
                r.createCell(0).setCellValue(p.requestNo());
                r.createCell(1).setCellValue(p.aoYen());
                if (!p.rateMissing()) {
                    r.createCell(2).setCellValue(p.rateYenPerM());
                }
                r.createCell(3).setCellValue(p.actualMeters());
                r.createCell(4).setCellValue(p.planMeters());
                r.createCell(5).setCellValue(p.actualYen());
                r.createCell(6).setCellValue(p.planYen());
            }
        }
    }
}
