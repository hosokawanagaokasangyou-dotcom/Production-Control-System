package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.nio.file.Path;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;
import jp.co.pm.ai.desktop.io.PoiWorkbookFileWriter;

/**
 * 加工実績・加工予定トレンド（日別・月別・KPI）を Excel ブックとして出力する。
 *
 * <p>工場現場での閲覧・印刷（A4横）に適した書式・配色（BIZ UDPゴシック、土日色、差異色、数値書式）を設定する。
 */
public final class ProcessingTrendWorkbookExporter {

    public static final String SHEET_SUMMARY = "加工トレンドサマリ";
    public static final String SHEET_DAILY = "日別明細";

    public static final String FONT_FAMILY =
            DispatchAladdinEntryWorkbookExporter.DEFAULT_WORKBOOK_FONT_FAMILY; // BIZ UDPゴシック
    public static final int FONT_SIZE_PT = 10;

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final DateTimeFormatter DATE_TIME_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
    private static final DateTimeFormatter TS_COMPACT = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
    private static final DateTimeFormatter YM_FMT = DateTimeFormatter.ofPattern("yyyy年M月");
    private static final String[] WEEKDAY_JA = {"月", "火", "水", "木", "金", "土", "日"};

    // スタイル色
    private static final byte[] COLOR_HEADER_FILL = new byte[] {(byte) 0xD9, (byte) 0xE1, (byte) 0xF2};
    private static final byte[] COLOR_SECTION_FILL = new byte[] {(byte) 0x1E, (byte) 0x3A, (byte) 0x8A};
    private static final byte[] COLOR_KPI_HEADER_FILL = new byte[] {(byte) 0xEE, (byte) 0xF2, (byte) 0xF7};
    private static final byte[] COLOR_SATURDAY_FILL = new byte[] {(byte) 0xEB, (byte) 0xF2, (byte) 0xFA};
    private static final byte[] COLOR_SUNDAY_FILL = new byte[] {(byte) 0xFD, (byte) 0xEC, (byte) 0xEC};
    private static final byte[] COLOR_TODAY_FILL = new byte[] {(byte) 0xFE, (byte) 0xF3, (byte) 0xC7};
    private static final byte[] COLOR_TOTAL_FILL = new byte[] {(byte) 0xE2, (byte) 0xE8, (byte) 0xF0};

    private static final byte[] COLOR_DIFF_POSITIVE = new byte[] {(byte) 0x15, (byte) 0x80, (byte) 0x3D};
    private static final byte[] COLOR_DIFF_NEGATIVE = new byte[] {(byte) 0xB9, (byte) 0x1C, (byte) 0x1C};

    public record ProcessingTrendExportRequest(
            ProcessingTrendAggregator.Result result,
            ProcessingTrendAggregator.MonthlyResult monthlyResult,
            ProcessingTrendAggregator.Filter filter,
            LocalDateTime exportedAt,
            String actualSourceLabel,
            String planSourceLabel,
            String aladdinSourceLabel,
            String dispatchSourceLabel,
            String loadStatsSummary,
            List<String> notices,
            String dailyReportSourceLabel) {

        public ProcessingTrendExportRequest {
            Objects.requireNonNull(result, "result");
            Objects.requireNonNull(filter, "filter");
            exportedAt = exportedAt != null ? exportedAt : LocalDateTime.now();
            notices = notices != null ? List.copyOf(notices) : List.of();
            dailyReportSourceLabel = dailyReportSourceLabel != null ? dailyReportSourceLabel : "";
        }

        public ProcessingTrendExportRequest(
                ProcessingTrendAggregator.Result result,
                ProcessingTrendAggregator.MonthlyResult monthlyResult,
                ProcessingTrendAggregator.Filter filter,
                LocalDateTime exportedAt,
                String actualSourceLabel,
                String planSourceLabel,
                String aladdinSourceLabel,
                String dispatchSourceLabel,
                String loadStatsSummary,
                List<String> notices) {
            this(
                    result,
                    monthlyResult,
                    filter,
                    exportedAt,
                    actualSourceLabel,
                    planSourceLabel,
                    aladdinSourceLabel,
                    dispatchSourceLabel,
                    loadStatsSummary,
                    notices,
                    "");
        }
    }

    public record ProcessingTrendExportResult(Path path, int dayRows, boolean empty) {}

    private ProcessingTrendWorkbookExporter() {}

    /**
     * ファイル名を提案する（例: 加工トレンド_20260901-20260930_日別_日報_全機械_全工程_20260907-072700.xlsx）。
     */
    public static String suggestFileName(
            ProcessingTrendAggregator.Filter filter, boolean isMonthly, LocalDateTime at) {
        LocalDateTime ts = at != null ? at : LocalDateTime.now();
        String gran = isMonthly ? "月別" : "日別";
        String actToken =
                filter.actualSource() == ProcessingTrendAggregator.ActualSource.DAILY_REPORT
                        ? "日報"
                        : "明細";
        String mach = safeToken(filter.hasMachine() ? filter.machine() : "全機械");
        String proc = safeToken(filter.hasProcess() ? filter.process() : "全工程");
        String pStart = filter.from().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String pEnd = filter.to().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        return String.format(
                "加工トレンド_%s-%s_%s_%s_%s_%s_%s.xlsx",
                pStart, pEnd, gran, actToken, mach, proc, ts.format(TS_COMPACT));
    }

    public static String safeToken(String s) {
        if (s == null || s.isBlank()) {
            return "全";
        }
        return s.replaceAll("[\\\\/:*?\"<>|\\s]+", "_");
    }

    /**
     * メモリ上に Excel ブックを構築する。
     */
    public static XSSFWorkbook buildWorkbook(ProcessingTrendExportRequest req) {
        Objects.requireNonNull(req, "req");
        XSSFWorkbook wb = new XSSFWorkbook();
        Styles s = new Styles(wb);

        buildSummarySheet(wb, s, req);
        buildDailySheet(wb, s, req);

        return wb;
    }

    /**
     * target ファイルにワークブックを原子的に保存する。
     */
    public static ProcessingTrendExportResult writeTo(
            Path target, ProcessingTrendExportRequest req, Map<String, String> ui) throws IOException {
        try (XSSFWorkbook wb = buildWorkbook(req)) {
            PoiWorkbookFileWriter.writeReplacing(target, wb, ui);
        }
        return new ProcessingTrendExportResult(
                target.toAbsolutePath().normalize(),
                req.result().days().size(),
                req.result().isEmpty());
    }

    // ---- サマリシート構築 --------------------------------------------------------------------

    private static void buildSummarySheet(XSSFWorkbook wb, Styles s, ProcessingTrendExportRequest req) {
        Sheet sheet = wb.createSheet(SHEET_SUMMARY);
        sheet.setDisplayGridlines(true);
        configurePrint(sheet);

        int r = 0;

        // タイトル
        Row titleRow = sheet.createRow(r++);
        titleRow.setHeightInPoints(28);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("加工トレンド集計サマリ");
        titleCell.setCellStyle(s.titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        // 出力日時
        Row dateRow = sheet.createRow(r++);
        Cell dateLbl = dateRow.createCell(0);
        dateLbl.setCellValue("出力日時: " + req.exportedAt().format(DATE_TIME_FMT));
        dateLbl.setCellStyle(s.mutedStyle);

        // セクション: 集計条件
        Row condSec = sheet.createRow(r++);
        condSec.setHeightInPoints(20);
        Cell condSecCell = condSec.createCell(0);
        condSecCell.setCellValue("■ 集計条件・データソース");
        condSecCell.setCellStyle(s.sectionHeaderStyle);
        sheet.addMergedRegion(new CellRangeAddress(r - 1, r - 1, 0, 5));

        String fromTo = req.filter().from().format(DATE_FMT) + " 〜 " + req.filter().to().format(DATE_FMT);
        addLabelValueRow(sheet, s, r++, "集計期間", fromTo);
        addLabelValueRow(sheet, s, r++, "対象機械", req.filter().hasMachine() ? req.filter().machine() : "（すべて）");
        addLabelValueRow(sheet, s, r++, "対象工程", req.filter().hasProcess() ? req.filter().process() : "（すべて）");
        addLabelValueRow(sheet, s, r++, "実績ソース", req.filter().actualSource().label());
        addLabelValueRow(sheet, s, r++, "予定ソース", req.filter().planSource().label());
        if (req.dailyReportSourceLabel() != null && !req.dailyReportSourceLabel().isBlank()) {
            addLabelValueRow(sheet, s, r++, "加工日報CSV", nz(req.dailyReportSourceLabel()));
        }
        addLabelValueRow(sheet, s, r++, "実績明細ファイル", nz(req.actualSourceLabel()));
        addLabelValueRow(sheet, s, r++, "アラジンファイル", nz(req.aladdinSourceLabel()));
        addLabelValueRow(sheet, s, r++, "配台ファイル", nz(req.dispatchSourceLabel()));
        if (req.loadStatsSummary() != null && !req.loadStatsSummary().isBlank()) {
            addLabelValueRow(sheet, s, r++, "読込データ規模", req.loadStatsSummary());
        }

        r++; // 空行

        // セクション: KPI サマリ
        boolean hasCompare =
                req.result().compareSourceLabel() != null
                        && !req.result().compareSourceLabel().isEmpty();
        boolean isDailyReport =
                req.filter().actualSource() == ProcessingTrendAggregator.ActualSource.DAILY_REPORT;

        Row kpiSec = sheet.createRow(r++);
        kpiSec.setHeightInPoints(20);
        Cell kpiSecCell = kpiSec.createCell(0);
        kpiSecCell.setCellValue("■ KPI サマリ（期間全体）");
        kpiSecCell.setCellStyle(s.sectionHeaderStyle);
        sheet.addMergedRegion(new CellRangeAddress(r - 1, r - 1, 0, hasCompare ? 7 : 5));

        Row kpiHead = sheet.createRow(r++);
        kpiHead.setHeightInPoints(22);
        String[] kpiHeaders;
        if (hasCompare) {
            String actLbl = isDailyReport ? "日報実績合計" : "明細実績合計";
            String compLbl = isDailyReport ? "実績明細合計" : "日報実績合計";
            kpiHeaders =
                    new String[] {
                        actLbl, compLbl, "差異合計", "予定合計", "期間進捗率", "残予定(当日以降)", "見込合計", "見込差異"
                    };
        } else {
            kpiHeaders =
                    new String[] {"実績合計", "予定合計", "期間進捗率", "残予定(当日以降)", "見込合計", "見込差異"};
        }
        for (int i = 0; i < kpiHeaders.length; i++) {
            Cell c = kpiHead.createCell(i);
            c.setCellValue(kpiHeaders[i]);
            c.setCellStyle(s.kpiHeaderStyle);
        }

        Row kpiVal = sheet.createRow(r++);
        kpiVal.setHeightInPoints(24);
        int kCol = 0;

        // 実績合計
        Cell cAct = kpiVal.createCell(kCol++);
        cAct.setCellValue(req.result().actualTotalM());
        cAct.setCellStyle(s.kpiValueNumberStyle);

        if (hasCompare) {
            // 比較実績合計
            Cell cComp = kpiVal.createCell(kCol++);
            cComp.setCellValue(req.result().compareActualTotalM());
            cComp.setCellStyle(s.kpiValueNumberStyle);

            // 明細差合計
            Cell cDiffComp = kpiVal.createCell(kCol++);
            double compDiff = req.result().actualCompareDiffTotalM();
            cDiffComp.setCellValue(compDiff);
            cDiffComp.setCellStyle(
                    compDiff >= 0 ? s.kpiPositiveDiffStyle : s.kpiNegativeDiffStyle);
        }

        // 予定合計
        Cell cPlan = kpiVal.createCell(kCol++);
        cPlan.setCellValue(req.result().planTotalM());
        cPlan.setCellStyle(s.kpiValueNumberStyle);

        // 進捗率
        Cell cProg = kpiVal.createCell(kCol++);
        if (req.result().progressDenominatorSufficient()) {
            cProg.setCellValue(String.format("%.1f %%", req.result().progressPct()));
        } else {
            cProg.setCellValue("—");
        }
        cProg.setCellStyle(s.kpiValueTextStyle);

        // 残予定
        Cell cRem = kpiVal.createCell(kCol++);
        cRem.setCellValue(req.result().remainingPlanM());
        cRem.setCellStyle(s.kpiValueNumberStyle);

        // 見込合計
        Cell cProj = kpiVal.createCell(kCol++);
        cProj.setCellValue(req.result().projectedTotalM());
        cProj.setCellStyle(s.kpiValueNumberStyle);

        // 見込差異
        Cell cDiff = kpiVal.createCell(kCol++);
        cDiff.setCellValue(req.result().projectedDiffM());
        cDiff.setCellStyle(req.result().projectedDiffM() >= 0 ? s.kpiPositiveDiffStyle : s.kpiNegativeDiffStyle);

        r++; // 空行

        // セクション: 月別サマリ表
        if (req.monthlyResult() != null && !req.monthlyResult().months().isEmpty()) {
            Row monthSec = sheet.createRow(r++);
            monthSec.setHeightInPoints(20);
            Cell monthSecCell = monthSec.createCell(0);
            monthSecCell.setCellValue("■ 月別トレンド集計");
            monthSecCell.setCellStyle(s.sectionHeaderStyle);
            sheet.addMergedRegion(new CellRangeAddress(r - 1, r - 1, 0, hasCompare ? 10 : 7));

            Row mHead = sheet.createRow(r++);
            mHead.setHeightInPoints(20);
            String[] mCols;
            if (hasCompare) {
                String actM = isDailyReport ? "日報実績 (m)" : "明細実績 (m)";
                String compM = isDailyReport ? "実績明細 (m)" : "日報実績 (m)";
                String actCumM = isDailyReport ? "日報累計 (m)" : "明細累計 (m)";
                String compCumM = isDailyReport ? "明細累計 (m)" : "日報累計 (m)";
                mCols =
                        new String[] {
                            "年月",
                            "期間日数",
                            actM,
                            compM,
                            "差異 (m)",
                            "予定 (m)",
                            actCumM,
                            compCumM,
                            "予定累計 (m)",
                            "見込累計 (m)"
                        };
            } else {
                mCols =
                        new String[] {
                            "年月", "期間日数", "実績 (m)", "予定 (m)", "実績累計 (m)", "予定累計 (m)", "見込累計 (m)"
                        };
            }
            for (int i = 0; i < mCols.length; i++) {
                Cell c = mHead.createCell(i);
                c.setCellValue(mCols[i]);
                c.setCellStyle(s.tableHeaderStyle);
            }

            for (ProcessingTrendAggregator.MonthPoint mp : req.monthlyResult().months()) {
                Row row = sheet.createRow(r++);
                row.setHeightInPoints(18);
                int mCol = 0;

                Cell cYm = row.createCell(mCol++);
                String ymLabel = mp.month().format(YM_FMT) + (mp.incomplete() ? " (部分)" : "");
                cYm.setCellValue(ymLabel);
                cYm.setCellStyle(mp.isCurrentMonth() ? s.todayTextCellStyle : s.textCellStyle);

                Cell cDays = row.createCell(mCol++);
                cDays.setCellValue(mp.daysInBucket() + " 日");
                cDays.setCellStyle(s.centerCellStyle);

                Cell cMAct = row.createCell(mCol++);
                cMAct.setCellValue(mp.actualM());
                cMAct.setCellStyle(s.numberCellStyle);

                if (hasCompare) {
                    Cell cMComp = row.createCell(mCol++);
                    cMComp.setCellValue(mp.compareActualM());
                    cMComp.setCellStyle(s.numberCellStyle);

                    Cell cMCompDiff = row.createCell(mCol++);
                    double mDiff = mp.actualCompareDiffM();
                    cMCompDiff.setCellValue(mDiff);
                    cMCompDiff.setCellStyle(
                            mDiff >= 0 ? s.diffPositiveStyle : s.diffNegativeStyle);
                }

                Cell cMPlan = row.createCell(mCol++);
                cMPlan.setCellValue(mp.planM());
                cMPlan.setCellStyle(s.numberCellStyle);

                Cell cMCumAct = row.createCell(mCol++);
                cMCumAct.setCellValue(mp.actualCumM());
                cMCumAct.setCellStyle(s.numberCellStyle);

                if (hasCompare) {
                    Cell cMCumComp = row.createCell(mCol++);
                    cMCumComp.setCellValue(mp.compareActualCumM());
                    cMCumComp.setCellStyle(s.numberCellStyle);
                }

                Cell cMCumPlan = row.createCell(mCol++);
                cMCumPlan.setCellValue(mp.planCumM());
                cMCumPlan.setCellStyle(s.numberCellStyle);

                Cell cMCumProj = row.createCell(mCol++);
                cMCumProj.setCellValue(mp.projectedCumM());
                cMCumProj.setCellStyle(s.numberCellStyle);
            }

            r++; // 空行
        }

        // 注意事項・備考
        if (!req.notices().isEmpty() || !req.result().warnings().isEmpty()) {
            Row noteSec = sheet.createRow(r++);
            Cell noteCell = noteSec.createCell(0);
            noteCell.setCellValue("■ 注意事項・備考");
            noteCell.setCellStyle(s.boldStyle);

            for (String note : req.notices()) {
                Row row = sheet.createRow(r++);
                Cell c = row.createCell(0);
                c.setCellValue("・ " + note);
                c.setCellStyle(s.warnTextStyle);
                sheet.addMergedRegion(new CellRangeAddress(r - 1, r - 1, 0, hasCompare ? 9 : 6));
            }
            for (String warn : req.result().warnings()) {
                Row row = sheet.createRow(r++);
                Cell c = row.createCell(0);
                c.setCellValue("・ " + warn);
                c.setCellStyle(s.warnTextStyle);
                sheet.addMergedRegion(new CellRangeAddress(r - 1, r - 1, 0, hasCompare ? 9 : 6));
            }
        }

        // 列幅設定
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setColumnWidth(1, 14 * 256);
        sheet.setColumnWidth(2, 16 * 256);
        sheet.setColumnWidth(3, 16 * 256);
        sheet.setColumnWidth(4, 16 * 256);
        sheet.setColumnWidth(5, 16 * 256);
        sheet.setColumnWidth(6, 16 * 256);
        sheet.setColumnWidth(7, 16 * 256);
        if (hasCompare) {
            sheet.setColumnWidth(8, 16 * 256);
            sheet.setColumnWidth(9, 16 * 256);
            sheet.setColumnWidth(10, 16 * 256);
        }
    }

    private static void addLabelValueRow(Sheet sheet, Styles s, int rowIdx, String label, String value) {
        Row row = sheet.createRow(rowIdx);
        row.setHeightInPoints(18);
        Cell c0 = row.createCell(0);
        c0.setCellValue(label);
        c0.setCellStyle(s.condLabelStyle);

        Cell c1 = row.createCell(1);
        c1.setCellValue(value != null ? value : "—");
        c1.setCellStyle(s.condValueStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 1, 5));
    }

    // ---- 日別明細シート構築 ------------------------------------------------------------------

    private static void buildDailySheet(XSSFWorkbook wb, Styles s, ProcessingTrendExportRequest req) {
        Sheet sheet = wb.createSheet(SHEET_DAILY);
        sheet.setDisplayGridlines(true);
        sheet.createFreezePane(0, 1);
        configurePrint(sheet);

        boolean hasCompare =
                req.result().compareSourceLabel() != null
                        && !req.result().compareSourceLabel().isEmpty();
        boolean isDailyReport =
                req.filter().actualSource() == ProcessingTrendAggregator.ActualSource.DAILY_REPORT;

        Row head = sheet.createRow(0);
        head.setHeightInPoints(22);
        String[] headers;
        if (hasCompare) {
            String actCol = isDailyReport ? "日報実績 (m)" : "明細実績 (m)";
            String compCol = isDailyReport ? "実績明細 (m)" : "日報実績 (m)";
            String actCumCol = isDailyReport ? "日報累計 (m)" : "明細累計 (m)";
            String compCumCol = isDailyReport ? "明細累計 (m)" : "日報累計 (m)";
            headers =
                    new String[] {
                        "日付",
                        "曜日",
                        actCol,
                        "7日移動平均 (m)",
                        compCol,
                        "差異 (m)",
                        "予定 (m)",
                        actCumCol,
                        compCumCol,
                        "予定累計 (m)",
                        "見込累計 (m)",
                        "備考"
                    };
        } else {
            headers =
                    new String[] {
                        "日付", "曜日", "実績 (m)", "7日移動平均 (m)", "予定 (m)",
                        "実績累計 (m)", "予定累計 (m)", "見込累計 (m)", "備考"
                    };
        }
        for (int i = 0; i < headers.length; i++) {
            Cell c = head.createCell(i);
            c.setCellValue(headers[i]);
            c.setCellStyle(s.tableHeaderStyle);
        }

        int r = 1;
        LocalDate today = req.result().today();
        double sumAct = 0.0;
        double sumCompAct = 0.0;
        double sumCompDiff = 0.0;
        double sumPlan = 0.0;

        for (ProcessingTrendAggregator.DayPoint dp : req.result().days()) {
            Row row = sheet.createRow(r++);
            row.setHeightInPoints(18);

            LocalDate date = dp.date();
            DayOfWeek dow = date.getDayOfWeek();
            boolean isToday = date.equals(today);
            boolean isSat = dow == DayOfWeek.SATURDAY;
            boolean isSun = dow == DayOfWeek.SUNDAY;

            CellStyle dateStyle =
                    isToday
                            ? s.todayDateCellStyle
                            : (isSun ? s.sunDateCellStyle : (isSat ? s.satDateCellStyle : s.dateCellStyle));
            CellStyle textStyle =
                    isToday
                            ? s.todayTextCellStyle
                            : (isSun ? s.sunTextCellStyle : (isSat ? s.satTextCellStyle : s.textCellStyle));
            CellStyle numStyle = isToday ? s.todayNumberCellStyle : s.numberCellStyle;

            int col = 0;

            // 日付
            Cell cDate = row.createCell(col++);
            cDate.setCellValue(date.format(DATE_FMT));
            cDate.setCellStyle(dateStyle);

            // 曜日
            Cell cDow = row.createCell(col++);
            cDow.setCellValue(WEEKDAY_JA[dow.getValue() - 1]);
            cDow.setCellStyle(textStyle);

            // 実績
            Cell cAct = row.createCell(col++);
            cAct.setCellValue(dp.actualM());
            cAct.setCellStyle(numStyle);
            sumAct += dp.actualM();

            // 7日移動平均
            Cell cMa7 = row.createCell(col++);
            cMa7.setCellValue(dp.actual7dMaM());
            cMa7.setCellStyle(numStyle);

            if (hasCompare) {
                // 比較実績
                Cell cCompAct = row.createCell(col++);
                cCompAct.setCellValue(dp.compareActualM());
                cCompAct.setCellStyle(numStyle);
                sumCompAct += dp.compareActualM();

                // 差異 (日報 − 実績明細)
                Cell cDiffComp = row.createCell(col++);
                double actCompDiff = dp.actualCompareDiffM();
                cDiffComp.setCellValue(actCompDiff);
                cDiffComp.setCellStyle(
                        actCompDiff >= 0 ? s.diffPositiveStyle : s.diffNegativeStyle);
                sumCompDiff += actCompDiff;
            }

            // 予定
            Cell cPlan = row.createCell(col++);
            cPlan.setCellValue(dp.planM());
            cPlan.setCellStyle(numStyle);
            sumPlan += dp.planM();

            // 実績累計
            Cell cCumAct = row.createCell(col++);
            cCumAct.setCellValue(dp.actualCumM());
            cCumAct.setCellStyle(numStyle);

            if (hasCompare) {
                // 比較実績累計
                Cell cCumCompAct = row.createCell(col++);
                cCumCompAct.setCellValue(dp.compareActualCumM());
                cCumCompAct.setCellStyle(numStyle);
            }

            // 予定累計
            Cell cCumPlan = row.createCell(col++);
            cCumPlan.setCellValue(dp.planCumM());
            cCumPlan.setCellStyle(numStyle);

            // 見込累計
            Cell cCumProj = row.createCell(col++);
            cCumProj.setCellValue(dp.projectedCumM());
            cCumProj.setCellStyle(numStyle);

            // 備考
            Cell cRemark = row.createCell(col++);
            StringBuilder rem = new StringBuilder();
            if (isToday) {
                rem.append("当日 ");
            }
            if (date.isAfter(today)) {
                rem.append("未来 ");
            }
            if (dp.usesPlanForProjection()) {
                rem.append("見込=予定");
            }
            cRemark.setCellValue(rem.toString().trim());
            cRemark.setCellStyle(textStyle);
        }

        // 合計行
        Row totalRow = sheet.createRow(r);
        totalRow.setHeightInPoints(20);
        int tCol = 0;

        Cell cTotLbl = totalRow.createCell(tCol++);
        cTotLbl.setCellValue("合計");
        cTotLbl.setCellStyle(s.totalLabelStyle);

        Cell cTotDow = totalRow.createCell(tCol++);
        cTotDow.setCellValue("");
        cTotDow.setCellStyle(s.totalLabelStyle);

        Cell cTotAct = totalRow.createCell(tCol++);
        cTotAct.setCellValue(sumAct);
        cTotAct.setCellStyle(s.totalNumberStyle);

        Cell cTotMa7 = totalRow.createCell(tCol++);
        cTotMa7.setCellValue("");
        cTotMa7.setCellStyle(s.totalLabelStyle);

        if (hasCompare) {
            Cell cTotCompAct = totalRow.createCell(tCol++);
            cTotCompAct.setCellValue(sumCompAct);
            cTotCompAct.setCellStyle(s.totalNumberStyle);

            Cell cTotCompDiff = totalRow.createCell(tCol++);
            cTotCompDiff.setCellValue(sumCompDiff);
            cTotCompDiff.setCellStyle(
                    sumCompDiff >= 0 ? s.totalPositiveDiffStyle : s.totalNegativeDiffStyle);
        }

        Cell cTotPlan = totalRow.createCell(tCol++);
        cTotPlan.setCellValue(sumPlan);
        cTotPlan.setCellStyle(s.totalNumberStyle);

        for (int c = tCol; c < headers.length; c++) {
            Cell cEmpty = totalRow.createCell(c);
            cEmpty.setCellValue("");
            cEmpty.setCellStyle(s.totalLabelStyle);
        }

        // 列幅設定
        for (int i = 0; i < headers.length; i++) {
            if (i == 1) {
                sheet.setColumnWidth(i, 8 * 256);
            } else if (i == headers.length - 1) {
                sheet.setColumnWidth(i, 16 * 256);
            } else {
                sheet.setColumnWidth(i, 15 * 256);
            }
        }
    }

    private static void configurePrint(Sheet sheet) {
        PrintSetup ps = sheet.getPrintSetup();
        ps.setLandscape(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        sheet.setAutobreaks(true);
    }

    private static String nz(String s) {
        return s != null && !s.isBlank() ? s : "—";
    }

    // ---- スタイル保持ヘルパークラス -----------------------------------------------------------

    private static final class Styles {
        final CellStyle titleStyle;
        final CellStyle boldStyle;
        final CellStyle mutedStyle;
        final CellStyle sectionHeaderStyle;
        final CellStyle condLabelStyle;
        final CellStyle condValueStyle;
        final CellStyle kpiHeaderStyle;
        final CellStyle kpiValueNumberStyle;
        final CellStyle kpiValueTextStyle;
        final CellStyle kpiPositiveDiffStyle;
        final CellStyle kpiNegativeDiffStyle;
        final CellStyle tableHeaderStyle;
        final CellStyle dateCellStyle;
        final CellStyle satDateCellStyle;
        final CellStyle sunDateCellStyle;
        final CellStyle todayDateCellStyle;
        final CellStyle textCellStyle;
        final CellStyle centerCellStyle;
        final CellStyle satTextCellStyle;
        final CellStyle sunTextCellStyle;
        final CellStyle todayTextCellStyle;
        final CellStyle numberCellStyle;
        final CellStyle todayNumberCellStyle;
        final CellStyle diffPositiveStyle;
        final CellStyle diffNegativeStyle;
        final CellStyle totalLabelStyle;
        final CellStyle totalNumberStyle;
        final CellStyle totalPositiveDiffStyle;
        final CellStyle totalNegativeDiffStyle;
        final CellStyle warnTextStyle;

        Styles(XSSFWorkbook wb) {
            DataFormat df = wb.createDataFormat();
            short numFmt = df.getFormat("#,##0.0 \"m\"");

            XSSFFont baseFont = wb.createFont();
            baseFont.setFontName(FONT_FAMILY);
            baseFont.setFontHeightInPoints((short) FONT_SIZE_PT);

            XSSFFont boldFont = wb.createFont();
            boldFont.setFontName(FONT_FAMILY);
            boldFont.setFontHeightInPoints((short) FONT_SIZE_PT);
            boldFont.setBold(true);

            XSSFFont titleFont = wb.createFont();
            titleFont.setFontName(FONT_FAMILY);
            titleFont.setFontHeightInPoints((short) 15);
            titleFont.setBold(true);

            XSSFFont secFont = wb.createFont();
            secFont.setFontName(FONT_FAMILY);
            secFont.setFontHeightInPoints((short) 11);
            secFont.setBold(true);

            XSSFFont mutedFont = wb.createFont();
            mutedFont.setFontName(FONT_FAMILY);
            mutedFont.setFontHeightInPoints((short) 9);
            mutedFont.setColor(new XSSFColor(new byte[] {(byte) 0x64, (byte) 0x74, (byte) 0x8B}));

            XSSFFont posFont = wb.createFont();
            posFont.setFontName(FONT_FAMILY);
            posFont.setFontHeightInPoints((short) FONT_SIZE_PT);
            posFont.setBold(true);
            posFont.setColor(new XSSFColor(COLOR_DIFF_POSITIVE));

            XSSFFont negFont = wb.createFont();
            negFont.setFontName(FONT_FAMILY);
            negFont.setFontHeightInPoints((short) FONT_SIZE_PT);
            negFont.setBold(true);
            negFont.setColor(new XSSFColor(COLOR_DIFF_NEGATIVE));

            XSSFFont kpiValFont = wb.createFont();
            kpiValFont.setFontName(FONT_FAMILY);
            kpiValFont.setFontHeightInPoints((short) 13);
            kpiValFont.setBold(true);

            XSSFFont kpiPosFont = wb.createFont();
            kpiPosFont.setFontName(FONT_FAMILY);
            kpiPosFont.setFontHeightInPoints((short) 13);
            kpiPosFont.setBold(true);
            kpiPosFont.setColor(new XSSFColor(COLOR_DIFF_POSITIVE));

            XSSFFont kpiNegFont = wb.createFont();
            kpiNegFont.setFontName(FONT_FAMILY);
            kpiNegFont.setFontHeightInPoints((short) 13);
            kpiNegFont.setBold(true);
            kpiNegFont.setColor(new XSSFColor(COLOR_DIFF_NEGATIVE));

            // タイトル
            titleStyle = wb.createCellStyle();
            titleStyle.setFont(titleFont);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            boldStyle = wb.createCellStyle();
            boldStyle.setFont(boldFont);

            mutedStyle = wb.createCellStyle();
            mutedStyle.setFont(mutedFont);

            warnTextStyle = wb.createCellStyle();
            warnTextStyle.setFont(negFont);

            // セクションヘッダー
            sectionHeaderStyle = wb.createCellStyle();
            sectionHeaderStyle.setFont(secFont);
            sectionHeaderStyle.setFillForegroundColor(new XSSFColor(COLOR_SECTION_FILL));
            sectionHeaderStyle.setFillPattern(FillPatternType.NO_FILL);
            sectionHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // 条件ラベル・値
            condLabelStyle = wb.createCellStyle();
            condLabelStyle.setFont(boldFont);
            condLabelStyle.setAlignment(HorizontalAlignment.RIGHT);
            condLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            condValueStyle = wb.createCellStyle();
            condValueStyle.setFont(baseFont);
            condValueStyle.setAlignment(HorizontalAlignment.LEFT);
            condValueStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // KPI
            kpiHeaderStyle = createBaseCell(wb, boldFont);
            kpiHeaderStyle.setFillForegroundColor(new XSSFColor(COLOR_KPI_HEADER_FILL));
            kpiHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            kpiHeaderStyle.setAlignment(HorizontalAlignment.CENTER);

            kpiValueNumberStyle = createBaseCell(wb, kpiValFont);
            kpiValueNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
            kpiValueNumberStyle.setDataFormat(numFmt);

            kpiValueTextStyle = createBaseCell(wb, kpiValFont);
            kpiValueTextStyle.setAlignment(HorizontalAlignment.CENTER);

            kpiPositiveDiffStyle = createBaseCell(wb, kpiPosFont);
            kpiPositiveDiffStyle.setAlignment(HorizontalAlignment.RIGHT);
            kpiPositiveDiffStyle.setDataFormat(numFmt);

            kpiNegativeDiffStyle = createBaseCell(wb, kpiNegFont);
            kpiNegativeDiffStyle.setAlignment(HorizontalAlignment.RIGHT);
            kpiNegativeDiffStyle.setDataFormat(numFmt);

            // 表ヘッダー
            tableHeaderStyle = createBaseCell(wb, boldFont);
            tableHeaderStyle.setFillForegroundColor(new XSSFColor(COLOR_HEADER_FILL));
            tableHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);

            // 表セル
            dateCellStyle = createBaseCell(wb, baseFont);
            dateCellStyle.setAlignment(HorizontalAlignment.CENTER);

            satDateCellStyle = createBaseCell(wb, baseFont);
            satDateCellStyle.setAlignment(HorizontalAlignment.CENTER);
            satDateCellStyle.setFillForegroundColor(new XSSFColor(COLOR_SATURDAY_FILL));
            satDateCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            sunDateCellStyle = createBaseCell(wb, baseFont);
            sunDateCellStyle.setAlignment(HorizontalAlignment.CENTER);
            sunDateCellStyle.setFillForegroundColor(new XSSFColor(COLOR_SUNDAY_FILL));
            sunDateCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            todayDateCellStyle = createBaseCell(wb, boldFont);
            todayDateCellStyle.setAlignment(HorizontalAlignment.CENTER);
            todayDateCellStyle.setFillForegroundColor(new XSSFColor(COLOR_TODAY_FILL));
            todayDateCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            textCellStyle = createBaseCell(wb, baseFont);
            textCellStyle.setAlignment(HorizontalAlignment.LEFT);

            centerCellStyle = createBaseCell(wb, baseFont);
            centerCellStyle.setAlignment(HorizontalAlignment.CENTER);

            satTextCellStyle = createBaseCell(wb, baseFont);
            satTextCellStyle.setAlignment(HorizontalAlignment.CENTER);
            satTextCellStyle.setFillForegroundColor(new XSSFColor(COLOR_SATURDAY_FILL));
            satTextCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            sunTextCellStyle = createBaseCell(wb, baseFont);
            sunTextCellStyle.setAlignment(HorizontalAlignment.CENTER);
            sunTextCellStyle.setFillForegroundColor(new XSSFColor(COLOR_SUNDAY_FILL));
            sunTextCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            todayTextCellStyle = createBaseCell(wb, boldFont);
            todayTextCellStyle.setAlignment(HorizontalAlignment.LEFT);
            todayTextCellStyle.setFillForegroundColor(new XSSFColor(COLOR_TODAY_FILL));
            todayTextCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            numberCellStyle = createBaseCell(wb, baseFont);
            numberCellStyle.setAlignment(HorizontalAlignment.RIGHT);
            numberCellStyle.setDataFormat(numFmt);

            todayNumberCellStyle = createBaseCell(wb, boldFont);
            todayNumberCellStyle.setAlignment(HorizontalAlignment.RIGHT);
            todayNumberCellStyle.setDataFormat(numFmt);
            todayNumberCellStyle.setFillForegroundColor(new XSSFColor(COLOR_TODAY_FILL));
            todayNumberCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            diffPositiveStyle = createBaseCell(wb, posFont);
            diffPositiveStyle.setAlignment(HorizontalAlignment.RIGHT);
            diffPositiveStyle.setDataFormat(numFmt);

            diffNegativeStyle = createBaseCell(wb, negFont);
            diffNegativeStyle.setAlignment(HorizontalAlignment.RIGHT);
            diffNegativeStyle.setDataFormat(numFmt);

            // 合計
            totalLabelStyle = createBaseCell(wb, boldFont);
            totalLabelStyle.setAlignment(HorizontalAlignment.CENTER);
            totalLabelStyle.setFillForegroundColor(new XSSFColor(COLOR_TOTAL_FILL));
            totalLabelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            totalNumberStyle = createBaseCell(wb, boldFont);
            totalNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
            totalNumberStyle.setDataFormat(numFmt);
            totalNumberStyle.setFillForegroundColor(new XSSFColor(COLOR_TOTAL_FILL));
            totalNumberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            totalPositiveDiffStyle = createBaseCell(wb, kpiPosFont);
            totalPositiveDiffStyle.setAlignment(HorizontalAlignment.RIGHT);
            totalPositiveDiffStyle.setDataFormat(numFmt);
            totalPositiveDiffStyle.setFillForegroundColor(new XSSFColor(COLOR_TOTAL_FILL));
            totalPositiveDiffStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            totalNegativeDiffStyle = createBaseCell(wb, kpiNegFont);
            totalNegativeDiffStyle.setAlignment(HorizontalAlignment.RIGHT);
            totalNegativeDiffStyle.setDataFormat(numFmt);
            totalNegativeDiffStyle.setFillForegroundColor(new XSSFColor(COLOR_TOTAL_FILL));
            totalNegativeDiffStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        private static XSSFCellStyle createBaseCell(XSSFWorkbook wb, XSSFFont font) {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setFont(font);
            cs.setBorderTop(BorderStyle.THIN);
            cs.setBorderBottom(BorderStyle.THIN);
            cs.setBorderLeft(BorderStyle.THIN);
            cs.setBorderRight(BorderStyle.THIN);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            return cs;
        }
    }
}
