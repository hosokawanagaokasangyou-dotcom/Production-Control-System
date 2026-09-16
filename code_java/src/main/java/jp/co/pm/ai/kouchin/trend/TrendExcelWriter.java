package jp.co.pm.ai.kouchin.trend;

import java.awt.GraphicsEnvironment;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDispBlanksAs;
import org.openxmlformats.schemas.drawingml.x2006.chart.STMarkerStyle;

/**
 * 月次トレンド Excel。シートは サマリ / 工程比較 / 負荷シフト / 付録_国分 / 付録_湖南 / 付録_シート区分。
 */
public final class TrendExcelWriter {

    public static final String SHEET_SUMMARY = "サマリ";
    public static final String SHEET_COMPARE = "工程比較";
    public static final String SHEET_SHIFT = "負荷シフト";
    public static final String SHEET_APP_KOKUBU = "付録_国分";
    public static final String SHEET_APP_KONAN = "付録_湖南";
    public static final String SHEET_APP_SHEETS = "付録_シート区分";

    static final String COLOR_KOKUBU = "C0504D";
    static final String COLOR_KONAN = "9BBB59";

    private static final String FILL_WARN = "FCE4D6";
    private static final String FILL_SHIFT = "FFF2CC";
    private static final String FILL_HEADER = "D6DCE4";
    private static final String FILL_KPI = "E2EFDA";
    private static final String FILL_SECTION = "D9E1F2";
    private static final String FILL_DIFF = "F2F2F2";
    private static final String COLOR_GRAY = "666666";
    private static final String COLOR_MUTED = "808080";
    private static final String COLOR_RED = "C00000";

    private static final Map<String, String> TAB =
            Map.of(
                    SHEET_SUMMARY, "1F4E79",
                    SHEET_COMPARE, "C00000",
                    SHEET_SHIFT, "7030A0",
                    SHEET_APP_KOKUBU, "A6A6A6",
                    SHEET_APP_KONAN, "A6A6A6",
                    SHEET_APP_SHEETS, "BFBFBF");

    private final XSSFWorkbook wb;
    private final String baseFont;
    private final String numFont;
    private final Map<String, XSSFFont> fonts = new HashMap<>();
    private final Map<String, XSSFCellStyle> styles = new HashMap<>();
    private final short fmtInt;

    private TrendExcelWriter(XSSFWorkbook wb) {
        this.wb = wb;
        this.baseFont = pickFont("BIZ UDPゴシック", "BIZ UDPGothic", "游ゴシック", "Yu Gothic");
        this.numFont = pickFont("BIZ UDゴシック", "BIZ UDGothic", "游ゴシック", "Yu Gothic");
        this.fmtInt = wb.createDataFormat().getFormat("#,##0");
    }

    public static XSSFWorkbook buildWorkbook(TrendModel model) {
        XSSFWorkbook wb = new XSSFWorkbook();
        TrendExcelWriter writer = new TrendExcelWriter(wb);
        writer.writeAll(
                model == null
                        ? new TrendModel(List.of(), List.of(), null, null, List.of(), List.of(), List.of(), List.of())
                        : model);
        return wb;
    }

    private void writeAll(TrendModel model) {
        List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months = model.months();
        List<TrendMonthData> kokubu = model.kokubu();
        List<TrendMonthData> konan = model.konan();
        TrendCrossMatrix gw =
                model.wage() != null
                        ? model.wage()
                        : TrendAggregate.buildGroupedCross(kokubu, konan, months, "wage");
        TrendCrossMatrix gq =
                model.qty() != null
                        ? model.qty()
                        : TrendAggregate.buildGroupedCross(kokubu, konan, months, "qty");
        List<String> focus =
                model.focusProcesses().isEmpty()
                        ? TrendAggregate.pickFocusProcesses(gw, gq, 6)
                        : model.focusProcesses();
        List<TrendShiftRow> shiftQty = TrendAggregate.rankShiftSuspects(gq, "qty");
        List<TrendShiftRow> shiftWage = TrendAggregate.rankShiftSuspects(gw, "wage");
        List<TrendShiftRow> shiftHome = shiftQty.isEmpty() ? shiftWage : shiftQty;
        jp.co.pm.ai.kouchin.verify.YearMonthKey[] pair =
                TrendAggregate.lastComparablePair(months, kokubu, konan);

        XSSFSheet home = wb.createSheet(SHEET_SUMMARY);
        tabColor(home, SHEET_SUMMARY);
        writeHome(home, model, months, kokubu, konan, focus, shiftHome, pair);

        XSSFSheet cmp = wb.createSheet(SHEET_COMPARE);
        tabColor(cmp, SHEET_COMPARE);
        writeProcessCompare(cmp, months, gw, gq, focus);

        XSSFSheet shift = wb.createSheet(SHEET_SHIFT);
        tabColor(shift, SHEET_SHIFT);
        writeLoadShift(shift, shiftQty, shiftWage, pair[0], pair[1]);

        XSSFSheet appK = wb.createSheet(SHEET_APP_KOKUBU);
        tabColor(appK, SHEET_APP_KOKUBU);
        writeAppendixFactory(appK, "国分 種類別", kokubu, months);

        XSSFSheet appN = wb.createSheet(SHEET_APP_KONAN);
        tabColor(appN, SHEET_APP_KONAN);
        writeAppendixFactory(appN, "湖南 種類別", konan, months);

        XSSFSheet appS = wb.createSheet(SHEET_APP_SHEETS);
        tabColor(appS, SHEET_APP_SHEETS);
        writeAppendixSheets(appS, months, kokubu, konan);
    }

    private void writeHome(
            XSSFSheet ws,
            TrendModel payload,
            List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months,
            List<TrendMonthData> kokubu,
            List<TrendMonthData> konan,
            List<String> focus,
            List<TrendShiftRow> shiftRows,
            jp.co.pm.ai.kouchin.verify.YearMonthKey[] pair) {
        put(ws, 0, 0, "後加工工賃 月次トレンド — ホーム", font(true, null, 14, false));
        if (!months.isEmpty()) {
            put(ws, 1, 0,
                    "対象期間: " + months.get(0).ymLabel() + " 〜 " + months.get(months.size() - 1).ymLabel()
                            + "（" + months.size() + "か月）",
                    font(false, null, 10.5, false));
        }
        put(ws, 2, 0, "まず「工程比較」で注目工程の国分↔湖南を確認。振替は断定せず「疑い」として見ます。", font(false, COLOR_GRAY, 10.5, false));
        put(ws, 4, 0, "【警告】", font(true, COLOR_RED, 10.5, false));
        List<String> warnings = payload.warnings();
        int r;
        if (!warnings.isEmpty()) {
            int n = Math.min(warnings.size(), 8);
            for (int i = 0; i < n; i++) {
                Cell c = put(ws, 5 + i, 0, warnings.get(i), font(false, null, 10.5, false));
                fill(c, FILL_WARN);
            }
            r = 5 + n;
        } else {
            put(ws, 5, 0, "（なし）", font(false, null, 10.5, false));
            r = 7;
        }
        r++;
        put(ws, r, 0, "【30秒サマリ】", font(true, null, 12, false));
        r++;
        jp.co.pm.ai.kouchin.verify.YearMonthKey prev = pair[0];
        jp.co.pm.ai.kouchin.verify.YearMonthKey latest = pair[1];
        if (latest != null && prev != null) {
            put(ws, r, 0, "比較月: " + prev.ymLabel() + " → " + latest.ymLabel() + "（両工場データあり）", font(false, null, 10.5, false));
            r++;
            String[] headers = {"指標", "国分", "湖南", "両工場合計"};
            for (int j = 0; j < headers.length; j++) {
                put(ws, r, j, headers[j], font(true, null, 10.5, false));
                fill(cell(ws, r, j), FILL_HEADER);
            }
            r++;
            for (String[] spec : new String[][] {{"wage", "加工賃", "千円"}, {"qty", "加工量", "km"}}) {
                String metric = spec[0];
                String label = spec[1];
                String unit = spec[2];
                Double tk0 = TrendAggregate.factoryMonthTotal(kokubu, prev, metric);
                Double tk1 = TrendAggregate.factoryMonthTotal(kokubu, latest, metric);
                Double th0 = TrendAggregate.factoryMonthTotal(konan, prev, metric);
                Double th1 = TrendAggregate.factoryMonthTotal(konan, latest, metric);
                put(ws, r, 0, label + "（" + latest.ymLabel() + "）[" + unit + "]", font(false, null, 10.5, false));
                num(ws, r, 1, round(tk1));
                num(ws, r, 2, round(th1));
                if (tk1 != null && th1 != null) {
                    num(ws, r, 3, round(tk1 + th1));
                }
                for (int c = 0; c < 4; c++) {
                    fill(cell(ws, r, c), FILL_KPI);
                }
                r++;
                put(ws, r, 0, label + " 前月差[" + unit + "]", font(false, null, 10.5, false));
                Double dk = (tk0 != null && tk1 != null) ? tk1 - tk0 : null;
                Double dh = (th0 != null && th1 != null) ? th1 - th0 : null;
                num(ws, r, 1, round(dk));
                num(ws, r, 2, round(dh));
                if (dk != null && dh != null) {
                    num(ws, r, 3, round(dk + dh));
                }
                r++;
            }
        } else {
            put(ws, r, 0, "両工場で比較できる連続2か月が不足しています", font(false, null, 10.5, false));
            r++;
        }
        r++;
        String shiftUnit = !shiftRows.isEmpty() && "qty".equals(shiftRows.get(0).metric()) ? "km" : "千円";
        put(ws, r, 0, "【振替・シフト疑い Top】（" + shiftUnit + "ベース前月差。異符号の相殺）", font(true, null, 10.5, false));
        r++;
        String[] sh =
                {"工程", "判定", "国分前月差(" + shiftUnit + ")", "湖南前月差(" + shiftUnit + ")", "両工場合計差(" + shiftUnit + ")", "相殺量(" + shiftUnit + ")"};
        for (int j = 0; j < sh.length; j++) {
            put(ws, r, j, sh[j], font(true, null, 10.5, false));
            fill(cell(ws, r, j), FILL_HEADER);
        }
        r++;
        List<TrendShiftRow> shown =
                shiftRows.stream().filter(x -> "強い相殺".equals(x.badge()) || "弱い相殺".equals(x.badge())).limit(5).toList();
        if (shown.isEmpty()) {
            put(ws, r, 0, "（該当なし）", font(false, null, 10.5, false));
            r++;
        }
        for (TrendShiftRow row : shown) {
            put(ws, r, 0, row.process(), font(false, null, 10.5, false));
            put(ws, r, 1, row.badge(), font(false, null, 10.5, false));
            num(ws, r, 2, row.dKokubu());
            num(ws, r, 3, row.dKonan());
            num(ws, r, 4, row.dCombined());
            num(ws, r, 5, row.offset());
            for (int c = 0; c < 6; c++) {
                fill(cell(ws, r, c), FILL_SHIFT);
            }
            r++;
        }
        r++;
        put(ws, r, 0, "【注目工程】→ 次タブ「工程比較」", font(true, null, 10.5, false));
        r++;
        put(ws, r, 0, focus.isEmpty() ? "（なし）" : String.join("、", focus), font(false, null, 10.5, false));
        r += 2;
        put(ws, r, 0, "【ナビ】", font(true, null, 10.5, false));
        r++;
        put(ws, r, 0, "1. 工程比較 … 注目工程ごとに国分(赤茶)と湖南(若草)を同じグラフで比較", font(false, null, 10.5, false));
        r++;
        put(ws, r, 0, "2. 負荷シフト … 動いた工程の一覧（前月差・判定）", font(false, null, 10.5, false));
        r++;
        put(ws, r, 0, "3. 付録_* … 工場別の全工程表（通常は不要）", font(false, null, 10.5, false));
        r += 2;
        put(ws, r, 0, "【使用ファイル】", font(true, null, 10.5, false));
        r++;
        r = writeFiles(ws, r, "国分", filesOf(kokubu));
        r = writeFiles(ws, r, "湖南", filesOf(konan));
        ws.setColumnWidth(0, 42 * 256);
        for (int c = 1; c <= 5; c++) {
            ws.setColumnWidth(c, 14 * 256);
        }
    }

    private static Map<jp.co.pm.ai.kouchin.verify.YearMonthKey, Path> filesOf(List<TrendMonthData> data) {
        Map<jp.co.pm.ai.kouchin.verify.YearMonthKey, Path> files = new LinkedHashMap<>();
        for (TrendMonthData m : data) {
            if (m != null && m.ym() != null && m.path() != null) {
                files.put(m.ym(), m.path());
            }
        }
        return files;
    }

    private int writeFiles(XSSFSheet ws, int r, String factory, Map<jp.co.pm.ai.kouchin.verify.YearMonthKey, Path> files) {
        if (files == null || files.isEmpty()) {
            return r;
        }
        for (var e : files.entrySet()) {
            Path p = e.getValue();
            String name = p == null ? "" : p.getFileName().toString();
            put(ws, r, 0, factory + " " + e.getKey().ymLabel() + ": " + name, font(false, null, 10.5, false));
            r++;
        }
        return r;
    }

    private void writeProcessCompare(
            XSSFSheet ws,
            List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months,
            TrendCrossMatrix gw,
            TrendCrossMatrix gq,
            List<String> focus) {
        put(ws, 0, 0, "工程比較（グループ合算・国分↔湖南）", font(true, null, 14, false));
        put(ws, 1, 0,
                "国分の「スライス1/3」はグループ「スライス」に合算。色は国分=赤茶(#C0504D)・湖南=若草(#9BBB59)で固定。"
                        + "グラフ系列は国分・湖南のみ（差は表参照）。ラベルは値のみ。欠損月は空欄（ゼロ埋めしない）。",
                font(false, COLOR_GRAY, 10.5, false));
        put(ws, 2, 0, "注目工程: " + String.join(", ", focus), font(false, null, 10.5, false));
        ws.setColumnWidth(0, 16 * 256);
        int row = 4;
        int n = months.size();
        for (String proc : focus) {
            put(ws, row, 0, "■ " + proc, font(true, null, 12, false));
            for (int c = 0; c <= n; c++) {
                fill(cell(ws, row, c), FILL_SECTION);
            }
            row++;
            put(ws, row, 0, "【加工量】（単位: km）", font(true, null, 10.5, false));
            int headerQ = row + 1;
            writeCompareBlock(ws, headerQ, months, gq, proc);
            addTwoSeriesChart(ws, proc + "・加工量 (km)", headerQ, 2, n, n + 2, headerQ, "km");
            int wageStart = headerQ + 16;
            put(ws, wageStart, 0, "【加工賃】（単位: 千円）", font(true, null, 10.5, false));
            int wageHeader = wageStart + 1;
            writeCompareBlock(ws, wageHeader, months, gw, proc);
            addTwoSeriesChart(ws, proc + "・加工賃 (千円)", wageHeader, 2, n, n + 2, wageHeader, "千円");
            row = wageHeader + 17;
        }
    }

    private void writeCompareBlock(
            XSSFSheet ws,
            int headerRow,
            List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months,
            TrendCrossMatrix g,
            String proc) {
        put(ws, headerRow, 0, "区分", font(true, null, 10.5, false));
        for (int j = 0; j < months.size(); j++) {
            put(ws, headerRow, j + 1, months.get(j).ymLabel(), font(true, null, 10.5, false));
        }
        for (int c = 0; c <= months.size(); c++) {
            fill(cell(ws, headerRow, c), FILL_HEADER);
        }
        put(ws, headerRow + 1, 0, "国分", font(false, null, 10.5, false));
        put(ws, headerRow + 2, 0, "湖南", font(false, null, 10.5, false));
        put(ws, headerRow + 3, 0, "差(国分-湖南)", font(false, COLOR_MUTED, 10.5, false));
        fill(cell(ws, headerRow + 3, 0), FILL_DIFF);
        for (int j = 0; j < months.size(); j++) {
            jp.co.pm.ai.kouchin.verify.YearMonthKey ym = months.get(j);
            Double k = val(g, "国分", proc, ym);
            Double h = val(g, "湖南", proc, ym);
            num(ws, headerRow + 1, j + 1, k);
            num(ws, headerRow + 2, j + 1, h);
            Double diff = (k != null && h != null) ? k - h : null;
            num(ws, headerRow + 3, j + 1, round(diff));
            fill(cell(ws, headerRow + 3, j + 1), FILL_DIFF);
        }
    }

    private void writeLoadShift(
            XSSFSheet ws,
            List<TrendShiftRow> shiftQty,
            List<TrendShiftRow> shiftWage,
            jp.co.pm.ai.kouchin.verify.YearMonthKey prev,
            jp.co.pm.ai.kouchin.verify.YearMonthKey latest) {
        put(ws, 0, 0, "負荷シフト（振替・シフト疑い一覧）", font(true, null, 14, false));
        put(ws, 1, 0,
                "判定は仮説です。「強い相殺」= 国分と湖南の前月差が異符号で大きさが近い。"
                        + "両工場合計差が大きければ需要変動の可能性が高いです。",
                font(false, COLOR_GRAY, 10.5, false));
        if (prev != null && latest != null) {
            put(ws, 2, 0, "比較: " + prev.ymLabel() + " → " + latest.ymLabel(), font(false, null, 10.5, false));
        }
        int row = 4;
        row = writeShiftTable(ws, row, "【加工量ベース】（単位: km）", shiftQty);
        row += 2;
        writeShiftTable(ws, row, "【加工賃ベース】（単位: 千円）", shiftWage);
        ws.setColumnWidth(0, 16 * 256);
        ws.setColumnWidth(1, 12 * 256);
        for (int c = 2; c < 10; c++) {
            ws.setColumnWidth(c, 12 * 256);
        }
    }

    private int writeShiftTable(XSSFSheet ws, int row, String title, List<TrendShiftRow> rows) {
        put(ws, row, 0, title, font(true, null, 10.5, false));
        row++;
        String[] headers =
                {"工程", "判定", "国分(前)", "国分(当)", "国分差", "湖南(前)", "湖南(当)", "湖南差", "合計差", "相殺量"};
        for (int j = 0; j < headers.length; j++) {
            put(ws, row, j, headers[j], font(true, null, 10.5, false));
            fill(cell(ws, row, j), FILL_HEADER);
        }
        row++;
        for (TrendShiftRow rec : rows) {
            put(ws, row, 0, rec.process(), font(false, null, 10.5, false));
            put(ws, row, 1, rec.badge(), font(false, null, 10.5, false));
            num(ws, row, 2, rec.kokubuPrev());
            num(ws, row, 3, rec.kokubuLatest());
            num(ws, row, 4, rec.dKokubu());
            num(ws, row, 5, rec.konanPrev());
            num(ws, row, 6, rec.konanLatest());
            num(ws, row, 7, rec.dKonan());
            num(ws, row, 8, rec.dCombined());
            num(ws, row, 9, rec.offset());
            if ("強い相殺".equals(rec.badge()) || "弱い相殺".equals(rec.badge())) {
                for (int c = 0; c < 10; c++) {
                    fill(cell(ws, row, c), FILL_SHIFT);
                }
            }
            row++;
        }
        return row;
    }

    private void writeAppendixFactory(
            XSSFSheet ws,
            String title,
            List<TrendMonthData> monthsData,
            List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months) {
        TrendAggregate.KindSeries wage = TrendAggregate.buildKindSeries(monthsData, months, "wage");
        TrendAggregate.KindSeries qty = TrendAggregate.buildKindSeries(monthsData, months, "qty");
        java.util.TreeSet<String> procs = new java.util.TreeSet<>();
        procs.addAll(wage.processes());
        procs.addAll(qty.processes());
        put(ws, 0, 0, title + "（付録・全工程・グラフなし）", font(true, null, 10.5, false));
        put(ws, 1, 0, "主分析は「工程比較」「負荷シフト」を使ってください。", font(false, null, 10.5, false));
        put(ws, 3, 0, "加工賃（単位: 千円）", font(false, null, 10.5, false));
        put(ws, 4, 0, "工程", font(true, null, 10.5, false));
        for (int j = 0; j < months.size(); j++) {
            put(ws, 4, j + 1, months.get(j).ymLabel(), font(true, null, 10.5, false));
        }
        for (int c = 0; c <= months.size(); c++) {
            fill(cell(ws, 4, c), FILL_HEADER);
        }
        int i = 5;
        for (String p : procs) {
            put(ws, i, 0, p, font(false, null, 10.5, false));
            for (int j = 0; j < months.size(); j++) {
                num(ws, i, j + 1, wage.values().getOrDefault(p, Map.of()).get(months.get(j)));
            }
            i++;
        }
        int start = 5 + procs.size() + 2;
        put(ws, start, 0, "加工量（単位: km）", font(true, null, 10.5, false));
        put(ws, start + 1, 0, "工程", font(true, null, 10.5, false));
        for (int j = 0; j < months.size(); j++) {
            put(ws, start + 1, j + 1, months.get(j).ymLabel(), font(true, null, 10.5, false));
        }
        for (int c = 0; c <= months.size(); c++) {
            fill(cell(ws, start + 1, c), FILL_HEADER);
        }
        int qi = start + 2;
        for (String p : procs) {
            put(ws, qi, 0, p, font(false, null, 10.5, false));
            for (int j = 0; j < months.size(); j++) {
                num(ws, qi, j + 1, qty.values().getOrDefault(p, Map.of()).get(months.get(j)));
            }
            qi++;
        }
        ws.setColumnWidth(0, 18 * 256);
    }

    private void writeAppendixSheets(
            XSSFSheet ws,
            List<jp.co.pm.ai.kouchin.verify.YearMonthKey> months,
            List<TrendMonthData> kokubu,
            List<TrendMonthData> konan) {
        put(ws, 0, 0, "付録: シート区分（工場並記）", font(true, null, 10.5, false));
        int row = 2;
        for (var pair : List.of(Map.entry("国分", kokubu), Map.entry("湖南", konan))) {
            TrendAggregate.SheetSeries wage = TrendAggregate.buildSheetSeries(pair.getValue(), months, "wage");
            put(ws, row, 0, pair.getKey(), font(true, null, 10.5, false));
            row++;
            put(ws, row, 0, "シート", font(true, null, 10.5, false));
            for (int j = 0; j < months.size(); j++) {
                put(ws, row, j + 1, months.get(j).ymLabel(), font(true, null, 10.5, false));
            }
            for (int c = 0; c <= months.size(); c++) {
                fill(cell(ws, row, c), FILL_HEADER);
            }
            row++;
            for (String name : wage.sheetNames()) {
                put(ws, row, 0, name, font(false, null, 10.5, false));
                for (int j = 0; j < months.size(); j++) {
                    num(ws, row, j + 1, wage.values().getOrDefault(name, Map.of()).get(months.get(j)));
                }
                row++;
            }
            row += 2;
        }
        ws.setColumnWidth(0, 18 * 256);
    }

    private void addTwoSeriesChart(
            XSSFSheet sheet,
            String title,
            int headerRow,
            int dataRows,
            int nMonths,
            int anchorCol0,
            int anchorRow1,
            String unit) {
        if (dataRows < 1 || nMonths < 1) {
            return;
        }
        int header0 = headerRow;
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(0, 0, 0, 0, anchorCol0, header0, anchorCol0 + 10, header0 + 14);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(title);
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);
        XDDFCategoryAxis bottom = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis left = chart.createValueAxis(AxisPosition.LEFT);
        left.setTitle(unit);
        left.setCrosses(AxisCrosses.AUTO_ZERO);
        Double ymax = null;
        for (int r = header0 + 1; r <= header0 + dataRows; r++) {
            for (int c = 1; c <= nMonths; c++) {
                Cell cell = cell(sheet, r, c);
                if (cell != null && cell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                    double v = cell.getNumericCellValue();
                    ymax = ymax == null ? v : Math.max(ymax, v);
                }
            }
        }
        if (ymax != null && ymax > 0) {
            left.setMinimum(0);
            left.setMaximum(ymax * 1.25);
        }
        XDDFCategoryDataSource cats =
                XDDFDataSourcesFactory.fromStringCellRange(
                        sheet, new CellRangeAddress(header0, header0, 1, nMonths));
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottom, left);
        String[] titles = {"国分", "湖南"};
        String[] colors = {COLOR_KOKUBU, COLOR_KONAN};
        for (int i = 0; i < dataRows; i++) {
            XDDFNumericalDataSource<Double> vals =
                    XDDFDataSourcesFactory.fromNumericCellRange(
                            sheet, new CellRangeAddress(header0 + 1 + i, header0 + 1 + i, 1, nMonths));
            var series = data.addSeries(cats, vals);
            series.setTitle(titles[i], null);
            if (series instanceof XDDFLineChartData.Series lineSeries) {
                lineSeries.setSmooth(false);
                lineSeries.setMarkerStyle(MarkerStyle.CIRCLE);
                XDDFSolidFillProperties fill =
                        new XDDFSolidFillProperties(XDDFColor.from(rgb(colors[i])));
                XDDFLineProperties line = new XDDFLineProperties();
                line.setFillProperties(fill);
                XDDFShapeProperties shape = new XDDFShapeProperties();
                shape.setLineProperties(line);
                lineSeries.setShapeProperties(shape);
            }
        }
        chart.plot(data);
        if (chart.getCTChart().getDispBlanksAs() == null) {
            chart.getCTChart().addNewDispBlanksAs();
        }
        chart.getCTChart().getDispBlanksAs().setVal(STDispBlanksAs.GAP);
        if (chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0) {
            CTLineChart line = chart.getCTChart().getPlotArea().getLineChartArray(0);
            for (int i = 0; i < line.sizeOfSerArray() && i < colors.length; i++) {
                CTLineSer ser = line.getSerArray(i);
                if (ser.getSpPr() == null) {
                    ser.addNewSpPr();
                }
                if (ser.getSpPr().getLn() == null) {
                    ser.getSpPr().addNewLn();
                }
                var ln = ser.getSpPr().getLn();
                if (ln.isSetSolidFill()) {
                    ln.unsetSolidFill();
                }
                ln.addNewSolidFill().addNewSrgbClr().setVal(rgb(colors[i]));
                if (ser.getMarker() == null) {
                    ser.addNewMarker();
                }
                if (ser.getMarker().getSymbol() == null) {
                    ser.getMarker().addNewSymbol();
                }
                ser.getMarker().getSymbol().setVal(STMarkerStyle.CIRCLE);
                CTDLbls dLbls = ser.isSetDLbls() ? ser.getDLbls() : ser.addNewDLbls();
                applyValueLabels(dLbls);
            }
            CTDLbls chartLabels = line.isSetDLbls() ? line.getDLbls() : line.addNewDLbls();
            applyValueLabels(chartLabels);
        }
    }

    private static void applyValueLabels(CTDLbls dLbls) {
        if (dLbls.getShowVal() == null) {
            dLbls.addNewShowVal();
        }
        dLbls.getShowVal().setVal(true);
        if (dLbls.getShowCatName() == null) {
            dLbls.addNewShowCatName();
        }
        dLbls.getShowCatName().setVal(false);
        if (dLbls.getShowSerName() == null) {
            dLbls.addNewShowSerName();
        }
        dLbls.getShowSerName().setVal(false);
        if (dLbls.getShowLegendKey() == null) {
            dLbls.addNewShowLegendKey();
        }
        dLbls.getShowLegendKey().setVal(false);
        if (dLbls.getShowPercent() == null) {
            dLbls.addNewShowPercent();
        }
        dLbls.getShowPercent().setVal(false);
        if (dLbls.getDLblPos() == null) {
            dLbls.addNewDLblPos();
        }
        dLbls.getDLblPos().setVal(STDLblPos.T);
        if (dLbls.getNumFmt() == null) {
            dLbls.addNewNumFmt();
        }
        dLbls.getNumFmt().setFormatCode("#,##0");
    }

    private static Double val(TrendCrossMatrix g, String factory, String proc, jp.co.pm.ai.kouchin.verify.YearMonthKey ym) {
        Map<String, Map<jp.co.pm.ai.kouchin.verify.YearMonthKey, Double>> fac = g.values().get(factory);
        if (fac == null) {
            return null;
        }
        Map<jp.co.pm.ai.kouchin.verify.YearMonthKey, Double> pm = fac.get(proc);
        if (pm == null) {
            return null;
        }
        return pm.get(ym);
    }

    private static Double round(Double v) {
        return v == null ? null : (double) TrendText.roundDisp(v);
    }

    private void tabColor(XSSFSheet ws, String name) {
        ws.setTabColor(new XSSFColor(rgb(TAB.get(name)), null));
    }

    private XSSFFont font(boolean bold, String colorHex, double size, boolean numeric) {
        String key = bold + "/" + colorHex + "/" + size + "/" + numeric;
        return fonts.computeIfAbsent(
                key,
                k -> {
                    XSSFFont f = wb.createFont();
                    f.setFontName(numeric ? numFont : baseFont);
                    f.setFontHeight(size);
                    f.setBold(bold);
                    if (colorHex != null) {
                        f.setColor(new XSSFColor(rgb(colorHex), null));
                    }
                    return f;
                });
    }

    private XSSFCellStyle style(XSSFFont font, String fill, boolean numeric) {
        String key = font.hashCode() + "/" + fill + "/" + numeric;
        return styles.computeIfAbsent(
                key,
                k -> {
                    XSSFCellStyle cs = wb.createCellStyle();
                    cs.setFont(font);
                    if (fill != null) {
                        cs.setFillForegroundColor(new XSSFColor(rgb(fill), null));
                        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                    if (numeric) {
                        cs.setDataFormat(fmtInt);
                    }
                    return cs;
                });
    }

    private Cell put(XSSFSheet ws, int row0, int col, String value, XSSFFont font) {
        Cell cell = cell(ws, row0, col);
        cell.setCellValue(value == null ? "" : value);
        cell.setCellStyle(style(font, null, false));
        return cell;
    }

    private void num(XSSFSheet ws, int row0, int col, Double value) {
        Cell cell = cell(ws, row0, col);
        cell.setCellStyle(style(font(false, null, 10.5, true), null, true));
        if (value == null) {
            return;
        }
        cell.setCellValue(value);
    }

    private void fill(Cell cell, String fill) {
        if (cell == null) {
            return;
        }
        XSSFFont f = wb.getFontAt(cell.getCellStyle().getFontIndex());
        boolean numeric = cell.getCellStyle().getDataFormat() == fmtInt;
        cell.setCellStyle(style(f, fill, numeric));
    }

    private static Cell cell(XSSFSheet ws, int row0, int col) {
        Row row = ws.getRow(row0);
        if (row == null) {
            row = ws.createRow(row0);
        }
        Cell c = row.getCell(col);
        return c != null ? c : row.createCell(col);
    }

    private static byte[] rgb(String hex) {
        return new byte[] {
            (byte) Integer.parseInt(hex.substring(0, 2), 16),
            (byte) Integer.parseInt(hex.substring(2, 4), 16),
            (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }

    private static String pickFont(String... candidates) {
        Set<String> available;
        try {
            available =
                    new HashSet<>(
                            Arrays.asList(
                                    GraphicsEnvironment.getLocalGraphicsEnvironment()
                                            .getAvailableFontFamilyNames(Locale.JAPAN)));
        } catch (Throwable t) {
            available = Set.of();
        }
        for (String name : candidates) {
            if (available.isEmpty() || available.contains(name)) {
                return name;
            }
        }
        return candidates[candidates.length - 1];
    }
}
