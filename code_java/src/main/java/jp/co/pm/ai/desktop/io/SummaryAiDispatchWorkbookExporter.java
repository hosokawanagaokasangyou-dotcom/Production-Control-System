package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.ui.DeliveryCalendarMainCell;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;

/**
 * 納期管理ビュー4表を {@code code/サマリ_AI配台.xlsx} の表デザインを参考に {@link
 * AppPaths#summaryAiDispatchXlsxPath(Map)} へ上書き出力する。
 */
public final class SummaryAiDispatchWorkbookExporter {

    /** 納期管理ビュー子タブ「アラ・実績・シス比較」 */
    public static final String SHEET_MAIN_COMPARE = "\u30a2\u30e9\u30fb\u5b9f\u7e3e\u30fb\u30b7\u30b9\u6bd4\u8f03";

    public static final String SHEET_DISPATCH = "\u914d\u53f0\u7d50\u679c";
    public static final String SHEET_ACTUALS = "\u52a0\u5de5\u5b9f\u7e3e";

    /** 納期管理ビュー子タブ「アラジン加工計画取得データ」 */
    public static final String SHEET_ALADDIN =
            "\u30a2\u30e9\u30b8\u30f3\u52a0\u5de5\u8a08\u753b\u53d6\u5f97\u30c7\u30fc\u30bf";

    /** 日付列の (シス配台) 数量を横集計した列（サマリ Excel のみ。UI 列とは別）。 */
    public static final String COL_SYSTEM_DISPATCH_QTY_TOTAL = "\u30b7\u30b9\u30c6\u30e0\u914d\u53f0\u6570\u91cf\u5408\u8a08";

    private static final Pattern CAL_DATE_HDR =
            Pattern.compile("(\\d{4})\u5e74(\\d{1,2})\u6708(\\d{1,2})\u65e5\\([\u6708\u706b\u6c34\u6728\u91d1\u571f\u65e5]\\)");

    private static final String TEMPLATE_BASENAME = "\u30b5\u30de\u30ea_AI\u914d\u53f0.xlsx";

    private static final String LEGACY_TEMPLATE_BASENAME = "\u30b5\u30de\u30ea_AI\u914d\u53f0.xlsm";

    private SummaryAiDispatchWorkbookExporter() {}

    /**
     * 4シートを出力先へ上書き保存する。親ディレクトリが無ければ作成する。
     *
     * @return 書き込んだ絶対パス
     */
    public static Path writeOverwrite(
            Map<String, String> ui,
            PlanInputTabularIo.TabularSheet mainCompare,
            PlanInputTabularIo.TabularSheet actuals,
            PlanInputTabularIo.TabularSheet aladdin,
            PlanInputTabularIo.TabularSheet dispatch)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path output = AppPaths.summaryAiDispatchXlsxPath(u);
        if (output.getParent() != null) {
            Files.createDirectories(output.getParent());
        }
        Path template = resolveTemplatePath(u);
        ensureOutputWorkbookExists(output, template);
        SummaryAiDispatchExportPrefs.ExportPrefs prefs = SummaryAiDispatchExportPrefs.load();
        try (Workbook wb = openWorkbookForUpdate(output, template)) {
            SheetStyles styles = SheetStyles.of(wb, prefs);
            replaceSheet(
                    wb,
                    SHEET_MAIN_COMPARE,
                    layoutForExport(mainCompare, SummaryAiDispatchExportPrefs.SheetKey.MAIN_COMPARE, prefs),
                    styles,
                    prefs.sheet(SummaryAiDispatchExportPrefs.SheetKey.MAIN_COMPARE).frozenColumnCount());
            replaceSheet(
                    wb,
                    SHEET_DISPATCH,
                    layoutForExport(dispatch, SummaryAiDispatchExportPrefs.SheetKey.DISPATCH, prefs),
                    styles,
                    prefs.sheet(SummaryAiDispatchExportPrefs.SheetKey.DISPATCH).frozenColumnCount());
            replaceSheet(
                    wb,
                    SHEET_ACTUALS,
                    layoutForExport(actuals, SummaryAiDispatchExportPrefs.SheetKey.ACTUALS, prefs),
                    styles,
                    prefs.sheet(SummaryAiDispatchExportPrefs.SheetKey.ACTUALS).frozenColumnCount());
            replaceSheet(
                    wb,
                    SHEET_ALADDIN,
                    layoutForExport(aladdin, SummaryAiDispatchExportPrefs.SheetKey.ALADDIN, prefs),
                    styles,
                    prefs.sheet(SummaryAiDispatchExportPrefs.SheetKey.ALADDIN).frozenColumnCount());
            Path tmp = output.resolveSibling(output.getFileName() + ".tmp");
            try (var out = Files.newOutputStream(tmp)) {
                wb.write(out);
            }
            Files.move(tmp, output, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        }
        return output.toAbsolutePath().normalize();
    }

    /** 各シートの行数要約（ログ用）。 */
    public static String rowCountSummary(
            PlanInputTabularIo.TabularSheet mainCompare,
            PlanInputTabularIo.TabularSheet dispatch,
            PlanInputTabularIo.TabularSheet actuals,
            PlanInputTabularIo.TabularSheet aladdin) {
        return SHEET_MAIN_COMPARE
                + "="
                + rowCount(mainCompare)
                + " "
                + SHEET_DISPATCH
                + "="
                + rowCount(dispatch)
                + " "
                + SHEET_ACTUALS
                + "="
                + rowCount(actuals)
                + " "
                + SHEET_ALADDIN
                + "="
                + rowCount(aladdin);
    }

    private static int rowCount(PlanInputTabularIo.TabularSheet sheet) {
        return sheet != null && sheet.rows() != null ? sheet.rows().size() : 0;
    }

    /** shaped JSON・配台 JSON から4表を読み、{@link #writeOverwrite} する（メイン表は空）。 */
    public static Path writeFromCachedSources(Map<String, String> ui) throws IOException {
        return writeFromPipelineArtifacts(ui, emptySheet());
    }

    /**
     * 段階2/段階3 直後など: ディスク上の配台・成形 JSON と、任意のメイン表スナップショットでサマリを更新する。
     *
     * @param mainCompare メイン表が未読込のときは {@link #emptySheet()} を渡す
     */
    public static Path writeFromPipelineArtifacts(
            Map<String, String> ui, PlanInputTabularIo.TabularSheet mainCompare) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        PlanInputTabularIo.TabularSheet dispatch = loadDispatch(u);
        PlanInputTabularIo.TabularSheet actuals = loadArrayTableOrEmpty(
                AppPaths.resolveShapedProcessingActualsJsonPath(u));
        PlanInputTabularIo.TabularSheet aladdin =
                loadArrayTableOrEmpty(AppPaths.resolveShapedAladdinPlanJsonPath(u));
        PlanInputTabularIo.TabularSheet main =
                mainCompare != null ? mainCompare : emptySheet();
        return writeOverwrite(u, main, actuals, aladdin, dispatch);
    }

    private static PlanInputTabularIo.TabularSheet loadDispatch(Map<String, String> ui) throws IOException {
        Path json = AppPaths.resolveResultDispatchTableJsonPath(ui);
        if (!Files.isRegularFile(json)) {
            return emptySheet();
        }
        return JsonTableIo.loadFlatTable(json).toTabularSheet();
    }

    private static PlanInputTabularIo.TabularSheet loadArrayTableOrEmpty(Path path) throws IOException {
        if (!Files.isRegularFile(path)) {
            return emptySheet();
        }
        JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(path);
        return new PlanInputTabularIo.TabularSheet(t.columns(), t.rows());
    }

    private static Path resolveTemplatePath(Map<String, String> ui) {
        Path code = AppPaths.resolveRepoRoot(ui != null ? ui : Map.of()).resolve("code");
        Path xlsx = code.resolve(TEMPLATE_BASENAME).normalize();
        if (Files.isRegularFile(xlsx)) {
            return xlsx;
        }
        Path xlsm = code.resolve(LEGACY_TEMPLATE_BASENAME).normalize();
        return Files.isRegularFile(xlsm) ? xlsm : null;
    }

    private static PlanInputTabularIo.TabularSheet emptySheet() {
        return new PlanInputTabularIo.TabularSheet(List.of(), List.of());
    }

    /** 出力先が無いときだけテンプレートをコピーする（既存 {@code サマリ_AI配台.xlsx} は上書き更新）。 */
    private static void ensureOutputWorkbookExists(Path output, Path template) throws IOException {
        if (Files.isRegularFile(output)) {
            return;
        }
        if (template != null && Files.isRegularFile(template)) {
            Files.copy(template, output);
            return;
        }
        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet(SHEET_MAIN_COMPARE);
            try (var out = Files.newOutputStream(output)) {
                wb.write(out);
            }
        }
    }

    private static Workbook openWorkbookForUpdate(Path output, Path template) throws IOException {
        if (Files.isRegularFile(output)) {
            return WorkbookFactory.create(output.toFile());
        }
        if (template != null && Files.isRegularFile(template)) {
            return WorkbookFactory.create(template.toFile());
        }
        return new XSSFWorkbook();
    }

    private static PlanInputTabularIo.TabularSheet layoutForExport(
            PlanInputTabularIo.TabularSheet data,
            SummaryAiDispatchExportPrefs.SheetKey sheetKey,
            SummaryAiDispatchExportPrefs.ExportPrefs prefs) {
        PlanInputTabularIo.TabularSheet src = data != null ? data : emptySheet();
        return SummaryAiDispatchExportColumnSupport.applySheetLayout(
                src, sheetKey, prefs.sheet(sheetKey).nonDateColumnOrder());
    }

    private static void replaceSheet(
            Workbook wb,
            String sheetName,
            PlanInputTabularIo.TabularSheet data,
            SheetStyles styles,
            int frozenColumnCount) {
        int idx = wb.getSheetIndex(sheetName);
        if (idx >= 0) {
            wb.removeSheetAt(idx);
        }
        Sheet sh = wb.createSheet(sheetName);
        writeTabular(sh, data != null ? data : emptySheet(), styles, frozenColumnCount);
    }

    private static final String EMPTY_SHEET_HINT =
            "\uff08\u30c7\u30fc\u30bf\u306a\u3057: \u7d0d\u671f\u7ba1\u7406\u30d3\u30e5\u30fc\u3067\u518d\u8aad\u307f\u8fbc\u307f\u5f8c\u306b\u51fa\u529b\uff09";

    private static void writeTabular(
            Sheet sh, PlanInputTabularIo.TabularSheet data, SheetStyles styles, int frozenColumnCount) {
        List<String> headers = data.headers() != null ? data.headers() : List.of();
        List<List<String>> rows = data.rows() != null ? data.rows() : List.of();
        if (headers.isEmpty() && rows.isEmpty()) {
            Row hr = sh.createRow(0);
            Cell cell = hr.createCell(0);
            cell.setCellValue(EMPTY_SHEET_HINT);
            cell.setCellStyle(styles.dataWrap());
            sh.autoSizeColumn(0);
            return;
        }
        Row hr = sh.createRow(0);
        hr.setHeightInPoints(20f);
        for (int c = 0; c < headers.size(); c++) {
            Cell cell = hr.createCell(c);
            cell.setCellValue(headers.get(c) != null ? headers.get(c) : "");
            cell.setCellStyle(styles.header());
        }
        int r = 1;
        for (List<String> rowVals : rows) {
            Row rr = sh.createRow(r++);
            for (int c = 0; c < headers.size(); c++) {
                Cell cell = rr.createCell(c);
                String v =
                        rowVals != null && c < rowVals.size() && rowVals.get(c) != null
                                ? rowVals.get(c)
                                : "";
                cell.setCellValue(v);
                cell.setCellStyle(styles.dataWrap());
            }
        }
        if (!headers.isEmpty()) {
            sh.setAutoFilter(new CellRangeAddress(0, 0, 0, headers.size() - 1));
            int frozenCols = Math.max(0, Math.min(frozenColumnCount, headers.size()));
            sh.createFreezePane(frozenCols, 1);
        }
        for (int c = 0; c < headers.size(); c++) {
            sh.autoSizeColumn(c);
            int w = sh.getColumnWidth(c);
            int max = 256 * 48;
            if (w > max) {
                sh.setColumnWidth(c, max);
            }
        }
    }

    /** 納期管理メイン表（ControlsFX 表示と同系のセル文字列）。日付列の直前に {@link #COL_SYSTEM_DISPATCH_QTY_TOTAL} を挿入する。 */
    public static PlanInputTabularIo.TabularSheet mainCompareFromUi(
            List<String> headers, List<? extends List<? extends DeliveryCalendarMainCell>> rows) {
        List<String> srcHdrs = headers != null ? new ArrayList<>(headers) : new ArrayList<>();
        List<String> outHdrs = new ArrayList<>(srcHdrs);
        int firstDate = indexOfFirstDateColumn(outHdrs);
        int insertAt = firstDate >= 0 ? firstDate : outHdrs.size();
        int totalColIdx = outHdrs.indexOf(COL_SYSTEM_DISPATCH_QTY_TOTAL);
        if (totalColIdx < 0) {
            outHdrs.add(insertAt, COL_SYSTEM_DISPATCH_QTY_TOTAL);
            totalColIdx = insertAt;
        }

        List<List<String>> outRows = new ArrayList<>();
        if (rows != null) {
            for (List<? extends DeliveryCalendarMainCell> row : rows) {
                outRows.add(buildMainCompareExportRow(srcHdrs, row, outHdrs, totalColIdx));
            }
        }
        return new PlanInputTabularIo.TabularSheet(outHdrs, outRows);
    }

    private static List<String> buildMainCompareExportRow(
            List<String> srcHdrs,
            List<? extends DeliveryCalendarMainCell> srcRow,
            List<String> outHdrs,
            int totalColIdx) {
        List<String> line = new ArrayList<>(outHdrs.size());
        for (int dst = 0; dst < outHdrs.size(); dst++) {
            if (dst == totalColIdx) {
                line.add(formatSystemDispatchQtyTotal(srcHdrs, srcRow));
            } else {
                int srcCol = dst < totalColIdx ? dst : dst - 1;
                DeliveryCalendarMainCell cell =
                        srcRow != null && srcCol >= 0 && srcCol < srcRow.size()
                                ? srcRow.get(srcCol)
                                : null;
                line.add(formatMainCell(cell));
            }
        }
        return line;
    }

    private static String formatSystemDispatchQtyTotal(
            List<String> srcHdrs, List<? extends DeliveryCalendarMainCell> row) {
        double sum = 0;
        for (int c = 0; c < srcHdrs.size(); c++) {
            if (!isDateColumnHeader(srcHdrs.get(c))) {
                continue;
            }
            DeliveryCalendarMainCell cell = row != null && c < row.size() ? row.get(c) : null;
            if (cell instanceof DeliveryCalendarMainCell.TripleQty t
                    && !tripleQtyHidden(t.dispatch())) {
                sum += ResultDispatchNormalizer.parseDouble(t.dispatch());
            }
        }
        if (sum <= 1e-3) {
            return "";
        }
        return ResultDispatchNormalizer.formatQty(sum);
    }

    private static int indexOfFirstDateColumn(List<String> headers) {
        for (int i = 0; i < headers.size(); i++) {
            if (isDateColumnHeader(headers.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static boolean isDateColumnHeader(String header) {
        return header != null && CAL_DATE_HDR.matcher(header).matches();
    }

    private static String formatMainCell(DeliveryCalendarMainCell cell) {
        if (cell == null) {
            return "";
        }
        if (cell instanceof DeliveryCalendarMainCell.PlainText p) {
            return p.text() != null ? p.text() : "";
        }
        if (cell instanceof DeliveryCalendarMainCell.TripleQty t) {
            return formatTripleForExcel(t);
        }
        return "";
    }

    private static String formatTripleForExcel(DeliveryCalendarMainCell.TripleQty t) {
        List<String> lines = new ArrayList<>(3);
        if (!tripleQtyHidden(t.plan())) {
            lines.add(SpreadsheetTabularSupport.deliveryCalendarPlanLineForInspector(t.plan()));
        }
        if (!tripleQtyHidden(t.actual())) {
            lines.add(SpreadsheetTabularSupport.deliveryCalendarActualLineForInspector(t.actual()));
        }
        if (!tripleQtyHidden(t.dispatch())) {
            lines.add(
                    SpreadsheetTabularSupport.deliveryCalendarDispatchLineForInspector(t.dispatch()));
        }
        return String.join("\n", lines);
    }

    private static boolean tripleQtyHidden(String qty) {
        if (qty == null || qty.isBlank()) {
            return true;
        }
        String s = qty.strip();
        if ("\u2014".equals(s) || "-".equals(s)) {
            return true;
        }
        try {
            double v = Double.parseDouble(s.replace(",", ""));
            return !Double.isNaN(v) && !Double.isInfinite(v) && v == 0d;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private record SheetStyles(CellStyle header, CellStyle dataWrap) {
        static SheetStyles of(Workbook wb, SummaryAiDispatchExportPrefs.ExportPrefs prefs) {
            SummaryAiDispatchExportPrefs.ExportPrefs p =
                    prefs != null ? prefs : SummaryAiDispatchExportPrefs.ExportPrefs.defaults();
            SummaryAiDispatchExportThemePalette palette =
                    SummaryAiDispatchExportThemePalette.forTheme(p.theme());
            boolean xssf = wb instanceof XSSFWorkbook;
            String fontName =
                    p.fontFamily() != null && !p.fontFamily().isBlank()
                            ? p.fontFamily()
                            : SummaryAiDispatchExportPrefs.DEFAULT_FONT_FAMILY;
            short fontPt = (short) p.fontSizePt();

            Font headerFont = wb.createFont();
            headerFont.setFontName(fontName);
            headerFont.setFontHeightInPoints(fontPt);
            headerFont.setBold(true);
            if (xssf) {
                XSSFColor hf = new XSSFColor(palette.headerFontRgb(), null);
                ((org.apache.poi.xssf.usermodel.XSSFFont) headerFont).setColor(hf);
            }

            Font dataFont = wb.createFont();
            dataFont.setFontName(fontName);
            dataFont.setFontHeightInPoints(fontPt);
            if (xssf) {
                XSSFColor df = new XSSFColor(palette.dataFontRgb(), null);
                ((org.apache.poi.xssf.usermodel.XSSFFont) dataFont).setColor(df);
            }

            CellStyle header = wb.createCellStyle();
            header.setFont(headerFont);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setBorderTop(BorderStyle.THIN);
            header.setBorderBottom(BorderStyle.THIN);
            header.setBorderLeft(BorderStyle.THIN);
            header.setBorderRight(BorderStyle.THIN);
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if (xssf) {
                XSSFCellStyle xh = (XSSFCellStyle) header;
                xh.setFillForegroundColor(new XSSFColor(palette.headerFillRgb(), null));
            } else {
                header.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            }

            CellStyle dataWrap = wb.createCellStyle();
            dataWrap.setFont(dataFont);
            dataWrap.setAlignment(HorizontalAlignment.LEFT);
            dataWrap.setVerticalAlignment(VerticalAlignment.TOP);
            dataWrap.setWrapText(true);
            dataWrap.setBorderTop(BorderStyle.THIN);
            dataWrap.setBorderBottom(BorderStyle.THIN);
            dataWrap.setBorderLeft(BorderStyle.THIN);
            dataWrap.setBorderRight(BorderStyle.THIN);
            dataWrap.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if (xssf) {
                XSSFCellStyle xd = (XSSFCellStyle) dataWrap;
                xd.setFillForegroundColor(new XSSFColor(palette.dataFillRgb(), null));
            }

            return new SheetStyles(header, dataWrap);
        }
    }
}
