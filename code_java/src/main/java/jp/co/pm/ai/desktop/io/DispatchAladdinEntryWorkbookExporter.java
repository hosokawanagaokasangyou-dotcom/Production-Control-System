package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchInteractiveConsolidator;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage3Support;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingKouteiNaiyoMasterLookup;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingPlanMachineLookup;

/**
 * アラジン入力用配台計画 Excel（機械名ごとのシート、日別2段セル）を
 * {@link AppPaths#aladdinEntryDispatchPlanXlsxPath} へ上書き出力し、操作者別世代フォルダへも保存する。
 */
public final class DispatchAladdinEntryWorkbookExporter {

    /** 操作者フォルダあたりの世代保持上限（ファイル数）。 */
    public static final int MAX_GENERATIONS_PER_USER = 20;

    /** 世代ファイル名: {@code アラジン入力用_配台計画_yyyyMMdd-HHmmss.xlsx} */
    private static final DateTimeFormatter GEN_TS = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");

    private static final String GEN_FILE_PREFIX = "アラジン入力用_配台計画_";

    private static final String[] FIXED_HEADERS = {
        "依頼NO", "契約NO", "工程名", "原反投入日", "回答納期", "完了日チェック",
        "換算数量", "加工完了数量", "配台合計", "数量チェック",
    };

    private static final int FIXED_COLUMN_COUNT = FIXED_HEADERS.length;

    /** 日付列幅（Excel 文字数）。POI は 1/256 文字単位。 */
    private static final double DATE_COLUMN_WIDTH_CHARS = 16.8;

    private static final String[] WEEKDAY_JA = {"月", "火", "水", "木", "金", "土", "日"};

    private static final String EMPTY_SHEET_NAME = "データなし";

    /** 各シートの表示ズーム（%）。 */
    private static final int SHEET_VIEW_ZOOM_PERCENT = 75;

    /** 日付セル上段（現アラ計）のフォントサイズ（pt）。 */
    private static final short ALADDIN_LINE_FONT_SIZE_PT = 9;

    /** 日付セル下段（シス計）のフォントサイズ（pt）。 */
    private static final short SYSTEM_LINE_FONT_SIZE_PT = 12;

    private DispatchAladdinEntryWorkbookExporter() {}

    /** 出力結果（最新固定パスと世代パス）。 */
    public record ExportResult(Path latestPath, Path generationPath) {}

    /**
     * 出力先。{@link #SHARED} はサマリ Excel と同フォルダ側（共有ドライブ想定）、
     * {@link #LOCAL} はリポジトリ {@code code/アラジン入力用配台計画}。
     */
    public enum Destination {
        SHARED,
        LOCAL
    }

    /**
     * ディスク上の 結果_配台表.json / shaped_aladdin_plan.json と目次情報からブックを組み立てて出力する。
     *
     * @param indexByTid {@code RequestFormOriginalIndexLookup.loadByIraiNoKey} の結果（null 可）
     */
    public static ExportResult writeFromCachedSources(
            Map<String, String> ui,
            Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> indexByTid)
            throws IOException {
        return writeFromCachedSources(ui, indexByTid, Destination.SHARED);
    }

    /**
     * @param destination {@link Destination#SHARED}（既定）または {@link Destination#LOCAL}
     */
    public static ExportResult writeFromCachedSources(
            Map<String, String> ui,
            Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> indexByTid,
            Destination destination)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path dispatchJson = AppPaths.resolveResultDispatchTableJsonPath(u);
        if (!Files.isRegularFile(dispatchJson)) {
            throw new IOException("結果_配台表.json が見つかりません: " + dispatchJson);
        }
        JsonTableIo.SheetTable table = JsonTableIo.loadFlatTable(dispatchJson);
        List<String> columns = new ArrayList<>(table.columns());
        List<Map<String, String>> rows = new ArrayList<>();
        for (Map<String, String> r : table.rows()) {
            rows.add(new java.util.LinkedHashMap<>(r));
        }
        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(columns, rows);
        if (ResultDispatchStage3Support.hasStage3ActualColumn(columns)) {
            ResultDispatchStage3Support.applyStage3DisplayQuantities(columns, rows);
            ResultDispatchStage3Support.removeRedundantActualColumnFromMaps(columns, rows);
        }

        AladdinShapedPlanQtyLookup.ShapedTable shaped =
                AladdinShapedPlanQtyLookup.loadShapedTable(
                        AppPaths.resolveShapedAladdinPlanJsonPath(u));
        Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup =
                AladdinShapedPlanQtyLookup.buildLookup(shaped.headers(), shaped.rows());

        DispatchAladdinEntrySheetBuilder.EntryWorkbook model =
                DispatchAladdinEntrySheetBuilder.build(
                        columns, rows, aladdinLookup, indexByTid, LocalDate.now());
        return write(u, model, destination);
    }

    /** モデルを共有側の最新固定パスへ上書きし、操作者別世代フォルダへコピー・剪定する。 */
    public static ExportResult write(
            Map<String, String> ui, DispatchAladdinEntrySheetBuilder.EntryWorkbook model)
            throws IOException {
        return write(ui, model, Destination.SHARED);
    }

    /** モデルを指定出力先の最新固定パスへ上書きし、操作者別世代フォルダへコピー・剪定する。 */
    public static ExportResult write(
            Map<String, String> ui,
            DispatchAladdinEntrySheetBuilder.EntryWorkbook model,
            Destination destination)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Destination dest = destination != null ? destination : Destination.SHARED;
        Path latest =
                dest == Destination.LOCAL
                        ? AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(u)
                        : AppPaths.aladdinEntryDispatchPlanXlsxPath(u);
        Path repoRoot = AppPaths.resolveRepoRoot(u);
        String tmpSuffix = dest == Destination.LOCAL ? ".local.tmp" : ".tmp";
        Path stagingTmp =
                repoRoot.resolve(AppPaths.ALADDIN_ENTRY_DISPATCH_PLAN_XLSX + tmpSuffix)
                        .toAbsolutePath()
                        .normalize();
        PostProcessingPlanMachineLookup.Snapshot machineSnap = loadMachineSnapshot(u);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles styles = Styles.of(wb);
            if (model == null || model.sheets().isEmpty()) {
                writeEmptySheet(wb, styles);
            } else {
                for (DispatchAladdinEntrySheetBuilder.MachineSheet ms : model.sheets()) {
                    writeMachineSheet(wb, ms, model.dates(), styles, machineSnap);
                }
            }
            try (var out = Files.newOutputStream(stagingTmp)) {
                wb.write(out);
            }
            Files.createDirectories(latest.getParent());
            Files.copy(stagingTmp, latest, StandardCopyOption.REPLACE_EXISTING);
        } finally {
            Files.deleteIfExists(stagingTmp);
        }
        Path generation = saveGenerationCopy(u, latest, dest);
        return new ExportResult(latest, generation);
    }

    /** 最新ファイルを操作者別世代フォルダへコピーし、上限超過分を古い順に削除する。 */
    private static Path saveGenerationCopy(
            Map<String, String> ui, Path latest, Destination destination) throws IOException {
        String operator = SummaryAiDispatchGenerationStore.resolveOperatorUser(ui);
        Path operatorDir = operatorGenerationDir(ui, operator, destination);
        Files.createDirectories(operatorDir);
        String fileName =
                GEN_FILE_PREFIX + GEN_TS.format(LocalDateTime.now()) + ".xlsx";
        Path generation = operatorDir.resolve(fileName);
        Files.copy(latest, generation, StandardCopyOption.REPLACE_EXISTING);
        pruneGenerations(operatorDir);
        return generation;
    }

    /** 操作者別世代フォルダ（共有側 {@link AppPaths#aladdinEntryDispatchPlanDir} 配下）。 */
    public static Path operatorGenerationDir(Map<String, String> ui, String operatorUser) {
        return operatorGenerationDir(ui, operatorUser, Destination.SHARED);
    }

    /** 操作者別世代フォルダ（出力先に応じた親配下）。 */
    public static Path operatorGenerationDir(
            Map<String, String> ui, String operatorUser, Destination destination) {
        Path parent =
                destination == Destination.LOCAL
                        ? AppPaths.aladdinEntryDispatchPlanLocalDir(ui)
                        : AppPaths.aladdinEntryDispatchPlanDir(ui);
        return parent
                .resolve(SummaryAiDispatchGenerationStore.sanitizeOperatorDirName(operatorUser))
                .toAbsolutePath()
                .normalize();
    }

    /** 現在の操作者の世代フォルダ名（サニタイズ済み。世代ダイアログの既定選択用）。 */
    public static String currentOperatorDirName(Map<String, String> ui) {
        return SummaryAiDispatchGenerationStore.sanitizeOperatorDirName(
                SummaryAiDispatchGenerationStore.resolveOperatorUser(ui));
    }

    /** 世代フォルダ内の xlsx を更新日時の古い順に削除して {@link #MAX_GENERATIONS_PER_USER} 件へ揃える。 */
    static void pruneGenerations(Path operatorDir) throws IOException {
        if (!Files.isDirectory(operatorDir)) {
            return;
        }
        List<Path> files;
        try (var stream = Files.list(operatorDir)) {
            files =
                    stream.filter(Files::isRegularFile)
                            .filter(p -> p.getFileName().toString().endsWith(".xlsx"))
                            .sorted(Comparator.comparingLong(DispatchAladdinEntryWorkbookExporter::lastModifiedMillis))
                            .collect(java.util.stream.Collectors.toCollection(ArrayList::new));
        }
        while (files.size() > MAX_GENERATIONS_PER_USER) {
            Files.deleteIfExists(files.removeFirst());
        }
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return Long.MAX_VALUE;
        }
    }

    private static void writeEmptySheet(XSSFWorkbook wb, Styles styles) {
        Sheet sh = wb.createSheet(EMPTY_SHEET_NAME);
        Row row = sh.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("（データなし: 配台結果を再読み込みしてから出力してください）");
        cell.setCellStyle(styles.data());
        sh.setColumnWidth(0, 256 * 60);
        applyPrintSetup(sh);
    }

    /** ページ設定: 用紙 A4・横向き・横 2 ページ・縦は自動・印刷タイトルは 1 行目と A〜I 列・余白は「狭い」・表示ズーム 75%。 */
    private static void applyPrintSetup(Sheet sh) {
        PrintSetup ps = sh.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);
        ps.setFitWidth((short) 2);
        ps.setFitHeight((short) 0);
        sh.setFitToPage(true);
        sh.setAutobreaks(true);
        sh.setZoom(SHEET_VIEW_ZOOM_PERCENT);
        // Excel の「狭い」余白プリセット相当（インチ）。
        sh.setMargin(Sheet.LeftMargin, 0.25);
        sh.setMargin(Sheet.RightMargin, 0.25);
        sh.setMargin(Sheet.TopMargin, 0.75);
        sh.setMargin(Sheet.BottomMargin, 0.75);
        sh.setMargin(Sheet.HeaderMargin, 0.3);
        sh.setMargin(Sheet.FooterMargin, 0.3);
        // 印刷タイトル: 1 行目（タイトル行）・固定列（タイトル列）。
        sh.setRepeatingRows(new CellRangeAddress(0, 0, -1, -1));
        sh.setRepeatingColumns(new CellRangeAddress(-1, -1, 0, FIXED_COLUMN_COUNT - 1));
    }

    private static PostProcessingPlanMachineLookup.Snapshot loadMachineSnapshot(
            Map<String, String> ui) {
        try {
            return PostProcessingPlanMachineLookup.snapshot(ui);
        } catch (IOException e) {
            return PostProcessingPlanMachineLookup.Snapshot.empty();
        }
    }

    /**
     * シート名: 加工計画DATA の機械コードが取れるときは {@code コード 機械名}（コンボ表示と同形式）、
     * 取れないときは機械名のみ。
     */
    static String sheetNameForMachine(
            DispatchAladdinEntrySheetBuilder.MachineSheet machineSheet,
            PostProcessingPlanMachineLookup.Snapshot machineSnap) {
        String machineName = machineSheet.machineName();
        String code =
                PostProcessingPlanMachineLookup.resolveMachineCodeFromName(machineSnap, machineName);
        String base =
                code.isEmpty()
                        ? machineName
                        : PostProcessingKouteiNaiyoMasterLookup.displayLabel(code, machineName);
        return WorkbookUtil.createSafeSheetName(base);
    }

    private static void writeMachineSheet(
            XSSFWorkbook wb,
            DispatchAladdinEntrySheetBuilder.MachineSheet machineSheet,
            List<LocalDate> dates,
            Styles styles,
            PostProcessingPlanMachineLookup.Snapshot machineSnap) {
        String name = sheetNameForMachine(machineSheet, machineSnap);
        if (wb.getSheet(name) != null) {
            name = WorkbookUtil.createSafeSheetName(name + " (" + (wb.getNumberOfSheets() + 1) + ")");
        }
        Sheet sh = wb.createSheet(name);
        LocalDate today = dates.isEmpty() ? LocalDate.now() : dates.getFirst();

        Row header = sh.createRow(0);
        header.setHeightInPoints(33f);
        for (int c = 0; c < FIXED_COLUMN_COUNT; c++) {
            Cell cell = header.createCell(c);
            cell.setCellValue(FIXED_HEADERS[c]);
            cell.setCellStyle(styles.header());
        }
        for (int i = 0; i < dates.size(); i++) {
            LocalDate d = dates.get(i);
            Cell cell = header.createCell(FIXED_COLUMN_COUNT + i);
            cell.setCellValue(dateHeaderLabel(d));
            cell.setCellStyle(styles.dateHeaderFor(d, d.equals(today)));
        }

        int r = 1;
        for (DispatchAladdinEntrySheetBuilder.EntryRow entry : machineSheet.rows()) {
            Row row = sh.createRow(r++);
            row.setHeightInPoints(33f);
            writeFixedCell(row, 0, entry.taskId(), styles.data());
            writeFixedCell(row, 1, entry.contractNo(), styles.data());
            writeFixedCell(row, 2, entry.processName(), styles.data());
            writeFixedCell(row, 3, entry.inputDate(), styles.data());
            writeFixedCell(row, 4, entry.kaitoNoki(), styles.data());
            String completionCheck = entry.completionDateCheckText();
            writeFixedCell(
                    row,
                    5,
                    completionCheck,
                    completionCheck.isEmpty()
                            ? styles.data()
                            : entry.completionDateCheckOk()
                                    ? styles.checkOk()
                                    : styles.checkNg());
            writeFixedCell(
                    row, 6, ResultDispatchNormalizer.formatQty(entry.conversionQty()), styles.qty());
            writeFixedCell(
                    row, 7, ResultDispatchNormalizer.formatQty(entry.completedQty()), styles.qty());
            writeFixedCell(
                    row, 8, ResultDispatchNormalizer.formatQty(entry.dispatchTotal()), styles.qty());
            writeFixedCell(
                    row,
                    9,
                    entry.quantityCheckText(),
                    entry.quantityOk() ? styles.checkOk() : styles.checkNg());
            for (int i = 0; i < dates.size(); i++) {
                LocalDate d = dates.get(i);
                DispatchAladdinEntrySheetBuilder.EntryCell ec = entry.cells().get(d);
                Cell cell = row.createCell(FIXED_COLUMN_COUNT + i);
                if (ec == null || ec.isEmpty()) {
                    cell.setCellValue("");
                    cell.setCellStyle(styles.dateCellFor(d, false));
                } else {
                    String cellText = ec.cellText();
                    cell.setCellStyle(styles.dateCellFor(d, ec.mismatch()));
                    XSSFRichTextString rich = styles.dateCellRichText(cellText);
                    if (rich != null) {
                        cell.setCellValue(rich);
                    } else {
                        cell.setCellValue(cellText);
                    }
                }
            }
        }

        int lastCol = FIXED_COLUMN_COUNT + dates.size() - 1;
        if (lastCol >= 0) {
            sh.setAutoFilter(new CellRangeAddress(0, 0, 0, Math.max(lastCol, 0)));
        }
        sh.createFreezePane(FIXED_COLUMN_COUNT, 1);

        int[] fixedWidths = {14, 12, 10, 15, 15, 11, 9, 9, 9, 13};
        for (int c = 0; c < FIXED_COLUMN_COUNT; c++) {
            sh.setColumnWidth(c, 256 * fixedWidths[c]);
        }
        for (int i = 0; i < dates.size(); i++) {
            sh.setColumnWidth(
                    FIXED_COLUMN_COUNT + i, (int) Math.round(256 * DATE_COLUMN_WIDTH_CHARS));
        }
        applyPrintSetup(sh);
    }

    private static void writeFixedCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    /** 日付ヘッダ: {@code M/d} + 改行 + {@code (曜)}。 */
    static String dateHeaderLabel(LocalDate d) {
        return d.getMonthValue()
                + "/"
                + d.getDayOfMonth()
                + "\n("
                + WEEKDAY_JA[d.getDayOfWeek().getValue() - 1]
                + ")";
    }

    /**
     * 日付セル 2 段表示: 上段（現アラ計）= {@link #ALADDIN_LINE_FONT_SIZE_PT}、
     * 下段（シス計）= {@link #SYSTEM_LINE_FONT_SIZE_PT}。
     */
    static XSSFRichTextString buildDateCellRichText(
            String text, Font aladdinLineFont, Font systemLineFont) {
        if (text == null
                || text.isBlank()
                || aladdinLineFont == null
                || systemLineFont == null) {
            return null;
        }
        int newline = text.indexOf('\n');
        if (newline < 0) {
            return null;
        }
        XSSFRichTextString rich = new XSSFRichTextString(text);
        rich.applyFont(0, newline, aladdinLineFont);
        rich.applyFont(newline + 1, text.length(), systemLineFont);
        return rich;
    }

    /** シート内スタイル一式。 */
    private record Styles(
            CellStyle header,
            CellStyle dateHeaderWeekday,
            CellStyle dateHeaderSaturday,
            CellStyle dateHeaderSunday,
            CellStyle dateHeaderToday,
            CellStyle data,
            CellStyle qty,
            CellStyle checkOk,
            CellStyle checkNg,
            CellStyle dateCell,
            CellStyle dateCellMismatch,
            CellStyle dateCellWeekend,
            Font aladdinLineFont,
            Font systemLineFont) {

        XSSFRichTextString dateCellRichText(String text) {
            return buildDateCellRichText(text, aladdinLineFont, systemLineFont);
        }

        CellStyle dateHeaderFor(LocalDate d, boolean isToday) {
            if (isToday) {
                return dateHeaderToday;
            }
            if (d.getDayOfWeek() == DayOfWeek.SATURDAY) {
                return dateHeaderSaturday;
            }
            if (d.getDayOfWeek() == DayOfWeek.SUNDAY) {
                return dateHeaderSunday;
            }
            return dateHeaderWeekday;
        }

        CellStyle dateCellFor(LocalDate d, boolean mismatch) {
            if (mismatch) {
                return dateCellMismatch;
            }
            if (d.getDayOfWeek() == DayOfWeek.SATURDAY || d.getDayOfWeek() == DayOfWeek.SUNDAY) {
                return dateCellWeekend;
            }
            return dateCell;
        }

        static Styles of(XSSFWorkbook wb) {
            String fontName = SummaryAiDispatchExportPrefs.DEFAULT_FONT_FAMILY;
            short defaultPt = (short) SummaryAiDispatchExportPrefs.DEFAULT_FONT_SIZE_PT;

            Font headerFont = wb.createFont();
            headerFont.setFontName(fontName);
            headerFont.setFontHeightInPoints(defaultPt);
            headerFont.setBold(true);
            Font dataFont = wb.createFont();
            dataFont.setFontName(fontName);
            dataFont.setFontHeightInPoints(defaultPt);
            Font ngFont = wb.createFont();
            ngFont.setFontName(fontName);
            ngFont.setFontHeightInPoints(defaultPt);
            ngFont.setBold(true);
            ngFont.setColor(org.apache.poi.ss.usermodel.IndexedColors.DARK_RED.getIndex());

            Font aladdinLineFont = wb.createFont();
            aladdinLineFont.setFontName(fontName);
            aladdinLineFont.setFontHeightInPoints(ALADDIN_LINE_FONT_SIZE_PT);
            Font systemLineFont = wb.createFont();
            systemLineFont.setFontName(fontName);
            systemLineFont.setFontHeightInPoints(SYSTEM_LINE_FONT_SIZE_PT);

            CellStyle header = borderedStyle(wb, headerFont, HorizontalAlignment.CENTER, true);
            fill(header, new byte[] {(byte) 0xD9, (byte) 0xE1, (byte) 0xF2}); // 薄青グレー

            CellStyle dateHeaderWeekday = borderedStyle(wb, headerFont, HorizontalAlignment.CENTER, true);
            fill(dateHeaderWeekday, new byte[] {(byte) 0xD9, (byte) 0xE1, (byte) 0xF2});
            CellStyle dateHeaderSaturday = borderedStyle(wb, headerFont, HorizontalAlignment.CENTER, true);
            fill(dateHeaderSaturday, new byte[] {(byte) 0xBD, (byte) 0xD7, (byte) 0xEE}); // 青系
            CellStyle dateHeaderSunday = borderedStyle(wb, headerFont, HorizontalAlignment.CENTER, true);
            fill(dateHeaderSunday, new byte[] {(byte) 0xF8, (byte) 0xCB, (byte) 0xAD}); // 赤系
            CellStyle dateHeaderToday = borderedStyle(wb, headerFont, HorizontalAlignment.CENTER, true);
            fill(dateHeaderToday, new byte[] {(byte) 0xC6, (byte) 0xEF, (byte) 0xCE}); // 緑系（当日）

            CellStyle data = borderedStyle(wb, dataFont, HorizontalAlignment.LEFT, false);
            CellStyle qty = borderedStyle(wb, dataFont, HorizontalAlignment.RIGHT, false);
            CellStyle checkOk = borderedStyle(wb, dataFont, HorizontalAlignment.CENTER, false);
            CellStyle checkNg = borderedStyle(wb, ngFont, HorizontalAlignment.CENTER, false);
            fill(checkNg, new byte[] {(byte) 0xFF, (byte) 0xC7, (byte) 0xCE}); // 赤系

            CellStyle dateCell = borderedStyle(wb, dataFont, HorizontalAlignment.RIGHT, true);
            CellStyle dateCellMismatch = borderedStyle(wb, dataFont, HorizontalAlignment.RIGHT, true);
            fill(dateCellMismatch, new byte[] {(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}); // 薄黄
            CellStyle dateCellWeekend = borderedStyle(wb, dataFont, HorizontalAlignment.RIGHT, true);
            fill(dateCellWeekend, new byte[] {(byte) 0xF2, (byte) 0xF2, (byte) 0xF2}); // 薄灰

            return new Styles(
                    header,
                    dateHeaderWeekday,
                    dateHeaderSaturday,
                    dateHeaderSunday,
                    dateHeaderToday,
                    data,
                    qty,
                    checkOk,
                    checkNg,
                    dateCell,
                    dateCellMismatch,
                    dateCellWeekend,
                    aladdinLineFont,
                    systemLineFont);
        }

        private static CellStyle borderedStyle(
                XSSFWorkbook wb, Font font, HorizontalAlignment align, boolean wrap) {
            CellStyle s = wb.createCellStyle();
            s.setFont(font);
            s.setAlignment(align);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
            s.setWrapText(wrap);
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
            return s;
        }

        private static void fill(CellStyle style, byte[] rgb) {
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(rgb, null));
        }
    }
}
