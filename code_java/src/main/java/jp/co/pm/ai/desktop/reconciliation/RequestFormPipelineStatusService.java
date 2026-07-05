package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.reconciliation.JuchuSheetColumnLayout.Col;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/**
 * 依頼書原本フォルダを走査し、受注ファイル転記状況とアラジン加工計画の有無を集約する。
 */
public final class RequestFormPipelineStatusService {

    private static final int JUCHU_SHEET_MAX_SCAN_ROWS = 20_000;
    private static final String JUCHU_SHEET_NAME = "受注ﾌｧｲﾙ";
    private static final int PLAN_DAY_COLUMNS = AladdinShapedPlanQtyLookup.PIPELINE_CHECK_PLAN_DAY_COLUMNS;
    /** UI の受注入力日フィルタ既定値（日）。 */
    public static final int DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS = 30;

    private static final Pattern ORIGINAL_SHEET_NAME =
            Pattern.compile("^[A-Z]+\\d+-\\d+$|^[A-Z]\\d+-\\d+-\\d+$");

    private RequestFormPipelineStatusService() {}

    public record PipelineStatusRow(
            String iraiNo,
            String originalFileName,
            boolean originalPresent,
            String user,
            boolean juchuRegistered,
            String rateDisplay,
            double ratePercent,
            int mismatchCount,
            String originalContractNoDisplay,
            String contractNoStatus,
            boolean aladdinPresent,
            List<String> planDayValues,
            JuchuTransferCoverageCheck.CoverageResult coverage,
            List<PlanEntry> planEntries,
            LocalDate juchuInputDate,
            String juchuInputDateDisplay,
            String juchuInputOperatorDisplay,
            LocalDate juchuAdjustDeliveryDate,
            String juchuAdjustDeliveryDateDisplay,
            String rawInputDateDisplay) {}

    public record ScanResult(
            List<PipelineStatusRow> rows,
            List<String> warnings,
            boolean aladdinJsonAvailable,
            List<String> planDateHeaders) {}

    public static ScanResult scan(Map<String, String> ui, JuchuHeaderAliasRegistry registry) {
        Map<String, String> env = ui != null ? ui : Map.of();
        JuchuHeaderAliasRegistry reg =
                registry != null ? registry : JuchuHeaderAliasRegistry.loadDefault();
        List<String> warnings = new ArrayList<>();
        String juchuPath = resolveJuchuFilePath(env);
        Map<String, Map<String, String>> dbRows = loadJuchuRows(juchuPath, reg, warnings);

        Path shapedPath = AppPaths.resolveShapedAladdinPlanJsonPath(env);
        boolean aladdinJsonAvailable = Files.isRegularFile(shapedPath);
        AladdinShapedPlanQtyLookup.ShapedTable shaped =
                AladdinShapedPlanQtyLookup.loadShapedTable(shapedPath);
        List<String> planDateHeaders =
                aladdinJsonAvailable
                        ? AladdinShapedPlanQtyLookup.extractSortedDateColumnHeaders(
                                shaped.headers(), PLAN_DAY_COLUMNS)
                        : List.of();
        if (!aladdinJsonAvailable) {
            warnings.add(
                    "shaped_aladdin_plan.json がありません。納期管理ビュー →"
                            + " アラジン加工計画取得データ で読込してください。");
        } else if (planDateHeaders.isEmpty() && !shaped.headers().isEmpty()) {
            warnings.add(
                    "shaped_aladdin_plan.json に日付列が見つかりません。"
                            + " アラジン加工計画取得データを再読込してください。");
        }

        List<Map<String, String>> rawRequests = loadOriginalRequests(env, warnings);
        List<PipelineStatusRow> rows = new ArrayList<>();
        Set<String> processedOriginalKeys = new HashSet<>();
        for (Map<String, String> raw : rawRequests) {
            String iraiNo = firstNonBlank(raw.get("依頼Ｎｏ"), raw.get("依頼No"), raw.get("依頼NO"));
            if (iraiNo.isBlank()) {
                continue;
            }
            String normKey = JuchuTransferValueNormalizer.normalizeKey(iraiNo);
            processedOriginalKeys.add(normKey);
            Map<String, String> originalDb =
                    RequestFormOriginalExtractor.buildDbDefaultsFromRaw(raw);
            Map<String, String> juchuDb = dbRows.get(normKey);
            rows.add(
                    buildRow(
                            iraiNo,
                            nullToEmpty(raw.get("_sourceFileName")),
                            true,
                            originalDb,
                            juchuDb,
                            reg,
                            juchuPath,
                            aladdinJsonAvailable,
                            shaped,
                            planDateHeaders));
        }
        for (Map.Entry<String, Map<String, String>> entry : dbRows.entrySet()) {
            if (processedOriginalKeys.contains(entry.getKey())) {
                continue;
            }
            Map<String, String> juchuDb = entry.getValue();
            String iraiNo =
                    firstNonBlank(
                            juchuDb.get("依頼No"),
                            juchuDb.get("依頼Ｎｏ"),
                            juchuDb.get("依頼NO"),
                            entry.getKey());
            rows.add(
                    buildRow(
                            iraiNo,
                            "",
                            false,
                            Map.of(),
                            juchuDb,
                            reg,
                            juchuPath,
                            aladdinJsonAvailable,
                            shaped,
                            planDateHeaders));
        }
        rows.sort(
                (a, b) -> {
                    int c = a.iraiNo().compareToIgnoreCase(b.iraiNo());
                    if (c != 0) {
                        return c;
                    }
                    return a.originalFileName().compareToIgnoreCase(b.originalFileName());
                });
        return new ScanResult(
                List.copyOf(rows), List.copyOf(warnings), aladdinJsonAvailable, planDateHeaders);
    }

    private static PipelineStatusRow buildRow(
            String iraiNo,
            String originalFileName,
            boolean originalPresent,
            Map<String, String> originalDb,
            Map<String, String> juchuDb,
            JuchuHeaderAliasRegistry reg,
            String juchuPath,
            boolean aladdinJsonAvailable,
            AladdinShapedPlanQtyLookup.ShapedTable shaped,
            List<String> planDateHeaders) {
        JuchuTransferCoverageCheck.CoverageResult coverage =
                JuchuTransferCoverageCheck.compare(originalDb, juchuDb, reg, juchuPath);
        String originalContractNoDisplay =
                JuchuTransferCoverageCheck.formatOriginalContractNoDisplay(
                        originalDb, originalPresent);
        String contractNoDisplay =
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(
                        juchuDb, coverage.juchuRowExists());
        List<PlanEntry> planEntries =
                aladdinJsonAvailable
                        ? AladdinShapedPlanQtyLookup.collectEntriesForTaskIdFromTable(
                                shaped.headers(), shaped.rows(), iraiNo)
                        : List.of();
        List<String> planDayValues =
                aladdinJsonAvailable
                        ? AladdinShapedPlanQtyLookup.aggregatePlanMetersByEntryDates(
                                planEntries, PLAN_DAY_COLUMNS)
                        : emptyPlanDayValues();
        String user =
                originalPresent
                        ? firstNonBlank(originalDb.get("ユーザー"))
                        : firstNonBlank(juchuDb != null ? juchuDb.get("ユーザー") : null);
        LocalDate juchuInputDate = parseJuchuInputDate(juchuDb);
        String juchuInputDateDisplay = formatJuchuDateFieldDisplay(juchuDb, Col.NYURYOKU_BI.dbKey());
        String juchuInputOperatorDisplay = formatJuchuInputOperatorDisplay(juchuDb);
        LocalDate juchuAdjustDeliveryDate =
                parseJuchuDateField(juchuDb, Col.CHOSEI_NOKI.dbKey());
        String juchuAdjustDeliveryDateDisplay =
                formatJuchuDateFieldDisplay(juchuDb, Col.CHOSEI_NOKI.dbKey());
        String rawInputDateDisplay =
                formatRawInputDateDisplay(originalDb, originalPresent, juchuDb);
        return new PipelineStatusRow(
                iraiNo,
                originalFileName,
                originalPresent,
                user,
                coverage.juchuRowExists(),
                coverage.rateDisplay(),
                coverage.ratePercent(),
                coverage.mismatchCount(),
                originalContractNoDisplay,
                contractNoDisplay,
                !planEntries.isEmpty(),
                planDayValues,
                coverage,
                List.copyOf(planEntries),
                juchuInputDate,
                juchuInputDateDisplay,
                juchuInputOperatorDisplay,
                juchuAdjustDeliveryDate,
                juchuAdjustDeliveryDateDisplay,
                rawInputDateDisplay);
    }

    /** 原反（材料）投入日の表示。受注「投入日」優先、なければ原本「投入日」。 */
    public static String formatRawInputDateDisplay(
            Map<String, String> originalDb,
            boolean originalPresent,
            Map<String, String> juchuDb) {
        String juchu = formatJuchuDateFieldDisplay(juchuDb, Col.TONYU_BI.dbKey());
        if (!juchu.isBlank()) {
            return juchu;
        }
        if (!originalPresent || originalDb == null) {
            return "";
        }
        String raw = originalDb.get("投入日");
        if (JuchuTransferValueNormalizer.isBlank(raw)) {
            return "";
        }
        String firstLine = raw.strip();
        int nl = firstLine.indexOf('\n');
        if (nl >= 0) {
            firstLine = firstLine.substring(0, nl).strip();
        }
        LocalDate parsed = JuchuTransferValueNormalizer.parseLocalDate(firstLine);
        if (parsed != null) {
            return parsed.getYear()
                    + "/"
                    + parsed.getMonthValue()
                    + "/"
                    + parsed.getDayOfMonth();
        }
        return firstLine;
    }

    /** 受注「入力担当／入力者」の表示文字列。未登録・空は空文字。 */
    public static String formatJuchuInputOperatorDisplay(Map<String, String> juchuDb) {
        if (juchuDb == null || juchuDb.isEmpty()) {
            return "";
        }
        return firstNonBlank(
                juchuDb.get(Col.NYURYOKU_TANTO.dbKey()),
                juchuDb.get("入力者"),
                juchuDb.get("入力担当"));
    }

    /** 受注ファイルの日付項目表示（{@code yyyy/M/d}）。未登録・空は空文字。 */
    public static String formatJuchuDateFieldDisplay(Map<String, String> juchuDb, String fieldKey) {
        if (juchuDb == null || juchuDb.isEmpty() || fieldKey == null || fieldKey.isBlank()) {
            return "";
        }
        String raw = juchuDb.get(fieldKey);
        if (JuchuTransferValueNormalizer.isBlank(raw)) {
            return "";
        }
        LocalDate parsed = parseJuchuDateField(juchuDb, fieldKey);
        if (parsed != null) {
            return parsed.getYear()
                    + "/"
                    + parsed.getMonthValue()
                    + "/"
                    + parsed.getDayOfMonth();
        }
        return raw.strip();
    }

    /** 受注「入力日」の表示文字列。未登録・空は空文字。 */
    public static String formatJuchuInputDateDisplay(Map<String, String> juchuDb) {
        return formatJuchuDateFieldDisplay(juchuDb, Col.NYURYOKU_BI.dbKey());
    }

    /** 受注ファイルの日付項目を解釈する。未登録・空・解釈不能は {@code null}。 */
    public static LocalDate parseJuchuDateField(Map<String, String> juchuDb, String fieldKey) {
        if (juchuDb == null || juchuDb.isEmpty() || fieldKey == null || fieldKey.isBlank()) {
            return null;
        }
        String raw = juchuDb.get(fieldKey);
        if (JuchuTransferValueNormalizer.isBlank(raw)) {
            return null;
        }
        return JuchuTransferValueNormalizer.parseLocalDate(raw);
    }

    /** 受注ﾌｧｲﾙ「入力日」を解釈する。未登録・空・解釈不能は {@code null}。 */
    public static LocalDate parseJuchuInputDate(Map<String, String> juchuDb) {
        return parseJuchuDateField(juchuDb, Col.NYURYOKU_BI.dbKey());
    }

    /**
     * 受注ﾌｧｲﾙの入力日が {@code excludeDays} 日以上前なら非表示対象。
     * 入力日 {@code null} または {@code excludeDays <= 0} は非表示にしない。
     */
    public static boolean shouldHideByJuchuInputDate(LocalDate inputDate, int excludeDays) {
        if (excludeDays <= 0 || inputDate == null) {
            return false;
        }
        LocalDate cutoff = LocalDate.now().minusDays(excludeDays);
        return !inputDate.isAfter(cutoff);
    }

    /** {@link #parseJuchuInputDate} と {@link #shouldHideByJuchuInputDate} の合成。 */
    static boolean shouldHideByJuchuInputDate(Map<String, String> juchuDb, int excludeDays) {
        return shouldHideByJuchuInputDate(parseJuchuInputDate(juchuDb), excludeDays);
    }

    /**
     * 調整納期が当日より前、または未設定なら非表示対象（「当日以降のみ表示」ON 時）。
     * 当日は表示対象。
     */
    public static boolean shouldHideByAdjustDeliveryBeforeToday(LocalDate adjustDeliveryDate) {
        if (adjustDeliveryDate == null) {
            return true;
        }
        return adjustDeliveryDate.isBefore(LocalDate.now());
    }

    static List<String> emptyPlanDayValues() {
        List<String> out = new ArrayList<>(PLAN_DAY_COLUMNS);
        for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
            out.add("");
        }
        return List.copyOf(out);
    }

    private static String resolveJuchuFilePath(Map<String, String> ui) {
        Optional<RequestFormInputSettingsStore.Settings> settings =
                RequestFormInputSettingsStore.load(ui);
        if (settings.isPresent()) {
            String saved = settings.get().paths().juchuFilePath();
            if (saved != null && !saved.isBlank()) {
                return saved.strip();
            }
        }
        return AppPaths.resolveRequestFormJuchuFile(ui).map(Path::toString).orElse("");
    }

    private static Map<String, Map<String, String>> loadJuchuRows(
            String juchuPath, JuchuHeaderAliasRegistry registry, List<String> warnings) {
        Map<String, Map<String, String>> dbRows = new HashMap<>();
        if (juchuPath == null || juchuPath.isBlank()) {
            warnings.add("受注ファイルパスが未設定です。");
            return dbRows;
        }
        File juchuFile = new File(juchuPath);
        if (!juchuFile.isFile()) {
            warnings.add("受注ファイルが見つかりません: " + juchuPath);
            return dbRows;
        }
        try (FileInputStream fis = new FileInputStream(juchuFile);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet(JUCHU_SHEET_NAME);
            if (sheet == null) {
                warnings.add("受注ﾌｧｲﾙ シートが見つかりません: " + juchuPath);
                return dbRows;
            }
            int firstDataRow = registry.headerRowIndexFor(juchuPath) + 1;
            int reqNoColIdx =
                    JuchuSheetColumnLayout.resolveTransferColumnIndex(
                            JuchuSheetColumnLayout.Col.IRAI_NO, registry, juchuPath);
            int lastDataRow = findLastPopulatedDataRow(sheet, firstDataRow, reqNoColIdx);
            for (int r = firstDataRow; r <= lastDataRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                String reqNo = readCellAsString(row.getCell(reqNoColIdx)).strip();
                if (reqNo.isEmpty()) {
                    continue;
                }
                Map<String, String> vals =
                        JuchuSheetColumnLayout.readDbValuesFromRow(row, registry, juchuPath);
                String normKey = JuchuTransferValueNormalizer.normalizeKey(reqNo);
                Map<String, String> existing = dbRows.get(normKey);
                if (existing != null) {
                    JuchuTransferCoverageCheck.mergeContractNoValues(existing, vals);
                } else {
                    dbRows.put(normKey, vals);
                }
            }
        } catch (Exception ex) {
            warnings.add("受注ファイル読込エラー: " + ex.getMessage());
        }
        return dbRows;
    }

    private static List<Map<String, String>> loadOriginalRequests(
            Map<String, String> ui, List<String> warnings) {
        List<Map<String, String>> rawRequests = new ArrayList<>();
        Path originalDir = AppPaths.resolveRequestFormOriginalDir(ui);
        if (!NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui)) {
            warnings.add("依頼書原本フォルダにアクセスできません: " + originalDir);
            return rawRequests;
        }
        File folder = originalDir.toFile();
        File[] files = listOriginalWorkbooks(folder);
        if (files == null || files.length == 0) {
            warnings.add("依頼書原本が見つかりません: " + originalDir);
            return rawRequests;
        }
        Path repoRoot = AppPaths.resolveRepoRoot(ui);
        File parseCacheRoot = repoRoot.resolve("preview_cache").toFile();
        if (!parseCacheRoot.exists()) {
            parseCacheRoot.mkdirs();
        }
        for (File file : files) {
            try {
                Optional<List<Map<String, String>>> cached =
                        RequestFormSourceCache.loadParseEntries(parseCacheRoot, file);
                List<Map<String, String>> parsed;
                if (cached.isPresent()) {
                    parsed = cached.get();
                } else {
                    parsed = parseOriginalWorkbook(file);
                    RequestFormSourceCache.saveParseEntries(parseCacheRoot, file, parsed);
                }
                for (Map<String, String> entry : parsed) {
                    Map<String, String> tagged = new HashMap<>(entry);
                    tagged.put("_sourceFileName", file.getName());
                    rawRequests.add(tagged);
                }
            } catch (Exception ex) {
                warnings.add("原本解析エラー " + file.getName() + ": " + ex.getMessage());
            }
        }
        return rawRequests;
    }

    private static List<Map<String, String>> parseOriginalWorkbook(File file) throws Exception {
        List<Map<String, String>> parsed = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                String sName = wb.getSheetName(s);
                if (ORIGINAL_SHEET_NAME.matcher(sName).matches()) {
                    parsed.add(
                            RequestFormOriginalExtractor.buildRawMapFromSheet(
                                    file, sName, wb.getSheetAt(s)));
                }
            }
        }
        return parsed;
    }

    private static File[] listOriginalWorkbooks(File folder) {
        return folder.listFiles(
                (dir, name) ->
                        name.endsWith(".xlsm")
                                && !name.startsWith("~$")
                                && !name.equals("加工依頼書入力.xlsm"));
    }

    private static int findLastPopulatedDataRow(Sheet sheet, int firstDataRow, int reqNoColIdx) {
        int poiLast = sheet.getLastRowNum();
        if (poiLast < firstDataRow) {
            return firstDataRow - 1;
        }
        int scanLimit = Math.min(poiLast, firstDataRow + JUCHU_SHEET_MAX_SCAN_ROWS);
        for (int r = scanLimit; r >= firstDataRow; r--) {
            Row row = sheet.getRow(r);
            if (row != null && rowHasReqNo(row, reqNoColIdx)) {
                return r;
            }
        }
        return firstDataRow - 1;
    }

    private static boolean rowHasReqNo(Row row, int reqNoColIdx) {
        Cell cell = row.getCell(reqNoColIdx);
        if (cell == null) {
            return false;
        }
        return switch (cell.getCellType()) {
            case STRING -> !cell.getStringCellValue().strip().isEmpty();
            case NUMERIC -> true;
            case FORMULA -> {
                try {
                    yield !cell.getStringCellValue().strip().isEmpty();
                } catch (Exception ex) {
                    try {
                        yield cell.getNumericCellValue() != 0.0d;
                    } catch (Exception ignored) {
                        yield false;
                    }
                }
            }
            default -> false;
        };
    }

    private static String readCellAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield cell.getStringCellValue();
                } catch (Exception ex) {
                    yield String.valueOf(cell.getNumericCellValue());
                }
            }
            default -> "";
        };
    }

    private static String firstNonBlank(String... values) {
        if (values == null) {
            return "";
        }
        for (String v : values) {
            if (v != null && !v.isBlank()) {
                return v.strip();
            }
        }
        return "";
    }

    private static String nullToEmpty(String val) {
        return val != null ? val : "";
    }
}
