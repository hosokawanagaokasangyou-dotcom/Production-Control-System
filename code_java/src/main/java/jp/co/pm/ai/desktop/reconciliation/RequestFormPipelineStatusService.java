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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
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

    /**
     * 依頼NO先頭が「2」の自社加工品（配台対象外）。前後空白は除いて判定する。
     */
    public static boolean isInHouseSelfProcessingIraiNo(String iraiNo) {
        if (iraiNo == null) {
            return false;
        }
        String s = iraiNo.strip();
        return !s.isEmpty() && s.charAt(0) == '2';
    }

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
            String rawInputDateDisplay,
            String indexResponseDate,
            String indexInputDate,
            String indexDeliveryDate,
            String indexDeliveryRemarks,
            String indexContractNo,
            String indexContractRemarks,
            RawInputDateCrossSourceCheck.CrossSourceResult rawInputDateCrossCheck) {}

    public record ScanResult(
            List<PipelineStatusRow> rows,
            List<String> warnings,
            boolean aladdinJsonAvailable,
            List<String> planDateHeaders,
            KonanDailyReportLookup dailyReportLookup) {}

    /** 走査の進捗通知（ワーカースレッドから呼ばれる。UI 更新は {@code Platform.runLater} 側で行う）。 */
    @FunctionalInterface
    public interface ScanProgressListener {
        /** @param fraction 0.0–1.0 */
        void onProgress(double fraction, String detail);
    }

    public static ScanResult scan(Map<String, String> ui, JuchuHeaderAliasRegistry registry) {
        return scan(ui, registry, null);
    }

    public static ScanResult scan(
            Map<String, String> ui,
            JuchuHeaderAliasRegistry registry,
            ScanProgressListener progress) {
        return scan(ui, registry, progress, 0);
    }

    /**
     * @param juchuInputHideDays 受注入力日がこの日数以上前の行は走査対象外（0 で無効）。UI の「受注入力日が N 日以上前を非表示」と同じ。
     */
    public static ScanResult scan(
            Map<String, String> ui,
            JuchuHeaderAliasRegistry registry,
            ScanProgressListener progress,
            int juchuInputHideDays) {
        Map<String, String> env = ui != null ? ui : Map.of();
        JuchuHeaderAliasRegistry reg =
                registry != null ? registry : JuchuHeaderAliasRegistry.loadDefault();
        List<String> warnings = new ArrayList<>();
        reportProgress(progress, 0.02, "受注ファイル読込中…");
        String juchuPath = resolveJuchuFilePath(env);
        Map<String, Map<String, String>> dbRows = loadJuchuRows(juchuPath, reg, warnings);

        reportProgress(progress, 0.08, "加工計画データ読込中…");
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
        AladdinShapedPlanQtyLookup.PipelineScanIndex shapedIndex =
                aladdinJsonAvailable
                        ? AladdinShapedPlanQtyLookup.buildPipelineScanIndex(
                                shaped.headers(), shaped.rows())
                        : AladdinShapedPlanQtyLookup.PipelineScanIndex.empty();

        List<Map<String, String>> rawRequests = loadOriginalRequests(env, warnings, progress);
        Map<String, Map<String, String>> excelRawByIraiKey = indexExcelRawByIraiKey(rawRequests);
        Path parseCacheRootPath = AppPaths.resolveRepoRoot(env).resolve("preview_cache");
        File parseCacheRoot = parseCacheRootPath.toFile();
        reportProgress(progress, 0.72, "加工日報読込中…");
        KonanDailyReportLookup dailyReport = KonanDailyReportLookup.load(env, warnings);
        List<PipelineStatusRow> rows = new ArrayList<>();
        Set<String> processedOriginalKeys = new HashSet<>();
        int rowWorkTotal = rawRequests.size() + dbRows.size();
        int rowWorkDone = 0;
        for (Map<String, String> raw : rawRequests) {
            String iraiNo = firstNonBlank(raw.get("依頼Ｎｏ"), raw.get("依頼No"), raw.get("依頼NO"));
            if (iraiNo.isBlank()) {
                continue;
            }
            String normKey = JuchuTransferValueNormalizer.normalizeKey(iraiNo);
            processedOriginalKeys.add(normKey);
            Map<String, String> juchuDb = dbRows.get(normKey);
            if (shouldSkipJuchuRowDuringScan(juchuDb, juchuInputHideDays)) {
                rowWorkDone++;
                reportRowProgress(progress, rowWorkDone, rowWorkTotal);
                continue;
            }
            Map<String, String> originalDb = buildOriginalDbFromRaw(raw);
            rows.add(
                    buildRow(
                            iraiNo,
                            resolveOriginalFileName(raw),
                            true,
                            originalDb,
                            juchuDb,
                            reg,
                            juchuPath,
                            aladdinJsonAvailable,
                            shapedIndex,
                            planDateHeaders,
                            RequestFormOriginalIndexSheetMeta.IndexSheetDisplay.fromRaw(raw),
                            resolveSheetInputDateRaw(raw)));
            rowWorkDone++;
            reportRowProgress(progress, rowWorkDone, rowWorkTotal);
        }
        for (Map.Entry<String, Map<String, String>> entry : dbRows.entrySet()) {
            if (processedOriginalKeys.contains(entry.getKey())) {
                continue;
            }
            Map<String, String> juchuDb = entry.getValue();
            if (shouldSkipJuchuRowDuringScan(juchuDb, juchuInputHideDays)) {
                rowWorkDone++;
                reportRowProgress(progress, rowWorkDone, rowWorkTotal);
                continue;
            }
            String iraiNo =
                    firstNonBlank(
                            juchuDb.get("依頼No"),
                            juchuDb.get("依頼Ｎｏ"),
                            juchuDb.get("依頼NO"),
                            entry.getKey());
            Map<String, String> linkedExcelRaw = excelRawByIraiKey.get(entry.getKey());
            Optional<Map<String, String>> linkedRaw =
                    linkedExcelRaw != null
                            ? Optional.of(linkedExcelRaw)
                            : Optional.empty();
            if (linkedRaw.isEmpty() && AppPaths.isRequestFormTpiPdfEnabled(env)) {
                linkedRaw = resolveLinkedTpiPdfRaw(iraiNo, env, parseCacheRoot, warnings);
            }
            if (linkedRaw.isPresent()) {
                Map<String, String> raw = linkedRaw.get();
                Map<String, String> originalDb = buildOriginalDbFromRaw(raw);
                rows.add(
                        buildRow(
                                iraiNo,
                                resolveOriginalFileName(raw),
                                true,
                                originalDb,
                                juchuDb,
                                reg,
                                juchuPath,
                                aladdinJsonAvailable,
                                shapedIndex,
                                planDateHeaders,
                                RequestFormOriginalIndexSheetMeta.IndexSheetDisplay.fromRaw(raw),
                                resolveSheetInputDateRaw(raw)));
                rowWorkDone++;
                reportRowProgress(progress, rowWorkDone, rowWorkTotal);
                continue;
            }
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
                            shapedIndex,
                            planDateHeaders,
                            RequestFormOriginalIndexSheetMeta.IndexSheetDisplay.empty(),
                            ""));
            rowWorkDone++;
            reportRowProgress(progress, rowWorkDone, rowWorkTotal);
        }
        reportProgress(progress, 0.98, "結果を整理中…");
        rows.sort(
                (a, b) -> {
                    int c = a.iraiNo().compareToIgnoreCase(b.iraiNo());
                    if (c != 0) {
                        return c;
                    }
                    return a.originalFileName().compareToIgnoreCase(b.originalFileName());
                });
        return new ScanResult(
                List.copyOf(rows),
                List.copyOf(warnings),
                aladdinJsonAvailable,
                planDateHeaders,
                dailyReport);
    }

    private static void reportProgress(ScanProgressListener progress, double fraction, String detail) {
        if (progress != null) {
            progress.onProgress(fraction, detail);
        }
    }

    private static void reportRowProgress(ScanProgressListener progress, int done, int total) {
        if (progress == null || total <= 0) {
            return;
        }
        if (done % 5 != 0 && done != total) {
            return;
        }
        double fraction = 0.74 + (0.22 * done / (double) total);
        reportProgress(progress, fraction, "行集計 " + done + "/" + total);
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
            AladdinShapedPlanQtyLookup.PipelineScanIndex shapedIndex,
            List<String> planDateHeaders,
            RequestFormOriginalIndexSheetMeta.IndexSheetDisplay indexSheet,
            String sheetInputDateRaw) {
        JuchuTransferCoverageCheck.CoverageResult coverage =
                JuchuTransferCoverageCheck.compare(originalDb, juchuDb, reg, juchuPath);
        String originalContractNoDisplay =
                JuchuTransferCoverageCheck.formatOriginalContractNoDisplay(
                        originalDb, originalPresent);
        String contractNoDisplay =
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(
                        juchuDb, coverage.juchuRowExists());
        List<PlanEntry> planEntries =
                aladdinJsonAvailable ? shapedIndex.planEntriesFor(iraiNo) : List.of();
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
        if (juchuAdjustDeliveryDate == null
                && !JuchuTransferValueNormalizer.isBlank(juchuAdjustDeliveryDateDisplay)) {
            juchuAdjustDeliveryDate =
                    JuchuTransferValueNormalizer.parseLocalDate(
                            juchuAdjustDeliveryDateDisplay.strip());
        }
        String rawInputDateDisplay =
                formatRawInputDateDisplay(originalDb, originalPresent, juchuDb);
        RequestFormOriginalIndexSheetMeta.IndexSheetDisplay idx =
                indexSheet != null
                        ? indexSheet
                        : RequestFormOriginalIndexSheetMeta.IndexSheetDisplay.empty();
        String aladdinRawInputDate =
                aladdinJsonAvailable ? shapedIndex.rawInputDateDisplayFor(iraiNo) : "";
        String juchuRawInputDate = formatJuchuDateFieldDisplay(juchuDb, Col.TONYU_BI.dbKey());
        RawInputDateCrossSourceCheck.CrossSourceResult rawInputDateCrossCheck =
                RawInputDateCrossSourceCheck.evaluate(
                        aladdinRawInputDate,
                        juchuRawInputDate,
                        idx.inputDate(),
                        sheetInputDateRaw,
                        aladdinJsonAvailable);
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
                rawInputDateDisplay,
                idx.responseDate(),
                idx.inputDate(),
                idx.deliveryDate(),
                idx.deliveryRemarks(),
                idx.contractNo(),
                idx.contractRemarks(),
                rawInputDateCrossCheck);
    }

    /** 目次マージ前の依頼シート投入日。メタ未設定時は原本 rawMap の「投入日」を使う。 */
    private static String resolveSheetInputDateRaw(Map<String, String> raw) {
        if (raw == null) {
            return "";
        }
        String meta = raw.get(RequestFormOriginalIndexSheetMeta.KEY_SHEET_INPUT_DATE);
        if (!JuchuTransferValueNormalizer.isBlank(meta)) {
            return meta.strip();
        }
        String direct = raw.get("投入日");
        return direct != null ? direct.strip() : "";
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

    static boolean shouldSkipJuchuRowDuringScan(Map<String, String> juchuDb, int hideDays) {
        return hideDays > 0 && shouldHideByJuchuInputDate(juchuDb, hideDays);
    }

    static Map<String, Map<String, String>> indexExcelRawByIraiKey(
            List<Map<String, String>> rawRequests) {
        Map<String, Map<String, String>> index = new HashMap<>();
        if (rawRequests == null) {
            return index;
        }
        for (Map<String, String> raw : rawRequests) {
            if (raw == null || isTpiPdfRaw(raw)) {
                continue;
            }
            String key =
                    JuchuTransferValueNormalizer.normalizeKey(
                            firstNonBlank(
                                    raw.get("依頼Ｎｏ"), raw.get("依頼No"), raw.get("依頼NO")));
            if (!key.isEmpty()) {
                index.putIfAbsent(key, raw);
            }
        }
        return index;
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

    /** 調整納期が当日以降（当日含む）。未設定は false。 */
    public static boolean isAdjustDeliveryOnOrAfterToday(LocalDate adjustDeliveryDate) {
        return adjustDeliveryDate != null && !adjustDeliveryDate.isBefore(LocalDate.now());
    }

    /**
     * 調整納期の {@link LocalDate}。スキャン時の解釈値を優先し、なければ表示文字列から再解釈する。
     * フィルタ・段階1確認免除で列表示と判定のずれを防ぐ。
     */
    public static LocalDate resolveAdjustDeliveryLocalDate(PipelineStatusRow row) {
        if (row == null) {
            return null;
        }
        LocalDate parsed = row.juchuAdjustDeliveryDate();
        if (parsed != null) {
            return parsed;
        }
        String display = row.juchuAdjustDeliveryDateDisplay();
        if (JuchuTransferValueNormalizer.isBlank(display)) {
            return null;
        }
        return JuchuTransferValueNormalizer.parseLocalDate(display.strip());
    }

    static List<String> emptyPlanDayValues() {
        List<String> out = new ArrayList<>(PLAN_DAY_COLUMNS);
        for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
            out.add("");
        }
        return List.copyOf(out);
    }

    private static String resolveJuchuFilePath(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        FactorySite site = GlobalInitSettingTarget.loadEffective(env);
        Optional<RequestFormInputSettingsStore.Settings> settings =
                RequestFormInputSettingsStore.load(env);
        if (settings.isPresent()) {
            String saved = settings.get().paths().juchuFilePath();
            if (saved != null
                    && !saved.isBlank()
                    && !AppPaths.factoryPathHintConflictsWithSite(saved, site)) {
                return saved.strip();
            }
        }
        return AppPaths.resolveRequestFormJuchuFile(env).map(Path::toString).orElse("");
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
            Map<String, String> ui, List<String> warnings, ScanProgressListener progress) {
        List<Map<String, String>> rawRequests = new ArrayList<>();
        Path repoRoot = AppPaths.resolveRepoRoot(ui);
        File parseCacheRoot = repoRoot.resolve("preview_cache").toFile();
        if (!parseCacheRoot.exists()) {
            parseCacheRoot.mkdirs();
        }
        RequestFormSourceCache.pruneStaleDiskCaches(parseCacheRoot);

        Path originalDir = AppPaths.resolveRequestFormOriginalDir(ui);
        File[] excelFiles = null;
        if (NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui)) {
            excelFiles = listOriginalWorkbooks(originalDir.toFile());
            if (excelFiles == null || excelFiles.length == 0) {
                warnings.add("Excel 依頼書原本が見つかりません: " + originalDir);
            }
        } else {
            warnings.add("依頼書原本フォルダにアクセスできません: " + originalDir);
        }

        File[] pdfFiles = listTpiPdfFiles(ui, warnings);
        int excelCount = excelFiles != null ? excelFiles.length : 0;
        int pdfCount = pdfFiles != null ? pdfFiles.length : 0;
        int totalSources = excelCount + pdfCount;
        int processedSources = 0;
        Set<String> scannedExcelCacheBaseNamesLower = new HashSet<>();
        FactorySite factorySite = GlobalInitSettingTarget.loadEffective(ui);

        if (excelFiles != null) {
            for (File file : excelFiles) {
                if (file != null && file.getName() != null) {
                    scannedExcelCacheBaseNamesLower.add(
                            file.getName().toLowerCase(java.util.Locale.ROOT));
                }
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
                processedSources++;
                reportOriginalFileProgress(progress, processedSources, totalSources, file.getName());
            }
        }

        Set<String> excelRawKeys = collectIraiNormKeys(rawRequests);
        appendExcelParseCacheFallback(
                rawRequests,
                excelRawKeys,
                parseCacheRoot,
                scannedExcelCacheBaseNamesLower,
                factorySite);
        excelRawKeys = collectIraiNormKeys(rawRequests);
        appendTpiPdfRawRequests(
                ui, rawRequests, excelRawKeys, parseCacheRoot, warnings, pdfFiles, progress, processedSources, totalSources);
        if (totalSources == 0) {
            reportProgress(progress, 0.72, "依頼書原本なし");
        }
        return rawRequests;
    }

    private static File[] listTpiPdfFiles(Map<String, String> ui, List<String> warnings) {
        if (!AppPaths.isRequestFormTpiPdfEnabled(ui)) {
            return null;
        }
        Optional<Path> tpiDirOpt = AppPaths.resolveRequestFormTpiPdfDir(ui);
        if (tpiDirOpt.isEmpty()) {
            return null;
        }
        String tpiPdfFolder = tpiDirOpt.get().toString();
        if (!NetworkSourceDirResolver.isRequestFormTpiPdfDirReachable(ui)) {
            warnings.add("TPI PDF フォルダにアクセスできません: " + tpiPdfFolder);
            return null;
        }
        return tpiDirOpt.get()
                .toFile()
                .listFiles(
                        (dir, name) ->
                                name != null
                                        && name.toLowerCase(java.util.Locale.ROOT).endsWith(".pdf")
                                        && !name.startsWith("~$"));
    }

    private static void reportOriginalFileProgress(
            ScanProgressListener progress, int processed, int total, String fileName) {
        if (progress == null || total <= 0) {
            return;
        }
        double fraction = 0.12 + (0.60 * processed / (double) total);
        String shortName = fileName != null && fileName.length() > 28 ? "…" + fileName.substring(fileName.length() - 27) : fileName;
        reportProgress(progress, fraction, "原本 " + processed + "/" + total + " " + shortName);
    }

    private static Set<String> collectIraiNormKeys(List<Map<String, String>> rawRequests) {
        Set<String> keys = new HashSet<>();
        for (Map<String, String> raw : rawRequests) {
            String key =
                    JuchuTransferValueNormalizer.normalizeKey(
                            firstNonBlank(
                                    raw.get("依頼Ｎｏ"), raw.get("依頼No"), raw.get("依頼NO")));
            if (!key.isEmpty()) {
                keys.add(key);
            }
        }
        return keys;
    }

    private static void appendTpiPdfRawRequests(
            Map<String, String> ui,
            List<Map<String, String>> rawRequests,
            Set<String> excelRawKeys,
            File parseCacheRoot,
            List<String> warnings,
            File[] pdfFiles,
            ScanProgressListener progress,
            int processedSources,
            int totalSources) {
        if (!AppPaths.isRequestFormTpiPdfEnabled(ui)) {
            return;
        }
        if (pdfFiles == null || pdfFiles.length == 0) {
            return;
        }
        Optional<Path> tpiDirOpt = AppPaths.resolveRequestFormTpiPdfDir(ui);
        if (tpiDirOpt.isEmpty()) {
            return;
        }
        File tpiDir = tpiDirOpt.get().toFile();
        int processed = processedSources;
        for (File pdf : pdfFiles) {
            try {
                Optional<List<Map<String, String>>> cached =
                        RequestFormSourceCache.loadParseEntries(parseCacheRoot, pdf);
                List<Map<String, String>> parsed;
                if (cached.isPresent()) {
                    parsed =
                            RequestFormTpiPdfSplitter.ensureSplitPdfs(
                                    pdf, cached.get(), parseCacheRoot, ui);
                    RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
                } else {
                    parsed =
                            RequestFormTpiPdfExtractor.extractEntriesWithSplit(
                                    pdf, ui, parseCacheRoot);
                    RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
                }
                for (Map<String, String> entry : parsed) {
                    String key =
                            JuchuTransferValueNormalizer.normalizeKey(
                                    firstNonBlank(
                                            entry.get("依頼Ｎｏ"),
                                            entry.get("依頼No"),
                                            entry.get("依頼NO")));
                    if (!key.isEmpty() && excelRawKeys.contains(key)) {
                        continue;
                    }
                    if (!RequestFormTpiPdfCatalog.shouldAutoAddScannedEntry(
                            tpiDir, pdf, parsed, entry)) {
                        continue;
                    }
                    Map<String, String> tagged = new HashMap<>(entry);
                    tagged.put("_sourceFileName", resolveOriginalFileName(tagged));
                    rawRequests.add(tagged);
                }
            } catch (Exception ex) {
                warnings.add("TPI PDF 解析エラー " + pdf.getName() + ": " + ex.getMessage());
            }
            processed++;
            reportOriginalFileProgress(progress, processed, totalSources, pdf.getName());
        }
    }

    static void appendExcelParseCacheFallback(
            List<Map<String, String>> rawRequests,
            Set<String> existingKeys,
            File parseCacheRoot,
            Set<String> scannedExcelBaseNamesLower,
            FactorySite factorySite) {
        if (rawRequests == null || existingKeys == null || parseCacheRoot == null) {
            return;
        }
        File parseDir = RequestFormSourceCache.parseDir(parseCacheRoot);
        File[] cacheFiles =
                parseDir.listFiles(
                        (dir, name) ->
                                name != null
                                        && name.toLowerCase(java.util.Locale.ROOT).endsWith(".json"));
        if (cacheFiles == null || cacheFiles.length == 0) {
            return;
        }
        boolean restrictToScannedFiles =
                scannedExcelBaseNamesLower != null && !scannedExcelBaseNamesLower.isEmpty();
        for (File cacheFile : cacheFiles) {
            Optional<List<Map<String, String>>> entries =
                    RequestFormSourceCache.loadExcelParseEntriesFromCacheFile(cacheFile);
            if (entries.isEmpty()) {
                continue;
            }
            String cacheStem = cacheFile.getName();
            if (cacheStem.toLowerCase(java.util.Locale.ROOT).endsWith(".json")) {
                cacheStem = cacheStem.substring(0, cacheStem.length() - 5);
            }
            String cacheSourceName = cacheStem + ".xlsm";
            if (restrictToScannedFiles) {
                if (!scannedExcelBaseNamesLower.contains(
                        cacheSourceName.toLowerCase(java.util.Locale.ROOT))) {
                    continue;
                }
            } else if (factorySite != null
                    && AppPaths.factoryPathHintConflictsWithSite(cacheSourceName, factorySite)) {
                continue;
            }
            for (Map<String, String> entry : entries.get()) {
                if (entry == null || isTpiPdfRaw(entry)) {
                    continue;
                }
                String key =
                        JuchuTransferValueNormalizer.normalizeKey(
                                firstNonBlank(
                                        entry.get("依頼Ｎｏ"),
                                        entry.get("依頼No"),
                                        entry.get("依頼NO")));
                if (key.isEmpty() || existingKeys.contains(key)) {
                    continue;
                }
                Map<String, String> tagged = new HashMap<>(entry);
                String sourceName =
                        firstNonBlank(entry.get("原本ファイル名"), cacheSourceName);
                if (factorySite != null
                        && AppPaths.factoryPathHintConflictsWithSite(sourceName, factorySite)) {
                    continue;
                }
                tagged.put("_sourceFileName", sourceName);
                rawRequests.add(tagged);
                existingKeys.add(key);
            }
        }
    }

    static Optional<Map<String, String>> resolveLinkedExcelOriginalRaw(
            String iraiNo, Map<String, String> ui, File parseCacheRoot, List<String> warnings) {
        if (iraiNo == null || iraiNo.isBlank() || parseCacheRoot == null) {
            return Optional.empty();
        }
        String normIrai = JuchuTransferValueNormalizer.normalizeKey(iraiNo);

        Path originalDir = AppPaths.resolveRequestFormOriginalDir(ui);
        if (NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui)) {
            File[] excelFiles = listOriginalWorkbooks(originalDir.toFile());
            if (excelFiles != null) {
                for (File file : excelFiles) {
                    try {
                        Optional<List<Map<String, String>>> cached =
                                RequestFormSourceCache.loadParseEntries(parseCacheRoot, file);
                        List<Map<String, String>> parsed;
                        if (cached.isPresent()) {
                            parsed = cached.get();
                        } else {
                            parsed = parseOriginalWorkbook(file);
                            RequestFormSourceCache.saveParseEntries(
                                    parseCacheRoot, file, parsed);
                        }
                        Optional<Map<String, String>> hit =
                                findExcelOriginalEntryInList(parsed, normIrai);
                        if (hit.isPresent()) {
                            return Optional.of(tagExcelOriginalRaw(hit.get(), file.getName()));
                        }
                    } catch (Exception ex) {
                        if (warnings != null) {
                            warnings.add(
                                    "Excel 原本照合エラー " + file.getName() + ": " + ex.getMessage());
                        }
                    }
                }
            }
        }

        File parseDir = RequestFormSourceCache.parseDir(parseCacheRoot);
        File[] cacheFiles =
                parseDir.listFiles(
                        (dir, name) ->
                                name != null
                                        && name.toLowerCase(java.util.Locale.ROOT).endsWith(".json"));
        if (cacheFiles != null) {
            for (File cacheFile : cacheFiles) {
                Optional<List<Map<String, String>>> entries =
                        RequestFormSourceCache.loadExcelParseEntriesFromCacheFile(cacheFile);
                if (entries.isEmpty()) {
                    continue;
                }
                Optional<Map<String, String>> hit =
                        findExcelOriginalEntryInList(entries.get(), normIrai);
                if (hit.isPresent()) {
                    String cacheStem = cacheFile.getName();
                    if (cacheStem.toLowerCase(java.util.Locale.ROOT).endsWith(".json")) {
                        cacheStem = cacheStem.substring(0, cacheStem.length() - 5);
                    }
                    String sourceName =
                            firstNonBlank(hit.get().get("原本ファイル名"), cacheStem + ".xlsm");
                    return Optional.of(tagExcelOriginalRaw(hit.get(), sourceName));
                }
            }
        }
        return Optional.empty();
    }

    private static Optional<Map<String, String>> findExcelOriginalEntryInList(
            List<Map<String, String>> entries, String normIrai) {
        if (entries == null || entries.isEmpty() || normIrai == null || normIrai.isBlank()) {
            return Optional.empty();
        }
        for (Map<String, String> entry : entries) {
            if (entry == null || isTpiPdfRaw(entry)) {
                continue;
            }
            String parsedIrai =
                    JuchuTransferValueNormalizer.normalizeKey(
                            firstNonBlank(
                                    entry.get("依頼Ｎｏ"),
                                    entry.get("依頼No"),
                                    entry.get("依頼NO")));
            if (normIrai.equals(parsedIrai)) {
                return Optional.of(entry);
            }
        }
        return Optional.empty();
    }

    private static Map<String, String> tagExcelOriginalRaw(
            Map<String, String> entry, String sourceFileName) {
        Map<String, String> tagged = new HashMap<>(entry);
        tagged.put(
                "_sourceFileName",
                firstNonBlank(sourceFileName, resolveOriginalFileName(tagged)));
        return tagged;
    }

    static Optional<Map<String, String>> resolveLinkedTpiPdfRaw(
            String iraiNo, Map<String, String> ui, File parseCacheRoot, List<String> warnings) {
        if (!AppPaths.isRequestFormTpiPdfEnabled(ui)) {
            return Optional.empty();
        }
        Optional<Path> tpiDirOpt = AppPaths.resolveRequestFormTpiPdfDir(ui);
        if (tpiDirOpt.isEmpty() || iraiNo == null || iraiNo.isBlank()) {
            return Optional.empty();
        }
        String tpiPdfFolder = tpiDirOpt.get().toString();
        if (!NetworkSourceDirResolver.isRequestFormTpiPdfDirReachable(ui)) {
            return Optional.empty();
        }
        Optional<File> linked =
                RequestFormTpiPdfCatalog.findForIraiNo(iraiNo, tpiPdfFolder);
        if (linked.isEmpty()) {
            linked =
                    RequestFormTpiPdfCatalog.findForIraiNoByPdfContent(
                            iraiNo, tpiPdfFolder, ui, parseCacheRoot);
        }
        if (linked.isEmpty()) {
            return Optional.empty();
        }
        Map<String, String> raw = loadTpiPdfRawLinked(linked.get(), iraiNo, ui, parseCacheRoot);
        if (raw.isEmpty()) {
            return Optional.empty();
        }
        Map<String, String> tagged = new HashMap<>(raw);
        tagged.put("_sourceFileName", resolveOriginalFileName(tagged));
        return Optional.of(tagged);
    }

    private static Map<String, String> loadTpiPdfRawLinked(
            File pdf, String iraiNo, Map<String, String> ui, File parseCacheRoot) {
        if (pdf == null || !pdf.isFile()) {
            return Map.of();
        }
        try {
            Optional<List<Map<String, String>>> cached =
                    RequestFormSourceCache.loadParseEntries(parseCacheRoot, pdf);
            List<Map<String, String>> parsed;
            if (cached.isPresent() && !cached.get().isEmpty()) {
                parsed =
                        RequestFormTpiPdfSplitter.ensureSplitPdfs(
                                pdf, cached.get(), parseCacheRoot, ui);
                RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
            } else {
                parsed =
                        RequestFormTpiPdfExtractor.extractEntriesWithSplit(pdf, ui, parseCacheRoot);
                RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
            }
            if (parsed.isEmpty()) {
                return Map.of();
            }
            Map<String, String> selected = selectLinkedTpiEntry(parsed, iraiNo, ui);
            return reconcileLinkedTpiRaw(selected, iraiNo);
        } catch (Exception ex) {
            return Map.of();
        }
    }

    private static Map<String, String> selectLinkedTpiEntry(
            List<Map<String, String>> entries, String iraiNo, Map<String, String> ui) {
        if (entries == null || entries.isEmpty()) {
            return Map.of();
        }
        String normIrai = JuchuTransferValueNormalizer.normalizeKey(iraiNo);
        for (Map<String, String> entry : entries) {
            String parsedIrai = entry.get("依頼Ｎｏ");
            if (parsedIrai == null
                    || parsedIrai.isBlank()
                    || !JuchuTransferValueNormalizer.normalizeKey(parsedIrai).equals(normIrai)) {
                continue;
            }
            Optional<Path> tpiDirOpt = AppPaths.resolveRequestFormTpiPdfDir(ui);
            File tpiDir = tpiDirOpt.map(Path::toFile).orElse(null);
            String pdfName = entry.get("原本ファイル名");
            File pdf =
                    pdfName != null && tpiDir != null ? new File(tpiDir, pdfName) : null;
            if (pdf != null
                    && !RequestFormTpiPdfCatalog.canLinkIraiInSharedPdf(
                            iraiNo, pdf, entries, tpiDir)) {
                return Map.of();
            }
            return entry;
        }
        return Map.of();
    }

    private static Map<String, String> reconcileLinkedTpiRaw(
            Map<String, String> parsed, String iraiNo) {
        if (parsed == null || parsed.isEmpty()) {
            return Map.of();
        }
        Map<String, String> raw = new java.util.LinkedHashMap<>(parsed);
        String parsedIrai = raw.get("依頼Ｎｏ");
        if (parsedIrai == null
                || parsedIrai.isBlank()
                || !JuchuTransferValueNormalizer.normalizeKey(parsedIrai)
                        .equals(JuchuTransferValueNormalizer.normalizeKey(iraiNo))) {
            return Map.of();
        }
        return raw;
    }

    static Map<String, String> buildOriginalDbFromRaw(Map<String, String> raw) {
        if (isTpiPdfRaw(raw)) {
            return RequestFormOriginalExtractor.buildTpiDbDefaultsFromRaw(raw);
        }
        return RequestFormOriginalExtractor.buildDbDefaultsFromRaw(raw);
    }

    static String resolveOriginalFileName(Map<String, String> raw) {
        if (raw == null || raw.isEmpty()) {
            return "";
        }
        String tagged = raw.get("_sourceFileName");
        if (tagged != null && !tagged.isBlank()) {
            return tagged;
        }
        String splitPath = raw.get(RequestFormTpiPdfFieldLayout.META_SPLIT_PDF_PATH);
        if (splitPath != null && !splitPath.isBlank()) {
            return new File(splitPath).getName();
        }
        String originalName = raw.get("原本ファイル名");
        return originalName != null ? originalName : "";
    }

    private static boolean isTpiPdfRaw(Map<String, String> raw) {
        return raw != null
                && RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF.equals(
                        raw.get(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND));
    }

    private static List<Map<String, String>> parseOriginalWorkbook(File file) throws Exception {
        return RequestFormOriginalWorkbookParser.parse(file);
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
