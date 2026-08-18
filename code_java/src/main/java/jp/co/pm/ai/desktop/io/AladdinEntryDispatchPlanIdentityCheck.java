package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.util.WorkbookUtil;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingKouteiNaiyoMasterLookup;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingPlanMachineLookup;

/**
 * 指定した（または操作者世代の最新）アラジン入力用配台計画（シス計）と、
 * ソース最新のアラジン加工計画が同一かを判定する。
 *
 * <p>比較対象はアラジン入力 Excel にシス計があるセルのみ。加工計画にあって Excel に無い依頼
 * （完了済みで配台表から落ちた行など）は NG にしない。完了済み行（未加工=0 かつ全数未加工でない）
 * は加工計画ルックアップから除外する。
 *
 * <p>Excel の再出力と shaped JSON の上書きは行わない。
 */
public final class AladdinEntryDispatchPlanIdentityCheck {

    public static final String BADGE_IDENTICAL = "配台計画と加工計画は同一";

    public static final String ERROR_NO_GENERATION = "ログイン中の操作者が生成した配台計画がありません";

    public static final String ERROR_NO_EXCEL = "比較する配台計画 Excel がありません";

    public static final String ERROR_NO_PLAN_SOURCE = "アラジン加工計画の最新ソースがありません";

    private static final int DIALOG_DIFF_LIMIT = 50;

    private static final Pattern GEN_TS =
            Pattern.compile("アラジン入力用_配台計画_(\\d{8})-\\d{6}\\.xlsx", Pattern.CASE_INSENSITIVE);

    private static final DateTimeFormatter GEN_DAY = DateTimeFormatter.ofPattern("yyyyMMdd");

    private AladdinEntryDispatchPlanIdentityCheck() {}

    /** Excel 1 セル分のシス計。 */
    public record SystemQty(
            String machineName, String taskId, String processName, LocalDate date, double qty) {}

    /** 1 キー分の差異。 */
    public record Diff(
            String machineName,
            String taskId,
            String processName,
            LocalDate date,
            double systemQty,
            double planQty) {}

    /** 判定結果。 */
    public record Result(
            boolean identical,
            boolean error,
            String badgeText,
            String message,
            List<Diff> diffs,
            Optional<Path> excelPath,
            Optional<Path> planSourcePath) {

        public String dialogBody() {
            if (error || identical || diffs == null || diffs.isEmpty()) {
                return message != null ? message : "";
            }
            StringBuilder sb = new StringBuilder();
            if (message != null && !message.isBlank()) {
                sb.append(message).append("\n\n");
            }
            int shown = 0;
            for (Diff d : diffs) {
                if (shown >= DIALOG_DIFF_LIMIT) {
                    sb.append("他 ").append(diffs.size() - DIALOG_DIFF_LIMIT).append("件");
                    break;
                }
                sb.append("機械=")
                        .append(d.machineName())
                        .append("  依頼NO=")
                        .append(d.taskId())
                        .append("  工程=")
                        .append(d.processName())
                        .append("  日付=")
                        .append(d.date())
                        .append("  シス計=")
                        .append(formatQty(d.systemQty()))
                        .append("  加工計画=")
                        .append(formatQty(d.planQty()))
                        .append('\n');
                shown++;
            }
            return sb.toString().stripTrailing();
        }
    }

    /**
     * 操作者世代の最新 Excel シス計と、ソース最新の加工計画を突合する。
     * 失敗時は {@link Result#error} が true（例外は投げない）。
     */
    public static Result evaluate(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Optional<Path> excel = newestOperatorGenerationXlsx(u);
        if (excel.isEmpty()) {
            return errorResult(ERROR_NO_GENERATION, Optional.empty(), Optional.empty());
        }
        return evaluate(u, excel.get());
    }

    /**
     * 指定した配台計画 Excel のシス計と、ソース最新の加工計画を突合する。
     * Excel の再出力と shaped JSON の上書きは行わない。
     */
    public static Result evaluate(Map<String, String> ui, Path excelPath) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (excelPath == null || !Files.isRegularFile(excelPath)) {
            return errorResult(
                    ERROR_NO_EXCEL,
                    excelPath != null ? Optional.of(excelPath) : Optional.empty(),
                    Optional.empty());
        }
        Optional<Path> excel = Optional.of(excelPath.toAbsolutePath().normalize());
        Optional<Path> planSource;
        try {
            planSource = newestPlanSourceFile(u);
        } catch (IOException ex) {
            return errorResult(
                    ex.getMessage() != null ? ex.getMessage() : ERROR_NO_PLAN_SOURCE,
                    excel,
                    Optional.empty());
        }
        if (planSource.isEmpty()) {
            return errorResult(ERROR_NO_PLAN_SOURCE, excel, Optional.empty());
        }
        Path planFile = planSource.get();
        String low = planFile.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            return errorResult("Parquet は未対応です: " + planFile, excel, planSource);
        }
        try {
            PlanInputTabularIo.TabularSheet tab =
                    AladdinProcessingPlanSourceReloader.readNewestAladdinTabularFromDisk(planFile);
            PlanInputTabularIo.TabularSheet activePlan = excludeCompletedPlanRows(tab);
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                    AladdinShapedPlanQtyLookup.buildLookup(activePlan.headers(), activePlan.rows());
            LocalDate ref = referenceDateForExcel(excel.get());
            List<SystemQty> system =
                    remapSheetMachines(
                            AladdinEntryDispatchPlanWorkbookReader.readSystemQtys(excel.get(), ref),
                            loadMachineSnapshot(u));
            Result compared = compare(system, lookup);
            String resultKey = compared.identical() ? "ok" : "mismatch";
            int diffCount = compared.diffs() != null ? compared.diffs().size() : 0;
            // 完了行除外前の読込表を履歴保存する（設計どおりチェック時に読んだ表）。
            IdentityCheckHistoryStore.save(
                    u,
                    excel.get(),
                    tab,
                    resultKey,
                    compared.badgeText(),
                    diffCount,
                    excel,
                    planSource);
            return new Result(
                    compared.identical(),
                    false,
                    compared.badgeText(),
                    compared.message(),
                    compared.diffs(),
                    excel,
                    planSource);
        } catch (IOException ex) {
            return errorResult(
                    ex.getMessage() != null ? ex.getMessage() : ex.toString(), excel, planSource);
        }
    }

    /** シス計があるキーだけを突合する。加工計画にあってシス計が無いセルは比較しない。 */
    public static Result compare(
            List<SystemQty> systemQtys,
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup) {
        Map<String, Map<String, Map<String, Map<String, Double>>>> plan =
                lookup != null ? lookup : Map.of();
        List<Diff> diffs = new ArrayList<>();
        if (systemQtys != null) {
            for (SystemQty s : systemQtys) {
                if (s == null || s.date() == null) {
                    continue;
                }
                if (Math.abs(s.qty()) <= DispatchAladdinEntrySheetBuilder.QTY_MATCH_EPS) {
                    continue;
                }
                String dateKey = isoSlash(s.date());
                double planQty =
                        AladdinShapedPlanQtyLookup.lookup(
                                plan,
                                s.machineName(),
                                JuchuTransferValueNormalizer.normalizeKey(s.taskId()),
                                dateKey,
                                s.processName());
                if (Math.abs(planQty - s.qty()) > DispatchAladdinEntrySheetBuilder.QTY_MATCH_EPS) {
                    diffs.add(
                            new Diff(
                                    nz(s.machineName()),
                                    nz(s.taskId()),
                                    nz(s.processName()),
                                    s.date(),
                                    s.qty(),
                                    planQty));
                }
            }
        }
        if (diffs.isEmpty()) {
            return new Result(
                    true, false, BADGE_IDENTICAL, BADGE_IDENTICAL, List.of(), Optional.empty(), Optional.empty());
        }
        String badge = "差異 " + diffs.size() + "件";
        return new Result(false, false, badge, badge, List.copyOf(diffs), Optional.empty(), Optional.empty());
    }

    static PlanInputTabularIo.TabularSheet excludeCompletedPlanRows(PlanInputTabularIo.TabularSheet tab) {
        if (tab == null || tab.headers() == null || tab.rows() == null || tab.rows().isEmpty()) {
            return tab != null ? tab : new PlanInputTabularIo.TabularSheet(List.of(), List.of());
        }
        int unprocIdx = indexOfHeader(tab.headers(), "未加工");
        int actualIdx = indexOfHeader(tab.headers(), "実加工数");
        int convIdx = indexOfHeader(tab.headers(), "換算数量");
        if (unprocIdx < 0 && actualIdx < 0 && convIdx < 0) {
            return tab;
        }
        List<List<String>> kept = new ArrayList<>();
        for (List<String> row : tab.rows()) {
            if (!isCompletedAladdinPlanRow(row, unprocIdx, actualIdx, convIdx)) {
                kept.add(row);
            }
        }
        return new PlanInputTabularIo.TabularSheet(tab.headers(), kept);
    }

    static boolean isCompletedAladdinPlanRow(
            List<String> row, int unprocIdx, int actualIdx, int convIdx) {
        double unprocessed = cellQty(row, unprocIdx);
        double actual = cellQty(row, actualIdx);
        double conversion = cellQty(row, convIdx);
        double eps = DispatchAladdinEntrySheetBuilder.QTY_MATCH_EPS;
        if (unprocessed > eps) {
            return false;
        }
        boolean fullUnprocessed = conversion > eps && actual <= eps;
        return !fullUnprocessed;
    }

    private static double cellQty(List<String> row, int idx) {
        if (row == null || idx < 0 || idx >= row.size()) {
            return 0d;
        }
        return JuchuTransferValueNormalizer.normalizeNumeric(row.get(idx));
    }

    /** ログイン中操作者の世代フォルダにある最新 xlsx。 */
    public static Optional<Path> newestOperatorGenerationXlsx(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String operator = OperatorUserPaths.resolveOperatorUser(u);
        Path dir = DispatchAladdinEntryWorkbookExporter.operatorGenerationDir(u, operator);
        if (!Files.isDirectory(dir)) {
            return Optional.empty();
        }
        try (var stream = Files.list(dir)) {
            return stream.filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xlsx"))
                    .max(Comparator.comparingLong(AladdinEntryDispatchPlanIdentityCheck::lastModifiedMillis));
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    static LocalDate referenceDateForExcel(Path excel) {
        if (excel != null && excel.getFileName() != null) {
            Matcher m = GEN_TS.matcher(excel.getFileName().toString());
            if (m.matches()) {
                try {
                    return LocalDate.parse(m.group(1), GEN_DAY);
                } catch (RuntimeException ignored) {
                    // fall through
                }
            }
        }
        if (excel != null) {
            try {
                Instant instant = Files.getLastModifiedTime(excel).toInstant();
                return instant.atZone(ZoneId.systemDefault()).toLocalDate();
            } catch (IOException ignored) {
                // fall through
            }
        }
        return LocalDate.now();
    }

    private static Optional<Path> newestPlanSourceFile(Map<String, String> ui) throws IOException {
        Path dir = AppPaths.resolveTaskInputSourceDir(ui);
        if (dir == null || !Files.isDirectory(dir)) {
            return Optional.empty();
        }
        return NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir)
                .map(p -> p.toAbsolutePath().normalize());
    }

    private static Result errorResult(String message, Optional<Path> excel, Optional<Path> plan) {
        String msg = message != null ? message : "同一化チェックに失敗しました";
        return new Result(false, true, msg, msg, List.of(), excel, plan);
    }


    private static String isoSlash(LocalDate d) {
        return String.format("%04d/%02d/%02d", d.getYear(), d.getMonthValue(), d.getDayOfMonth());
    }


    private static String formatQty(double qty) {
        if (Math.abs(qty - Math.rint(qty)) < 1e-9) {
            return Long.toString(Math.round(qty));
        }
        return Double.toString(qty);
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }

    static List<SystemQty> remapSheetMachines(
            List<SystemQty> systemQtys, PostProcessingPlanMachineLookup.Snapshot snap) {
        if (systemQtys == null || systemQtys.isEmpty()) {
            return systemQtys != null ? systemQtys : List.of();
        }
        List<SystemQty> out = new ArrayList<>(systemQtys.size());
        for (SystemQty s : systemQtys) {
            if (s == null) {
                continue;
            }
            String machine = resolveMachineNameFromSheet(s.machineName(), snap);
            out.add(new SystemQty(machine, s.taskId(), s.processName(), s.date(), s.qty()));
        }
        return List.copyOf(out);
    }

    static String resolveMachineNameFromSheet(
            String sheetName, PostProcessingPlanMachineLookup.Snapshot snap) {
        if (sheetName == null || sheetName.isBlank()) {
            return "";
        }
        String trimmed = sheetName.strip();
        if (snap == null || !snap.loaded()) {
            return trimmed;
        }
        String normSheet = PostProcessingPlanMachineLookup.normalizeMachineNameKey(trimmed);
        String codeByName = snap.machineNameToCode().get(normSheet);
        if (codeByName != null) {
            String name = snap.machineCodeToName().getOrDefault(codeByName, "");
            return name.isBlank() ? trimmed : name;
        }
        for (Map.Entry<String, String> e : snap.machineCodeToName().entrySet()) {
            String label =
                    PostProcessingKouteiNaiyoMasterLookup.displayLabel(e.getKey(), e.getValue());
            String safe = WorkbookUtil.createSafeSheetName(label);
            if (trimmed.equals(label)
                    || trimmed.equals(safe)
                    || normSheet.equals(PostProcessingPlanMachineLookup.normalizeMachineNameKey(label))
                    || normSheet.equals(PostProcessingPlanMachineLookup.normalizeMachineNameKey(safe))) {
                return e.getValue() != null && !e.getValue().isBlank() ? e.getValue() : e.getKey();
            }
        }
        return trimmed;
    }

    private static PostProcessingPlanMachineLookup.Snapshot loadMachineSnapshot(
            Map<String, String> ui) {
        try {
            return PostProcessingPlanMachineLookup.snapshot(ui);
        } catch (IOException e) {
            return PostProcessingPlanMachineLookup.Snapshot.empty();
        }
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }

    private static int indexOfHeader(List<String> headers, String... titles) {
        if (headers == null || titles == null) {
            return -1;
        }
        for (String title : titles) {
            for (int i = 0; i < headers.size(); i++) {
                String h = headers.get(i);
                if (h != null && title.equals(h.strip())) {
                    return i;
                }
            }
        }
        return -1;
    }

}
