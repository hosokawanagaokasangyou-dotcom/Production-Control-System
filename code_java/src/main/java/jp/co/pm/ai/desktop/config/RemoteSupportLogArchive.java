package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.net.InetAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 現工場の共有 DATA 同階層 {@code remote_log/<操作者>/} へ、実行ログを世代保存および日次追記する。
 */
public final class RemoteSupportLogArchive {

    private static final Logger LOG = Logger.getLogger(RemoteSupportLogArchive.class.getName());

    /** 世代保持日数。 */
    public static final int RETENTION_DAYS = 3;

    public static final String UI_RUN_LOG_FILENAME = "ui_run_log.txt";
    public static final String META_JSON_FILENAME = "meta.json";
    public static final String UI_DAILY_DIR_NAME = "ui_daily";

    public static final String STAGE_KOUCHIN = "kouchin";
    public static final String STAGE_KOUCHIN_TREND = "kouchin-trend";
    public static final String STAGE_SESSION = "session";
    public static final String STAGE_RUNTIME = "runtime";

    static final long DAILY_FLUSH_DELAY_MS = 1_500L;
    static final long EVENT_ARCHIVE_DELAY_MS = 2_000L;

    private static final DateTimeFormatter GEN_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
    private static final DateTimeFormatter DAY = DateTimeFormatter.ISO_LOCAL_DATE;
    private static final Pattern GEN_DIR_PREFIX =
            Pattern.compile("^(\\d{8})-\\d{6}_");
    private static final Pattern DAILY_FILE =
            Pattern.compile("^(\\d{4}-\\d{2}-\\d{2})\\.txt$");
    private static final ObjectMapper JSON = new ObjectMapper();

    private static final ExecutorService EXEC =
            Executors.newSingleThreadExecutor(
                    r -> {
                        Thread t = new Thread(r, "pm-ai-remote-support-log");
                        t.setDaemon(true);
                        return t;
                    });

    private static final ScheduledExecutorService SCHED =
            Executors.newSingleThreadScheduledExecutor(
                    r -> {
                        Thread t = new Thread(r, "pm-ai-remote-support-log-sched");
                        t.setDaemon(true);
                        return t;
                    });

    private static final Object DAILY_LOCK = new Object();
    private static final StringBuilder DAILY_BUF = new StringBuilder();
    private static Map<String, String> dailyUi = Map.of();
    private static String dailyOperator = "";
    private static ScheduledFuture<?> dailyFlushFuture;

    private static final Object EVENT_LOCK = new Object();
    private static ScheduledFuture<?> eventArchiveFuture;

    private RemoteSupportLogArchive() {}

    /**
     * MainShell の段階スクリプト定数から stageId を返す。対象外は null。
     */
    public static String stageIdForMainShellScript(
            String script, String stage1, String stage2, String stage21) {
        if (script == null) {
            return null;
        }
        if (Objects.equals(script, stage1)) {
            return "stage1";
        }
        if (Objects.equals(script, stage2)) {
            return "stage2";
        }
        if (Objects.equals(script, stage21)) {
            return "stage2.1";
        }
        return null;
    }

    /** 段階1／2／2.1 のときだけ Python 実行ログと配台 JSON を同梱する。 */
    public static boolean copiesPipelineArtifacts(String stageId) {
        if (stageId == null || stageId.isBlank()) {
            return false;
        }
        String s = stageId.strip();
        return "stage1".equals(s) || "stage2".equals(s) || "stage2.1".equals(s);
    }

    /**
     * 開発者がユーザーPCの実行時問題を追うために世代スナップショットを取る行。
     * {@code [remote_log]} 自身は再帰防止で除外する。
     */
    public static boolean isDiagnosticRuntimeLine(String line) {
        if (line == null || line.isBlank()) {
            return false;
        }
        if (line.contains("[remote_log]")) {
            return false;
        }
        if (line.contains("[kouchin]")) {
            return true;
        }
        if (line.contains("[kouchin-trend]")) {
            return true;
        }
        if (line.contains("読取不可")) {
            return true;
        }
        if (line.contains("ドロップされたファイルがありません")) {
            return true;
        }
        return line.contains("Exception");
    }

    public static String eventIdForDiagnosticLine(String line) {
        if (line != null && line.contains("[kouchin-trend]")) {
            return STAGE_KOUCHIN_TREND;
        }
        if (line != null && line.contains("[kouchin]")) {
            return STAGE_KOUCHIN;
        }
        return STAGE_RUNTIME;
    }

    /** {@link AppPaths#KEY_PM_AI_REMOTE_LOG} が無効化されていなければ true。 */
    public static boolean isEnabled(Map<String, String> ui) {
        return AppPaths.isTruthyUiEnv(ui, AppPaths.KEY_PM_AI_REMOTE_LOG, true);
    }

    public static String generationDirName(LocalDateTime when, String stageId) {
        LocalDateTime t = when != null ? when : LocalDateTime.now();
        String sid = (stageId == null || stageId.isBlank()) ? "stage" : stageId.strip();
        return GEN_TS.format(t) + "_" + sid.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    /**
     * 世代フォルダ名の先頭日付が retentionDays より古ければ削除対象。
     *
     * @return 削除すべきとき true。日付が取れないフォルダは false（残す）
     */
    public static boolean isGenerationExpired(
            String dirName, LocalDate today, int retentionDays) {
        if (dirName == null || dirName.isBlank() || today == null || retentionDays < 0) {
            return false;
        }
        Matcher m = GEN_DIR_PREFIX.matcher(dirName.strip());
        if (!m.find()) {
            return false;
        }
        try {
            LocalDate day =
                    LocalDate.parse(m.group(1), DateTimeFormatter.BASIC_ISO_DATE);
            return day.isBefore(today.minusDays(retentionDays));
        } catch (DateTimeParseException ex) {
            return false;
        }
    }

    public static boolean isDailyUiLogExpired(
            String fileName, LocalDate today, int retentionDays) {
        if (fileName == null || fileName.isBlank() || today == null || retentionDays < 0) {
            return false;
        }
        Matcher m = DAILY_FILE.matcher(fileName.strip());
        if (!m.matches()) {
            return false;
        }
        try {
            LocalDate day = LocalDate.parse(m.group(1), DAY);
            return day.isBefore(today.minusDays(retentionDays));
        } catch (DateTimeParseException ex) {
            return false;
        }
    }

    public static Path resolveDailyUiLogFile(
            Map<String, String> ui, String operatorName, LocalDate day) {
        LocalDate d = day != null ? day : LocalDate.now(ZoneId.systemDefault());
        return AppPaths.resolveRemoteLogRoot(ui)
                .resolve(OperatorUserPaths.sanitizeOperatorDirName(operatorName))
                .resolve(UI_DAILY_DIR_NAME)
                .resolve(DAY.format(d) + ".txt");
    }

    static String resolveOperatorName(Map<String, String> ui) {
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (session != null && !session.isBlank()) {
            return session.strip();
        }
        return OperatorUserPaths.resolveOperatorUser(ui);
    }

    /**
     * 実行・ログ1行を日次ファイルへバッファする。共有フォルダへの実書込は短遅延後。
     */
    public static void offerDailyUiLogLine(Map<String, String> ui, String line) {
        if (!isEnabled(ui) || line == null || line.isEmpty()) {
            return;
        }
        String operator = resolveOperatorName(ui);
        if (operator == null || operator.isBlank()) {
            return;
        }
        Map<String, String> uiCopy = ui != null ? Map.copyOf(ui) : Map.of();
        synchronized (DAILY_LOCK) {
            dailyUi = uiCopy;
            dailyOperator = operator;
            DAILY_BUF.append(line);
            if (!line.endsWith("\n")) {
                DAILY_BUF.append('\n');
            }
            if (dailyFlushFuture != null) {
                dailyFlushFuture.cancel(false);
            }
            dailyFlushFuture =
                    SCHED.schedule(
                            RemoteSupportLogArchive::flushDailyBufferQuietly,
                            DAILY_FLUSH_DELAY_MS,
                            TimeUnit.MILLISECONDS);
        }
    }

    /** バッファ中の日次ログをすぐ書く（終了時）。 */
    public static void flushDailyUiLogNow() {
        ScheduledFuture<?> pending;
        synchronized (DAILY_LOCK) {
            pending = dailyFlushFuture;
            dailyFlushFuture = null;
        }
        if (pending != null) {
            pending.cancel(false);
        }
        flushDailyBufferQuietly();
    }

    /**
     * 診断イベント後に世代フォルダを遅延保存する（連続する [kouchin] 行を1世代にまとめる）。
     */
    public static void scheduleArchiveAfterEventAsync(
            Map<String, String> ui,
            String stageId,
            Integer exitCode,
            Throwable error,
            Supplier<String> uiLogSupplier,
            Consumer<String> logConsumer) {
        if (!isEnabled(ui) || stageId == null || stageId.isBlank() || uiLogSupplier == null) {
            return;
        }
        Map<String, String> uiCopy = ui != null ? Map.copyOf(ui) : Map.of();
        String stage = stageId.strip();
        Integer code = exitCode;
        Throwable err = error;
        Consumer<String> log = logConsumer != null ? logConsumer : s -> {};
        synchronized (EVENT_LOCK) {
            if (eventArchiveFuture != null) {
                eventArchiveFuture.cancel(false);
            }
            eventArchiveFuture =
                    SCHED.schedule(
                            () -> {
                                String uiLog;
                                try {
                                    uiLog = uiLogSupplier.get();
                                } catch (Exception ex) {
                                    uiLog = "";
                                }
                                archiveAfterStageAsync(
                                        uiCopy, stage, code, err, uiLog != null ? uiLog : "", log);
                            },
                            EVENT_ARCHIVE_DELAY_MS,
                            TimeUnit.MILLISECONDS);
        }
    }

    /**
     * アプリ終了時: 日次バッファを掃き出し、session 世代を同期保存する。
     */
    public static void flushOnShutdown(
            Map<String, String> ui, String operatorName, String uiLogText) {
        flushDailyUiLogNow();
        if (!isEnabled(ui)) {
            return;
        }
        String operator =
                operatorName != null && !operatorName.isBlank()
                        ? operatorName.strip()
                        : resolveOperatorName(ui);
        if (operator == null || operator.isBlank()) {
            return;
        }
        try {
            archiveAfterStage(
                    ui,
                    operator,
                    STAGE_SESSION,
                    null,
                    null,
                    uiLogText,
                    LocalDateTime.now(ZoneId.systemDefault()));
        } catch (Exception ex) {
            LOG.log(Level.WARNING, "remote_log session 保存失敗", ex);
        }
    }

    /**
     * 段階終了後に非同期でアーカイブする。失敗は logConsumer / JUL のみ（段階結果には影響しない）。
     */
    public static void archiveAfterStageAsync(
            Map<String, String> ui,
            String stageId,
            Integer exitCode,
            Throwable error,
            String uiLogText,
            Consumer<String> logConsumer) {
        if (!isEnabled(ui) || stageId == null || stageId.isBlank()) {
            return;
        }
        String operator = resolveOperatorName(ui);
        if (operator == null || operator.isBlank()) {
            if (logConsumer != null) {
                logConsumer.accept("[remote_log] 操作者が未選択のためスキップしました。");
            }
            return;
        }
        Map<String, String> uiCopy = ui != null ? Map.copyOf(ui) : Map.of();
        String uiText = uiLogText != null ? uiLogText : "";
        String stage = stageId.strip();
        Integer code = exitCode;
        String errMsg = error != null ? String.valueOf(error.getMessage()) : null;
        Consumer<String> log = logConsumer != null ? logConsumer : s -> {};
        EXEC.execute(
                () -> {
                    try {
                        Path gen =
                                archiveAfterStage(
                                        uiCopy,
                                        operator,
                                        stage,
                                        code,
                                        errMsg,
                                        uiText,
                                        LocalDateTime.now(ZoneId.systemDefault()));
                        if (gen != null) {
                            log.accept("[remote_log] 保存: " + gen.toAbsolutePath().normalize());
                        }
                    } catch (Exception ex) {
                        LOG.log(Level.WARNING, "remote_log 保存失敗", ex);
                        log.accept(
                                "[remote_log] 保存失敗: "
                                        + (ex.getMessage() != null
                                                ? ex.getMessage()
                                                : ex.toString()));
                    }
                });
    }

    /**
     * 同期書込（テスト・診断用）。成功時は世代フォルダパス、スキップ時は null。
     */
    public static Path archiveAfterStage(
            Map<String, String> ui,
            String operatorName,
            String stageId,
            Integer exitCode,
            String errorMessage,
            String uiLogText,
            LocalDateTime when)
            throws IOException {
        if (!isEnabled(ui)) {
            return null;
        }
        if (operatorName == null || operatorName.isBlank()) {
            return null;
        }
        if (stageId == null || stageId.isBlank()) {
            return null;
        }
        Path root = AppPaths.resolveRemoteLogRoot(ui);
        Path userDir = root.resolve(OperatorUserPaths.sanitizeOperatorDirName(operatorName));
        Files.createDirectories(userDir);
        Path genDir = userDir.resolve(generationDirName(when, stageId));
        Files.createDirectories(genDir);

        Path uiLog = genDir.resolve(UI_RUN_LOG_FILENAME);
        Files.writeString(
                uiLog,
                uiLogText != null ? uiLogText : "",
                StandardCharsets.UTF_8);

        Path execSrc = AppPaths.resolveExecutionLogTxtPath(ui);
        boolean copiedExec = false;
        Path dispatchJsonSrc = AppPaths.resolveResultDispatchTableStage2JsonPath(ui);
        boolean copiedDispatchJson = false;
        if (copiesPipelineArtifacts(stageId)) {
            if (Files.isRegularFile(execSrc)) {
                Files.copy(
                        execSrc,
                        genDir.resolve(AppPaths.EXECUTION_LOG_TXT),
                        StandardCopyOption.REPLACE_EXISTING);
                copiedExec = true;
            }
            if (Files.isRegularFile(dispatchJsonSrc)) {
                Files.copy(
                        dispatchJsonSrc,
                        genDir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                        StandardCopyOption.REPLACE_EXISTING);
                copiedDispatchJson = true;
            }
        }

        Map<String, Object> meta = new LinkedHashMap<>();
        meta.put("format_version", 1);
        meta.put("stage_id", stageId.strip());
        meta.put("operator", operatorName.strip());
        meta.put("exit_code", exitCode);
        if (errorMessage != null && !errorMessage.isBlank()) {
            meta.put("error", errorMessage.strip());
        }
        meta.put(
                "saved_at",
                (when != null ? when : LocalDateTime.now()).format(DateTimeFormatter.ISO_LOCAL_DATE_TIME));
        meta.put("host", resolveHostNameQuietly());
        meta.put(
                "factory",
                GlobalInitSettingTarget.loadEffective(ui != null ? ui : Map.of()).name());
        meta.put("os_user", nullToEmpty(System.getProperty("user.name")));
        meta.put("java_io_tmpdir", nullToEmpty(System.getProperty("java.io.tmpdir")));
        meta.put(
                "app_version",
                AppVersionInfo.resolveDisplayedVersion(
                        Path.of(System.getProperty("user.dir", ".")),
                        ui != null ? ui : Map.of()));
        meta.put("ui_log_chars", uiLogText != null ? uiLogText.length() : 0);
        meta.put("execution_log_copied", copiedExec);
        meta.put(
                "execution_log_source",
                execSrc.toAbsolutePath().normalize().toString());
        meta.put("result_dispatch_json_copied", copiedDispatchJson);
        meta.put(
                "result_dispatch_json_source",
                dispatchJsonSrc.toAbsolutePath().normalize().toString());
        Files.writeString(
                genDir.resolve(META_JSON_FILENAME),
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(meta) + "\n",
                StandardCharsets.UTF_8);

        pruneExpiredGenerations(
                userDir,
                (when != null ? when : LocalDateTime.now()).toLocalDate(),
                RETENTION_DAYS);
        return genDir.toAbsolutePath().normalize();
    }

    public static Path appendDailyUiLog(
            Map<String, String> ui, String operatorName, String text, LocalDate day)
            throws IOException {
        if (!isEnabled(ui)) {
            return null;
        }
        if (operatorName == null || operatorName.isBlank()) {
            return null;
        }
        Path file = resolveDailyUiLogFile(ui, operatorName, day);
        Files.createDirectories(file.getParent());
        boolean fresh = !Files.isRegularFile(file);
        String body = text != null ? text : "";
        if (!body.isEmpty() && !body.endsWith("\n")) {
            body = body + "\n";
        }
        if (fresh) {
            String header = dailyHeader(ui, operatorName, day != null ? day : LocalDate.now());
            Files.writeString(
                    file,
                    header + body,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.WRITE);
        } else if (!body.isEmpty()) {
            Files.writeString(
                    file,
                    body,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        }
        pruneExpiredDailyUiLogs(
                file.getParent(),
                day != null ? day : LocalDate.now(ZoneId.systemDefault()),
                RETENTION_DAYS);
        return file.toAbsolutePath().normalize();
    }

    /** ユーザーフォルダ内の期限切れ世代を削除。削除したパス一覧を返す。 */
    public static List<Path> pruneExpiredGenerations(
            Path userDir, LocalDate today, int retentionDays) throws IOException {
        List<Path> removed = new ArrayList<>();
        if (userDir == null || !Files.isDirectory(userDir)) {
            return removed;
        }
        List<Path> children = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(userDir)) {
            for (Path child : stream) {
                if (Files.isDirectory(child)
                        && !UI_DAILY_DIR_NAME.equals(child.getFileName().toString())) {
                    children.add(child);
                }
            }
        }
        children.sort(Comparator.comparing(p -> p.getFileName().toString()));
        for (Path child : children) {
            String name = child.getFileName().toString();
            if (!isGenerationExpired(name, today, retentionDays)) {
                continue;
            }
            deleteRecursively(child);
            removed.add(child.toAbsolutePath().normalize());
        }
        return removed;
    }

    public static List<Path> pruneExpiredDailyUiLogs(
            Path dailyDir, LocalDate today, int retentionDays) throws IOException {
        List<Path> removed = new ArrayList<>();
        if (dailyDir == null || !Files.isDirectory(dailyDir)) {
            return removed;
        }
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dailyDir, "*.txt")) {
            for (Path child : stream) {
                if (Files.isRegularFile(child)) {
                    files.add(child);
                }
            }
        }
        for (Path child : files) {
            if (!isDailyUiLogExpired(child.getFileName().toString(), today, retentionDays)) {
                continue;
            }
            Files.deleteIfExists(child);
            removed.add(child.toAbsolutePath().normalize());
        }
        return removed;
    }

    private static void flushDailyBufferQuietly() {
        String chunk;
        Map<String, String> ui;
        String operator;
        synchronized (DAILY_LOCK) {
            if (DAILY_BUF.length() == 0) {
                return;
            }
            chunk = DAILY_BUF.toString();
            DAILY_BUF.setLength(0);
            ui = dailyUi;
            operator = dailyOperator;
        }
        try {
            appendDailyUiLog(ui, operator, chunk, LocalDate.now(ZoneId.systemDefault()));
        } catch (Exception ex) {
            LOG.log(Level.WARNING, "remote_log 日次保存失敗", ex);
        }
    }

    private static String dailyHeader(
            Map<String, String> ui, String operatorName, LocalDate day) {
        return "# pm-ai-desktop ui_daily "
                + DAY.format(day)
                + "\n# operator="
                + operatorName.strip()
                + " factory="
                + GlobalInitSettingTarget.loadEffective(ui != null ? ui : Map.of()).name()
                + " host="
                + resolveHostNameQuietly()
                + " os_user="
                + nullToEmpty(System.getProperty("user.name"))
                + " app_version="
                + AppVersionInfo.resolveDisplayedVersion(
                        Path.of(System.getProperty("user.dir", ".")),
                        ui != null ? ui : Map.of())
                + " java_io_tmpdir="
                + nullToEmpty(System.getProperty("java.io.tmpdir"))
                + "\n";
    }

    private static void deleteRecursively(Path root) throws IOException {
        if (root == null || !Files.exists(root)) {
            return;
        }
        if (Files.isRegularFile(root)) {
            Files.deleteIfExists(root);
            return;
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path child : stream) {
                deleteRecursively(child);
            }
        }
        Files.deleteIfExists(root);
    }

    private static String resolveHostNameQuietly() {
        try {
            return InetAddress.getLocalHost().getHostName();
        } catch (Exception ex) {
            return "";
        }
    }

    private static String nullToEmpty(String s) {
        return s != null ? s : "";
    }
}
