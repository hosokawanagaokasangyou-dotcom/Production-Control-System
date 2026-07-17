package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.net.InetAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
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
import java.util.function.Consumer;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * サマリ Excel 同階層の {@code remote_log/<操作者>/} へ、段階終了時の実行ログを世代保存する。
 */
public final class RemoteSupportLogArchive {

    private static final Logger LOG = Logger.getLogger(RemoteSupportLogArchive.class.getName());

    /** 世代保持日数。 */
    public static final int RETENTION_DAYS = 3;

    public static final String UI_RUN_LOG_FILENAME = "ui_run_log.txt";
    public static final String META_JSON_FILENAME = "meta.json";

    private static final DateTimeFormatter GEN_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
    private static final Pattern GEN_DIR_PREFIX =
            Pattern.compile("^(\\d{8})-\\d{6}_");
    private static final ObjectMapper JSON = new ObjectMapper();

    private static final ExecutorService EXEC =
            Executors.newSingleThreadExecutor(
                    r -> {
                        Thread t = new Thread(r, "pm-ai-remote-support-log");
                        t.setDaemon(true);
                        return t;
                    });

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
        String operator = FactoryOperatorUserStore.sessionOperatorName();
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
        if (Files.isRegularFile(execSrc)) {
            Files.copy(
                    execSrc,
                    genDir.resolve(AppPaths.EXECUTION_LOG_TXT),
                    StandardCopyOption.REPLACE_EXISTING);
            copiedExec = true;
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
        meta.put("ui_log_chars", uiLogText != null ? uiLogText.length() : 0);
        meta.put("execution_log_copied", copiedExec);
        meta.put(
                "execution_log_source",
                execSrc.toAbsolutePath().normalize().toString());
        Files.writeString(
                genDir.resolve(META_JSON_FILENAME),
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(meta) + "\n",
                StandardCharsets.UTF_8);

        pruneExpiredGenerations(userDir, LocalDate.now(ZoneId.systemDefault()), RETENTION_DAYS);
        return genDir.toAbsolutePath().normalize();
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
                if (Files.isDirectory(child)) {
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
}
