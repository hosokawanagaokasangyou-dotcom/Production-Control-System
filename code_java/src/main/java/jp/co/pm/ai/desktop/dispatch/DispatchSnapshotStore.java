package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.net.InetAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;

/**
 * ローカルの段階成果物を工場共有へ世代コピーする。書き込み正本はローカルのまま。
 * 一覧・削除・公開の失敗は呼び出し側が握る。壊れた世代は読み飛ばし、アプリを止めない。
 */
public final class DispatchSnapshotStore {

    public static final int FORMAT_VERSION = 1;
    public static final int RETENTION_DAYS = 30;
    public static final String META_JSON = "meta.json";

    private static final DateTimeFormatter GEN_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss-SSS", Locale.ROOT);
    private static final DateTimeFormatter DAY = DateTimeFormatter.BASIC_ISO_DATE;
    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchSnapshotStore() {}

    public record PublishResult(Path generationDir, List<String> copied, List<String> missing) {}

    public record SnapshotRef(
            String operatorDir,
            String generationDir,
            String stage,
            String host,
            String savedAt,
            List<String> missing,
            int formatVersion) {}

    /**
     * ローカル output の現在の一組を共有へコピーする。共有不可なら例外。呼び出し側は配台成功を失敗にしない。
     */
    public static PublishResult publishLocalOutput(Map<String, String> ui, String stageId)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path localDir = AppPaths.defaultPlanningOutputDir(u);
        Path plan;
        Path member;
        try {
            plan = Stage2OutputNaming.newestPrimaryPlanJson(localDir);
            member = Stage2OutputNaming.pairedMemberJson(plan);
        } catch (IOException ex) {
            throw ex;
        }
        Path dispatch = AppPaths.resolveResultDispatchTableJsonPath(u);
        Path shapedAladdin = AppPaths.resolveShapedAladdinPlanJsonPath(u);
        Path shapedActuals = AppPaths.resolveShapedProcessingActualsJsonPath(u);
        String operator = OperatorUserPaths.resolveOperatorUser(u);
        return publish(
                AppPaths.resolveDispatchSnapshotRoot(u),
                operator,
                stageId,
                LocalDateTime.now(),
                plan,
                member,
                dispatch,
                shapedAladdin,
                shapedActuals);
    }

    public static PublishResult publish(
            Path snapshotRoot,
            String operatorName,
            String stageId,
            LocalDateTime when,
            Path planJson,
            Path memberJson,
            Path dispatchJson,
            Path shapedAladdin,
            Path shapedActuals)
            throws IOException {
        if (snapshotRoot == null) {
            throw new IOException("snapshot root is null");
        }
        String operatorDir = OperatorUserPaths.sanitizeOperatorDirName(operatorName);
        LocalDateTime t = when != null ? when : LocalDateTime.now();
        String stage = sanitizeToken(stageId == null || stageId.isBlank() ? "stage" : stageId);
        String host = sanitizeToken(hostName());
        String genName = GEN_TS.format(t) + "_" + host + "_" + stage;
        Path genDir = snapshotRoot.resolve(operatorDir).resolve(genName);
        boolean committed = false;
        try {
            Files.createDirectories(genDir);

            List<String> copied = new ArrayList<>();
            List<String> missing = new ArrayList<>();
            copyIfPresent(planJson, genDir, copied, missing);
            Path memberSameRun = memberOnSameStamp(planJson, memberJson);
            if (memberSameRun == null) {
                String expected = Stage2OutputNaming.expectedMemberFileName(planJson);
                missing.add(expected != null ? expected : "(none)");
            } else {
                copyIfPresent(memberSameRun, genDir, copied, missing);
            }
            copyNamed(dispatchJson, genDir, AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME, copied, missing);
            copyNamed(shapedAladdin, genDir, AppPaths.SHAPED_ALADDIN_PLAN_JSON_BASENAME, copied, missing);
            copyNamed(
                    shapedActuals,
                    genDir,
                    AppPaths.SHAPED_PROCESSING_ACTUALS_JSON_BASENAME,
                    copied,
                    missing);

            Map<String, Object> meta = new LinkedHashMap<>();
            meta.put("format_version", FORMAT_VERSION);
            meta.put("operator", operatorName == null ? "" : operatorName.strip());
            meta.put("operator_dir", operatorDir);
            meta.put("stage", stage);
            meta.put("host", host);
            meta.put("saved_at", t.toString());
            meta.put("copied", copied);
            meta.put("missing", missing);
            Files.writeString(
                    genDir.resolve(META_JSON),
                    JSON.writerWithDefaultPrettyPrinter().writeValueAsString(meta) + "\n",
                    StandardCharsets.UTF_8);
            committed = true;
            pruneExpired(snapshotRoot, t.toLocalDate(), RETENTION_DAYS);
            return new PublishResult(genDir, List.copyOf(copied), List.copyOf(missing));
        } catch (IOException ex) {
            if (!committed) {
                deleteTreeQuiet(genDir);
            }
            throw ex;
        }
    }

    /** 読める世代だけ返す。壊れた meta や未知の format_version は飛ばす。 */
    public static List<SnapshotRef> list(Path snapshotRoot) {
        List<SnapshotRef> out = new ArrayList<>();
        if (snapshotRoot == null || !Files.isDirectory(snapshotRoot)) {
            return out;
        }
        try (DirectoryStream<Path> operators = Files.newDirectoryStream(snapshotRoot)) {
            for (Path operator : operators) {
                if (!Files.isDirectory(operator)) {
                    continue;
                }
                Path opName = operator.getFileName();
                if (opName == null) {
                    continue;
                }
                try (DirectoryStream<Path> gens = Files.newDirectoryStream(operator)) {
                    for (Path gen : gens) {
                        SnapshotRef ref = readRef(opName.toString(), gen);
                        if (ref != null) {
                            out.add(ref);
                        }
                    }
                } catch (IOException ex) {
                    // その操作者フォルダだけ飛ばす
                }
            }
        } catch (IOException ex) {
            return List.of();
        }
        out.sort(Comparator.comparing(SnapshotRef::generationDir).reversed());
        return out;
    }

    /**
     * 自分の世代フォルダだけ消す。他者の名前や、世代以外のフォルダは消さない。
     *
     * @return 削除できたとき true
     */
    public static boolean deleteOwn(Path snapshotRoot, String operatorName, String generationDir) {
        if (snapshotRoot == null || generationDir == null || generationDir.isBlank()) {
            return false;
        }
        if (!isGenerationDirName(generationDir)) {
            return false;
        }
        String operatorDir = OperatorUserPaths.sanitizeOperatorDirName(operatorName);
        Path target = snapshotRoot.resolve(operatorDir).resolve(generationDir).normalize();
        Path operatorRoot = snapshotRoot.resolve(operatorDir).normalize();
        if (!target.startsWith(operatorRoot) || target.equals(operatorRoot)) {
            return false;
        }
        if (!Files.isDirectory(target)) {
            return false;
        }
        try {
            try (var walk = Files.walk(target)) {
                walk.sorted(Comparator.reverseOrder()).forEach(p -> {
                    try {
                        Files.deleteIfExists(p);
                    } catch (IOException ex) {
                        throw new IllegalStateException(ex);
                    }
                });
            }
            return !Files.exists(target);
        } catch (IOException | IllegalStateException ex) {
            return false;
        }
    }

    public static boolean isExpired(String generationDir, LocalDate today, int retentionDays) {
        if (generationDir == null || today == null || retentionDays < 0) {
            return false;
        }
        if (generationDir.length() < 8) {
            return false;
        }
        String dayText = generationDir.substring(0, 8);
        try {
            LocalDate day = LocalDate.parse(dayText, DAY);
            return day.isBefore(today.minusDays(retentionDays));
        } catch (DateTimeParseException ex) {
            return false;
        }
    }

    public static Path generationDir(Path snapshotRoot, String operatorDir, String generationDir) {
        if (snapshotRoot == null || operatorDir == null || generationDir == null) {
            return null;
        }
        if (!isGenerationDirName(generationDir)) {
            return null;
        }
        String operator = operatorDir.strip();
        if (!operator.equals(OperatorUserPaths.sanitizeOperatorDirName(operator))) {
            return null;
        }
        Path root = snapshotRoot.toAbsolutePath().normalize();
        Path dir = root.resolve(operator).resolve(generationDir).normalize();
        if (!dir.startsWith(root)) {
            return null;
        }
        return dir;
    }

    static boolean isGenerationDirName(String name) {
        if (name == null || name.isBlank()) {
            return false;
        }
        return name.matches("\\d{8}-\\d{6}-\\d{3}_[^\\\\/:*?\"<>|]+_[^\\\\/:*?\"<>|]+");
    }

    private static void pruneExpired(Path snapshotRoot, LocalDate today, int retentionDays) {
        if (snapshotRoot == null || !Files.isDirectory(snapshotRoot)) {
            return;
        }
        try (DirectoryStream<Path> operators = Files.newDirectoryStream(snapshotRoot)) {
            for (Path operator : operators) {
                if (!Files.isDirectory(operator)) {
                    continue;
                }
                Path opName = operator.getFileName();
                if (opName == null) {
                    continue;
                }
                try (DirectoryStream<Path> gens = Files.newDirectoryStream(operator)) {
                    for (Path gen : gens) {
                        Path fn = gen.getFileName();
                        if (fn == null || !Files.isDirectory(gen)) {
                            continue;
                        }
                        if (isExpired(fn.toString(), today, retentionDays)) {
                            deleteOwn(snapshotRoot, opName.toString(), fn.toString());
                        }
                    }
                } catch (IOException ex) {
                    // 次の操作者へ
                }
            }
        } catch (IOException ex) {
            // 削除できなくても公開済み世代は残す
        }
    }

    private static SnapshotRef readRef(String operatorDir, Path gen) {
        if (!Files.isDirectory(gen)) {
            return null;
        }
        Path fn = gen.getFileName();
        if (fn == null || !isGenerationDirName(fn.toString())) {
            return null;
        }
        Path metaPath = gen.resolve(META_JSON);
        if (!Files.isRegularFile(metaPath)) {
            return new SnapshotRef(operatorDir, fn.toString(), "", "", "", List.of(), 0);
        }
        try {
            JsonNode meta = JSON.readTree(Files.readString(metaPath, StandardCharsets.UTF_8));
            int version = meta.path("format_version").asInt(0);
            if (version > FORMAT_VERSION) {
                return null;
            }
            List<String> missing = new ArrayList<>();
            JsonNode missingNode = meta.get("missing");
            if (missingNode != null && missingNode.isArray()) {
                for (JsonNode n : missingNode) {
                    missing.add(n.asText(""));
                }
            }
            return new SnapshotRef(
                    operatorDir,
                    fn.toString(),
                    meta.path("stage").asText(""),
                    meta.path("host").asText(""),
                    meta.path("saved_at").asText(""),
                    List.copyOf(missing),
                    version);
        } catch (IOException | RuntimeException ex) {
            return null;
        }
    }

    /** ファイル名のスタンプが計画と一致する人員だけ。別スタンプは隣にあっても使わない。 */
    private static Path memberOnSameStamp(Path planJson, Path memberJson) {
        String expected = Stage2OutputNaming.expectedMemberFileName(planJson);
        if (expected == null) {
            return null;
        }
        if (memberJson != null
                && memberJson.getFileName() != null
                && expected.equals(memberJson.getFileName().toString())
                && Files.isRegularFile(memberJson, LinkOption.NOFOLLOW_LINKS)) {
            return memberJson;
        }
        return Stage2OutputNaming.pairedMemberJson(planJson);
    }

    private static void copyIfPresent(Path source, Path genDir, List<String> copied, List<String> missing)
            throws IOException {
        if (source == null || !Files.isRegularFile(source, LinkOption.NOFOLLOW_LINKS)) {
            missing.add(source == null ? "(none)" : source.getFileName().toString());
            return;
        }
        Path name = source.getFileName();
        if (name == null) {
            missing.add("(none)");
            return;
        }
        Files.copy(source, genDir.resolve(name.toString()), StandardCopyOption.REPLACE_EXISTING);
        copied.add(name.toString());
    }

    private static void copyNamed(
            Path source, Path genDir, String name, List<String> copied, List<String> missing)
            throws IOException {
        if (source == null || !Files.isRegularFile(source, LinkOption.NOFOLLOW_LINKS)) {
            missing.add(name);
            return;
        }
        Files.copy(source, genDir.resolve(name), StandardCopyOption.REPLACE_EXISTING);
        copied.add(name);
    }

    private static String hostName() {
        try {
            String host = InetAddress.getLocalHost().getHostName();
            return host == null || host.isBlank() ? "pc" : host.strip();
        } catch (Exception ex) {
            return "pc";
        }
    }

    private static void deleteTreeQuiet(Path dir) {
        if (dir == null || !Files.exists(dir)) {
            return;
        }
        try (var walk = Files.walk(dir)) {
            for (Path p : walk.sorted((a, b) -> b.getNameCount() - a.getNameCount()).toList()) {
                Files.deleteIfExists(p);
            }
        } catch (IOException ex) {
            // 残骸が残っても公開失敗としては呼び出し側に返す
        }
    }

    private static String sanitizeToken(String raw) {
        String t = raw == null ? "" : raw.strip().replaceAll("[\\\\/:*?\"<>|\\s]+", "_");
        if (t.isEmpty()) {
            return "x";
        }
        if (t.length() > 24) {
            return t.substring(0, 24);
        }
        return t;
    }
}
