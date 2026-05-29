package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Stream;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;

/**
 * 段階2.1 成果物（{@code output/stage21/}）をメイン output 正本へ反映する。
 */
public final class Stage21OutputPromoter {

    private static final String OVERTIME_OVERRIDES_BASENAME = "overtime_simulation_overrides.json";

    private Stage21OutputPromoter() {}

    public record Result(
            int filesCopied,
            Path mainDispatchJson,
            Path mainPlanJson,
            Path mainMemberJson,
            List<Path> copiedPaths) {}

    /**
     * {@code stage21/} 内の最新段階2.1 成果物を {@link AppPaths#resolveResultDispatchTableDir} へ上書きコピーする。
     */
    public static Result promoteToMainOutput(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path stage21Dir = AppPaths.resolveStage21OutputDir(u);
        Path mainDir = AppPaths.resolveResultDispatchTableDir(u);
        if (!Files.isDirectory(stage21Dir)) {
            throw new IOException("段階2.1 出力フォルダがありません: " + stage21Dir);
        }
        Files.createDirectories(mainDir);

        List<Path> copied = new ArrayList<>();
        Path mainDispatch = mainDir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Path stage21Dispatch = stage21Dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        if (Files.isRegularFile(stage21Dispatch)) {
            copyReplace(stage21Dispatch, mainDispatch, copied);
        }

        Path overtimeSrc = stage21Dir.resolve(OVERTIME_OVERRIDES_BASENAME);
        if (Files.isRegularFile(overtimeSrc)) {
            copyReplace(overtimeSrc, mainDir.resolve(OVERTIME_OVERRIDES_BASENAME), copied);
        }

        Path planJson = Stage2OutputNaming.newestPrimaryPlanJson(stage21Dir);
        if (planJson != null) {
            copyArtifactFamily(stage21Dir, mainDir, planJson, copied);
        }
        Path planXlsx = Stage2OutputNaming.newestPrimaryPlanXlsx(stage21Dir);
        if (planXlsx != null) {
            copyArtifactFamily(stage21Dir, mainDir, planXlsx, copied);
        }

        Path memberJson = Stage2OutputNaming.newestPrimaryMemberJson(stage21Dir);
        if (memberJson != null) {
            copyArtifactFamily(stage21Dir, mainDir, memberJson, copied);
        }
        Path memberXlsx = Stage2OutputNaming.newestPrimaryMemberXlsx(stage21Dir);
        if (memberXlsx != null) {
            copyArtifactFamily(stage21Dir, mainDir, memberXlsx, copied);
        }

        if (copied.isEmpty()) {
            throw new IOException("段階2.1 出力に反映対象ファイルがありません: " + stage21Dir);
        }

        Path mainPlan =
                planJson != null
                        ? mainDir.resolve(planJson.getFileName())
                        : null;
        Path mainMember =
                memberJson != null
                        ? mainDir.resolve(memberJson.getFileName())
                        : null;
        return new Result(
                copied.size(),
                Files.isRegularFile(mainDispatch) ? mainDispatch : null,
                mainPlan,
                mainMember,
                List.copyOf(copied));
    }

    private static void copyArtifactFamily(
            Path sourceDir, Path targetDir, Path primary, List<Path> copied) throws IOException {
        String familyPrefix = artifactFamilyPrefix(primary);
        if (familyPrefix == null) {
            copyReplace(primary, targetDir.resolve(primary.getFileName()), copied);
            return;
        }
        try (Stream<Path> stream = Files.list(sourceDir)) {
            stream.filter(Files::isRegularFile)
                    .filter(p -> belongsToArtifactFamily(p, familyPrefix))
                    .forEach(
                            p -> {
                                try {
                                    copyReplace(p, targetDir.resolve(p.getFileName()), copied);
                                } catch (IOException ex) {
                                    throw new PromotionIOException(ex);
                                }
                            });
        } catch (PromotionIOException ex) {
            throw ex.ioException();
        }
    }

    private static void copyReplace(Path source, Path target, List<Path> copied) throws IOException {
        Objects.requireNonNull(source);
        Objects.requireNonNull(target);
        Files.createDirectories(target.getParent());
        Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
        copied.add(target.toAbsolutePath().normalize());
    }

    static String artifactFamilyPrefix(Path artifactPath) {
        if (artifactPath == null) {
            return null;
        }
        Path fn = artifactPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        if (Stage2OutputNaming.acceptsPrimaryPlanJson(artifactPath)
                || Stage2OutputNaming.acceptsPrimaryPlanXlsx(artifactPath)) {
            if (name.startsWith(Stage2OutputNaming.PLAN_PREFIX)) {
                int stemEnd = Stage2OutputNaming.PLAN_PREFIX.length() + Stage2OutputNaming.STAMP_DIGITS;
                if (name.length() >= stemEnd) {
                    return name.substring(0, stemEnd);
                }
            }
            int dot = name.lastIndexOf('.');
            return dot > 0 ? name.substring(0, dot) : name;
        }
        if (Stage2OutputNaming.acceptsPrimaryMemberJson(artifactPath)
                || Stage2OutputNaming.acceptsPrimaryMemberXlsx(artifactPath)) {
            if (name.startsWith(Stage2OutputNaming.MEMBER_PREFIX)) {
                int stemEnd =
                        Stage2OutputNaming.MEMBER_PREFIX.length() + Stage2OutputNaming.STAMP_DIGITS;
                if (name.length() >= stemEnd) {
                    return name.substring(0, stemEnd);
                }
            }
            int dot = name.lastIndexOf('.');
            return dot > 0 ? name.substring(0, dot) : name;
        }
        return null;
    }

    static boolean belongsToArtifactFamily(Path candidate, String familyPrefix) {
        if (candidate == null || familyPrefix == null || familyPrefix.isBlank()) {
            return false;
        }
        Path fn = candidate.getFileName();
        if (fn == null) {
            return false;
        }
        String name = fn.toString();
        if (name.startsWith(familyPrefix)) {
            return true;
        }
        if (familyPrefix.startsWith(Stage2OutputNaming.PLAN_PREFIX)) {
            return name.startsWith("production_plan_multi_day_")
                    && familyPrefix.startsWith("production_plan_multi_day_");
        }
        if (familyPrefix.startsWith("member_schedule_")) {
            return name.startsWith("member_schedule_");
        }
        return false;
    }

    private static final class PromotionIOException extends RuntimeException {
        PromotionIOException(IOException cause) {
            super(cause);
        }

        IOException ioException() {
            return (IOException) super.getCause();
        }
    }
}
