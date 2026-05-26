package jp.co.pm.ai.desktop.dispatch.rules.paths;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;

/** Path resolution for {@code dispatch_special_rules/} work directory. */
public final class DispatchRulePaths {

    public static final String DISPATCH_SPECIAL_RULES_DIR = "dispatch_special_rules";
    public static final String DISPATCH_SPECIAL_RULES_FILENAME = "dispatch_special_rules.json";
    public static final String KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON =
            "PM_AI_DISPATCH_SPECIAL_RULES_JSON";
    public static final String KEY_PM_AI_DISPATCH_RULE_ENGINE = "PM_AI_DISPATCH_RULE_ENGINE";
    public static final String KEY_PM_AI_DISPATCH_RULE_LEGACY_FALLBACK =
            "PM_AI_DISPATCH_RULE_LEGACY_FALLBACK";
    public static final String KEY_PM_AI_DISPATCH_RULE_HISTORY_MAX =
            "PM_AI_DISPATCH_RULE_HISTORY_MAX";
    public static final String KEY_PM_AI_DISPATCH_RULE_RUN_SNAPSHOT_MAX =
            "PM_AI_DISPATCH_RULE_RUN_SNAPSHOT_MAX";
    public static final String KEY_PM_AI_DISPATCH_RULE_BLOCK_ON_CONFLICT =
            "PM_AI_DISPATCH_RULE_BLOCK_ON_CONFLICT";

    private DispatchRulePaths() {}

    public static Path workDirectory(Map<String, String> ui) {
        Path summary = AppPaths.summaryAiDispatchXlsxPath(ui);
        Path parent = summary.getParent();
        if (parent == null) {
            return Path.of(".").resolve(DISPATCH_SPECIAL_RULES_DIR);
        }
        return parent.resolve(DISPATCH_SPECIAL_RULES_DIR);
    }

    public static Path workJsonPath(Map<String, String> ui) {
        return workDirectory(ui).resolve(DISPATCH_SPECIAL_RULES_FILENAME);
    }

    public static Path bundledTemplatePath(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui)
                .resolve("code")
                .resolve("json")
                .resolve(DISPATCH_SPECIAL_RULES_DIR)
                .resolve(DISPATCH_SPECIAL_RULES_FILENAME);
    }

    public static Path historyDirectory(Map<String, String> ui) {
        return workDirectory(ui).resolve("history");
    }

    public static Path runSnapshotsDirectory(Map<String, String> ui) {
        return workDirectory(ui).resolve("run_snapshots");
    }

    public static Path resolveWorkJson(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = u.get(KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON);
        if (explicit != null && !explicit.isBlank()) {
            Path p = Path.of(explicit.trim());
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        Path work = workJsonPath(u);
        if (Files.isRegularFile(work)) {
            return work.toAbsolutePath().normalize();
        }
        return work.toAbsolutePath().normalize();
    }

    public static boolean ensureWorkJsonFromRepoIfMissing(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path target = workJsonPath(u);
        if (Files.isRegularFile(target)) {
            return true;
        }
        Path source = bundledTemplatePath(u);
        if (!Files.isRegularFile(source)) {
            return false;
        }
        try {
            Files.createDirectories(target.getParent());
            Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
            return Files.isRegularFile(target);
        } catch (IOException ex) {
            return false;
        }
    }

    public static Optional<Path> resolveDefaultWorkJson(Map<String, String> ui) {
        if (ensureWorkJsonFromRepoIfMissing(ui)) {
            return Optional.of(workJsonPath(ui).toAbsolutePath().normalize());
        }
        Path bundled = bundledTemplatePath(ui);
        if (Files.isRegularFile(bundled)) {
            return Optional.of(bundled.toAbsolutePath().normalize());
        }
        return Optional.empty();
    }
}
