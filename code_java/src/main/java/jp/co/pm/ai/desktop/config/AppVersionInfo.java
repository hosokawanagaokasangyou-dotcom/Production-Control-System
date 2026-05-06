package jp.co.pm.ai.desktop.config;

import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

/** Resolves {@link AppPaths#VERSION_TXT_FILE_NAME} for the run tab label. */
public final class AppVersionInfo {

    private AppVersionInfo() {}

    /**
     * Version string for the run/log tab. Portable: {@code pm-ai-data/version.txt}; dev: repo root
     * {@code version.txt}.
     */
    public static String resolveDisplayedVersion(Path cwd, Map<String, String> ui) {
        Path c = cwd.toAbsolutePath().normalize();
        Path portableMarker =
                c.resolve("pm-ai-data")
                        .resolve("code")
                        .resolve("python")
                        .resolve("task_extract_stage1.py");
        try {
            if (Files.isRegularFile(portableMarker)) {
                Path vf = c.resolve("pm-ai-data").resolve(AppPaths.VERSION_TXT_FILE_NAME);
                return PortableBundleSelfUpdater.parseVersionFile(vf)
                        .map(BigDecimal::toPlainString)
                        .orElse("•s–¾");
            }
            Path repo = AppPaths.resolveRepoRoot(ui != null ? ui : Map.of());
            Path vf = repo.resolve(AppPaths.VERSION_TXT_FILE_NAME);
            return PortableBundleSelfUpdater.parseVersionFile(vf)
                    .map(BigDecimal::toPlainString)
                    .orElse("•s–¾");
        } catch (Exception e) {
            return "•s–¾";
        }
    }
}
