package jp.co.pm.ai.desktop.config;

import java.nio.file.Path;
import java.util.Map;

/**
 * Paths under repository {@code init_setting/} for package-time UI defaults (session + table columns).
 */
public final class InitSettingPaths {

    /** Snapshot compatible with session-state JSON keys (written by export action). */
    public static final String SESSION_DEFAULTS_FILE = "session_defaults.json";

    /** Same shape as ~/.pm-ai-desktop/table-column-order.json */
    public static final String TABLE_COLUMN_DEFAULTS_FILE = "table_column_defaults.json";

    private InitSettingPaths() {}

    /** {@code <repo>/init_setting} */
    public static Path resolveRepoInitSettingDir(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui).resolve("init_setting");
    }

    /** {@code <cwd>/pm-ai-data/init_setting} next to portable exe. */
    public static Path portableBundleInitSettingDir() {
        return Path.of(System.getProperty("user.dir", "."))
                .toAbsolutePath()
                .normalize()
                .resolve("pm-ai-data")
                .resolve("init_setting");
    }

    /** {@code <cwd>/init_setting} when running from repository root. */
    public static Path cwdInitSettingDir() {
        return Path.of(System.getProperty("user.dir", "."))
                .toAbsolutePath()
                .normalize()
                .resolve("init_setting");
    }
}
