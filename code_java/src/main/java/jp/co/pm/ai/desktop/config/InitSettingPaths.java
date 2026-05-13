package jp.co.pm.ai.desktop.config;

import java.nio.file.Path;
import java.util.Locale;
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

    /**
     * 工場別グローバル既定（湖南／国分）。{@link GlobalInitSettingTarget} の選択に対応。
     *
     * <p>ファイル名は {@link FactorySite#name()} の小文字を付与する（{@code session_defaults_konan.json} 等）。
     */
    public static String sessionDefaultsFileForFactory(FactorySite site) {
        if (site == null) {
            site = FactorySite.KONAN;
        }
        return "session_defaults_" + site.name().toLowerCase(Locale.ROOT) + ".json";
    }

    /** 工場別の列順・表設定マージ用 JSON（{@link #TABLE_COLUMN_DEFAULTS_FILE} と同形状）。 */
    public static String tableColumnDefaultsFileForFactory(FactorySite site) {
        if (site == null) {
            site = FactorySite.KONAN;
        }
        return "table_column_defaults_" + site.name().toLowerCase(Locale.ROOT) + ".json";
    }

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
