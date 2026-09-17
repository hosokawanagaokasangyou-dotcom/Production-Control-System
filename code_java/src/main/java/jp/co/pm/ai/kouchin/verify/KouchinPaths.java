package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * 検証に使う各データソースのフォルダ。
 * 既定値は {@link AppPaths} の UNC パスで、環境変数（またはUIのセッション保存値）で上書きできる。
 *
 * <p>{@code outputDir} は単一実行時のファイル名生成用（国分固定ルート）。実際の二重保存は
 * {@link KouchinOutputDirs#resolveAll} が行う。
 *
 * <p>①東レCSVは国分・湖南でフォルダが分かれる。取り込みは現工場、検証は現工場優先で他工場へフォールバックする。
 */
public record KouchinPaths(
        Path torayCsvDir,
        Path konanTorayCsvDir,
        Path kokubuNagaokaDir,
        Path kokubuAladdinDir,
        Path konanShisanDir,
        Path konanAladdinDir,
        Path konanMonthlyDir,
        Path outputDir,
        Path judgmentDir,
        FactorySite factorySite) {

    public KouchinPaths(
            Path torayCsvDir,
            Path kokubuNagaokaDir,
            Path kokubuAladdinDir,
            Path konanShisanDir,
            Path konanAladdinDir,
            Path konanMonthlyDir,
            Path outputDir,
            Path judgmentDir) {
        this(
                torayCsvDir,
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_TORAY_CSV_DIR),
                kokubuNagaokaDir,
                kokubuAladdinDir,
                konanShisanDir,
                konanAladdinDir,
                konanMonthlyDir,
                outputDir,
                judgmentDir,
                null);
    }

    public static KouchinPaths defaults() {
        return defaults(null);
    }

    public static KouchinPaths defaults(FactorySite site) {
        return new KouchinPaths(
                Paths.get(AppPaths.DEFAULT_KOUCHIN_TORAY_CSV_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_TORAY_CSV_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KOKUBU_NAGAOKA_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KOKUBU_ALADDIN_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_SHISAN_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_ALADDIN_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_KONAN_MONTHLY_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_BASE_DIR),
                Paths.get(AppPaths.DEFAULT_KOUCHIN_BASE_DIR),
                site);
    }

    /**
     * 環境変数マップから生成する。未設定・空文字のキーは既定値を使う。
     */
    public static KouchinPaths fromEnv(Map<String, String> env) {
        FactorySite site = AppPaths.currentDispatchFactorySite(env == null ? Map.of() : env);
        KouchinPaths d = defaults(site);
        if (env == null || env.isEmpty()) {
            return d;
        }
        return new KouchinPaths(
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, d.torayCsvDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, d.konanTorayCsvDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_NAGAOKA_DIR, d.kokubuNagaokaDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, d.kokubuAladdinDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KONAN_SHISAN_DIR, d.konanShisanDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR, d.konanAladdinDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_KONAN_MONTHLY_DIR, d.konanMonthlyDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR, d.outputDir),
                pick(env, AppPaths.KEY_PM_AI_KOUCHIN_JUDGMENT_DIR, d.judgmentDir),
                site);
    }

    private static Path pick(Map<String, String> env, String key, Path fallback) {
        String value = env.get(key);
        return value == null || value.isBlank() ? fallback : Paths.get(value.trim());
    }

    public Path source1Dir(FactoryId factory) {
        return factory == FactoryId.KONAN ? konanTorayCsvDir : torayCsvDir;
    }

    public Path importTorayCsvDir() {
        return factorySite == FactorySite.KONAN ? konanTorayCsvDir : torayCsvDir;
    }

    /** 検証用①。現工場フォルダにCSVがあればそれを使い、無ければ他工場へフォールバックする。 */
    public Path resolveTorayCsvDir() {
        Path preferred = importTorayCsvDir();
        if (hasTorayCsv(preferred)) {
            return preferred;
        }
        Path other = factorySite == FactorySite.KONAN ? torayCsvDir : konanTorayCsvDir;
        if (hasTorayCsv(other)) {
            return other;
        }
        return preferred;
    }

    static boolean hasTorayCsv(Path dir) {
        if (dir == null || !Files.isDirectory(dir)) {
            return false;
        }
        try {
            FileDiscovery.findTorayCsv(dir);
            return true;
        } catch (RuntimeException e) {
            return false;
        }
    }

    public Path source2Dir(FactoryId factory) {
        return factory == FactoryId.KOKUBU ? kokubuNagaokaDir : konanShisanDir;
    }

    public Path source3Dir(FactoryId factory) {
        return factory == FactoryId.KOKUBU ? kokubuAladdinDir : konanAladdinDir;
    }

    public Path monthlyDir(FactoryId factory) {
        return factory == FactoryId.KOKUBU ? null : konanMonthlyDir;
    }
}
