package jp.co.pm.ai.desktop.bridge;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.PortableBundleSelfUpdater;

/**
 * 段階1/2・プローブ用 Python 実行ファイルの解決。{@link jp.co.pm.ai.desktop.MainShellController} とヘッドレス CLI が同じ
 * 既定順（同梱 embed → バンドル相対 → OS の {@code python}/{@code python3}）を参照する。
 */
public final class StagePythonExecutable {

    private StagePythonExecutable() {}

    /**
     * 優先順: {@link AppPaths#KEY_PM_AI_PYTHON}（正規化後）→ 同梱 embed / バンドル既定 → {@link #defaultPythonCommandName()}.
     */
    public static Path resolve(Map<String, String> ui) {
        String raw = ui != null ? ui.get(AppPaths.KEY_PM_AI_PYTHON) : null;
        String normalized = AppPaths.normalizePmAiPythonExecutable(raw);
        if (normalized != null && !normalized.isBlank()) {
            return Path.of(normalized.trim());
        }
        return Path.of(defaultOsPythonString());
    }

    /** {@link #resolve(Map)} で {@code PM_AI_PYTHON} が空のときと同じパス（シェル未結線時の退避）。 */
    public static Path defaultPythonPathWhenShellMissing() {
        return Path.of(defaultOsPythonString());
    }

    /**
     * 環境変数タブのブートストラップ既定用 {@code PM_AI_PYTHON} 文字列（相対パスを含み得る）。
     *
     * @see jp.co.pm.ai.desktop.MainShellController#bootstrapDefaultValueForKey
     */
    public static String defaultPmAiPythonForBootstrap() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Optional<Path> portableEmbed = AppPaths.findPortablePythonEmbedExecutable(cwd);
        if (portableEmbed.isPresent()) {
            return portableEmbed.get().toString();
        }
        if (PortableBundleSelfUpdater.isPortableBundleLayout(cwd)) {
            return Path.of("pm-ai-data", "runtime", "python-embed", "python.exe").toString();
        }
        return defaultOsPythonString();
    }

    private static String defaultOsPythonString() {
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Optional<Path> portableEmbed = AppPaths.findPortablePythonEmbedExecutable(cwd);
        if (portableEmbed.isPresent()) {
            return portableEmbed.get().toString();
        }
        Path bundledWin =
                cwd.resolve("pm-ai-data")
                        .resolve("runtime")
                        .resolve("python-embed")
                        .resolve("python.exe");
        if (Files.isRegularFile(bundledWin)) {
            return bundledWin.toAbsolutePath().normalize().toString();
        }
        return defaultPythonCommandName();
    }

    private static String defaultPythonCommandName() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win")
                ? "python"
                : "python3";
    }
}
