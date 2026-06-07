package jp.co.pm.ai.desktop.io;

import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;

/** .rdp に書き込む接続先リモート起動プログラム設定（環境変数タブと UI から解決）。 */
public final class RdpCompanionLauncher {

    /** 接続先サーバー上の起動プログラム（.rdp の alternate shell）。 */
    public static final String KEY_PM_AI_RDP_COMPANION_PROGRAM = "PM_AI_RDP_COMPANION_PROGRAM";

    /** 接続先プログラムへ渡す引数（alternate shell の引数）。 */
    public static final String KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS = "PM_AI_RDP_COMPANION_PROGRAM_ARGS";

    private RdpCompanionLauncher() {}

    public static Optional<String> resolveRemoteProgramPath(Map<String, String> ui) {
        String raw = trimFromUi(ui, KEY_PM_AI_RDP_COMPANION_PROGRAM);
        return raw.isEmpty() ? Optional.empty() : Optional.of(raw);
    }

    public static String resolveRemoteProgramArgs(Map<String, String> ui) {
        return trimFromUi(ui, KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS);
    }

    public static Optional<String> formatEmbeddedSummary(Map<String, String> ui) {
        if (!isEmbedStartupInProfileEnabled(ui)) {
            return Optional.empty();
        }
        Optional<String> program = resolveRemoteProgramPath(ui);
        if (program.isEmpty()) {
            return Optional.empty();
        }
        String args = resolveRemoteProgramArgs(ui);
        if (args.isBlank()) {
            return program;
        }
        return Optional.of(program.get() + " " + args);
    }

    /** {@link AppPaths#KEY_PM_AI_RDP_EMBED_STARTUP_IN_PROFILE} が有効なときのみ .rdp へ組込。 */
    public static boolean isEmbedStartupInProfileEnabled(Map<String, String> ui) {
        String raw = trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_EMBED_STARTUP_IN_PROFILE);
        if (raw.isEmpty()) {
            return false;
        }
        String v = raw.toLowerCase(Locale.ROOT);
        return "1".equals(v) || "true".equals(v) || "on".equals(v) || "yes".equals(v);
    }

    private static String trimFromUi(Map<String, String> ui, String key) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(key);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(key);
        }
        return raw != null ? raw.trim() : "";
    }
}
