package jp.co.pm.ai.desktop.io;

import java.util.Map;
import java.util.Optional;

/** .rdp に書き込む接続先リモート起動プログラム設定（環境変数タブと UI から解決）。 */
public final class RdpCompanionLauncher {

    /** 接続先サーバー上の起動プログラム（.rdp の remoteapplicationprogram）。 */
    public static final String KEY_PM_AI_RDP_COMPANION_PROGRAM = "PM_AI_RDP_COMPANION_PROGRAM";

    /** 接続先プログラムへ渡す引数（.rdp の remoteapplicationcmdline）。 */
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

    private static String trimFromUi(Map<String, String> ui, String key) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(key);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(key);
        }
        return raw != null ? raw.trim() : "";
    }
}
