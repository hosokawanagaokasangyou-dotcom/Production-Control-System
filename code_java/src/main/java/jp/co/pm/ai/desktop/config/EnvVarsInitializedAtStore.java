package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Optional;

/**
 * 環境変数タブを ui_ref 既定へ初期化した日時（{@code ~/.pm-ai-desktop/env-vars-initialized-at.txt}）。
 *
 * <p>{@link jp.co.pm.ai.desktop.MainShellController#applyEnvRowsFullBundledResetAndPersist} 実行時に記録する。
 */
public final class EnvVarsInitializedAtStore {

    private static final DateTimeFormatter DISPLAY_FMT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private EnvVarsInitializedAtStore() {}

    private static Path storePath() {
        return AppPaths.resolveDesktopAppHomeDir().resolve("env-vars-initialized-at.txt");
    }

    public static Path storePathForTests() {
        return storePath();
    }

    /** 現在時刻を記録する（ISO-8601 instant 文字列）。 */
    public static void recordNow() {
        try {
            Files.createDirectories(storePath().getParent());
            Files.writeString(storePath(), Instant.now().toString(), StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }

    public static Optional<Instant> load() {
        try {
            if (!Files.isRegularFile(storePath())) {
                return Optional.empty();
            }
            String raw = Files.readString(storePath(), StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(Instant.parse(raw));
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }

    /** ツールバー表示用（未記録時は {@code —}）。 */
    public static String formatForToolbar() {
        return load().map(DISPLAY_FMT::format).orElse("—");
    }

    /** 環境変数 ui_ref 既定への初期化が一度でも記録されているか。 */
    public static boolean isRecorded() {
        return load().isPresent();
    }
}
