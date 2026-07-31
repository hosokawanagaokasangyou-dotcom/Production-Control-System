package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.HexFormat;
import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;
import java.util.function.Predicate;

/**
 * 環境変数タブを ui_ref 既定へ初期化した日時（{@code ~/.pm-ai-desktop/env-vars-initialized-at.txt}）。
 *
 * <p>{@link jp.co.pm.ai.desktop.MainShellController#applyEnvRowsFullBundledResetAndPersist} 実行時に記録する。
 *
 * <p>起動時照合はテンプレート再構築ではなく {@link #recordEnvFingerprint(Map, Predicate)} で保存した
 * SHA-256 フィンガープリントを正とする（配台ルックアップ表パス等の表示補完差を避ける）。
 */
public final class EnvVarsInitializedAtStore {

    private static final DateTimeFormatter DISPLAY_FMT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private EnvVarsInitializedAtStore() {}

    private static Path storePath() {
        return AppPaths.resolveDesktopAppHomeDir().resolve("env-vars-initialized-at.txt");
    }

    private static Path fingerprintPath() {
        return AppPaths.resolveDesktopAppHomeDir().resolve("env-vars-initialized-fingerprint.sha256");
    }

    private static Path baselineCanonicalPath() {
        return AppPaths.resolveDesktopAppHomeDir().resolve("env-vars-initialized-baseline.txt");
    }

    public static Path storePathForTests() {
        return storePath();
    }

    public static Path fingerprintPathForTests() {
        return fingerprintPath();
    }

    /** 現在時刻を記録する（ISO-8601 instant 文字列）。 */
    public static void recordNow() {
        try {
            Files.createDirectories(storePath().getParent());
            Files.writeString(storePath(), Instant.now().toString(), StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }

    /**
     * 初期化完了直後の環境変数タブ値のフィンガープリント（キー昇順 {@code key=value} 行の SHA-256）を保存する。
     */
    public static void recordEnvFingerprint(Map<String, String> env, Predicate<String> includeKey) {
        try {
            CanonicalEnvBody canonical = canonicalEnvBody(env, includeKey);
            if (canonical.digest().isEmpty()) {
                return;
            }
            Files.createDirectories(fingerprintPath().getParent());
            Files.writeString(fingerprintPath(), canonical.digest(), StandardCharsets.UTF_8);
            Files.writeString(baselineCanonicalPath(), canonical.body(), StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }

    static Map<String, String> loadRecordedBaselineEnv(Predicate<String> includeKey) {
        try {
            if (!Files.isRegularFile(baselineCanonicalPath())) {
                return Map.of();
            }
            String raw = Files.readString(baselineCanonicalPath(), StandardCharsets.UTF_8);
            if (raw.isBlank()) {
                return Map.of();
            }
            TreeMap<String, String> map = new TreeMap<>();
            for (String line : raw.split("\n")) {
                if (line.isBlank()) {
                    continue;
                }
                int eq = line.indexOf('=');
                if (eq <= 0) {
                    continue;
                }
                String k = line.substring(0, eq).trim();
                String v = line.substring(eq + 1);
                if (!k.isEmpty() && (includeKey == null || includeKey.test(k))) {
                    map.put(k, v);
                }
            }
            return map;
        } catch (Exception ignored) {
            return Map.of();
        }
    }

    /**
     * 記録済みベースラインと比較し、{@code includeKey} に含まれるキーだけが一致するか（RDP 実行時設定の差分は無視する移行判定用）。
     */
    public static boolean matchesRecordedBaselineForKeys(
            Map<String, String> current, Predicate<String> includeKey) {
        if (current == null || includeKey == null) {
            return false;
        }
        Map<String, String> baseline = loadRecordedBaselineEnv(k -> true);
        if (baseline.isEmpty()) {
            return false;
        }
        TreeMap<String, String> cur = canonicalEnvEntries(current, includeKey);
        for (Map.Entry<String, String> e : cur.entrySet()) {
            String k = e.getKey();
            String cv = e.getValue() != null ? e.getValue() : "";
            String bv = baseline.getOrDefault(k, "");
            if (!cv.equals(bv)) {
                return false;
            }
        }
        return true;
    }

    public static Optional<String> loadEnvFingerprint() {
        try {
            if (!Files.isRegularFile(fingerprintPath())) {
                return Optional.empty();
            }
            String raw = Files.readString(fingerprintPath(), StandardCharsets.UTF_8).trim();
            return raw.isEmpty() ? Optional.empty() : Optional.of(raw);
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }

    public static boolean envFingerprintMatches(Map<String, String> env, Predicate<String> includeKey) {
        Optional<String> stored = loadEnvFingerprint();
        if (stored.isEmpty()) {
            return false;
        }
        return stored.get().equals(canonicalEnvBody(env, includeKey).digest());
    }

    private static TreeMap<String, String> canonicalEnvEntries(
            Map<String, String> env, Predicate<String> includeKey) {
        Predicate<String> keyFilter = includeKey != null ? includeKey : k -> true;
        TreeMap<String, String> canonical = new TreeMap<>();
        if (env == null) {
            return canonical;
        }
        for (Map.Entry<String, String> e : env.entrySet()) {
            String k = e.getKey() != null ? e.getKey().trim() : "";
            if (k.isEmpty() || !keyFilter.test(k)) {
                continue;
            }
            canonical.put(k, e.getValue() != null ? e.getValue().trim() : "");
        }
        return canonical;
    }

    private static CanonicalEnvBody canonicalEnvBody(Map<String, String> env, Predicate<String> includeKey) {
        TreeMap<String, String> canonical = canonicalEnvEntries(env, includeKey);
        if (canonical.isEmpty()) {
            return new CanonicalEnvBody("", "");
        }
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, String> e : canonical.entrySet()) {
            sb.append(e.getKey()).append('=').append(e.getValue()).append('\n');
        }
        String body = sb.toString();
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] hash = md.digest(body.getBytes(StandardCharsets.UTF_8));
            return new CanonicalEnvBody(body, HexFormat.of().formatHex(hash));
        } catch (Exception ignored) {
            return new CanonicalEnvBody(body, "");
        }
    }

    private record CanonicalEnvBody(String body, String digest) {}

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
