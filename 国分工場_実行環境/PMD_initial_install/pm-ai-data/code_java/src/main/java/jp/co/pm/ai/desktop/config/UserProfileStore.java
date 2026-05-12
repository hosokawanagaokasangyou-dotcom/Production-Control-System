package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * ユーザープロファイル（{@code ~/.pm-ai-desktop/user-profiles/*.json}）。アップデート同期の対象外で、配布の上書きを受けない。
 * 初回インストール／バージョンアップ用 ZIP のワークスペースミラーでも {@code user-profiles} ディレクトリは同梱しない（{@code package_workspace_copy.ps1} / {@code fast_package_app.ps1}）。
 *
 * <p><strong>JSON エンベロープ（トップレベル）</strong>
 *
 * <ul>
 *   <li>{@code schemaVersion} — プロファイルファイル形式の版（現在は 1）。{@link #SUPPORTED_PROFILE_SCHEMA_MAX} を超えるファイルはより新しいアプリ向けとして拒否する。
 *   <li>{@code profileExportFormat} — 任意。エクスポート時に {@link #PROFILE_EXPORT_FORMAT} を書き込む（将来の包み拡張のための予約）。
 *   <li>{@code session} — {@link DesktopSessionStateStore#toJsonObject} 形状のオブジェクト（欠落時は空オブジェクトとして復元）。
 *   <li>{@code tableColumnOrder} — 列順ストアのオブジェクト（欠落時は空オブジェクト）。
 *   <li>{@code displayName}, {@code savedAt} — 一覧表示用。
 *   <li>{@code exportedAt}, {@code sourceAppVersion} — ファイルへエクスポートしたときのみ付与（デバッグ・移行の目安）。
 * </ul>
 *
 * <p>セッション JSON 内の未知キー・欠落キーは {@link DesktopSessionStateStore#desktopSessionFromStoredJson} の解釈に従う。
 */
public final class UserProfileStore {

    public static final int MAX_PROFILES = 100;

    /** このビルドが書き出す {@code schemaVersion}（プロファイル包み）。 */
    public static final int CURRENT_PROFILE_SCHEMA_VERSION = 1;

    /**
     * このビルドが読み込める {@code schemaVersion} の上限。それより大きいファイルは「アプリを更新してください」とする。
     */
    public static final int SUPPORTED_PROFILE_SCHEMA_MAX = 1;

    /** エクスポート JSON の {@code profileExportFormat} に書く値（包みの将来拡張用）。 */
    public static final int PROFILE_EXPORT_FORMAT = 1;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path DIRECTORY =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "user-profiles");

    /** 名前省略時の表示・{@code displayName} 保存に使う日時（秒まで）。 */
    private static final DateTimeFormatter STAMP =
            DateTimeFormatter.ofPattern("yyyyMMddHHmmss").withLocale(Locale.JAPAN);

    private UserProfileStore() {}

    public static Path directory() {
        return DIRECTORY;
    }

    /** 一覧表示用（逆順＝新しいファイル名が先頭になりやすい）。 */
    public record ListedProfile(Path path, String displayLabel, String savedAtIso) {}

    public static List<ListedProfile> listProfiles() throws IOException {
        if (!Files.isDirectory(DIRECTORY)) {
            return List.of();
        }
        List<ListedProfile> out = new ArrayList<>();
        try (Stream<Path> stream = Files.list(DIRECTORY)) {
            stream.filter(p -> Files.isRegularFile(p) && p.getFileName().toString().endsWith(".json"))
                    .sorted(Comparator.comparing(p -> p.getFileName().toString(), Comparator.reverseOrder()))
                    .forEach(
                            p -> {
                                try {
                                    JsonNode root = JSON.readTree(p.toFile());
                                    if (root == null || !root.isObject()) {
                                        return;
                                    }
                                    String dn = root.path("displayName").asText("").strip();
                                    String sa = root.path("savedAt").asText("").strip();
                                    String label =
                                            dn.isEmpty()
                                                    ? (sa.isEmpty() ? p.getFileName().toString() : sa)
                                                    : dn;
                                    out.add(new ListedProfile(p, label, sa));
                                } catch (IOException ignored) {
                                }
                            });
        }
        return List.copyOf(out);
    }

    public static int countProfiles() throws IOException {
        if (!Files.isDirectory(DIRECTORY)) {
            return 0;
        }
        try (Stream<Path> stream = Files.list(DIRECTORY)) {
            return (int)
                    stream.filter(p -> Files.isRegularFile(p) && p.getFileName().toString().endsWith(".json"))
                            .count();
        }
    }

    /**
     * @param optionalDisplayName 空なら {@link #STAMP} に相当する文字列を名前として保存
     * @return 書き込んだファイルパス
     */
    public static Path saveProfile(String optionalDisplayName, ObjectNode sessionJson, JsonNode tableColumnRoot)
            throws IOException {
        if (countProfiles() >= MAX_PROFILES) {
            throw new IllegalStateException("プロファイルは最大 " + MAX_PROFILES + " 個までです。");
        }
        Files.createDirectories(DIRECTORY);
        String stamp = STAMP.format(LocalDateTime.now());
        String displayStored =
                optionalDisplayName != null && !optionalDisplayName.isBlank()
                        ? optionalDisplayName.strip()
                        : stamp;
        String base =
                optionalDisplayName != null && !optionalDisplayName.isBlank()
                        ? stamp + "_" + sanitizeFileSegment(optionalDisplayName.strip())
                        : stamp;
        Path dest = uniqueDestination(base);
        ObjectNode root = buildProfileRootJson(displayStored, LocalDateTime.now().toString(), sessionJson, tableColumnRoot);
        JSON.writerWithDefaultPrettyPrinter().writeValue(dest.toFile(), root);
        return dest;
    }

    /**
     * 任意パスへプロファイル JSON を書き出す（チーム共有・バックアップ用）。{@code exportedAt} と {@code sourceAppVersion} を付与する。
     *
     * @param destination 保存先ファイル（親ディレクトリが無ければ作成する）
     * @param displayName 表示名（任意）
     * @param savedAtIso 元の {@code savedAt} を維持したいときに渡す。空なら現在時刻
     * @param sourceAppVersion {@link AppVersionInfo#resolveDisplayedVersion} 等（任意）
     */
    public static void exportProfileToFile(
            Path destination,
            String displayName,
            String savedAtIso,
            ObjectNode sessionJson,
            JsonNode tableColumnRoot,
            String sourceAppVersion)
            throws IOException {
        String saved = savedAtIso != null && !savedAtIso.isBlank() ? savedAtIso.strip() : LocalDateTime.now().toString();
        String dn = displayName != null ? displayName.strip() : "";
        ObjectNode root = buildProfileRootJson(dn, saved, sessionJson, tableColumnRoot);
        root.put("profileExportFormat", PROFILE_EXPORT_FORMAT);
        root.put("exportedAt", LocalDateTime.now().toString());
        root.put("sourceAppVersion", sourceAppVersion != null ? sourceAppVersion : "");
        Path parent = destination.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(destination.toFile(), root);
    }

    private static ObjectNode buildProfileRootJson(
            String displayName, String savedAtIso, ObjectNode sessionJson, JsonNode tableColumnRoot) {
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", CURRENT_PROFILE_SCHEMA_VERSION);
        root.put("savedAt", savedAtIso);
        root.put("displayName", displayName != null ? displayName : "");
        root.set("session", sessionJson != null ? sessionJson.deepCopy() : JSON.createObjectNode());
        root.set(
                "tableColumnOrder",
                tableColumnRoot != null && tableColumnRoot.isObject()
                        ? tableColumnRoot.deepCopy()
                        : JSON.createObjectNode());
        return root;
    }

    private static Path uniqueDestination(String base) {
        Path p = DIRECTORY.resolve(base + ".json");
        if (!Files.isRegularFile(p)) {
            return p;
        }
        for (int i = 2; i < 10_000; i++) {
            Path q = DIRECTORY.resolve(base + "_" + i + ".json");
            if (!Files.isRegularFile(q)) {
                return q;
            }
        }
        throw new IllegalStateException("ユニークなプロファイル名を生成できませんでした。");
    }

    private static String sanitizeFileSegment(String s) {
        String t = s.replaceAll("[^a-zA-Z0-9._\\-\\u3040-\\u30ff\\u3400-\\u9fff]", "_");
        if (t.length() > 80) {
            t = t.substring(0, 80);
        }
        return t.isEmpty() ? "profile" : t;
    }

    public record LoadedProfile(String displayName, String savedAt, DesktopSessionState session, JsonNode tableColumnOrder) {}

    public static LoadedProfile loadProfile(Path file) throws IOException {
        JsonNode root = JSON.readTree(file.toFile());
        return loadProfileFromJson(root);
    }

    /** クラスパスやストリームから読んだ JSON 用。 */
    public static LoadedProfile loadProfileFromJson(JsonNode root) throws IOException {
        if (root == null || !root.isObject()) {
            throw new IOException("プロファイルが読めません（オブジェクトではありません）。");
        }
        ObjectNode obj = (ObjectNode) root;
        migrateProfileEnvelope(obj);
        String displayName = obj.path("displayName").asText("");
        String savedAt = obj.path("savedAt").asText("");
        JsonNode sess = obj.get("session");
        if (sess != null && !sess.isObject()) {
            throw new IOException("プロファイルの session がオブジェクトではありません。");
        }
        DesktopSessionState state =
                DesktopSessionStateStore.desktopSessionFromStoredJson(
                        sess != null && sess.isObject() ? sess : JSON.createObjectNode());
        JsonNode tables = obj.get("tableColumnOrder");
        if (tables == null || !tables.isObject()) {
            tables = JSON.createObjectNode();
        }
        return new LoadedProfile(displayName, savedAt, state, tables);
    }

    /**
     * {@code schemaVersion} を現在の版へ段階マイグレーションし、将来の互換用フックを置く。
     *
     * @throws IOException 未対応のより新しい {@code schemaVersion} のとき
     */
    private static void migrateProfileEnvelope(ObjectNode root) throws IOException {
        int ver = root.path("schemaVersion").asInt(0);
        if (ver == 0) {
            root.put("schemaVersion", 1);
            ver = 1;
        }
        if (ver > SUPPORTED_PROFILE_SCHEMA_MAX) {
            throw new IOException(
                    "このファイルはより新しいバージョンのアプリで作成されています。アプリを更新してください。(schemaVersion="
                            + ver
                            + ")");
        }
        while (ver < CURRENT_PROFILE_SCHEMA_VERSION) {
            applyProfileEnvelopeMigration(root, ver, ver + 1);
            ver++;
            root.put("schemaVersion", ver);
        }
    }

    /** 将来 {@code schemaVersion} を上げたときに段階変換を追加する。 */
    private static void applyProfileEnvelopeMigration(ObjectNode root, int fromVer, int toVer) {
        if (fromVer == toVer) {
            return;
        }
        // 現状 from→to は定義なし（CURRENT==SUPPORTED==1）
    }

    public static void deleteProfile(Path file) throws IOException {
        Files.deleteIfExists(file);
    }
}
