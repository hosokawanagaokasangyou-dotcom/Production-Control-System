package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Writes {@link InitSettingPaths#resolveRepoInitSettingDir(Map)} for package defaults export（工場別ファイル名は
 * {@link InitSettingPaths#sessionDefaultsFileForFactory}／{@link InitSettingPaths#tableColumnDefaultsFileForFactory}）。
 */
public final class InitSettingPersistence {

    private static final ObjectMapper JSON = new ObjectMapper();

    private InitSettingPersistence() {}

    /**
     * {@link GlobalInitSettingTarget} の選択工場向けに、{@code session_defaults_<工場>.json}、{@code
     * table_column_defaults_<工場>.json}、{@code juchu_header_aliases_<工場>.json} をリポジトリ {@code init_setting/} に保存する。
     */
    public static void savePackageDefaults(Map<String, String> ui, DesktopSessionState state)
            throws IOException {
        savePackageDefaults(ui, state, GlobalInitSettingTarget.load(), null);
    }

    /**
     * @param initSettingTarget 書き出し先ファイル名に使う工場（null のとき湖南）
     * @param juchuHeaderAliasRegistry 非 null のとき列定義ウィザード設定をその内容で書き出す
     */
    public static void savePackageDefaults(
            Map<String, String> ui,
            DesktopSessionState state,
            FactorySite initSettingTarget,
            JuchuHeaderAliasRegistry juchuHeaderAliasRegistry)
            throws IOException {
        if (state == null) {
            return;
        }
        FactorySite t = initSettingTarget != null ? initSettingTarget : FactorySite.KONAN;
        Path dir = InitSettingPaths.resolveRepoInitSettingDir(ui);
        Files.createDirectories(dir);
        Path sessionDest = dir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(t));
        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(
                        sessionDest.toFile(),
                        DesktopSessionStateStore.toJsonObjectForGlobalInitSetting(state));

        Path tableDest = dir.resolve(InitSettingPaths.tableColumnDefaultsFileForFactory(t));
        JsonNode merged = TableColumnOrderPersistence.mergedTableColumnDefaultsRootForExport();
        if (merged != null && merged.isObject()) {
            JSON.writerWithDefaultPrettyPrinter().writeValue(tableDest.toFile(), merged);
        }

        Path juchuDest = dir.resolve(InitSettingPaths.juchuHeaderAliasesFileForFactory(t));
        JuchuHeaderAliasRegistry registry =
                juchuHeaderAliasRegistry != null
                        ? juchuHeaderAliasRegistry
                        : JuchuHeaderAliasRegistry.loadForFactory(t, ui != null ? ui : Map.of());
        registry.exportToJsonFile(juchuDest);
    }

    /**
     * ポータル自動バージョンアップで正本→{@code pm-ai-data} 同期のあと、バンドル由来の
     * {@code pm-ai-data/init_setting} をリポジトリ {@code init_setting/} へ上書きコピーする。
     *
     * <p>{@link DesktopSessionStateStore#applyPortableUpgradeBundledPolicyToSessionStore(Map)} が
     * {@link InitSettingPaths#resolveRepoInitSettingDir(Map)} をマージ最終層に含められるようにする。
     *
     * @param pmAiDataRoot 実行ディレクトリ直下の {@code pm-ai-data}（同期済み）
     */
    public static void applyPortableUpgradeOverwriteFromPmAiData(Path pmAiDataRoot, Map<String, String> ui)
            throws IOException {
        String configuredRepoRoot = ui != null ? ui.get(AppPaths.KEY_PM_AI_REPO_ROOT) : null;
        Path destinationRepoRoot =
                configuredRepoRoot != null && !configuredRepoRoot.isBlank()
                        ? Path.of(configuredRepoRoot)
                        : null;
        applyPortableUpgradeOverwriteFromPmAiData(pmAiDataRoot, destinationRepoRoot);
    }

    /**
     * ポータブル VU 用に、コピー元とコピー先のリポジトリ根を明示して既定設定を反映する。
     * {@code sourceRepoRoot} と {@code destinationRepoRoot} が同一の場合はコピーを省略する。
     */
    public static void applyPortableUpgradeOverwriteFromPmAiData(
            Path pmAiDataRoot, Path destinationRepoRoot) throws IOException {
        if (pmAiDataRoot == null) {
            return;
        }
        Path srcDir = pmAiDataRoot.resolve("init_setting");
        if (!Files.isDirectory(srcDir)) {
            return;
        }
        Path dstDir =
                destinationRepoRoot != null
                        ? destinationRepoRoot.toAbsolutePath().normalize().resolve("init_setting")
                        : InitSettingPaths.resolveRepoInitSettingDir(Map.of());
        Files.createDirectories(dstDir);
        copyIfRegularFile(srcDir, dstDir, InitSettingPaths.SESSION_DEFAULTS_FILE);
        copyIfRegularFile(srcDir, dstDir, InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
        for (FactorySite site : FactorySite.values()) {
            copyIfRegularFile(srcDir, dstDir, InitSettingPaths.sessionDefaultsFileForFactory(site));
            copyIfRegularFile(srcDir, dstDir, InitSettingPaths.tableColumnDefaultsFileForFactory(site));
            copyIfRegularFile(srcDir, dstDir, InitSettingPaths.juchuHeaderAliasesFileForFactory(site));
        }
    }

    private static void copyIfRegularFile(Path srcDir, Path dstDir, String fileName) throws IOException {
        Path src = srcDir.resolve(fileName);
        Path dst = dstDir.resolve(fileName);
        if (!Files.isRegularFile(src) || sameFile(src, dst)) {
            return;
        }
        Files.copy(src, dst, StandardCopyOption.REPLACE_EXISTING);
    }

    private static boolean sameFile(Path left, Path right) {
        try {
            return Files.exists(right) && Files.isSameFile(left, right);
        } catch (IOException ignored) {
            return left.toAbsolutePath().normalize().equals(right.toAbsolutePath().normalize());
        }
    }
}
