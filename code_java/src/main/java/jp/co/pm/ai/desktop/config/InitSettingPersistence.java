package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Writes {@link InitSettingPaths#resolveRepoInitSettingDir(Map)} for package defaults export（工場別ファイル名は
 * {@link InitSettingPaths#sessionDefaultsFileForFactory}／{@link InitSettingPaths#tableColumnDefaultsFileForFactory}）。
 */
public final class InitSettingPersistence {

    private static final ObjectMapper JSON = new ObjectMapper();

    private InitSettingPersistence() {}

    /**
     * {@link GlobalInitSettingTarget} の選択工場向けに、{@code session_defaults_<工場>.json} と {@code
     * table_column_defaults_<工場>.json} をリポジトリ {@code init_setting/} に保存する。
     */
    public static void savePackageDefaults(Map<String, String> ui, DesktopSessionState state)
            throws IOException {
        savePackageDefaults(ui, state, GlobalInitSettingTarget.load());
    }

    /**
     * @param initSettingTarget 書き出し先ファイル名に使う工場（null のとき湖南）
     */
    public static void savePackageDefaults(
            Map<String, String> ui, DesktopSessionState state, FactorySite initSettingTarget)
            throws IOException {
        if (state == null) {
            return;
        }
        FactorySite t = initSettingTarget != null ? initSettingTarget : FactorySite.KONAN;
        Path dir = InitSettingPaths.resolveRepoInitSettingDir(ui);
        Files.createDirectories(dir);
        Path sessionDest = dir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(t));
        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(sessionDest.toFile(), DesktopSessionStateStore.toJsonObject(state));

        Path tableDest = dir.resolve(InitSettingPaths.tableColumnDefaultsFileForFactory(t));
        JsonNode merged = TableColumnOrderPersistence.mergedTableColumnDefaultsRootForExport();
        if (merged != null && merged.isObject()) {
            JSON.writerWithDefaultPrettyPrinter().writeValue(tableDest.toFile(), merged);
        }
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
        if (pmAiDataRoot == null) {
            return;
        }
        Path srcDir = pmAiDataRoot.resolve("init_setting");
        if (!Files.isDirectory(srcDir)) {
            return;
        }
        Path dstDir = InitSettingPaths.resolveRepoInitSettingDir(ui);
        Files.createDirectories(dstDir);
        copyIfRegularFile(srcDir, dstDir, InitSettingPaths.SESSION_DEFAULTS_FILE);
        copyIfRegularFile(srcDir, dstDir, InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
        for (FactorySite site : FactorySite.values()) {
            copyIfRegularFile(srcDir, dstDir, InitSettingPaths.sessionDefaultsFileForFactory(site));
            copyIfRegularFile(srcDir, dstDir, InitSettingPaths.tableColumnDefaultsFileForFactory(site));
        }
    }

    private static void copyIfRegularFile(Path srcDir, Path dstDir, String fileName) throws IOException {
        Path src = srcDir.resolve(fileName);
        if (!Files.isRegularFile(src)) {
            return;
        }
        Files.copy(src, dstDir.resolve(fileName), StandardCopyOption.REPLACE_EXISTING);
    }
}
