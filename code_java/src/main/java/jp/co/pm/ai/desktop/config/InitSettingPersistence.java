package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/** Writes {@link InitSettingPaths#resolveRepoInitSettingDir(Map)} for package defaults export. */
public final class InitSettingPersistence {

    private static final ObjectMapper JSON = new ObjectMapper();

    private InitSettingPersistence() {}

    /**
     * Saves session_defaults.json and table_column_defaults.json under repository {@code init_setting/}.
     */
    public static void savePackageDefaults(Map<String, String> ui, DesktopSessionState state)
            throws IOException {
        if (state == null) {
            return;
        }
        Path dir = InitSettingPaths.resolveRepoInitSettingDir(ui);
        Files.createDirectories(dir);
        Path sessionDest = dir.resolve(InitSettingPaths.SESSION_DEFAULTS_FILE);
        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(sessionDest.toFile(), DesktopSessionStateStore.toJsonObject(state));

        Path tableDest = dir.resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
        JsonNode merged = TableColumnOrderPersistence.mergedTableColumnDefaultsRootForExport();
        if (merged != null && merged.isObject()) {
            JSON.writerWithDefaultPrettyPrinter().writeValue(tableDest.toFile(), merged);
        }
    }
}
