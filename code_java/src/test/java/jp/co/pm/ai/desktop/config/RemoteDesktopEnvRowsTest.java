package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import jp.co.pm.ai.desktop.EnvVarRow;

class RemoteDesktopEnvRowsTest {

    @Test
    void bootstrap_appliesRdpDefaultsWhenBlank(@TempDir Path fakeRepo) throws Exception {
        Path summary = fakeRepo.resolve("code").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(summary.getParent());
        Files.writeString(summary, "x", StandardCharsets.UTF_8);

        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        RemoteDesktopEnvRows.populateFromUiRef(rows);
        for (EnvVarRow row : rows) {
            if (AppPaths.KEY_PM_AI_REPO_ROOT.equals(row.getName().strip())) {
                row.setValue(fakeRepo.toString());
            }
            if (AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR.equals(row.getName().strip())) {
                row.setValue("");
            }
        }
        RemoteDesktopEnvRows.applyRdpLauncherEmptyDefaults(rows);

        Map<String, String> ui = RemoteDesktopEnvRows.collectMap(rows);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_RELEASE_DIR,
                ui.get(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR));
        assertTrue(ui.containsKey(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR));
    }

    @Test
    void syncRowValue_updatesExistingRow() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow row = new EnvVarRow();
        row.setName(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR);
        row.setValue("old");
        rows.add(row);

        RemoteDesktopEnvRows.syncRowValue(rows, AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR, "new");
        assertEquals("new", rows.get(0).getValue());
    }
}
