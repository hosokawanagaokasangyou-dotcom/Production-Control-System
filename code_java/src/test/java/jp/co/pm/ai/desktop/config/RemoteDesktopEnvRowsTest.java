package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import jp.co.pm.ai.desktop.EnvVarRow;
import jp.co.pm.ai.desktop.io.RdpPreviewSettings;
import jp.co.pm.ai.desktop.io.RdpSecurityDialogAutomator;

class RemoteDesktopEnvRowsTest {

    @Test
    void populateFromUiRef_includesOnlyRdpRelevantKeys() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        RemoteDesktopEnvRows.populateFromUiRef(rows);

        Set<String> names =
                rows.stream()
                        .map(r -> r.getName() != null ? r.getName().strip() : "")
                        .filter(n -> !n.isEmpty())
                        .collect(Collectors.toSet());

        assertFalse(names.contains(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
        assertFalse(names.contains("GEMINI_MODEL_TRY_ORDER"));
        assertFalse(names.contains(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR));
        assertTrue(names.contains(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR));
        assertTrue(names.contains(AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR));
        assertTrue(RemoteDesktopEnvRows.relevantEnvKeys().containsAll(names));
    }

    @Test
    void applyFromSession_filtersDispatchKeys(@TempDir Path userHome) {
        String priorHome = System.getProperty("user.home");
        String priorDir = AppPaths.desktopAppHomeDirName();
        try {
            System.setProperty("user.home", userHome.toString());
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            DesktopSessionStateStore.patchUiEnvRowsAndTheme(
                    List.of(
                            new UiEnvRowSnapshot(
                                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK, "C:\\master.xlsm", ""),
                            new UiEnvRowSnapshot(
                                    AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, "\\\\share\\rpa", "")),
                    DesktopTheme.LIGHT.storedId());
            DesktopSessionState session = DesktopSessionStateStore.load();

            ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
            RemoteDesktopEnvRows.applyFromSession(rows, session);

            Set<String> names =
                    rows.stream()
                            .map(r -> r.getName() != null ? r.getName().strip() : "")
                            .collect(Collectors.toSet());
            assertEquals(Set.of(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR), names);
        } finally {
            if (priorHome != null) {
                System.setProperty("user.home", priorHome);
            } else {
                System.clearProperty("user.home");
            }
            AppPaths.setDesktopAppHomeDirName(priorDir);
        }
    }

    @Test
    void snapshot_excludesIrrelevantRows() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow dispatch = new EnvVarRow();
        dispatch.setName(AppPaths.KEY_PM_AI_OUTPUT_DIR);
        dispatch.setValue("C:\\out");
        rows.add(dispatch);
        EnvVarRow rdp = new EnvVarRow();
        rdp.setName(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR);
        rdp.setValue("\\\\share\\rpa");
        rows.add(rdp);

        List<UiEnvRowSnapshot> snap = RemoteDesktopEnvRows.snapshot(rows);
        assertEquals(1, snap.size());
        assertEquals(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, snap.get(0).name());
    }

    @Test
    void mergeMissingFromUiRef_skipsDispatchKeys() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow existing = new EnvVarRow();
        existing.setName(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR);
        existing.setValue("keep");
        rows.add(existing);

        RemoteDesktopEnvRows.mergeMissingFromUiRef(rows);

        Set<String> names =
                rows.stream()
                        .map(r -> r.getName() != null ? r.getName().strip() : "")
                        .filter(n -> !n.isEmpty())
                        .collect(Collectors.toSet());
        assertFalse(names.contains(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
        assertTrue(names.contains(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR));
        assertTrue(names.contains(AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR));
    }

    @Test
    void syncRowValue_ignoresIrrelevantKey() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        RemoteDesktopEnvRows.syncRowValue(rows, AppPaths.KEY_PM_AI_MASTER_WORKBOOK, "x");
        assertTrue(rows.isEmpty());

        RemoteDesktopEnvRows.syncRowValue(
                rows, RdpPreviewSettings.KEY_PM_AI_RDP_PREVIEW_IN_TAB, "1");
        assertEquals(1, rows.size());
        assertEquals(RdpPreviewSettings.KEY_PM_AI_RDP_PREVIEW_IN_TAB, rows.get(0).getName());
    }

    @Test
    void pruneIrrelevantRows_removesDispatchRows() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow blank = new EnvVarRow();
        rows.add(blank);
        EnvVarRow dispatch = new EnvVarRow();
        dispatch.setName("GEMINI_MODEL");
        rows.add(dispatch);
        EnvVarRow rdp = new EnvVarRow();
        rdp.setName(RdpSecurityDialogAutomator.KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG);
        rows.add(rdp);

        RemoteDesktopEnvRows.pruneIrrelevantRows(rows);

        assertEquals(2, rows.size());
        assertTrue(rows.stream().anyMatch(r -> r.getName() == null || r.getName().isBlank()));
        assertTrue(
                rows.stream()
                        .anyMatch(
                                r ->
                                        RdpSecurityDialogAutomator
                                                .KEY_PM_AI_RDP_AUTO_CONFIRM_SECURITY_DIALOG
                                                .equals(
                                                        r.getName() != null
                                                                ? r.getName().strip()
                                                                : "")));
    }

    @Test
    void bootstrap_appliesRdpDefaultsWhenBlank(@TempDir Path userHome) {
        String priorHome = System.getProperty("user.home");
        String priorDir = AppPaths.desktopAppHomeDirName();
        try {
            System.setProperty("user.home", userHome.toString());
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            DesktopSessionStateStore.patchUiEnvRowsAndTheme(
                    List.of(
                            new UiEnvRowSnapshot(
                                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK, "C:\\master.xlsm", ""),
                            new UiEnvRowSnapshot(
                                    AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, "", "")),
                    DesktopTheme.LIGHT.storedId());
            DesktopSessionState session = DesktopSessionStateStore.load();

            ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
            RemoteDesktopEnvRows.bootstrap(rows, session);

            Map<String, String> ui = RemoteDesktopEnvRows.collectMap(rows);
            assertFalse(ui.containsKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
            assertEquals(
                    AppPaths.DEFAULT_PM_AI_RPA_LAUNCHER_DEPLOY_DIR,
                    ui.get(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR));
            assertTrue(ui.containsKey(AppPaths.KEY_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR));
        } finally {
            if (priorHome != null) {
                System.setProperty("user.home", priorHome);
            } else {
                System.clearProperty("user.home");
            }
            AppPaths.setDesktopAppHomeDirName(priorDir);
        }
    }

    @Test
    void migrateLegacyRdpEnvKeys_movesPmdKeysToRpaKeys() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow deploy = new EnvVarRow();
        deploy.setName(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR);
        deploy.setValue("\\\\legacy\\deploy");
        rows.add(deploy);
        EnvVarRow store = new EnvVarRow();
        store.setName(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR);
        store.setValue("\\\\legacy\\data");
        rows.add(store);

        RemoteDesktopEnvRows.migrateLegacyRdpEnvKeys(rows);

        Map<String, String> ui = RemoteDesktopEnvRows.collectMap(rows);
        assertFalse(ui.containsKey(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR));
        assertFalse(ui.containsKey(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR));
        assertEquals("\\\\legacy\\deploy", ui.get(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR));
        assertEquals(
                "\\\\legacy\\data", ui.get(AppPaths.KEY_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR));
    }

    @Test
    void syncRowValue_updatesExistingRow() {
        ObservableList<EnvVarRow> rows = FXCollections.observableArrayList();
        EnvVarRow row = new EnvVarRow();
        row.setName(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR);
        row.setValue("old");
        rows.add(row);

        RemoteDesktopEnvRows.syncRowValue(rows, AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, "new");
        assertEquals("new", rows.get(0).getValue());
    }
}
