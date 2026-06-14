package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javafx.collections.ObservableList;

import jp.co.pm.ai.desktop.EnvVarRow;
import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;

/** リモートデスクトップ専用シェルの環境変数表（{@link EnvVarRow}）の読込・永続化補助。 */
public final class RemoteDesktopEnvRows {

    private RemoteDesktopEnvRows() {}

    public static void bootstrap(ObservableList<EnvVarRow> rows, DesktopSessionState session) {
        if (rows == null) {
            return;
        }
        if (session != null && session.uiEnvRows() != null && !session.uiEnvRows().isEmpty()) {
            applyFromSession(rows, session);
        } else {
            populateFromUiRef(rows);
        }
        applyRdpLauncherEmptyDefaults(rows);
        if (rows.isEmpty()) {
            rows.add(new EnvVarRow());
        }
    }

    public static void populateFromUiRef(ObservableList<EnvVarRow> rows) {
        List<EnvVarRow> list = new ArrayList<>();
        for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
            String key = e.key() != null ? e.key().trim() : "";
            if (key.isEmpty()) {
                continue;
            }
            EnvVarRow row = new EnvVarRow();
            row.setName(key);
            row.setValue(e.value() != null ? e.value() : "");
            row.setDescription(EnvVarDocs.mergeDescriptions(e.description(), key));
            list.add(row);
        }
        rows.setAll(list);
    }

    public static void applyFromSession(ObservableList<EnvVarRow> rows, DesktopSessionState session) {
        if (session == null || session.uiEnvRows() == null || session.uiEnvRows().isEmpty()) {
            return;
        }
        List<EnvVarRow> restored = new ArrayList<>(session.uiEnvRows().size());
        for (UiEnvRowSnapshot snap : session.uiEnvRows()) {
            String name = snap.name() != null ? snap.name().trim() : "";
            if (name.isEmpty()) {
                continue;
            }
            EnvVarRow row = new EnvVarRow();
            row.setName(snap.name() != null ? snap.name() : "");
            row.setValue(snap.value() != null ? snap.value() : "");
            String desc = snap.description() != null ? snap.description() : "";
            if (desc.isBlank()) {
                desc = EnvVarDocs.mergeDescriptions("", name);
            }
            row.setDescription(desc);
            restored.add(row);
        }
        rows.setAll(restored);
    }

    /** {@link DesktopUiEnvMapLoader} と同様の RDP 専用既定を、値が空の行へだけ適用する。 */
    public static void applyRdpLauncherEmptyDefaults(ObservableList<EnvVarRow> rows) {
        Map<String, String> ui = collectMap(rows);
        if (ui.getOrDefault(AppPaths.KEY_PM_AI_REPO_ROOT, "").isBlank()) {
            setRowValueIfBlank(
                    rows,
                    AppPaths.KEY_PM_AI_REPO_ROOT,
                    AppPaths.resolveRepoRoot(ui).toString());
        }
        setRowValueIfBlank(
                rows,
                AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR,
                AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR);
        setRowValueIfBlank(
                rows,
                AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR,
                AppPaths.defaultRdpLauncherSharedDataDir().toString());
        setRowValueIfBlank(
                rows,
                AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR,
                AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_RELEASE_DIR);
    }

    public static Map<String, String> collectMap(ObservableList<EnvVarRow> rows) {
        LinkedHashMap<String, String> map = new LinkedHashMap<>();
        if (rows == null) {
            return map;
        }
        for (EnvVarRow row : rows) {
            String key = row.getName() != null ? row.getName().trim() : "";
            if (key.isEmpty() || key.startsWith("#")) {
                continue;
            }
            map.put(key, row.getValue() != null ? row.getValue() : "");
        }
        return map;
    }

    public static List<UiEnvRowSnapshot> snapshot(ObservableList<EnvVarRow> rows) {
        if (rows == null) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>(rows.size());
        for (EnvVarRow row : rows) {
            String key = row.getName() != null ? row.getName().trim() : "";
            if (key.isEmpty()) {
                continue;
            }
            out.add(
                    new UiEnvRowSnapshot(
                            key,
                            row.getValue() != null ? row.getValue() : "",
                            row.getDescription() != null ? row.getDescription() : ""));
        }
        return List.copyOf(out);
    }

    public static void mergeMissingFromUiRef(ObservableList<EnvVarRow> rows) {
        if (rows == null) {
            return;
        }
        Set<String> existing = new HashSet<>();
        for (EnvVarRow r : rows) {
            String name = r.getName() != null ? r.getName().trim() : "";
            if (!name.isEmpty()) {
                existing.add(name);
            }
        }
        for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
            String key = e.key() != null ? e.key().trim() : "";
            if (key.isEmpty() || existing.contains(key)) {
                continue;
            }
            EnvVarRow row = new EnvVarRow();
            row.setName(key);
            row.setValue(e.value() != null ? e.value() : "");
            row.setDescription(EnvVarDocs.mergeDescriptions(e.description(), key));
            rows.add(row);
            existing.add(key);
        }
    }

    public static void syncRowValue(ObservableList<EnvVarRow> rows, String envKey, String value) {
        if (rows == null || envKey == null || envKey.isBlank()) {
            return;
        }
        String key = envKey.strip();
        String v = value != null ? value : "";
        for (EnvVarRow row : rows) {
            if (key.equals(row.getName() != null ? row.getName().strip() : "")) {
                row.setValue(v);
                return;
            }
        }
        EnvVarRow row = new EnvVarRow();
        row.setName(key);
        row.setValue(v);
        row.setDescription(EnvVarDocs.mergeDescriptions("", key));
        rows.add(row);
    }

    private static void setRowValueIfBlank(ObservableList<EnvVarRow> rows, String key, String value) {
        if (key == null || key.isBlank() || value == null || value.isBlank()) {
            return;
        }
        for (EnvVarRow row : rows) {
            if (key.equals(row.getName() != null ? row.getName().strip() : "")) {
                if (row.getValue() == null || row.getValue().isBlank()) {
                    row.setValue(value);
                }
                return;
            }
        }
    }
}
