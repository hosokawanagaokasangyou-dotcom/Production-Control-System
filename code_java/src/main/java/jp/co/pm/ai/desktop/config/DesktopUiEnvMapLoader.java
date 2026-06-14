package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;

/** 環境変数タブを持たないシェル向けの env マップ読込（RDP 専用シェルは環境変数タブから永続化）。 */
public final class DesktopUiEnvMapLoader {

    private DesktopUiEnvMapLoader() {}

    public static Map<String, String> loadInitialMap() {
        LinkedHashMap<String, String> map = new LinkedHashMap<>();
        for (WorkbookEnvSheetReader.RowEntry e : UiRefEnvDefaults.loadOrEmpty()) {
            String key = e.key() != null ? e.key().trim() : "";
            if (key.isEmpty()) {
                continue;
            }
            map.put(key, e.value() != null ? e.value() : "");
        }
        DesktopSessionState session = DesktopSessionStateStore.load();
        if (session.uiEnvRows() != null) {
            for (UiEnvRowSnapshot row : session.uiEnvRows()) {
                String name = row.name() != null ? row.name().trim() : "";
                if (name.isEmpty()) {
                    continue;
                }
                map.put(name, row.value() != null ? row.value() : "");
            }
        }
        if (map.getOrDefault(AppPaths.KEY_PM_AI_REPO_ROOT, "").isBlank()) {
            map.put(AppPaths.KEY_PM_AI_REPO_ROOT, AppPaths.resolveRepoRoot(map).toString());
        }
        if (AppPaths.usesRemoteDesktopAppHome()) {
            String portableKey = AppPaths.KEY_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR;
            if (map.getOrDefault(portableKey, "").isBlank()) {
                map.put(portableKey, AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR);
            }
            String operatorStoreKey = AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR;
            if (map.getOrDefault(operatorStoreKey, "").isBlank()) {
                map.put(operatorStoreKey, AppPaths.defaultRdpLauncherSharedDataDir().toString());
            }
            String deployKey = AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR;
            if (map.getOrDefault(deployKey, "").isBlank()) {
                map.put(deployKey, AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_RELEASE_DIR);
            }
            // PM_AI_RDP_LAUNCHER_INI は空のまま（操作者別 {操作者名}_RPA設定.ini を resolveRdpLauncherIni が解決）
        }
        return Map.copyOf(map);
    }

    public static void persistEnvMapAndTheme(Map<String, String> envMap, String uiThemeId) {
        DesktopSessionStateStore.patchUiEnvMapAndTheme(envMap, uiThemeId);
    }
}
