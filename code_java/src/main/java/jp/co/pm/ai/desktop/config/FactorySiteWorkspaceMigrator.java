package jp.co.pm.ai.desktop.config;

import java.nio.file.Files;
import java.util.List;
import java.util.Map;

/**
 * 初回操作者ログイン時に旧永続化（session-state 工場項目・GlobalInitSettingTarget）から operator-local へ seed。
 */
public final class FactorySiteWorkspaceMigrator {

    private FactorySiteWorkspaceMigrator() {}

    public static void migrateIfNeeded(
            String operatorName,
            FactorySite currentSite,
            List<UiEnvRowSnapshot> uiEnvRows,
            DesktopSessionState sessionState,
            Map<String, String> ui) {
        if (FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName).isEmpty()) {
            return;
        }
        if (currentSite == null || currentSite == FactorySite.RDP_LAUNCHER) {
            return;
        }
        try {
            var marker = AppPaths.operatorLocalMigrationMarkerPath(operatorName);
            if (Files.isRegularFile(marker)) {
                return;
            }
            Files.createDirectories(marker.getParent());
            if (FactorySiteWorkspaceStore.loadLastFactorySite(operatorName).isEmpty()) {
                FactorySiteWorkspaceStore.saveLastFactorySite(
                        operatorName,
                        GlobalInitSettingTarget.load());
            }
            if (FactorySiteWorkspaceStore.load(operatorName, currentSite).isEmpty()) {
                DesktopSessionState fragment =
                        sessionState != null
                                ? sessionState.extractFactoryScopedFields()
                                : DesktopSessionState.empty();
                FactorySiteWorkspaceStore.save(
                        operatorName,
                        currentSite,
                        new FactorySiteWorkspaceSnapshot(
                                uiEnvRows != null ? uiEnvRows : List.of(), fragment));
            }
            Files.writeString(marker, java.time.Instant.now().toString());
        } catch (Exception ignored) {
        }
    }
}
