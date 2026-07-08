package jp.co.pm.ai.desktop.config;

import java.util.List;

/** 工場別ワークスペース（環境変数行 + 工場スコープ session 断片）。 */
public record FactorySiteWorkspaceSnapshot(
        List<UiEnvRowSnapshot> uiEnvRows, DesktopSessionState sessionFragment) {

    public FactorySiteWorkspaceSnapshot {
        uiEnvRows = uiEnvRows != null ? List.copyOf(uiEnvRows) : List.of();
        sessionFragment = sessionFragment != null ? sessionFragment : DesktopSessionState.empty();
    }
}
