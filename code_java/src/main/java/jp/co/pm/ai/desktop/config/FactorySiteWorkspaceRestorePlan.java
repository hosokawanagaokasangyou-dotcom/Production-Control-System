package jp.co.pm.ai.desktop.config;

import java.util.Optional;

/**
 * 工場切替・起動復元で環境変数をどう載せるかの方針。
 *
 * <p>保存済み {@code uiEnvRows} があるときはそれを正本として復元し、工場 UNC 既定で上書きしない。
 * 無いときは ui_ref 既定＋工場 overlay で初期化する。いずれの場合も {@code init_setting} を環境変数より先に適用する。
 */
public record FactorySiteWorkspaceRestorePlan(
        boolean applyInitSettingBeforeEnv,
        boolean restoreSavedUiEnvRows,
        boolean bundledEnvReset,
        boolean overlayFactoryNetworkDefaults,
        boolean applySessionFragment,
        boolean preserveEnvInitializationInSessionFragment) {

    public static FactorySiteWorkspaceRestorePlan of(
            Optional<FactorySiteWorkspaceSnapshot> workspace) {
        boolean present = workspace != null && workspace.isPresent();
        boolean hasEnv = present && workspace.get().hasUiEnvRows();
        return new FactorySiteWorkspaceRestorePlan(
                true, hasEnv, !hasEnv, true, present, !hasEnv);
    }
}
