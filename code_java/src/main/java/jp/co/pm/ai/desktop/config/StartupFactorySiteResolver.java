package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.util.Map;
import java.util.Optional;

/**
 * 起動スプラッシュ表示前に、メインシェル起動後と同等の利用工場を解決する。
 *
 * <p>永続ファイルのみの {@link GlobalInitSettingTarget#load()} では、セッション環境変数の UNC 推定や操作者ワークスペースの
 * 前回工場とずれることがある。
 */
public final class StartupFactorySiteResolver {

    private StartupFactorySiteResolver() {}

    /** スプラッシュの工場バッジ・テーマ用（永続ファイルは変更しない）。 */
    public static FactorySite resolveForSplash() {
        Optional<FactorySite> fromFollowUp =
                PortableBundleUpgradeFollowUp.readIfPresent()
                        .flatMap(PortableBundleUpgradeFollowUp::factorySiteOrEmpty);
        if (fromFollowUp.isPresent()) {
            return fromFollowUp.get();
        }
        Map<String, String> ui = DesktopUiEnvMapLoader.loadForStartupFactoryInference();
        FactorySite base = GlobalInitSettingTarget.peekEffective(ui);
        return resolveOperatorWorkspaceLastFactory(ui, base).orElse(base);
    }

    /**
     * {@link jp.co.pm.ai.desktop.MainShellController} の起動時工場復元（{@code finalizeOperatorLocalWorkspaceAfterSessionEstablished}）と同じ
     * 判定で、操作者ローカルに保存した前回工場を優先する。
     */
    static Optional<FactorySite> resolveOperatorWorkspaceLastFactory(
            Map<String, String> ui, FactorySite current) {
        if (current == null || current == FactorySite.RDP_LAUNCHER) {
            return Optional.empty();
        }
        Map<String, String> env = ui != null ? ui : Map.of();
        try {
            FactoryOperatorUserStore.configureFromUi(env, current);
            String operator = FactoryOperatorUserStore.lastSelectedForFactory(current);
            if (operator.isBlank()
                    || FactoryOperatorUserStore.isGuestOperator(operator)
                    || !FactoryOperatorUserStore.wouldRestoreSessionFromLocalLastSelected(
                            current, operator)) {
                return Optional.empty();
            }
            Optional<FactorySite> last = FactorySiteWorkspaceStore.loadLastFactorySite(operator);
            if (last.isEmpty() || last.get() == current) {
                return Optional.empty();
            }
            FactorySite target = last.get();
            if (!FactorySiteOperatorAccess.isOperatorAllowedForFactory(env, target, operator)) {
                return Optional.empty();
            }
            return Optional.of(target);
        } catch (IOException ignored) {
            return Optional.empty();
        }
    }
}
