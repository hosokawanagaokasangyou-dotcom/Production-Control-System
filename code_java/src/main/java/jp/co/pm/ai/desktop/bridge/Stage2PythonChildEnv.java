package jp.co.pm.ai.desktop.bridge;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;

/**
 * Python 子（段階1/2）へ渡す env の共通加工（レガシーキー除去・一覧スキップに応じたネットワークソース解決・pause 無効化）。
 *
 * <p>配台計画タブからの {@code PM_AI_PLAN_INPUT_PATH} 補完は JavaFX 側のみのため {@link
 * jp.co.pm.ai.desktop.MainShellController#childEnvForPython} が先にオーバーレイする。
 */
public final class Stage2PythonChildEnv {

    /** 環境タブから廃止されたが OS に残り得るキー。Python 子には渡さない。 */
    public static final Set<String> LEGACY_WORKBOOK_KEYS_STRIPPED_FOR_PYTHON_CHILD =
            Set.of("TASK_INPUT_WORKBOOK", "PM_AI_TASK_INPUT_WORKBOOK");

    private Stage2PythonChildEnv() {}

    public static void stripLegacyWorkbookKeys(Map<String, String> m) {
        if (m == null) {
            return;
        }
        for (String k : LEGACY_WORKBOOK_KEYS_STRIPPED_FOR_PYTHON_CHILD) {
            m.remove(k);
        }
    }

    public static void ensureSkipWorkbookEnvSheetDefault(Map<String, String> m) {
        if (m == null) {
            return;
        }
        String skip = m.get(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET);
        if (skip == null || skip.isBlank()) {
            m.put(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET, "1");
        }
    }

    /**
     * ネットワークソース解決を適用し、子プロセスの pause を無効化する。
     *
     * @return 解決結果（ログ用）。{@code m} はインプレース更新される。
     */
    public static NetworkSourceDirResolver.Result applyNetworkSourceAndChildPause(
            Map<String, String> m, boolean skipTaskDirListing, boolean skipActualDetailDirListing) {
        NetworkSourceDirResolver.Result netRes =
                NetworkSourceDirResolver.resolve(m, skipTaskDirListing, skipActualDetailDirListing);
        NetworkSourceDirResolver.applyToEnv(m, netRes);
        m.put(AppPaths.KEY_PM_AI_CMD_PAUSE_ON_ERROR, "0");
        return netRes;
    }

    /**
     * ヘッドレス同一検証 CLI 用: OS 環境をコピーし、Python 子向けの最低限のマージを行う（配台計画タブの補完はしない）。
     */
    public static HashMap<String, String> headlessBaseFromSystemEnv() {
        HashMap<String, String> m = new HashMap<>(System.getenv());
        stripLegacyWorkbookKeys(m);
        ensureSkipWorkbookEnvSheetDefault(m);
        boolean taskReach = NetworkSourceDirResolver.isTaskInputSourceDirReachable(m);
        boolean actReach = NetworkSourceDirResolver.isActualDetailSourceDirReachable(m);
        applyNetworkSourceAndChildPause(m, !taskReach, !actReach);
        return m;
    }
}
