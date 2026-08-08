package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.util.Map;

/**
 * 工場コンボ・工場切替の可否（2 段: サマリ Excel 親フォルダ到達 → ユーザー管理 bin 照合）。
 */
public final class FactorySiteOperatorAccess {

    private FactorySiteOperatorAccess() {}

    /** 当該工場のサマリ Excel 親フォルダにディレクトリ一覧できるか。 */
    public static boolean isFactorySummaryFolderReachable(Map<String, String> ui, FactorySite site) {
        if (site == null || site == FactorySite.RDP_LAUNCHER) {
            return false;
        }
        var shared = AppPaths.summarySharedDataDirForFactory(ui != null ? ui : Map.of(), site);
        return NetworkSourceDirResolver.isDirectoryListingReachable(shared);
    }

    /**
     * ログイン中操作者が当該工場のユーザー管理一覧に含まれるか（第 2 段）。
     * 呼び出し前に {@link FactoryOperatorUserStore#configureForCurrentApp(Map, FactorySite)} 済みであること。
     */
    public static boolean isSessionOperatorInFactoryUserManagement(FactorySite site) {
        if (site == null || site == FactorySite.RDP_LAUNCHER) {
            return false;
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (session.isBlank() || FactoryOperatorUserStore.isGuestOperator(session)) {
            return true;
        }
        try {
            return FactoryOperatorUserStore.loginChoicesForFactory(site).contains(session);
        } catch (IOException ex) {
            return false;
        }
    }

    /** 工場ユーザーとして当該工場を利用可能か（短絡評価）。 */
    public static boolean isSessionOperatorAllowedForFactory(Map<String, String> ui, FactorySite site) {
        if (site == null || site == FactorySite.RDP_LAUNCHER) {
            return false;
        }
        if (!isFactorySummaryFolderReachable(ui, site)) {
            return false;
        }
        return isSessionOperatorInFactoryUserManagement(site);
    }

    /** コンボ不可理由（ログ・Tooltip 用）。到達可なら empty。 */
    public static String comboBlockReasonJa(Map<String, String> ui, FactorySite site) {
        if (site == null || site == FactorySite.RDP_LAUNCHER) {
            return "";
        }
        if (!isFactorySummaryFolderReachable(ui, site)) {
            var shared = AppPaths.summarySharedDataDirForFactory(ui != null ? ui : Map.of(), site);
            return "共有 DATA フォルダにアクセスできません（当該工場のユーザーではありません）: " + shared;
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (!session.isBlank()
                && !FactoryOperatorUserStore.isGuestOperator(session)
                && !isSessionOperatorInFactoryUserManagement(site)) {
            return "操作者「" + session + "」は " + site.displayLabelJa() + " のユーザー管理に未登録です";
        }
        return "";
    }
}
