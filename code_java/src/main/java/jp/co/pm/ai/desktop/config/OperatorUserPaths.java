package jp.co.pm.ai.desktop.config;

import java.util.Map;

/**
 * 操作者名から共有フォルダ配下のディレクトリ名を解決する（アラジン入力用配台計画の世代フォルダ等）。
 */
public final class OperatorUserPaths {

    public static final String UNKNOWN_OPERATOR_DIR = "unknown";

    private OperatorUserPaths() {}

    /** 環境変数 {@link AppPaths#KEY_PM_AI_OPERATOR_USER} またはセッション操作者名。 */
    public static String resolveOperatorUser(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String fromUi = u.getOrDefault(AppPaths.KEY_PM_AI_OPERATOR_USER, "").strip();
        if (!fromUi.isEmpty()) {
            return fromUi;
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        return session.isBlank() ? UNKNOWN_OPERATOR_DIR : session;
    }

    /** ファイルシステム用に操作者名をサニタイズする。 */
    public static String sanitizeOperatorDirName(String operatorUser) {
        if (operatorUser == null || operatorUser.isBlank()) {
            return UNKNOWN_OPERATOR_DIR;
        }
        String t = operatorUser.strip().replaceAll("[\\\\/:*?\"<>|]", "_");
        if (t.equals(".") || t.equals("..") || t.contains("..")) {
            return UNKNOWN_OPERATOR_DIR;
        }
        if (t.length() > 40) {
            t = t.substring(0, 40);
        }
        return t.isEmpty() ? UNKNOWN_OPERATOR_DIR : t;
    }
}
