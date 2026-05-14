package jp.co.pm.ai.planning.stage2;

import java.util.Locale;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.config.AppPaths;

/** 段階2関連の環境変数トークン解釈（Python planning_core の truthy / 無効トークンに概ね整合）。 */
public final class Stage2EnvParsing {

    private static final Set<String> OFF =
            Set.of("0", "false", "no", "off", "none", "n");

    private Stage2EnvParsing() {}

    public static boolean envEnabled(String key, Map<String, String> ui, boolean defaultWhenUnset) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(key);
        if (raw == null || raw.isBlank()) {
            return defaultWhenUnset;
        }
        return !OFF.contains(raw.strip().toLowerCase(Locale.ROOT));
    }

    public static boolean stage2WriteExcel(Map<String, String> ui) {
        return envEnabled(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, ui, true);
    }

    /**
     * {@code PM_AI_STAGE2_ENGINE=java} かつ本キーが真のとき、配台本体は Python 子プロセス（正本）へ委譲する。
     */
    public static boolean javaDelegatesPythonDispatch(Map<String, String> ui) {
        return envEnabled(AppPaths.KEY_PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH, ui, false);
    }

    /**
     * {@link AppPaths#KEY_PM_AI_DISPATCH_ENGINE} が {@code java}（大小無視）のとき真。未設定・空・その他の値は偽。
     */
    public static boolean dispatchCoreExplicitJava(Map<String, String> ui) {
        return "java".equalsIgnoreCase(trimDispatchToken(ui, AppPaths.KEY_PM_AI_DISPATCH_ENGINE));
    }

    /**
     * {@link AppPaths#KEY_PM_AI_DISPATCH_ENGINE} が {@code python}（大小無視）のとき真。未設定・空は偽（従来の委譲フラグのみ参照）。
     */
    public static boolean dispatchCoreExplicitPython(Map<String, String> ui) {
        return "python".equalsIgnoreCase(trimDispatchToken(ui, AppPaths.KEY_PM_AI_DISPATCH_ENGINE));
    }

    private static String trimDispatchToken(Map<String, String> ui, String key) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(key);
        return raw != null ? raw.strip() : "";
    }
}
