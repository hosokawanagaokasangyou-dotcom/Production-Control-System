package jp.co.pm.ai.desktop.ui;

/** 表の行ホバー暗転（フォーカス行以外を薄く表示）のグローバル ON/OFF。 */
public final class UiRowHoverDimmingSettings {

    public static final boolean DEFAULT_ENABLED = true;

    private static volatile boolean enabled = DEFAULT_ENABLED;

    private UiRowHoverDimmingSettings() {}

    public static boolean enabled() {
        return enabled;
    }

    public static void setEnabled(boolean on) {
        enabled = on;
    }
}
