package jp.co.pm.ai.desktop;

/** 起動スプラッシュの文言・スタイル差し替え。 */
public record StartupSplashBranding(
        String title,
        String subtitleJa,
        String subtitleEn,
        String statusText,
        String rootStyleClass,
        String backgroundResource) {

    private static final String DEFAULT_BACKGROUND =
            "/jp/co/pm/ai/desktop/images/splash-background.png";

    public static final StartupSplashBranding PMD =
            new StartupSplashBranding(
                    "工程管理 AI 配台",
                    "発泡樹脂（ペフ）· ロール加工 · 配台管理",
                    "PEF FOAM · ROLL SLICE SLIT PACK",
                    "発泡樹脂のロール加工・配台システムを起動しています…",
                    "",
                    DEFAULT_BACKGROUND);

    public static final StartupSplashBranding REMOTE_DESKTOP_LAUNCHER =
            new StartupSplashBranding(
                    RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE,
                    "RDP 接続 · Aladdin RPA · 共有フォルダランチャー",
                    "REMOTE DESKTOP · RPA LAUNCHER",
                    "リモートデスクトップ環境を準備しています…",
                    "splash-app-rdp-launcher",
                    DEFAULT_BACKGROUND);
}
