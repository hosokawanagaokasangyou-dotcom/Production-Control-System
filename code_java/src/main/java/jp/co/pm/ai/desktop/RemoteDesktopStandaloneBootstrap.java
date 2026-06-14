package jp.co.pm.ai.desktop;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.runtime.WindowsLauncherUserDir;

/** リモートデスクトップ配布用アプリ起動時の PC ローカル設定ルート切替。PMD 本体は呼ばない。 */
public final class RemoteDesktopStandaloneBootstrap {

    private static volatile boolean activated;

    private RemoteDesktopStandaloneBootstrap() {}

    public static void activate() {
        if (activated) {
            return;
        }
        synchronized (RemoteDesktopStandaloneBootstrap.class) {
            if (activated) {
                return;
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            WindowsLauncherUserDir.alignWithPackagedLauncherIfWindows();
            activated = true;
        }
    }

    public static boolean isActivated() {
        return activated;
    }
}
