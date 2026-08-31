package jp.co.pm.ai.desktop.config;

/**
 * 環境変数初期化が必要なとき、ボタン操作を待たず自動実行するか。
 */
public final class EnvVarsAutoInitialization {

    private EnvVarsAutoInitialization() {}

    /**
     * 起動時照合が終わり、初期化待ちで、ゲストでも初期化実行中でもないとき true。
     */
    public static boolean shouldRun(
            boolean startupCheckCompleted,
            boolean initializationPending,
            boolean guestSession,
            boolean envResetInProgress) {
        return startupCheckCompleted
                && initializationPending
                && !guestSession
                && !envResetInProgress;
    }
}
