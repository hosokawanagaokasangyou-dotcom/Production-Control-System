package jp.co.pm.ai.desktop.config;

/**
 * 起動時の利用工場。スプラッシュと本番で同じ正本を使う。
 *
 * <p>正本は {@link LastLaunchedFactorySiteStore}（前回起動時に書き出した JSON）。ファイルが無い初回は湖南工場。
 */
public final class StartupFactorySiteResolver {

    private StartupFactorySiteResolver() {}

    /** スプラッシュ・本番共通の採用工場。この呼び出しでは永続ファイルを変更しない。 */
    public static FactorySite resolve() {
        return LastLaunchedFactorySiteStore.load();
    }

    /** スプラッシュの工場バッジ・テーマ用。{@link #resolve()} と同じ。 */
    public static FactorySite resolveForSplash() {
        return resolve();
    }

    /**
     * ディスク上の永続工場が採用工場と異なるとき、起動時に工場切替（ワークスペース復元）が必要。
     */
    public static boolean requiresStartupSwitch(FactorySite persisted, FactorySite adopted) {
        if (adopted == null || adopted == FactorySite.RDP_LAUNCHER) {
            return false;
        }
        if (persisted == null || persisted == FactorySite.RDP_LAUNCHER) {
            return true;
        }
        return persisted != adopted;
    }
}
