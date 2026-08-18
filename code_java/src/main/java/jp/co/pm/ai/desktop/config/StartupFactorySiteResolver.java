package jp.co.pm.ai.desktop.config;

/**
 * 起動スプラッシュ表示前の利用工場。
 *
 * <p>正本は {@link LastLaunchedFactorySiteStore}（前回起動時に書き出した JSON）。ファイルが無い初回は湖南工場。
 */
public final class StartupFactorySiteResolver {

    private StartupFactorySiteResolver() {}

    /** スプラッシュの工場バッジ・テーマ用（この呼び出しでは永続ファイルを変更しない）。 */
    public static FactorySite resolveForSplash() {
        return LastLaunchedFactorySiteStore.load();
    }
}
