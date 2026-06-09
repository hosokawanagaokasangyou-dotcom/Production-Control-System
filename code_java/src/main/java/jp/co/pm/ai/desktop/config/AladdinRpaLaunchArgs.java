package jp.co.pm.ai.desktop.config;

/** Aladdin RPA Studio 起動時のコマンドライン引数契約（Java / C# ランチャー共通）。 */
public final class AladdinRpaLaunchArgs {

    public static final String ID_FLAG = "--id";
    public static final String PASSWORD_FLAG = "--password";
    /** Aladdin RPA シナリオ（.ardrpa）指定。直後のトークンがシナリオパス。 */
    public static final String SCENARIO_FLAG = "--scenario";
    /** シナリオなし／終了後も RPA プロセスを維持する（Aladdin RPA Studio 向け）。 */
    public static final String ETERNAL_FLAG = "--eternal";

    private AladdinRpaLaunchArgs() {}
}
