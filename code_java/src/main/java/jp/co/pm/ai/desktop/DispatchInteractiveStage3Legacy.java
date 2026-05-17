package jp.co.pm.ai.desktop;

/**
 * 段階3（配台試行）の現行実装の置き場所を示すマーカー。
 *
 * <p>配台試行は {@link DispatchInteractiveTabController#onDispatchTrialAction}（内部 {@link
 * DispatchInteractiveTabController#startDispatchTrial()}）から起動し、段階2と同一エンジンで単一フェーズ実行する。
 *
 * <p>ロジックを追加・分割するときは、このクラスの Javadoc にエントリを追記するか、専用クラスへ切り出す。
 */
public final class DispatchInteractiveStage3Legacy {

    private DispatchInteractiveStage3Legacy() {}

    /** 現行の段階3（配台試行）UI エントリ。 */
    public static final String CURRENT_TRIAL_HANDLER =
            "DispatchInteractiveTabController#onDispatchTrialAction / startDispatchTrial()";
}
