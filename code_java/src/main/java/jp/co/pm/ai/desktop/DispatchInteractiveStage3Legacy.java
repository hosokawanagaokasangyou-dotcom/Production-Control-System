package jp.co.pm.ai.desktop;

/**
 * 段階3（配台試行）の現行実装の置き場所を示すマーカー。
 *
 * <p>二相（案B）のエントリは {@link DispatchInteractiveTabController#onStage3EquipmentTrialAction} /
 * {@link DispatchInteractiveTabController#onStage3PeopleTrialAction} /
 * {@link DispatchInteractiveTabController#onStage3BothTrialAction}（内部 {@link DispatchInteractiveTabController#startDispatchTrial}）。
 *
 * <p>新しい段階3ロジックを追加するときは、当面このクラスの Javadoc に新エントリポイントを追記するか、専用クラスへ切り出す。
 */
public final class DispatchInteractiveStage3Legacy {

    private DispatchInteractiveStage3Legacy() {}

    /** 現行の段階3（配台試行）UI エントリ。 */
    public static final String CURRENT_TRIAL_HANDLER =
            "DispatchInteractiveTabController#startDispatchTrial(String)";
}
