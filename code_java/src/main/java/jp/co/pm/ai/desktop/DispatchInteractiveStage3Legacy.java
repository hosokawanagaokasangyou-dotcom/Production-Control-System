package jp.co.pm.ai.desktop;

/**
 * 段階3（配台試行）の現行実装の置き場所を示すマーカー。
 *
 * <p>再設計・ステップビルド中も、従来の配台試行ロジックは {@link DispatchInteractiveTabController#onDispatchTrialAction}
 * に残してあり、Git 履歴および当メソッド本文を正として戻せる。
 *
 * <p>新しい段階3ロジックを追加するときは、当面このクラスの Javadoc に新エントリポイントを追記するか、専用クラスへ切り出す。
 */
public final class DispatchInteractiveStage3Legacy {

    private DispatchInteractiveStage3Legacy() {}

    /** 現行の段階3（配台試行）UI エントリ。 */
    public static final String CURRENT_TRIAL_HANDLER = "DispatchInteractiveTabController#onDispatchTrialAction";
}
