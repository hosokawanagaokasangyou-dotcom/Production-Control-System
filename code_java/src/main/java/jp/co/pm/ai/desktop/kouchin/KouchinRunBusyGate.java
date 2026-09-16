package jp.co.pm.ai.desktop.kouchin;

/**
 * 段階実行と工賃実行の Busy OR。Dialog なしの純ロジック。
 */
public final class KouchinRunBusyGate {

    private KouchinRunBusyGate() {}

    public static boolean overlayVisible(boolean stageBusy, boolean dispatchTrialBusy, boolean kouchinBusy) {
        return stageBusy || dispatchTrialBusy || kouchinBusy;
    }

    public static boolean cancelTargetsKouchin(boolean kouchinBusy) {
        return kouchinBusy;
    }

    public static boolean canStart(boolean stageBusy, boolean kouchinBusy) {
        return !stageBusy && !kouchinBusy;
    }
}
