package jp.co.pm.ai.desktop.ui;

/**
 * 工場切替プログレスの表示方針（起動シーケンス中は起動側モーダルに任せる）。
 */
public final class FactorySiteSwitchBusySupport {

    private FactorySiteSwitchBusySupport() {}

    /**
     * 工場切替後のタブ再読込中も進捗モーダルを維持するか。
     *
     * @param startupSequenceActive 起動シーケンス中なら false（起動側ダイアログに任せる）
     * @param backgroundLoadStarted タブ再読込チェーンが開始されたとき true
     */
    public static boolean keepBusyDialogForPostSwitchTabLoad(
            boolean startupSequenceActive, boolean backgroundLoadStarted) {
        return !startupSequenceActive && backgroundLoadStarted;
    }

    /** 起動後読込の状況文言を工場切替ダイアログへ載せる。空なら既定文言。 */
    public static String resolveTabLoadStatus(String startupBackgroundLoadMessage) {
        if (startupBackgroundLoadMessage == null || startupBackgroundLoadMessage.isBlank()) {
            return FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD;
        }
        return startupBackgroundLoadMessage;
    }

    /** 子ウィンドウをオーナー中央に置く X。 */
    public static double centerX(double ownerX, double ownerWidth, double childWidth) {
        return ownerX + (ownerWidth - childWidth) / 2.0;
    }

    /** 子ウィンドウをオーナー中央に置く Y。 */
    public static double centerY(double ownerY, double ownerHeight, double childHeight) {
        return ownerY + (ownerHeight - childHeight) / 2.0;
    }
}
