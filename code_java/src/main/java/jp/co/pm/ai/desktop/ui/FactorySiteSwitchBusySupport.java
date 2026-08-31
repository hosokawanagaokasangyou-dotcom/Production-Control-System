package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.scene.Parent;
import javafx.stage.Stage;
import jp.co.pm.ai.desktop.MainShellTabId;

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

    /**
     * ワークスペース適用直後も同一の進捗 Stage を表示したままにするか。
     *
     * <p>操作者のブロッキングダイアログを重ねると FX スレッドが詰まるため、そのときだけ false。
     * 起動シーケンス中は起動側ダイアログに任せる。
     */
    public static boolean keepBusyVisibleThroughFinish(
            boolean startupSequenceActive, boolean operatorBlockingDialogNeeded) {
        return !startupSequenceActive && !operatorBlockingDialogNeeded;
    }

    public static final int WORK_CONNECT = 0;
    public static final int WORK_SAVE = 1;
    public static final int WORK_LOAD = 2;
    public static final int WORK_ENV = 3;
    public static final int WORK_REFRESH_REQUEST_FORM = 4;
    public static final int WORK_REFRESH_PIPELINE = 5;
    public static final int WORK_REFRESH_REMOTE = 6;
    public static final int WORK_STABILIZE = 7;
    public static final int WORK_MATCH = 8;
    public static final int WORK_FINISH = 9;
    public static final int WORK_UNIT_COUNT = 10;

    public static final int POST_ATTENDANCE_COMPANY = 0;
    public static final int POST_ATTENDANCE_MEMBER = 1;
    public static final int POST_ATTENDANCE_MACHINE = 2;
    public static final int POST_ATTENDANCE_MASTER = 3;
    public static final int POST_BACKGROUND_LOAD = 4;
    public static final int POST_WORK_COUNT = 5;

    /** 切替本処理の進捗文言。状況を描画してから重い処理へ入る。 */
    public static String statusForWorkUnit(int unit) {
        return switch (unit) {
            case WORK_CONNECT -> FactorySiteSwitchBusyDialog.STATUS_CONNECT;
            case WORK_SAVE -> FactorySiteSwitchBusyDialog.STATUS_SAVING;
            case WORK_LOAD -> FactorySiteSwitchBusyDialog.STATUS_LOADING;
            case WORK_ENV -> FactorySiteSwitchBusyDialog.STATUS_ENV;
            case WORK_REFRESH_REQUEST_FORM -> FactorySiteSwitchBusyDialog.STATUS_REFRESH_REQUEST_FORM;
            case WORK_REFRESH_PIPELINE -> FactorySiteSwitchBusyDialog.STATUS_REFRESH_PIPELINE;
            case WORK_REFRESH_REMOTE -> FactorySiteSwitchBusyDialog.STATUS_REFRESH_REMOTE;
            case WORK_STABILIZE -> FactorySiteSwitchBusyDialog.STATUS_STABILIZE;
            case WORK_MATCH -> FactorySiteSwitchBusyDialog.STATUS_MATCH;
            case WORK_FINISH -> FactorySiteSwitchBusyDialog.STATUS_OPERATOR;
            default -> FactorySiteSwitchBusyDialog.STATUS_SAVING;
        };
    }

    /** 切替後の勤怠・タブ再読込の進捗文言。 */
    public static String statusForPostSwitchWork(int unit) {
        return switch (unit) {
            case POST_ATTENDANCE_COMPANY -> FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_COMPANY;
            case POST_ATTENDANCE_MEMBER -> FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MEMBER;
            case POST_ATTENDANCE_MACHINE -> FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MACHINE;
            case POST_ATTENDANCE_MASTER -> FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MASTER;
            case POST_BACKGROUND_LOAD -> FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD;
            default -> FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD;
        };
    }

    /** 起動後読込の状況文言を工場切替ダイアログへ載せる。空なら既定文言。 */
    public static String resolveTabLoadStatus(String startupBackgroundLoadMessage) {
        if (startupBackgroundLoadMessage == null || startupBackgroundLoadMessage.isBlank()) {
            return FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD;
        }
        return startupBackgroundLoadMessage;
    }

    /**
     * 工場切替の進捗文言に対応するメインシェルタブを返す。
     *
     * <p>切替本体の保存・接続確認など、特定タブに対応しない文言は空を返す。
     */
    public static Optional<MainShellTabId> targetTabForStatus(String status) {
        if (status == null || status.isBlank()) {
            return Optional.empty();
        }
        if (FactorySiteSwitchBusyDialog.STATUS_ENV.equals(status)
                || FactorySiteSwitchBusyDialog.STATUS_STABILIZE.equals(status)
                || FactorySiteSwitchBusyDialog.STATUS_MATCH.equals(status)) {
            return Optional.of(MainShellTabId.ENV);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_COMPANY.equals(status)
                || status.contains("会社カレンダー")) {
            return Optional.of(MainShellTabId.COMPANY_CALENDAR);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MEMBER.equals(status)
                || status.contains("メンバー勤怠")) {
            return Optional.of(MainShellTabId.MEMBER_ATTENDANCE);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MACHINE.equals(status)
                || status.contains("機械カレンダー")) {
            return Optional.of(MainShellTabId.MACHINE_CALENDAR);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MASTER.equals(status)
                || status.contains("マスタ配台シート")) {
            return Optional.of(MainShellTabId.MASTER_DISPATCH_SHEETS);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_REFRESH_REQUEST_FORM.equals(status)
                || status.contains("原本転記")) {
            return Optional.of(MainShellTabId.REQUEST_FORM_INPUT);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_REFRESH_PIPELINE.equals(status)
                || status.contains("計画確認")) {
            return Optional.of(MainShellTabId.REQUEST_FORM_PIPELINE_CHECK);
        }
        if (FactorySiteSwitchBusyDialog.STATUS_REFRESH_REMOTE.equals(status)
                || status.contains("リモートデスクトップ")) {
            return Optional.of(MainShellTabId.REMOTE_DESKTOP);
        }
        return Optional.empty();
    }

    /** 子ウィンドウをオーナー中央に置く X。 */
    public static double centerX(double ownerX, double ownerWidth, double childWidth) {
        return ownerX + (ownerWidth - childWidth) / 2.0;
    }

    /** 子ウィンドウをオーナー中央に置く Y。 */
    public static double centerY(double ownerY, double ownerHeight, double childHeight) {
        return ownerY + (ownerHeight - childHeight) / 2.0;
    }

    /**
     * 進捗 Stage を {@link javafx.stage.Stage#show()} する前に CSS／レイアウトを確定させ、
     * 最初のパルスでウィンドウサイズが 0 のまま出ないようにする。
     */
    public static void realizeStageForImmediateShow(Stage stage) {
        if (stage == null || stage.getScene() == null || stage.getScene().getRoot() == null) {
            return;
        }
        Parent root = stage.getScene().getRoot();
        root.applyCss();
        root.autosize();
        root.layout();
        stage.sizeToScene();
    }
}
