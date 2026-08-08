package jp.co.pm.ai.desktop;

import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;

/**
 * 起動・環境確定後にタブデータを順次バックグラウンド読込する。
 *
 * <p>順序: リモートデスクトップ → 会社カレンダー → メンバー勤怠 → 機械カレンダー → 原本転記 → 計画確認。
 */
final class StartupTabBackgroundLoadCoordinator {

    interface Host {
        void setStartupBackgroundLoadStatus(String message);

        void appendStartupBackgroundLog(String line);

        RemoteDesktopTabController remoteDesktopTab();

        CompanyCalendarTabController companyCalendarTab();

        MemberAttendanceTabController memberAttendanceTab();

        MachineCalendarTabController machineCalendarTab();

        RequestFormInputTabController requestFormInputTab();

        RequestFormPipelineCheckTabController requestFormPipelineCheckTab();

        void onStartupBackgroundLoadFinished();

        void setStartupTabBackgroundLoadActive(boolean active);

        boolean isStartupTabBackgroundLoadActive();
    }

    private static final int STEP_REMOTE = 1;
    private static final int STEP_COMPANY = 2;
    private static final int STEP_MEMBER = 3;
    private static final int STEP_MACHINE = 4;
    private static final int STEP_REQUEST_FORM = 5;
    private static final int STEP_PIPELINE_CHECK = 6;
    private static final int STEP_COUNT = 6;

    private final Host host;
    private final AtomicBoolean runScheduled = new AtomicBoolean(false);

    StartupTabBackgroundLoadCoordinator(Host host) {
        this.host = host;
    }

    /** 未実行なら読込チェーンを開始する。 */
    void scheduleIfIdle() {
        if (!runScheduled.compareAndSet(false, true)) {
            return;
        }
        host.setStartupTabBackgroundLoadActive(true);
        Platform.runLater(this::beginRemoteDesktop);
    }

    /** 工場切替・ワークスペース復元後に再読込する。 */
    void resetAndSchedule() {
        runScheduled.set(false);
        scheduleIfIdle();
    }

    private void beginRemoteDesktop() {
        setStatus(STEP_REMOTE, "リモートデスクトップ");
        host.appendStartupBackgroundLog("[startup-bg] リモートデスクトップを読込中…");
        RemoteDesktopTabController rdp = host.remoteDesktopTab();
        if (rdp == null) {
            beginCompanyCalendar();
            return;
        }
        rdp.scheduleBackgroundPreload(this::beginCompanyCalendar);
    }

    private void beginCompanyCalendar() {
        setStatus(STEP_COMPANY, "会社カレンダー");
        host.appendStartupBackgroundLog("[startup-bg] 会社カレンダーを読込中…");
        CompanyCalendarTabController tab = host.companyCalendarTab();
        if (tab == null) {
            beginMemberAttendance();
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(() -> finishStep("会社カレンダー", ok, this::beginMemberAttendance)));
    }

    private void beginMemberAttendance() {
        setStatus(STEP_MEMBER, "メンバー勤怠");
        host.appendStartupBackgroundLog("[startup-bg] メンバー勤怠を読込中…");
        MemberAttendanceTabController tab = host.memberAttendanceTab();
        if (tab == null) {
            beginMachineCalendar();
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(() -> finishStep("メンバー勤怠", ok, this::beginMachineCalendar)));
    }

    private void beginMachineCalendar() {
        setStatus(STEP_MACHINE, "機械カレンダー");
        host.appendStartupBackgroundLog("[startup-bg] 機械カレンダーを読込中…");
        MachineCalendarTabController tab = host.machineCalendarTab();
        if (tab == null) {
            beginRequestFormInput();
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(() -> finishStep("機械カレンダー", ok, this::beginRequestFormInput)));
    }

    private void beginRequestFormInput() {
        setStatus(STEP_REQUEST_FORM, "原本転記");
        host.appendStartupBackgroundLog("[startup-bg] 原本転記を読込中…");
        RequestFormInputTabController tab = host.requestFormInputTab();
        if (tab == null) {
            beginPipelineCheck();
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(() -> finishStep("原本転記", ok, this::beginPipelineCheck)));
    }

    private void beginPipelineCheck() {
        setStatus(STEP_PIPELINE_CHECK, "計画確認");
        host.appendStartupBackgroundLog("[startup-bg] 計画確認を走査中…");
        RequestFormPipelineCheckTabController tab = host.requestFormPipelineCheckTab();
        if (tab == null) {
            completeAll();
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(() -> finishStep("計画確認", ok, this::completeAll)));
    }

    private void finishStep(String label, boolean ok, Runnable next) {
        host.appendStartupBackgroundLog(
                "[startup-bg] " + label + (ok ? " 読込完了" : " 読込失敗（続行）"));
        next.run();
    }

    private void completeAll() {
        host.setStartupTabBackgroundLoadActive(false);
        host.setStartupBackgroundLoadStatus("");
        host.appendStartupBackgroundLog("[startup-bg] 起動後バックグラウンド読込完了");
        host.onStartupBackgroundLoadFinished();
    }

    private void setStatus(int step, String label) {
        host.setStartupBackgroundLoadStatus(
                "起動後読込 (" + step + "/" + STEP_COUNT + "): " + label + "…");
    }
}
