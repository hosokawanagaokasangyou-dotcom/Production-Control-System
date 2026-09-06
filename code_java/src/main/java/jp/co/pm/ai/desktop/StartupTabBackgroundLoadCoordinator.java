package jp.co.pm.ai.desktop;

import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.BooleanSupplier;

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

        /** 環境変数の起動チェック完了かつ初期化済みのときのみ true。 */
        boolean canScheduleStartupBackgroundLoad();

        /** 工場切替完了後。env 差分ブロック中でも true になりうる。 */
        boolean canScheduleFactorySwitchBackgroundLoad();
    }

    private static final int STEP_REMOTE = 1;
    private static final int STEP_COMPANY = 2;
    private static final int STEP_MEMBER = 3;
    private static final int STEP_MACHINE = 4;
    private static final int STEP_REQUEST_FORM = 5;
    private static final int STEP_PIPELINE_CHECK = 6;
    private static final int STEP_COUNT = 6;

    /** ユーザーがモーダルを閉じたあとのステップ間待機（UI 操作を優先）。 */
    static final long DEFERRED_STEP_YIELD_MS = 150L;

    private final Host host;
    private final AtomicBoolean runScheduled = new AtomicBoolean(false);
    /** {@link #cancelForFactorySwitch()} で増加。実行中チェーンが無効になったら後続ステップを進めない。 */
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private volatile long activeRunGeneration = -1L;
    /** 起動時チェックを閉じ、低優先度で読込を継続中。 */
    private final AtomicBoolean deferredLowPriority = new AtomicBoolean(false);

    StartupTabBackgroundLoadCoordinator(Host host) {
        this.host = host;
    }

    /**
     * 工場切替開始時に呼ぶ。進行中の起動後バックグラウンド読込チェーンを中断し、工場切替を優先する。
     */
    void cancelForFactorySwitch() {
        cancel("[startup-bg] 工場切替のためバックグラウンド読込を中断");
    }

    /**
     * 起動時チェックの「バックグラウンドで続行」。読込チェーンは止めず、優先度を下げて継続する。
     *
     * @return 進行中の読込を初めて BG 継続へ切り替えたとき {@code true}
     */
    boolean deferToBackgroundByUser() {
        if (!host.isStartupTabBackgroundLoadActive()) {
            return false;
        }
        if (!deferredLowPriority.compareAndSet(false, true)) {
            return false;
        }
        host.appendStartupBackgroundLog(
                "[startup-bg] ユーザー操作により優先度を下げてバックグラウンド読込を継続");
        return true;
    }

    boolean isDeferredLowPriority() {
        return deferredLowPriority.get();
    }

    private boolean cancel(String logLine) {
        loadGeneration.incrementAndGet();
        runScheduled.set(false);
        activeRunGeneration = -1L;
        deferredLowPriority.set(false);
        if (!host.isStartupTabBackgroundLoadActive()) {
            return false;
        }
        host.setStartupTabBackgroundLoadActive(false);
        host.setStartupBackgroundLoadStatus("");
        host.appendStartupBackgroundLog(logLine);
        return true;
    }

    private boolean isRunObsolete() {
        return activeRunGeneration != loadGeneration.get();
    }

    /** 未実行なら読込チェーンを開始する。環境変数初期化未完了時は何もしない。 */
    void scheduleIfIdle() {
        scheduleIfIdle(host::canScheduleStartupBackgroundLoad);
    }

    private void scheduleIfIdle(BooleanSupplier gate) {
        if (!gate.getAsBoolean()) {
            return;
        }
        if (!runScheduled.compareAndSet(false, true)) {
            return;
        }
        deferredLowPriority.set(false);
        activeRunGeneration = loadGeneration.get();
        host.setStartupTabBackgroundLoadActive(true);
        Platform.runLater(this::beginRemoteDesktop);
    }

    /** 工場切替・ワークスペース復元後に再読込する。 */
    void resetAndSchedule() {
        runScheduled.set(false);
        scheduleIfIdle();
    }

    /** 工場切替完了後に再読込する（env 差分ブロック中でも実行）。 */
    void resetAndScheduleAfterFactorySwitch() {
        runScheduled.set(false);
        scheduleIfIdle(host::canScheduleFactorySwitchBackgroundLoad);
    }

    private void beginRemoteDesktop() {
        if (isRunObsolete()) {
            return;
        }
        setStatus(STEP_REMOTE, "リモートデスクトップ");
        host.appendStartupBackgroundLog("[startup-bg] リモートデスクトップを読込中…");
        RemoteDesktopTabController rdp = host.remoteDesktopTab();
        if (rdp == null) {
            runNextStep(this::beginCompanyCalendar);
            return;
        }
        rdp.scheduleBackgroundPreload(() -> runNextStep(this::beginCompanyCalendar));
    }

    private void beginCompanyCalendar() {
        if (isRunObsolete()) {
            return;
        }
        setStatus(STEP_COMPANY, "会社カレンダー");
        host.appendStartupBackgroundLog("[startup-bg] 会社カレンダーを読込中…");
        CompanyCalendarTabController tab = host.companyCalendarTab();
        if (tab == null) {
            runNextStep(this::beginMemberAttendance);
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(
                        () -> finishStep("会社カレンダー", ok, this::beginMemberAttendance)));
    }

    private void beginMemberAttendance() {
        if (isRunObsolete()) {
            return;
        }
        setStatus(STEP_MEMBER, "メンバー勤怠");
        host.appendStartupBackgroundLog("[startup-bg] メンバー勤怠を読込中…");
        MemberAttendanceTabController tab = host.memberAttendanceTab();
        if (tab == null) {
            runNextStep(this::beginMachineCalendar);
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(
                        () -> finishStep("メンバー勤怠", ok, this::beginMachineCalendar)));
    }

    private void beginMachineCalendar() {
        if (isRunObsolete()) {
            return;
        }
        setStatus(STEP_MACHINE, "機械カレンダー");
        host.appendStartupBackgroundLog("[startup-bg] 機械カレンダーを読込中…");
        MachineCalendarTabController tab = host.machineCalendarTab();
        if (tab == null) {
            runNextStep(this::beginRequestFormInput);
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(
                        () -> finishStep("機械カレンダー", ok, this::beginRequestFormInput)));
    }

    private void beginRequestFormInput() {
        if (isRunObsolete()) {
            return;
        }
        setStatus(STEP_REQUEST_FORM, "原本転記");
        host.appendStartupBackgroundLog("[startup-bg] 原本転記を読込中…");
        RequestFormInputTabController tab = host.requestFormInputTab();
        if (tab == null) {
            runNextStep(this::beginPipelineCheck);
            return;
        }
        tab.preloadInBackground(
                ok -> Platform.runLater(
                        () -> finishStep("原本転記", ok, this::beginPipelineCheck)));
    }

    private void beginPipelineCheck() {
        if (isRunObsolete()) {
            return;
        }
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
        if (isRunObsolete()) {
            return;
        }
        host.appendStartupBackgroundLog(
                "[startup-bg] " + label + (ok ? " 読込完了" : " 読込失敗（続行）"));
        runNextStep(next);
    }

    /**
     * 次ステップへ進む。ユーザーが BG 継続へ切り替えたあとは短い yield を入れて UI を優先する。
     */
    private void runNextStep(Runnable next) {
        if (next == null) {
            return;
        }
        if (!deferredLowPriority.get()) {
            next.run();
            return;
        }
        Thread yielder =
                new Thread(
                        () -> {
                            try {
                                Thread.sleep(DEFERRED_STEP_YIELD_MS);
                            } catch (InterruptedException e) {
                                Thread.currentThread().interrupt();
                                return;
                            }
                            Platform.runLater(
                                    () -> {
                                        if (!isRunObsolete()) {
                                            next.run();
                                        }
                                    });
                        },
                        "startup-bg-deferred-yield");
        yielder.setDaemon(true);
        yielder.setPriority(Thread.MIN_PRIORITY);
        yielder.start();
    }

    private void completeAll() {
        if (isRunObsolete()) {
            return;
        }
        deferredLowPriority.set(false);
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
