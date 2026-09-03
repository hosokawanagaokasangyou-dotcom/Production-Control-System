package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class StartupTabBackgroundLoadCoordinatorTest {

    @BeforeAll
    static void initFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void scheduleIfIdle_skipsWhenEnvNotReady() {
        AtomicBoolean canStartup = new AtomicBoolean(false);
        AtomicBoolean canFactorySwitch = new AtomicBoolean(false);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canStartup, canFactorySwitch, active));

        coordinator.scheduleIfIdle();
        coordinator.resetAndSchedule();

        assertFalse(active.get(), "環境変数未初期化時は読込チェーンを開始しない");
    }

    @Test
    void resetAndScheduleAfterFactorySwitch_allowsWhenStartupGateBlocked() {
        AtomicBoolean canStartup = new AtomicBoolean(false);
        AtomicBoolean canFactorySwitch = new AtomicBoolean(true);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canStartup, canFactorySwitch, active));

        coordinator.resetAndSchedule();
        assertFalse(active.get(), "起動時ゲートが閉じているときは通常の再スケジュールは開始しない");

        coordinator.resetAndScheduleAfterFactorySwitch();
        assertTrue(active.get(), "工場切替後は専用ゲートで再スケジュールできる");
    }

    @Test
    void cancelForFactorySwitch_clearsActiveAndAllowsReschedule() {
        AtomicBoolean canStartup = new AtomicBoolean(true);
        AtomicBoolean canFactorySwitch = new AtomicBoolean(true);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canStartup, canFactorySwitch, active));

        coordinator.scheduleIfIdle();
        assertTrue(active.get(), "読込チェーン開始で active になる");

        coordinator.cancelForFactorySwitch();
        assertFalse(active.get(), "中断後は active が false");

        coordinator.resetAndSchedule();
        assertTrue(active.get(), "工場切替後に再スケジュール可能");
    }

    @Test
    void cancelByUser_stopsChainAndReportsCancelled() {
        AtomicBoolean canStartup = new AtomicBoolean(true);
        AtomicBoolean canFactorySwitch = new AtomicBoolean(true);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canStartup, canFactorySwitch, active));

        coordinator.scheduleIfIdle();
        assertTrue(active.get(), "読込チェーン開始で active になる");

        assertTrue(coordinator.cancelByUser(), "進行中の読込を中断したら true");
        assertFalse(active.get(), "キャンセル後は active が false");
        assertFalse(coordinator.cancelByUser(), "既に停止しているときは false");
    }

    private static final class StubHost implements StartupTabBackgroundLoadCoordinator.Host {
        private final AtomicBoolean canStartup;
        private final AtomicBoolean canFactorySwitch;
        private final AtomicBoolean active;

        StubHost(
                AtomicBoolean canStartup,
                AtomicBoolean canFactorySwitch,
                AtomicBoolean active) {
            this.canStartup = canStartup;
            this.canFactorySwitch = canFactorySwitch;
            this.active = active;
        }

        @Override
        public void setStartupBackgroundLoadStatus(String message) {}

        @Override
        public void appendStartupBackgroundLog(String line) {}

        @Override
        public RemoteDesktopTabController remoteDesktopTab() {
            return null;
        }

        @Override
        public CompanyCalendarTabController companyCalendarTab() {
            return null;
        }

        @Override
        public MemberAttendanceTabController memberAttendanceTab() {
            return null;
        }

        @Override
        public MachineCalendarTabController machineCalendarTab() {
            return null;
        }

        @Override
        public RequestFormInputTabController requestFormInputTab() {
            return null;
        }

        @Override
        public RequestFormPipelineCheckTabController requestFormPipelineCheckTab() {
            return null;
        }

        @Override
        public void onStartupBackgroundLoadFinished() {}

        @Override
        public void setStartupTabBackgroundLoadActive(boolean activeFlag) {
            active.set(activeFlag);
        }

        @Override
        public boolean isStartupTabBackgroundLoadActive() {
            return active.get();
        }

        @Override
        public boolean canScheduleStartupBackgroundLoad() {
            return canStartup.get();
        }

        @Override
        public boolean canScheduleFactorySwitchBackgroundLoad() {
            return canFactorySwitch.get();
        }
    }
}
