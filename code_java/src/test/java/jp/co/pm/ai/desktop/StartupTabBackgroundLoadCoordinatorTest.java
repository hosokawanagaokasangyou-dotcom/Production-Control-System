package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.atomic.AtomicBoolean;

import org.junit.jupiter.api.Test;

class StartupTabBackgroundLoadCoordinatorTest {

    @Test
    void scheduleIfIdle_skipsWhenEnvNotReady() {
        AtomicBoolean canSchedule = new AtomicBoolean(false);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canSchedule, active));

        coordinator.scheduleIfIdle();
        coordinator.resetAndSchedule();

        assertFalse(active.get(), "環境変数未初期化時は読込チェーンを開始しない");
    }

    @Test
    void cancelForFactorySwitch_clearsActiveAndAllowsReschedule() {
        AtomicBoolean canSchedule = new AtomicBoolean(true);
        AtomicBoolean active = new AtomicBoolean(false);
        StartupTabBackgroundLoadCoordinator coordinator =
                new StartupTabBackgroundLoadCoordinator(
                        new StubHost(canSchedule, active));

        coordinator.scheduleIfIdle();
        assertTrue(active.get(), "読込チェーン開始で active になる");

        coordinator.cancelForFactorySwitch();
        assertFalse(active.get(), "中断後は active が false");

        coordinator.resetAndSchedule();
        assertTrue(active.get(), "工場切替後に再スケジュール可能");
    }

    private static final class StubHost implements StartupTabBackgroundLoadCoordinator.Host {
        private final AtomicBoolean canSchedule;
        private final AtomicBoolean active;

        StubHost(AtomicBoolean canSchedule, AtomicBoolean active) {
            this.canSchedule = canSchedule;
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
            return canSchedule.get();
        }
    }
}
