package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class PipelineExecutionTimingHistoryLockTest {

    @TempDir
    Path temp;

    @Test
    void acquireReleaseAndLockBesideHistoryJson() throws Exception {
        Path history =
                temp.resolve(AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON).toAbsolutePath().normalize();
        Files.createDirectories(temp);
        assertFalse(PipelineExecutionTimingHistoryLock.isLocked(history));

        try (PipelineExecutionTimingHistoryLock.AcquiredLock lock =
                PipelineExecutionTimingHistoryLock.tryAcquire(history).orElseThrow()) {
            assertTrue(PipelineExecutionTimingHistoryLock.isLocked(history));
            var info = PipelineExecutionTimingHistoryLock.readLockInfo(history).orElseThrow();
            assertEquals(history.toString(), info.historyPath());
            assertFalse(info.host().isBlank());
            assertTrue(info.hostIp() != null);
        }
        assertFalse(PipelineExecutionTimingHistoryLock.isLocked(history));
    }

    @Test
    void breakStaleLockIfNeeded_removesExpiredLock(@TempDir Path temp) throws Exception {
        Path history =
                temp.resolve(AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON).toAbsolutePath().normalize();
        Path lock = PipelineExecutionTimingHistoryLock.lockFilePath(history);
        Files.createDirectories(lock.getParent());
        Instant stale = Instant.now().minusSeconds(600);
        String payload =
                "version=1\n"
                        + "history="
                        + history
                        + "\n"
                        + "host=other-pc\n"
                        + "hostIp=10.0.0.99\n"
                        + "pid=99999\n"
                        + "user=test\n"
                        + "startedAt="
                        + stale
                        + "\n";
        Files.writeString(lock, payload, StandardCharsets.UTF_8);
        assertTrue(PipelineExecutionTimingHistoryLock.breakStaleLockIfNeeded(lock));
        assertFalse(Files.exists(lock));
    }

    @Test
    void lockFileNameBesideHistoryJson() {
        Path history = temp.resolve("shared").resolve(AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON);
        Path lock = PipelineExecutionTimingHistoryLock.lockFilePath(history);
        assertEquals(
                AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON + PipelineExecutionTimingHistoryLock.LOCK_SUFFIX,
                lock.getFileName().toString());
        assertEquals(history.getParent(), lock.getParent());
    }
}
