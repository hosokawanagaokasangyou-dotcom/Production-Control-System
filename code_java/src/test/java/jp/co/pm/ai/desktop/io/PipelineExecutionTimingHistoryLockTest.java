package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

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
    void lockFileNameBesideHistoryJson() {
        Path history = temp.resolve("shared").resolve(AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON);
        Path lock = PipelineExecutionTimingHistoryLock.lockFilePath(history);
        assertEquals(
                AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON + PipelineExecutionTimingHistoryLock.LOCK_SUFFIX,
                lock.getFileName().toString());
        assertEquals(history.getParent(), lock.getParent());
    }
}
