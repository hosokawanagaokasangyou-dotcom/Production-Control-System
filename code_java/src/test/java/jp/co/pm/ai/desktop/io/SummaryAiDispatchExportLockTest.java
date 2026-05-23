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

class SummaryAiDispatchExportLockTest {

    @TempDir
    Path temp;

    @Test
    void acquireReleaseAndForceRemove() throws Exception {
        Path workbook = temp.resolve("サマリ_AI配台.xlsx");
        Files.createDirectories(temp);
        assertFalse(SummaryAiDispatchExportLock.isLocked(workbook));

        try (SummaryAiDispatchExportLock.AcquiredLock lock =
                SummaryAiDispatchExportLock.tryAcquire(workbook).orElseThrow()) {
            assertTrue(SummaryAiDispatchExportLock.isLocked(workbook));
            var info = SummaryAiDispatchExportLock.readLockInfo(workbook).orElseThrow();
            assertEquals(workbook.toAbsolutePath().normalize().toString(), info.workbook());
            assertFalse(info.host().isBlank());
        }
        assertFalse(SummaryAiDispatchExportLock.isLocked(workbook));

        try (SummaryAiDispatchExportLock.AcquiredLock lock =
                SummaryAiDispatchExportLock.tryAcquire(workbook).orElseThrow()) {
            assertTrue(SummaryAiDispatchExportLock.tryAcquire(workbook).isEmpty());
        }
        SummaryAiDispatchExportLock.forceRemove(workbook);
        assertFalse(SummaryAiDispatchExportLock.isLocked(workbook));
    }

    @Test
    void lockFileNameBesideWorkbook() {
        Path workbook = temp.resolve("code").resolve("サマリ_AI配台.xlsx");
        Path lock = SummaryAiDispatchExportLock.lockFilePath(workbook);
        assertEquals("サマリ_AI配台.xlsx.export.lock", lock.getFileName().toString());
        assertEquals(workbook.getParent(), lock.getParent());
    }

    @Test
    void expiredLockIsIgnoredAndRemoved() throws Exception {
        Path workbook = temp.resolve("サマリ_AI配台.xlsx");
        Path lock = SummaryAiDispatchExportLock.lockFilePath(workbook);
        Files.createDirectories(temp);
        Instant staleStarted =
                Instant.now().minus(SummaryAiDispatchExportLock.LOCK_MAX_AGE).minusSeconds(60);
        String payload =
                "version=1\n"
                        + "workbook="
                        + workbook.toAbsolutePath().normalize()
                        + "\nhost=stale-pc\npid=1\nuser=test\nstartedAt="
                        + staleStarted
                        + "\n";
        Files.writeString(lock, payload, StandardCharsets.UTF_8);

        assertFalse(SummaryAiDispatchExportLock.isLocked(workbook));
        assertFalse(Files.isRegularFile(lock));
        assertTrue(SummaryAiDispatchExportLock.readLockInfo(workbook).isEmpty());
        assertTrue(SummaryAiDispatchExportLock.tryAcquire(workbook).isPresent());
    }

    @Test
    void tryAcquireRemovesExpiredLockBeforeCreate() throws Exception {
        Path workbook = temp.resolve("サマリ_AI配台.xlsx");
        Path lock = SummaryAiDispatchExportLock.lockFilePath(workbook);
        Files.createDirectories(temp);
        Instant staleStarted =
                Instant.now().minus(SummaryAiDispatchExportLock.LOCK_MAX_AGE).minusSeconds(30);
        Files.writeString(
                lock,
                "version=1\nworkbook="
                        + workbook.toAbsolutePath().normalize()
                        + "\nhost=old\npid=9\nuser=u\nstartedAt="
                        + staleStarted
                        + "\n",
                StandardCharsets.UTF_8);

        try (SummaryAiDispatchExportLock.AcquiredLock acquired =
                SummaryAiDispatchExportLock.tryAcquire(workbook).orElseThrow()) {
            assertTrue(SummaryAiDispatchExportLock.isLocked(workbook));
            assertEquals(lock, acquired.lockPath());
        }
    }
}
