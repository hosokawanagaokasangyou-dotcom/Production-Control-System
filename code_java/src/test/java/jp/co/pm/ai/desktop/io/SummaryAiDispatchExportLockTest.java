package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

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
}
