package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

class SummaryAiDispatchExportLockTest {

    @TempDir
    Path temp;

    @BeforeEach
    void isolateOperatorStore() throws Exception {
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore",
                temp.resolve("operators.json").toString());
        FactoryOperatorUserStore.resetStoreForTests();
    }

    @AfterEach
    void clearOperatorStoreProperty() {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        FactoryOperatorUserStore.clearSessionOperatorName();
    }

    @Test
    void acquireReleaseAndForceRemove() throws Exception {
        Path workbook = temp.resolve("サマリ_AI配台.xlsx");
        Files.createDirectories(temp);
        FactoryOperatorUserStore.writeRawJsonForTests(
                """
                {
                  "schemaVersion": 1,
                  "factories": {
                    "KONAN": {
                      "names": ["砂田", "古家", "図司", "細川"],
                      "lastSelected": "砂田"
                    },
                    "KOKUBU": {
                      "names": ["砂田", "古家", "図司", "細川"],
                      "lastSelected": ""
                    }
                  }
                }
                """);
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        assertFalse(SummaryAiDispatchExportLock.isLocked(workbook));

        try (SummaryAiDispatchExportLock.AcquiredLock lock =
                SummaryAiDispatchExportLock.tryAcquire(workbook).orElseThrow()) {
            assertTrue(SummaryAiDispatchExportLock.isLocked(workbook));
            var info = SummaryAiDispatchExportLock.readLockInfo(workbook).orElseThrow();
            assertEquals(workbook.toAbsolutePath().normalize().toString(), info.workbook());
            assertFalse(info.host().isBlank());
            assertEquals("砂田", info.operator());
            assertEquals("砂田@" + info.host(), info.displayCreator());
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
