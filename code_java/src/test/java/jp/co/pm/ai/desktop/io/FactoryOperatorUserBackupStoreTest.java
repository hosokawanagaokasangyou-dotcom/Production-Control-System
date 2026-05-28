package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

class FactoryOperatorUserBackupStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolate() throws Exception {
        System.setProperty("pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.bin").toString());
        System.setProperty("pm.ai.test.factoryOperatorUserBackupRoot", tmp.resolve("backups").toString());
        FactoryOperatorUserStore.resetStoreForTests();
    }

    @AfterEach
    void clearProps() {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        System.clearProperty("pm.ai.test.factoryOperatorUserBackupRoot");
    }

    private static Map<String, String> ui() {
        return Map.of(
                AppPaths.KEY_PM_AI_REPO_ROOT,
                "/tmp/repo",
                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                "/tmp/shared/summary.xlsx");
    }

    @Test
    void createManualBackupAndRestore() throws Exception {
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "テスト");
        FactoryOperatorUserStore.ensureStoreFileOnDisk();

        var created = FactoryOperatorUserBackupStore.createManualBackup(ui(), "snapshot-1");
        assertEquals("snapshot-1", created.label());
        List<FactoryOperatorUserBackupStore.FactoryOperatorUserBackupEntry> index =
                FactoryOperatorUserBackupStore.loadIndex(ui());
        assertEquals(1, index.size());
        assertTrue(Files.isRegularFile(created.resolveBackupFile(FactoryOperatorUserBackupStore.resolveBackupsRoot(ui()))));

        FactoryOperatorUserStore.addName(FactorySite.KONAN, "変更後");
        assertTrue(FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("変更後"));

        FactoryOperatorUserBackupStore.restoreFromBackup(created, ui());
        assertTrue(FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("テスト"));
        assertTrue(!FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("変更後"));
    }

    @Test
    void trimRemovesOldestBackups() throws Exception {
        FactoryOperatorUserStore.ensureStoreFileOnDisk();
        for (int i = 0; i < FactoryOperatorUserBackupStore.MAX_BACKUP_GENERATIONS + 2; i++) {
            FactoryOperatorUserBackupStore.createManualBackup(ui(), "b-" + i);
        }
        List<FactoryOperatorUserBackupStore.FactoryOperatorUserBackupEntry> index =
                FactoryOperatorUserBackupStore.loadIndex(ui());
        assertEquals(FactoryOperatorUserBackupStore.MAX_BACKUP_GENERATIONS, index.size());
        assertFalse(index.stream().anyMatch(e -> "b-0".equals(e.label())));
        assertTrue(index.stream().anyMatch(e -> "b-31".equals(e.label())));
    }
}
