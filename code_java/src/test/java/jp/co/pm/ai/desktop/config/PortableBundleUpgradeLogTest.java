package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableBundleUpgradeLogTest {

    @Test
    void open_writesHeaderAndLines(@TempDir Path tmp) throws Exception {
        Path install = tmp.resolve("install");
        Path pm = install.resolve("pm-ai-data");
        Files.createDirectories(pm);
        PortableBundleUpgradeLog log = PortableBundleUpgradeLog.open(install, pm);
        Path file = log.logFile();
        log.appendLine("[portable-sync] 同期: code/a.txt");
        log.close(true, "完了");
        String text = Files.readString(file, StandardCharsets.UTF_8);
        assertTrue(text.contains("ポータルバージョンアップ ログ開始"));
        assertTrue(text.contains("同期: code/a.txt"));
        assertTrue(text.contains("正常終了"));
        assertEquals(
                pm.resolve(PortableBundleUpgradeLog.LOG_DIR_UNDER_PM_AI_DATA).normalize(),
                file.getParent().normalize());
    }

    @Test
    void resolveLogDirectory_usesPmAiDataCodeLog(@TempDir Path tmp) {
        Path install = tmp.resolve("install");
        Path pm = install.resolve("pm-ai-data");
        Path dir = PortableBundleUpgradeLog.resolveLogDirectory(install, pm);
        assertEquals(
                pm.resolve("code/log").toAbsolutePath().normalize(),
                dir.toAbsolutePath().normalize());
    }
}
