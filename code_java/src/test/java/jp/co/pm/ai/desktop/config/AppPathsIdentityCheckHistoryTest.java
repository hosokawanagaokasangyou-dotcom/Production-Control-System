package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppPathsIdentityCheckHistoryTest {

    @Test
    void resolveIdentityCheckHistoryRoot_isSiblingOfSummaryWorkbook(@TempDir Path temp) {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        temp.resolve("shared").resolve("サマリ.xlsx").toString());
        Path root = AppPaths.resolveIdentityCheckHistoryRoot(ui);
        assertEquals(
                temp.resolve("shared").resolve("同一化チェック履歴").toAbsolutePath().normalize(),
                root.toAbsolutePath().normalize());
    }
}
