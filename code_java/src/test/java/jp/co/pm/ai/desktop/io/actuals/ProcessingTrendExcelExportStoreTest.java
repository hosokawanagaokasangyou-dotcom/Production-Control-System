package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class ProcessingTrendExcelExportStoreTest {

    @Test
    void resolveDirectory_underPmAiCacheExports(@TempDir Path repo) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        Path dir = ProcessingTrendExcelExportStore.resolveDirectory(ui);
        assertEquals(
                repo.resolve(".pm-ai-cache").resolve("exports").resolve("processing-trend").normalize(),
                dir.normalize());
    }

    @Test
    void deleteAllXlsx_removesOnlyXlsx(@TempDir Path dir) throws Exception {
        Path a = dir.resolve("old.xlsx");
        Path b = dir.resolve("keep.txt");
        Path c = dir.resolve("other.XLSX");
        Files.writeString(a, "a", StandardCharsets.UTF_8);
        Files.writeString(b, "b", StandardCharsets.UTF_8);
        Files.writeString(c, "c", StandardCharsets.UTF_8);

        ProcessingTrendExcelExportStore.deleteAllXlsx(dir);

        assertFalse(Files.exists(a));
        assertFalse(Files.exists(c));
        assertTrue(Files.exists(b));
    }

    @Test
    void findNewestXlsx_returnsLatestByMtime(@TempDir Path dir) throws Exception {
        Path older = dir.resolve("a.xlsx");
        Path newer = dir.resolve("b.xlsx");
        Files.writeString(older, "1", StandardCharsets.UTF_8);
        Thread.sleep(30);
        Files.writeString(newer, "2", StandardCharsets.UTF_8);

        Optional<Path> found = ProcessingTrendExcelExportStore.findNewestXlsx(dir);
        assertTrue(found.isPresent());
        assertEquals(newer.getFileName(), found.get().getFileName());
    }

    @Test
    void findNewestXlsx_emptyWhenNone(@TempDir Path dir) {
        assertTrue(ProcessingTrendExcelExportStore.findNewestXlsx(dir).isEmpty());
    }

    @Test
    void prepareTarget_clearsOldThenReturnsPath(@TempDir Path dir) throws Exception {
        Files.writeString(dir.resolve("stale.xlsx"), "old", StandardCharsets.UTF_8);
        Path target = ProcessingTrendExcelExportStore.prepareTarget(dir, "fresh-20260908.xlsx");
        assertEquals(dir.resolve("fresh-20260908.xlsx"), target);
        assertFalse(Files.exists(dir.resolve("stale.xlsx")));
        assertFalse(Files.exists(target), "prepare はパスを返すだけでまだ書かない");
    }
}
