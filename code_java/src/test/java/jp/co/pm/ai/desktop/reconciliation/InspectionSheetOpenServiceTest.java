package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

class InspectionSheetOpenServiceTest {

    private String priorHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        GlobalInitSettingTarget.save(FactorySite.KONAN);
    }

    @AfterEach
    void tearDown() {
        InspectionSheetOpenService.joinBackgroundRebuildForTest();
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void rebuildAndFind(@TempDir Path root) throws Exception {
        Map<String, String> ui = uiWithSheet(writeSampleWorkbook(root));
        InspectionSheetOpenService.RebuildResult rebuilt = InspectionSheetOpenService.rebuild(ui, null);
        assertEquals(1, rebuilt.rows().size());
        List<InspectionSheetIndexStore.Row> hits = InspectionSheetOpenService.find(ui, "c8-9");
        assertEquals(1, hits.size());
        assertEquals("C8-9", hits.get(0).iraiNo());
        assertTrue(Files.isRegularFile(InspectionSheetIndexStore.indexFile(FactorySite.KONAN)));
    }

    @Test
    void startBackgroundRebuild_writesCsv(@TempDir Path root) throws Exception {
        Map<String, String> ui = uiWithSheet(writeSampleWorkbook(root));
        InspectionSheetOpenService.RebuildResult rebuilt =
                InspectionSheetOpenService.startBackgroundRebuild(ui).get(20, TimeUnit.SECONDS);
        assertEquals(1, rebuilt.rows().size());
        assertTrue(Files.isRegularFile(InspectionSheetIndexStore.indexFile(FactorySite.KONAN)));
        List<InspectionSheetIndexStore.Row> hits = InspectionSheetOpenService.find(ui, "C8-9");
        assertEquals(1, hits.size());
    }

    @Test
    void startBackgroundRebuild_coalescesInFlight(@TempDir Path root) throws Exception {
        Map<String, String> ui = uiWithSheet(writeSampleWorkbook(root));
        CountDownLatch inScan = new CountDownLatch(1);
        CountDownLatch release = new CountDownLatch(1);
        InspectionSheetIndexScanner.Progress hold =
                (done, total) -> {
                    inScan.countDown();
                    try {
                        assertTrue(release.await(20, TimeUnit.SECONDS));
                    } catch (InterruptedException ex) {
                        Thread.currentThread().interrupt();
                    }
                };
        try {
            CompletableFuture<InspectionSheetOpenService.RebuildResult> first =
                    InspectionSheetOpenService.startBackgroundRebuild(ui, hold);
            assertTrue(inScan.await(20, TimeUnit.SECONDS));
            CompletableFuture<InspectionSheetOpenService.RebuildResult> second =
                    InspectionSheetOpenService.startBackgroundRebuild(ui);
            assertSame(first, second);
            release.countDown();
            assertEquals(1, first.get(20, TimeUnit.SECONDS).rows().size());
        } finally {
            release.countDown();
        }
    }

    @Test
    void startBackgroundRebuild_unreachableDir_completesEmpty(@TempDir Path tmp) throws Exception {
        Path missing = tmp.resolve("no-such-inspection-dir");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_INSPECTION_SHEET_DIR, missing.toString());
        InspectionSheetOpenService.RebuildResult result =
                InspectionSheetOpenService.startBackgroundRebuild(ui).get(5, TimeUnit.SECONDS);
        assertTrue(result.rows().isEmpty());
    }

    @Test
    void find_emptyIndex_joinsBackgroundRebuild(@TempDir Path root) throws Exception {
        Map<String, String> ui = uiWithSheet(writeSampleWorkbook(root));
        CountDownLatch inScan = new CountDownLatch(1);
        CountDownLatch release = new CountDownLatch(1);
        InspectionSheetIndexScanner.Progress hold =
                (done, total) -> {
                    inScan.countDown();
                    try {
                        assertTrue(release.await(20, TimeUnit.SECONDS));
                    } catch (InterruptedException ex) {
                        Thread.currentThread().interrupt();
                    }
                };
        try {
            CompletableFuture<InspectionSheetOpenService.RebuildResult> warmup =
                    InspectionSheetOpenService.startBackgroundRebuild(ui, hold);
            assertTrue(inScan.await(20, TimeUnit.SECONDS));
            Thread finder =
                    new Thread(
                            () -> {
                                try {
                                    List<InspectionSheetIndexStore.Row> hits =
                                            InspectionSheetOpenService.find(ui, "c8-9");
                                    assertEquals(1, hits.size());
                                } catch (Exception ex) {
                                    throw new RuntimeException(ex);
                                }
                            },
                            "find-during-warmup");
            finder.start();
            Thread.sleep(200);
            assertTrue(finder.isAlive(), "find はバックグラウンド索引完了を待つ");
            release.countDown();
            finder.join(20_000);
            assertEquals(1, warmup.get(20, TimeUnit.SECONDS).rows().size());
        } finally {
            release.countDown();
        }
    }

    private static Map<String, String> uiWithSheet(Path root) {
        return Map.of(AppPaths.KEY_PM_AI_INSPECTION_SHEET_DIR, root.toString());
    }

    private static Path writeSampleWorkbook(Path root) throws Exception {
        Path month = root.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        Path xlsx = month.resolve("2026_C8-9(SEC済)完了.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("検査表");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日");
            r2.createCell(28).setCellValue(46261);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№");
            r3.createCell(3).setCellValue("C8-9");
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }
        return root;
    }
}
