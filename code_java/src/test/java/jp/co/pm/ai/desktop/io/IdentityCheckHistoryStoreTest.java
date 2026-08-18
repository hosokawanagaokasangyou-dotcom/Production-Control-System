package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.ZoneOffset;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class IdentityCheckHistoryStoreTest {

    @Test
    void save_writesExcelPlanJsonAndMeta(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path excel = temp.resolve("src.xlsx");
        Files.write(excel, new byte[] {1, 2, 3, 4});
        PlanInputTabularIo.TabularSheet tab =
                new PlanInputTabularIo.TabularSheet(
                        List.of("機械名", "依頼NO"), List.of(List.of("M1", "T1")));

        Optional<Path> dest =
                IdentityCheckHistoryStore.save(
                        ui,
                        excel,
                        tab,
                        "mismatch",
                        "差異 2件",
                        2,
                        Optional.of(excel),
                        Optional.of(temp.resolve("plan-source.xlsx")));

        assertTrue(dest.isPresent());
        Path dir = dest.get();
        assertTrue(Files.isRegularFile(dir.resolve(IdentityCheckHistoryStore.EXCEL_FILE)));
        assertTrue(Files.isRegularFile(dir.resolve(IdentityCheckHistoryStore.PLAN_JSON_FILE)));
        assertTrue(Files.isRegularFile(dir.resolve(IdentityCheckHistoryStore.META_FILE)));
        JsonTableIo.ArrayTable loaded =
                JsonTableIo.loadArrayTable(dir.resolve(IdentityCheckHistoryStore.PLAN_JSON_FILE));
        assertEquals(List.of("機械名", "依頼NO"), loaded.columns());
        assertEquals(1, loaded.rows().size());
        Optional<IdentityCheckHistoryStore.Meta> meta = IdentityCheckHistoryStore.readMeta(dir);
        assertTrue(meta.isPresent());
        assertEquals("mismatch", meta.get().result());
        assertEquals(2, meta.get().diffCount());
    }

    @Test
    void save_returnsEmptyWhenExcelMissing(@TempDir Path temp) {
        Map<String, String> ui = testUi(temp);
        PlanInputTabularIo.TabularSheet tab =
                new PlanInputTabularIo.TabularSheet(List.of("依頼NO"), List.of(List.of("T1")));

        Optional<Path> dest =
                IdentityCheckHistoryStore.save(
                        ui,
                        temp.resolve("missing.xlsx"),
                        tab,
                        "ok",
                        "同一",
                        0,
                        Optional.empty(),
                        Optional.empty());

        assertTrue(dest.isEmpty());
    }

    @Test
    void prune_keepsNewestTwentyByFolderNameDespiteMtime(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path opDir = IdentityCheckHistoryStore.resolveOperatorDir(ui, "テスト太郎");
        Files.createDirectories(opDir);
        for (int i = 0; i < 22; i++) {
            Path d =
                    IdentityCheckHistoryStore.uniqueSnapshotDir(
                            opDir, OffsetDateTime.of(2026, 8, 18, 10, 0, i, 0, ZoneOffset.ofHours(9)));
            Files.createDirectories(d);
            Files.writeString(d.resolve(IdentityCheckHistoryStore.META_FILE), "{}");
            // 意図的に mtime を逆順にして、フォルダ名順 prune であることを検証する。
            Files.setLastModifiedTime(
                    d, java.nio.file.attribute.FileTime.fromMillis(100_000L - i));
        }
        IdentityCheckHistoryStore.prune(opDir);
        try (var stream = Files.list(opDir)) {
            List<String> names =
                    stream.filter(Files::isDirectory)
                            .map(p -> p.getFileName().toString())
                            .sorted()
                            .toList();
            assertEquals(20, names.size());
            assertEquals("20260818-100002", names.getFirst());
            assertEquals("20260818-100021", names.getLast());
        }
    }

    @Test
    void prune_deletesIncompleteDirsWithoutMeta(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path opDir = IdentityCheckHistoryStore.resolveOperatorDir(ui, "テスト太郎");
        Files.createDirectories(opDir);
        Path incomplete = opDir.resolve("20260818-120000");
        Files.createDirectories(incomplete);
        Files.write(incomplete.resolve(IdentityCheckHistoryStore.EXCEL_FILE), new byte[] {1});
        Path complete =
                IdentityCheckHistoryStore.uniqueSnapshotDir(
                        opDir, OffsetDateTime.of(2026, 8, 18, 12, 0, 1, 0, ZoneOffset.ofHours(9)));
        Files.createDirectories(complete);
        Files.writeString(complete.resolve(IdentityCheckHistoryStore.META_FILE), "{}");

        IdentityCheckHistoryStore.prune(opDir);

        assertFalse(Files.exists(incomplete));
        assertTrue(Files.isDirectory(complete));
    }

    @Test
    void listNewestFirst_ordersByFolderNameNotMtime(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path opDir = IdentityCheckHistoryStore.resolveOperatorDir(ui, "テスト太郎");
        Files.createDirectories(opDir);
        Path older =
                IdentityCheckHistoryStore.uniqueSnapshotDir(
                        opDir, OffsetDateTime.of(2026, 8, 18, 9, 0, 0, 0, ZoneOffset.ofHours(9)));
        Path newer =
                IdentityCheckHistoryStore.uniqueSnapshotDir(
                        opDir, OffsetDateTime.of(2026, 8, 18, 10, 0, 0, 0, ZoneOffset.ofHours(9)));
        Files.createDirectories(older);
        Files.createDirectories(newer);
        Files.writeString(
                older.resolve(IdentityCheckHistoryStore.META_FILE),
                """
                {"savedAt":"2026-08-18T09:00:00+09:00","operator":"テスト太郎","result":"ok","badgeText":"","diffCount":0,"excelSourcePath":"","planSourcePath":"","excelFileName":"配台計画.xlsx","planJsonFileName":"加工計画.json"}
                """);
        Files.writeString(
                newer.resolve(IdentityCheckHistoryStore.META_FILE),
                """
                {"savedAt":"2026-08-18T10:00:00+09:00","operator":"テスト太郎","result":"mismatch","badgeText":"","diffCount":1,"excelSourcePath":"","planSourcePath":"","excelFileName":"配台計画.xlsx","planJsonFileName":"加工計画.json"}
                """);
        Files.setLastModifiedTime(newer, java.nio.file.attribute.FileTime.fromMillis(1_000L));
        Files.setLastModifiedTime(older, java.nio.file.attribute.FileTime.fromMillis(9_000L));

        List<IdentityCheckHistoryStore.SnapshotRef> list =
                IdentityCheckHistoryStore.listNewestFirst(ui, "テスト太郎");
        assertEquals(2, list.size());
        assertEquals("mismatch", list.getFirst().meta().result());
        assertEquals("ok", list.get(1).meta().result());
    }

    @Test
    void prune_keepsNewestTwenty(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path opDir = IdentityCheckHistoryStore.resolveOperatorDir(ui, "テスト太郎");
        Files.createDirectories(opDir);
        for (int i = 0; i < 22; i++) {
            Path d =
                    IdentityCheckHistoryStore.uniqueSnapshotDir(
                            opDir, OffsetDateTime.of(2026, 8, 18, 10, 0, i, 0, ZoneOffset.ofHours(9)));
            Files.createDirectories(d);
            Files.writeString(d.resolve(IdentityCheckHistoryStore.META_FILE), "{}");
        }
        IdentityCheckHistoryStore.prune(opDir);
        try (var stream = Files.list(opDir)) {
            assertEquals(20, stream.filter(Files::isDirectory).count());
        }
    }

    @Test
    void listNewestFirst_andOperatorDirs(@TempDir Path temp) throws Exception {
        Map<String, String> ui = testUi(temp);
        Path excel = temp.resolve("a.xlsx");
        Files.write(excel, new byte[] {9});
        PlanInputTabularIo.TabularSheet tab =
                new PlanInputTabularIo.TabularSheet(List.of("依頼NO"), List.of(List.of("T1")));
        IdentityCheckHistoryStore.save(
                ui, excel, tab, "ok", "同一", 0, Optional.of(excel), Optional.empty());

        List<IdentityCheckHistoryStore.SnapshotRef> list =
                IdentityCheckHistoryStore.listNewestFirst(ui, "テスト太郎");
        assertFalse(list.isEmpty());
        assertEquals("ok", list.getFirst().meta().result());
        assertTrue(IdentityCheckHistoryStore.listOperatorDirNames(ui).contains("テスト太郎"));
    }

    private static Map<String, String> testUi(Path temp) {
        Map<String, String> ui = new HashMap<>();
        ui.put(
                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                temp.resolve("shared").resolve("サマリ.xlsx").toString());
        ui.put(AppPaths.KEY_PM_AI_OPERATOR_USER, "テスト太郎");
        return ui;
    }
}
