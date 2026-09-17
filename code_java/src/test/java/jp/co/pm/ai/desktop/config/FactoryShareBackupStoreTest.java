package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FactoryShareBackupStoreTest {

    @Test
    void shouldSkip_backupDirAndExcelLock() {
        assertTrue(FactoryShareBackupStore.shouldSkipRelative(Path.of("バックアップ", "old.txt")));
        assertTrue(FactoryShareBackupStore.shouldSkipRelative(Path.of("remote_log", "バックアップ", "old.txt")));
        assertTrue(FactoryShareBackupStore.shouldSkipRelative(Path.of("~$lock.xlsx")));
        assertTrue(FactoryShareBackupStore.shouldSkipRelative(Path.of("remote_log", "~$a.xlsx")));
        assertFalse(FactoryShareBackupStore.shouldSkipRelative(Path.of("remote_log", "ui.txt")));
        assertFalse(FactoryShareBackupStore.shouldSkipRelative(Path.of("サマリ_AI配台.xlsx")));
    }

    @Test
    void resolveSharedBackupRoot_usesEachFactoryShare(@TempDir Path temp) {
        Path konanShare = temp.resolve("konan-share");
        Path kokubuShare = temp.resolve("kokubu-share");
        Map<String, String> konanUi =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KONAN.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        konanShare.toString());
        Map<String, String> kokubuUi =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KOKUBU.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        kokubuShare.toString());
        assertEquals(
                konanShare.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME).toAbsolutePath().normalize(),
                FactoryShareBackupStore.resolveSharedBackupRoot(konanUi, FactorySite.KONAN));
        assertEquals(
                kokubuShare.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME).toAbsolutePath().normalize(),
                FactoryShareBackupStore.resolveSharedBackupRoot(kokubuUi, FactorySite.KOKUBU));
    }

    @Test
    void resolveSource_konanUiWithoutHint_doesNotReuseSameFolderForKokubu(@TempDir Path temp)
            throws Exception {
        Path konanOverride = temp.resolve("generic-share");
        Files.createDirectories(konanOverride);
        Map<String, String> konanUi =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KONAN.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        konanOverride.toString());
        Path konanSrc = FactoryShareBackupStore.resolveSource(konanUi, FactorySite.KONAN);
        Path kokubuSrc = FactoryShareBackupStore.resolveSource(konanUi, FactorySite.KOKUBU);
        assertEquals(konanOverride.toAbsolutePath().normalize(), konanSrc);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR).toAbsolutePath().normalize(), kokubuSrc);
        assertNotEquals(konanSrc, kokubuSrc);
    }

    @Test
    void resolveSource_konanProductionUnc_resolvesKokubuToFactoryDefault() {
        Map<String, String> konanUi =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KONAN.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        AppPaths.DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KONAN);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR).toAbsolutePath().normalize(),
                FactoryShareBackupStore.resolveSource(konanUi, FactorySite.KONAN));
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR).toAbsolutePath().normalize(),
                FactoryShareBackupStore.resolveSource(konanUi, FactorySite.KOKUBU));
    }

    @Test
    void backup_writesLocalAndSharedAndSkipsBackupFolder(@TempDir Path temp) throws Exception {
        Path konan = temp.resolve("konan");
        Path kokubu = temp.resolve("kokubu");
        Files.createDirectories(konan.resolve("remote_log").resolve("古家"));
        Files.writeString(
                konan.resolve("remote_log").resolve("古家").resolve("note.txt"),
                "k",
                StandardCharsets.UTF_8);
        Files.createDirectories(konan.resolve("バックアップ").resolve("old"));
        Files.writeString(
                konan.resolve("バックアップ").resolve("old").resolve("skip.txt"),
                "no",
                StandardCharsets.UTF_8);
        Files.writeString(konan.resolve("~$book.xlsx"), "lock", StandardCharsets.UTF_8);
        Files.createDirectories(kokubu);
        Files.writeString(kokubu.resolve("b.txt"), "kb", StandardCharsets.UTF_8);

        Path localRoot = temp.resolve("local-backups");
        Map<FactorySite, Path> sources = new EnumMap<>(FactorySite.class);
        sources.put(FactorySite.KONAN, konan);
        sources.put(FactorySite.KOKUBU, kokubu);

        FactoryShareBackupStore.Result result =
                FactoryShareBackupStore.backup(
                        sources,
                        localRoot,
                        site ->
                                (site == FactorySite.KONAN ? konan : kokubu)
                                        .resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME),
                        LocalDateTime.of(2026, 9, 18, 7, 52, 0));

        assertEquals("20260918-075200", result.generationId());
        Path localKonan =
                localRoot.resolve("20260918-075200").resolve(FactorySite.KONAN.name());
        Path sharedKonan =
                konan.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME)
                        .resolve("20260918-075200")
                        .resolve(FactorySite.KONAN.name());
        assertTrue(Files.isRegularFile(localKonan.resolve("remote_log").resolve("古家").resolve("note.txt")));
        assertTrue(Files.isRegularFile(sharedKonan.resolve("remote_log").resolve("古家").resolve("note.txt")));
        assertFalse(Files.exists(localKonan.resolve("バックアップ")));
        assertFalse(Files.exists(localKonan.resolve("~$book.xlsx")));
        assertTrue(
                Files.isRegularFile(
                        localRoot.resolve("20260918-075200").resolve(FactorySite.KOKUBU.name()).resolve("b.txt")));
        assertTrue(
                Files.isRegularFile(
                        kokubu.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME)
                                .resolve("20260918-075200")
                                .resolve(FactorySite.KOKUBU.name())
                                .resolve("b.txt")));
        assertTrue(result.ok());
        assertNull(result.sites().get(0).localError());
        assertEquals("", result.failureSummaryJa());
    }

    @Test
    void backup_missingSource_doesNotPruneOlderGenerations(@TempDir Path temp) throws Exception {
        Path localRoot = temp.resolve("local-backups");
        for (String name :
                List.of(
                        "20260910-010000",
                        "20260911-010000",
                        "20260912-010000",
                        "20260913-010000",
                        "20260914-010000")) {
            Path gen = localRoot.resolve(name);
            Files.createDirectories(gen);
            Files.writeString(gen.resolve("keep.txt"), "k", StandardCharsets.UTF_8);
        }
        Path missing = temp.resolve("missing-konan");
        Path missingKokubu = temp.resolve("missing-kokubu");
        Map<FactorySite, Path> sources = new EnumMap<>(FactorySite.class);
        sources.put(FactorySite.KONAN, missing);
        sources.put(FactorySite.KOKUBU, missingKokubu);

        FactoryShareBackupStore.Result result =
                FactoryShareBackupStore.backup(
                        sources,
                        localRoot,
                        site -> temp.resolve("shared-" + site.name()),
                        LocalDateTime.of(2026, 9, 18, 7, 52, 0));

        assertFalse(result.ok());
        assertTrue(result.failureSummaryJa().contains("湖南"));
        assertTrue(result.failureSummaryJa().contains("国分"));
        assertTrue(Files.isDirectory(localRoot.resolve("20260910-010000")));
        assertFalse(Files.exists(localRoot.resolve("20260918-075200")));
    }

    @Test
    void backup_sameTimestamp_usesNextFreeGeneration(@TempDir Path temp) throws Exception {
        Path konan = temp.resolve("konan");
        Path kokubu = temp.resolve("kokubu");
        Files.createDirectories(konan);
        Files.writeString(konan.resolve("a.txt"), "1", StandardCharsets.UTF_8);
        Files.createDirectories(kokubu);
        Files.writeString(kokubu.resolve("b.txt"), "2", StandardCharsets.UTF_8);
        Path localRoot = temp.resolve("local-backups");
        Map<FactorySite, Path> sources = new EnumMap<>(FactorySite.class);
        sources.put(FactorySite.KONAN, konan);
        sources.put(FactorySite.KOKUBU, kokubu);
        LocalDateTime when = LocalDateTime.of(2026, 9, 18, 7, 52, 0);
        FactoryShareBackupStore.Result first =
                FactoryShareBackupStore.backup(
                        sources,
                        localRoot,
                        site ->
                                (site == FactorySite.KONAN ? konan : kokubu)
                                        .resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME),
                        when);
        Files.writeString(konan.resolve("a.txt"), "1b", StandardCharsets.UTF_8);
        FactoryShareBackupStore.Result second =
                FactoryShareBackupStore.backup(
                        sources,
                        localRoot,
                        site ->
                                (site == FactorySite.KONAN ? konan : kokubu)
                                        .resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME),
                        when);
        assertEquals("20260918-075200", first.generationId());
        assertEquals("20260918-075201", second.generationId());
        assertTrue(
                Files.isRegularFile(
                        localRoot.resolve("20260918-075200").resolve("KONAN").resolve("a.txt")));
        assertEquals(
                "1b",
                Files.readString(
                        localRoot.resolve("20260918-075201").resolve("KONAN").resolve("a.txt"),
                        StandardCharsets.UTF_8));
    }

    @Test
    void copyTree_destInsideSource_doesNotNestBackupFolder(@TempDir Path temp) throws Exception {
        Path src = temp.resolve("share");
        Files.createDirectories(src.resolve("data"));
        Files.writeString(src.resolve("data").resolve("a.txt"), "a", StandardCharsets.UTF_8);
        Path dest =
                src.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME)
                        .resolve("20260918-075200")
                        .resolve("KONAN");
        int copied = FactoryShareBackupStore.copyTree(src, dest, true);
        assertTrue(copied >= 1);
        assertTrue(Files.isRegularFile(dest.resolve("data").resolve("a.txt")));
        assertFalse(Files.exists(dest.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME)));
        assertFalse(
                Files.exists(
                        dest.resolve(AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME)
                                .resolve("20260918-075200")));
    }

    @Test
    void prune_keepsNewestGenerations(@TempDir Path temp) throws Exception {
        Path root = temp.resolve("backups");
        for (String name :
                List.of(
                        "20260910-010000",
                        "20260911-010000",
                        "20260912-010000",
                        "20260913-010000",
                        "20260914-010000",
                        "20260915-010000")) {
            Path gen = root.resolve(name);
            Files.createDirectories(gen);
            Files.writeString(gen.resolve("x.txt"), "x", StandardCharsets.UTF_8);
        }
        List<Path> removed =
                FactoryShareBackupStore.pruneOldGenerations(
                        root, FactoryShareBackupStore.MAX_GENERATIONS);
        assertEquals(1, removed.size());
        assertFalse(Files.exists(root.resolve("20260910-010000")));
        assertTrue(Files.isDirectory(root.resolve("20260915-010000")));
        assertTrue(Files.isDirectory(root.resolve("20260911-010000")));
    }
}
