package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class DispatchSnapshotStoreTest {

    @TempDir Path temp;

    @Test
    void publishCopiesSetAndRecordsMissing() throws Exception {
        Path plan = temp.resolve("計画2609230900000001.json");
        Path member = temp.resolve("人員2609230900000001.json");
        Path dispatch = temp.resolve("結果_配台表.json");
        Files.writeString(plan, "{}", StandardCharsets.UTF_8);
        Files.writeString(member, "{}", StandardCharsets.UTF_8);
        Files.writeString(dispatch, "{}", StandardCharsets.UTF_8);
        Path root = temp.resolve("share");

        DispatchSnapshotStore.PublishResult published =
                DispatchSnapshotStore.publish(
                        root,
                        "森岡",
                        "stage2",
                        LocalDateTime.of(2026, 9, 23, 9, 20, 1, 2_000_000),
                        plan,
                        member,
                        dispatch,
                        temp.resolve("shaped_aladdin_plan.json"),
                        temp.resolve("shaped_processing_actuals.json"));

        assertTrue(Files.isRegularFile(published.generationDir().resolve("計画2609230900000001.json")));
        assertTrue(Files.isRegularFile(published.generationDir().resolve("人員2609230900000001.json")));
        assertTrue(Files.isRegularFile(published.generationDir().resolve("結果_配台表.json")));
        assertTrue(published.missing().contains("shaped_aladdin_plan.json"));
        assertTrue(published.missing().contains("shaped_processing_actuals.json"));
        assertTrue(Files.isRegularFile(published.generationDir().resolve("meta.json")));

        List<DispatchSnapshotStore.SnapshotRef> listed = DispatchSnapshotStore.list(root);
        assertEquals(1, listed.size());
        assertEquals("森岡", listed.get(0).operatorDir());
        assertTrue(listed.get(0).missing().contains("shaped_aladdin_plan.json"));
    }

    @Test
    void publishKeepsPlanAndMemberOnTheSameStamp() throws Exception {
        Path planOld = temp.resolve("計画2609230900000001.json");
        Path planNew = temp.resolve("計画2609230900000002.json");
        Path memberOld = temp.resolve("人員2609230900000001.json");
        Path memberNew = temp.resolve("人員2609230900000002.json");
        Files.writeString(planOld, "{}", StandardCharsets.UTF_8);
        Files.writeString(planNew, "{}", StandardCharsets.UTF_8);
        Files.writeString(memberOld, "{}", StandardCharsets.UTF_8);
        Files.writeString(memberNew, "{}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(planNew, java.nio.file.attribute.FileTime.fromMillis(2_000));
        Files.setLastModifiedTime(memberOld, java.nio.file.attribute.FileTime.fromMillis(9_000));

        DispatchSnapshotStore.PublishResult published =
                DispatchSnapshotStore.publish(
                        temp.resolve("share"),
                        "森岡",
                        "stage2",
                        LocalDateTime.of(2026, 9, 23, 9, 30),
                        planNew,
                        memberOld,
                        null,
                        null,
                        null);

        assertTrue(Files.isRegularFile(published.generationDir().resolve("計画2609230900000002.json")));
        assertFalse(Files.exists(published.generationDir().resolve("人員2609230900000001.json")));
        assertTrue(Files.isRegularFile(published.generationDir().resolve("人員2609230900000002.json")));
        assertFalse(published.missing().contains("人員2609230900000002.json"));
    }

    @Test
    void listSkipsUnknownFormatAndCorruptMeta() throws Exception {
        Path root = temp.resolve("share");
        Path future = root.resolve("森岡").resolve("20260923-090000-001_pc_stage2");
        Files.createDirectories(future);
        Files.writeString(
                future.resolve("meta.json"),
                "{\"format_version\":99}\n",
                StandardCharsets.UTF_8);
        Path corrupt = root.resolve("森岡").resolve("20260923-090000-002_pc_stage2");
        Files.createDirectories(corrupt);
        Files.writeString(corrupt.resolve("meta.json"), "{", StandardCharsets.UTF_8);
        Path plain = root.resolve("森岡").resolve("20260923-090000-003_pc_stage2");
        Files.createDirectories(plain);

        List<DispatchSnapshotStore.SnapshotRef> listed = DispatchSnapshotStore.list(root);
        assertEquals(1, listed.size());
        assertEquals("20260923-090000-003_pc_stage2", listed.get(0).generationDir());
    }

    @Test
    void pruneDropsExpiredAndDeleteOwnRefusesOtherOperator() throws Exception {
        Path root = temp.resolve("share");
        Path oldGen = root.resolve("森岡").resolve("20200101-000000-001_pc_stage2");
        Path fresh = root.resolve("森岡").resolve("20260923-090000-001_pc_stage2");
        Path otherOld = root.resolve("細川").resolve("20200101-000000-001_pc_stage2");
        Path otherFresh = root.resolve("細川").resolve("20260923-090000-001_pc_stage2");
        Files.createDirectories(oldGen);
        Files.createDirectories(fresh);
        Files.createDirectories(otherOld);
        Files.createDirectories(otherFresh);
        Files.writeString(oldGen.resolve("meta.json"), "{\"format_version\":1}\n", StandardCharsets.UTF_8);
        Files.writeString(fresh.resolve("meta.json"), "{\"format_version\":1}\n", StandardCharsets.UTF_8);
        Files.writeString(otherOld.resolve("meta.json"), "{\"format_version\":1}\n", StandardCharsets.UTF_8);
        Files.writeString(otherFresh.resolve("meta.json"), "{\"format_version\":1}\n", StandardCharsets.UTF_8);

        Path plan = temp.resolve("計画2609230900000001.json");
        Files.writeString(plan, "{}", StandardCharsets.UTF_8);
        DispatchSnapshotStore.publish(
                root,
                "森岡",
                "stage2",
                LocalDateTime.of(2026, 9, 23, 10, 0),
                plan,
                null,
                null,
                null,
                null);

        assertFalse(Files.exists(oldGen));
        assertFalse(Files.exists(otherOld));
        assertTrue(Files.isDirectory(fresh));
        assertTrue(Files.isDirectory(otherFresh));
        assertTrue(DispatchSnapshotStore.deleteOwn(root, "森岡", "20260923-090000-001_pc_stage2"));
        assertFalse(Files.exists(fresh));
        assertTrue(Files.isDirectory(otherFresh));
        assertFalse(DispatchSnapshotStore.deleteOwn(root, "森岡", "../細川/20260923-090000-001_pc_stage2"));
        assertTrue(Files.isDirectory(otherFresh));
        assertTrue(DispatchSnapshotStore.isExpired("20200101-000000-001_pc_stage2", LocalDate.of(2026, 9, 23), 30));
        assertFalse(DispatchSnapshotStore.isExpired("not-a-generation", LocalDate.of(2026, 9, 23), 30));
    }

    @Test
    void localDisplayPathIsNotTheSnapshotRootAndListenerFailureDoesNotStopOthers() {
        Path root = temp.resolve("snaps");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_DISPATCH_SNAPSHOT_DIR, root.toString());
        DispatchResultSelection selection = new DispatchResultSelection();
        Path local = DispatchResultPaths.dispatchJson(ui, selection);
        assertFalse(local != null && local.startsWith(root));

        assertTrue(
                selection.selectSnapshot(
                        "細川",
                        "20260923-090000-001_pc_stage2",
                        DispatchResultSelection.badgeTextFor(false, "細川", "2026-09-23T09:20"),
                        "共有",
                        false));
        Path shown = DispatchResultPaths.dispatchJson(ui, selection);
        assertEquals(
                root.resolve("細川")
                        .resolve("20260923-090000-001_pc_stage2")
                        .resolve("結果_配台表.json")
                        .normalize(),
                shown);
        assertEquals("他者の結果: 細川 2026-09-23T09:20", selection.badgeText());
        assertEquals(" ［他者の結果］", DispatchResultSelection.tabMark(false, false));
        assertEquals("", DispatchResultSelection.tabMark(true, false));

        AtomicInteger calls = new AtomicInteger();
        selection.addListener(
                () -> {
                    throw new IllegalStateException("tab failed");
                });
        selection.addListener(calls::incrementAndGet);
        assertTrue(selection.selectLocal());
        assertEquals(1, calls.get());
        assertTrue(selection.isLocalLatest());
        DispatchResultPaths.planJson(ui, selection);
    }

    @Test
    void vetoBlocksChange() {
        DispatchResultSelection selection = new DispatchResultSelection();
        selection.setChangeAllowed(() -> false);
        assertFalse(
                selection.selectSnapshot(
                        "細川", "20260923-090000-001_pc_stage2", "他者の結果", "x", false));
        assertTrue(selection.isLocalLatest());
        selection.setChangeAllowed(() -> true);
        assertTrue(
                selection.selectSnapshot(
                        "細川", "20260923-090000-001_pc_stage2", "他者の結果", "x", false));
        assertFalse(selection.isLocalLatest());
    }

    @Test
    void generationDirStaysInsideSnapshotRoot() {
        Path root = temp.resolve("share");
        String gen = "20260923-090000-001_pc_stage2";
        Path ok = DispatchSnapshotStore.generationDir(root, "森岡", gen);
        assertTrue(ok.startsWith(root.toAbsolutePath().normalize()));
        assertNull(DispatchSnapshotStore.generationDir(root, "..", gen));
        assertNull(DispatchSnapshotStore.generationDir(root, "a/../../etc", gen));
        assertNull(DispatchSnapshotStore.generationDir(root, "森岡", "../20260923-090000-001_pc_stage2"));
    }
}
