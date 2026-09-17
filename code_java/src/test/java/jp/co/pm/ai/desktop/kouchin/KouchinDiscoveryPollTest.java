package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class KouchinDiscoveryPollTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("検証タブ表示中は3秒おきに検出する")
    void pollIntervalIsThreeSeconds() {
        assertEquals(3, KouchinVerifyTabController.DISCOVERY_POLL_SECONDS);
    }

    @Test
    @DisplayName("ポーリングはタブ表示中かつ検証未実行のときだけ")
    void pollOnlyWhileTabSelectedAndIdle() {
        assertTrue(KouchinVerifyTabController.shouldPollDiscovery(true, false));
        assertFalse(KouchinVerifyTabController.shouldPollDiscovery(false, false));
        assertFalse(KouchinVerifyTabController.shouldPollDiscovery(true, true));
        assertFalse(KouchinVerifyTabController.shouldPollDiscovery(false, true));
    }

    @Test
    @DisplayName("サイレント再検出では検出中ステータスを出さない")
    void silentPollDoesNotShowDetectingStatus() {
        assertTrue(KouchinVerifyTabController.shouldShowDetectingStatus(false, false));
        assertFalse(KouchinVerifyTabController.shouldShowDetectingStatus(true, false));
        assertFalse(KouchinVerifyTabController.shouldShowDetectingStatus(false, true));
        assertFalse(KouchinVerifyTabController.shouldShowDetectingStatus(true, true));
    }

    @Test
    @DisplayName("検出中のスキャンが終わるまで次のポーリングは開始しない")
    void skipOverlappingDiscoveryScan() {
        assertTrue(KouchinVerifyTabController.shouldStartDiscoveryScan(false));
        assertFalse(KouchinVerifyTabController.shouldStartDiscoveryScan(true));
    }

    @Test
    @DisplayName("ファイル更新日時が変わると検出表スナップショットも変わる")
    void snapshotChangesWhenFileModified() throws Exception {
        Path file = tmp.resolve("RVSHEET202608.csv");
        Files.writeString(file, "a");
        Files.setLastModifiedTime(file, FileTime.from(Instant.parse("2026-08-15T01:23:45Z")));
        jp.co.pm.ai.kouchin.verify.KouchinDiscovery.Row row = new jp.co.pm.ai.kouchin.verify.KouchinDiscovery.Row(
                "①東レCSV", file.getFileName().toString(), file.toString(), "2026年8月度", false, "");
        var before = List.of(KouchinVerifyTabController.DiscoveryLine.of("国分/湖南共通", row));
        Files.setLastModifiedTime(file, FileTime.from(Instant.parse("2026-09-17T01:00:00Z")));
        var after = List.of(KouchinVerifyTabController.DiscoveryLine.of("国分/湖南共通", row));
        assertFalse(KouchinVerifyTabController.sameDiscoverySnapshot(before, after));
        assertTrue(KouchinVerifyTabController.sameDiscoverySnapshot(before, before));
        assertEquals("2026/08/15 10:23:45", before.get(0).getModifiedAt());
        assertEquals("2026/09/17 10:00:00", after.get(0).getModifiedAt());
    }

    @Test
    @DisplayName("選択時に3秒ポーリングを開始し、非選択で止める")
    void sourceStartsPollOnSelectAndStopsOnDeselect() throws Exception {
        String src = Files.readString(Path.of(
                "src/main/java/jp/co/pm/ai/desktop/kouchin/KouchinVerifyTabController.java"));
        assertTrue(src.contains("startDiscoveryPoll()"), src);
        assertTrue(src.contains("stopDiscoveryPoll()"), src);
        assertTrue(src.contains("onDiscoveryPollTick()"), src);
        assertTrue(src.contains("Duration.seconds(DISCOVERY_POLL_SECONDS)"), src);
        assertTrue(src.contains("FileDiscovery.invalidateListingCache()"), src);
        int selected = src.indexOf("public void onMainShellTabSelected()");
        int deselected = src.indexOf("public void onMainShellTabDeselected()");
        assertTrue(selected >= 0 && deselected > selected, "選択/非選択メソッドが無い");
        String selectBody = src.substring(selected, deselected);
        String deselectBody = src.substring(deselected, src.indexOf("public void reloadDiscovery()"));
        assertTrue(selectBody.contains("startDiscoveryPoll()"), selectBody);
        assertTrue(deselectBody.contains("stopDiscoveryPoll()"), deselectBody);
    }
}
