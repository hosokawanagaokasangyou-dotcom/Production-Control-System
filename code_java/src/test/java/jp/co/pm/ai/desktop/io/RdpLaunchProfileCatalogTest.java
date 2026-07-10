package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RdpLaunchProfileCatalogTest {

    @Test
    void loadBundledDefaults_hasExampleProfiles() {
        List<RdpLaunchProfile> profiles = RdpLaunchProfileCatalog.loadBundledDefaults();
        assertFalse(profiles.isEmpty());
        assertTrue(
                profiles.stream()
                        .anyMatch(
                                p ->
                                        p.number() == RdpRemoteLauncherIni.SLOT_SIGN_OUT
                                                && p.isSignOutOnlyProfile()));
        assertTrue(
                profiles.stream()
                        .anyMatch(p -> p.number() == 2 && p.name().contains("工程マスタ")));
    }

    @Test
    void saveAndLoad_roundTrip(@TempDir Path tmp) throws Exception {
        Path json = tmp.resolve("RDP起動プロファイル.json");
        List<RdpLaunchProfile> original =
                List.of(
                        new RdpLaunchProfile(
                                1,
                                "テスト",
                                "説明文",
                                "マスタ更新",
                                true,
                                null,
                                false,
                                1280,
                                800,
                                true,
                                null));
        RdpLaunchProfileCatalog.save(json, original);
        List<RdpLaunchProfile> loaded = RdpLaunchProfileCatalog.load(json);
        assertEquals(2, loaded.size());
        RdpLaunchProfile loadedProfile1 =
                loaded.stream().filter(p -> p.number() == 1).findFirst().orElseThrow();
        assertEquals(original.getFirst(), loadedProfile1);
        assertTrue(Files.readString(json, StandardCharsets.UTF_8).contains("テスト"));
    }

    @Test
    void save_preservesHighNumberedProfileMetadata(@TempDir Path tmp) throws Exception {
        Path json = tmp.resolve("RDP起動プロファイル.json");
        List<RdpLaunchProfile> original =
                List.of(
                        RdpLaunchProfile.empty(1),
                        RdpLaunchProfile.empty(2),
                        new RdpLaunchProfile(
                                3, "アラジン RPA", "説明", "マスタ更新", null, null, null, null, null, null, null));
        RdpLaunchProfileCatalog.save(json, original);
        List<RdpLaunchProfile> loaded = RdpLaunchProfileCatalog.load(json);
        assertEquals(4, loaded.size());
        assertEquals(
                "アラジン RPA",
                loaded.stream().filter(p -> p.number() == 3).findFirst().orElseThrow().name());
        assertTrue(Files.readString(json, StandardCharsets.UTF_8).contains("アラジン RPA"));
    }

    @Test
    void ensureCount_fillsMissingNumbers() {
        List<RdpLaunchProfile> ensured =
                RdpLaunchProfileCatalog.ensureCount(
                        List.of(new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, null)), 3);
        assertEquals(4, ensured.size());
        assertEquals(99, ensured.get(0).number());
        assertEquals(1, ensured.get(1).number());
        assertEquals(2, ensured.get(2).number());
        assertEquals("B", ensured.get(2).name());
        assertEquals(3, ensured.get(3).number());
    }

    @Test
    void displayLabel_includesNumber() {
        RdpLaunchProfile profile =
                new RdpLaunchProfile(2, "アラジン 工程マスタ取得", "", "", null, null, null, null, null, null, null);
        assertEquals("2: アラジン 工程マスタ取得", profile.displayLabel());
    }

    @Test
    void saveAndLoad_preservesDeletedFlag(@TempDir Path tmp) throws Exception {
        Path json = tmp.resolve("RDP起動プロファイル.json");
        List<RdpLaunchProfile> original =
                List.of(
                        new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null),
                        new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, true));
        RdpLaunchProfileCatalog.save(json, original);
        List<RdpLaunchProfile> loaded = RdpLaunchProfileCatalog.load(json);
        assertEquals(3, loaded.size());
        RdpLaunchProfile loaded1 =
                loaded.stream().filter(p -> p.number() == 1).findFirst().orElseThrow();
        RdpLaunchProfile loaded2 =
                loaded.stream().filter(p -> p.number() == 2).findFirst().orElseThrow();
        assertFalse(loaded1.isDeleted());
        assertTrue(loaded2.isDeleted());
        assertTrue(Files.readString(json, StandardCharsets.UTF_8).contains("\"deleted\" : true"));
    }

    @Test
    void activeAndDeletedProfiles_filtersCorrectly() {
        List<RdpLaunchProfile> all =
                List.of(
                        new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null),
                        new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, true),
                        new RdpLaunchProfile(3, "C", "", "", null, null, null, null, null, null, null));
        assertEquals(3, RdpLaunchProfileCatalog.countActive(all));
        assertEquals(3, RdpLaunchProfileCatalog.activeProfiles(all).size());
        assertEquals(1, RdpLaunchProfileCatalog.deletedProfiles(all).size());
        assertEquals(2, RdpLaunchProfileCatalog.deletedProfiles(all).getFirst().number());
    }

    @Test
    void canSoftDelete_requiresMoreThanOneActive() {
        List<RdpLaunchProfile> oneActive =
                List.of(
                        RdpLaunchProfile.signOutOnlyDefault(),
                        new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null));
        assertFalse(RdpLaunchProfileCatalog.canSoftDelete(oneActive));
        List<RdpLaunchProfile> twoActive =
                List.of(
                        RdpLaunchProfile.signOutOnlyDefault(),
                        new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null),
                        new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, null));
        assertTrue(RdpLaunchProfileCatalog.canSoftDelete(twoActive));
    }

    @Test
    void withDeleted_roundTrip() {
        RdpLaunchProfile original =
                new RdpLaunchProfile(1, "A", "desc", "cat", null, null, false, 1280, 800, false, null);
        RdpLaunchProfile deleted = original.withDeleted(true);
        assertTrue(deleted.isDeleted());
        assertEquals("1: A（削除済）", deleted.displayLabel());
        RdpLaunchProfile restored = deleted.withDeleted(false);
        assertFalse(restored.isDeleted());
        assertEquals(original, restored);
    }
}
