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
        assertEquals(1, loaded.size());
        assertEquals(original.getFirst(), loaded.getFirst());
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
        assertEquals(3, loaded.size());
        assertEquals("アラジン RPA", loaded.get(2).name());
        assertTrue(Files.readString(json, StandardCharsets.UTF_8).contains("アラジン RPA"));
    }

    @Test
    void ensureCount_fillsMissingNumbers() {
        List<RdpLaunchProfile> ensured =
                RdpLaunchProfileCatalog.ensureCount(
                        List.of(new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, null)), 3);
        assertEquals(3, ensured.size());
        assertEquals(1, ensured.get(0).number());
        assertEquals("B", ensured.get(1).name());
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
        assertEquals(2, loaded.size());
        assertFalse(loaded.get(0).isDeleted());
        assertTrue(loaded.get(1).isDeleted());
        assertTrue(Files.readString(json, StandardCharsets.UTF_8).contains("\"deleted\" : true"));
    }

    @Test
    void activeAndDeletedProfiles_filtersCorrectly() {
        List<RdpLaunchProfile> all =
                List.of(
                        new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null),
                        new RdpLaunchProfile(2, "B", "", "", null, null, null, null, null, null, true),
                        new RdpLaunchProfile(3, "C", "", "", null, null, null, null, null, null, null));
        assertEquals(2, RdpLaunchProfileCatalog.countActive(all));
        assertEquals(2, RdpLaunchProfileCatalog.activeProfiles(all).size());
        assertEquals(1, RdpLaunchProfileCatalog.deletedProfiles(all).size());
        assertEquals(2, RdpLaunchProfileCatalog.deletedProfiles(all).getFirst().number());
    }

    @Test
    void canSoftDelete_requiresMoreThanOneActive() {
        List<RdpLaunchProfile> oneActive =
                List.of(new RdpLaunchProfile(1, "A", "", "", null, null, null, null, null, null, null));
        assertFalse(RdpLaunchProfileCatalog.canSoftDelete(oneActive));
        List<RdpLaunchProfile> twoActive =
                List.of(
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
