package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

class FactorySiteWorkspaceStoreTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    private String priorHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        AppPaths.setDesktopAppHomeDirName(tmp.toString());
        FactorySiteWorkspaceStore.resetForTests();
        GlobalInitSettingTarget.save(FactorySite.KONAN);
    }

    @AfterEach
    void tearDown() {
        FactorySiteWorkspaceStore.resetForTests();
        AppPaths.setDesktopAppHomeDirName(priorHome);
    }

    @Test
    void saveAndLoad_roundTripsUiEnvRowsAndSessionFragment() throws Exception {
        List<UiEnvRowSnapshot> rows =
                List.of(
                        new UiEnvRowSnapshot(
                                AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                                "\\\\konan\\box\\original",
                                ""));
        ObjectNode sessionJson = JSON.createObjectNode();
        sessionJson.put("planInputPath", "C:\\plan\\konan.xlsx");
        DesktopSessionState session = DesktopSessionStateStore.parseSessionFragment(sessionJson);

        FactorySiteWorkspaceStore.save(
                "砂田", FactorySite.KONAN, new FactorySiteWorkspaceSnapshot(rows, session));

        Optional<FactorySiteWorkspaceSnapshot> loaded =
                FactorySiteWorkspaceStore.load("砂田", FactorySite.KONAN);
        assertTrue(loaded.isPresent());
        assertEquals(
                "\\\\konan\\box\\original",
                loaded.get().uiEnvRows().get(0).value());
        assertEquals("C:\\plan\\konan.xlsx", loaded.get().sessionFragment().planInputPath());
    }

    @Test
    void konanAndKokubuSnapshotsDoNotMix() throws Exception {
        ObjectNode konanSession = JSON.createObjectNode();
        konanSession.put("planInputPath", "konan-plan");
        ObjectNode kokubuSession = JSON.createObjectNode();
        kokubuSession.put("planInputPath", "kokubu-plan");

        FactorySiteWorkspaceStore.save(
                "砂田",
                FactorySite.KONAN,
                new FactorySiteWorkspaceSnapshot(
                        List.of(
                                new UiEnvRowSnapshot(
                                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                                        "\\\\konan\\original",
                                        "")),
                        DesktopSessionStateStore.parseSessionFragment(konanSession)));
        FactorySiteWorkspaceStore.save(
                "砂田",
                FactorySite.KOKUBU,
                new FactorySiteWorkspaceSnapshot(
                        List.of(
                                new UiEnvRowSnapshot(
                                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                                        "\\\\kokubu\\original",
                                        "")),
                        DesktopSessionStateStore.parseSessionFragment(kokubuSession)));

        FactorySiteWorkspaceSnapshot konan =
                FactorySiteWorkspaceStore.load("砂田", FactorySite.KONAN).orElseThrow();
        FactorySiteWorkspaceSnapshot kokubu =
                FactorySiteWorkspaceStore.load("砂田", FactorySite.KOKUBU).orElseThrow();

        assertEquals("\\\\konan\\original", konan.uiEnvRows().get(0).value());
        assertEquals("\\\\kokubu\\original", kokubu.uiEnvRows().get(0).value());
        assertNotEquals(konan.sessionFragment().planInputPath(), kokubu.sessionFragment().planInputPath());
    }

    @Test
    void saveLastFactorySite_persistsToDisk() throws Exception {
        FactorySiteWorkspaceStore.saveLastFactorySite("砂田", FactorySite.KOKUBU);
        assertEquals(
                FactorySite.KOKUBU,
                FactorySiteWorkspaceStore.loadLastFactorySite("砂田").orElseThrow());
        Path path = AppPaths.operatorLastFactorySitePath("砂田");
        assertTrue(Files.isRegularFile(path));
        assertEquals("KOKUBU", Files.readString(path).trim());
    }

    @Test
    void extractAndMergeFactoryScopedFields_preservesGlobalWindowTheme() throws Exception {
        ObjectNode globalJson = JSON.createObjectNode();
        globalJson.put("planInputPath", "plan-konan");
        globalJson.put("windowWidth", 1200);
        globalJson.put("uiTheme", "dark");
        DesktopSessionState global = DesktopSessionStateStore.parseSessionFragment(globalJson);

        ObjectNode factoryJson = JSON.createObjectNode();
        factoryJson.put("planInputPath", "plan-kokubu");
        DesktopSessionState factoryOnly = DesktopSessionStateStore.parseSessionFragment(factoryJson);

        DesktopSessionState merged = global.mergeFactoryScopedFrom(factoryOnly);
        assertEquals("plan-kokubu", merged.planInputPath());
        assertEquals(1200d, merged.windowWidth());
        assertEquals("dark", merged.uiTheme());
    }
}
