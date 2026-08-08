package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

class FactorySiteWorkspaceMigratorTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    private String priorHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        FactorySiteWorkspaceStore.resetForTests();
        GlobalInitSettingTarget.save(FactorySite.KONAN);
    }

    @AfterEach
    void tearDown() {
        FactorySiteWorkspaceStore.resetForTests();
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void migrateIfNeeded_runsOnceAndSeedsWorkspace(@TempDir Path tmp) throws Exception {
        ObjectNode sessionJson = JSON.createObjectNode();
        sessionJson.put("planInputPath", "seed-plan.xlsx");
        DesktopSessionState session = DesktopSessionStateStore.parseSessionFragment(sessionJson);
        List<UiEnvRowSnapshot> rows =
                List.of(
                        new UiEnvRowSnapshot(
                                AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, "\\\\seed\\original", ""));

        FactorySiteWorkspaceMigrator.migrateIfNeeded(
                "砂田", FactorySite.KONAN, rows, session, Map.of());

        Path marker = AppPaths.operatorLocalMigrationMarkerPath("砂田");
        assertTrue(Files.isRegularFile(marker));
        Optional<FactorySiteWorkspaceSnapshot> loaded =
                FactorySiteWorkspaceStore.load("砂田", FactorySite.KONAN);
        assertTrue(loaded.isPresent());
        assertEquals("seed-plan.xlsx", loaded.get().sessionFragment().planInputPath());
        assertEquals("\\\\seed\\original", loaded.get().uiEnvRows().get(0).value());

        FactorySiteWorkspaceStore.save(
                "砂田",
                FactorySite.KONAN,
                new FactorySiteWorkspaceSnapshot(
                        List.of(
                                new UiEnvRowSnapshot(
                                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                                        "\\\\updated\\original",
                                        "")),
                        session));

        FactorySiteWorkspaceMigrator.migrateIfNeeded(
                "砂田", FactorySite.KONAN, rows, session, Map.of());

        assertEquals(
                "\\\\updated\\original",
                FactorySiteWorkspaceStore.load("砂田", FactorySite.KONAN)
                        .orElseThrow()
                        .uiEnvRows()
                        .get(0)
                        .value());
    }

    @Test
    void migrateIfNeeded_doesNotSeedWorkspaceWhenUiEnvRowsEmpty() throws Exception {
        ObjectNode sessionJson = JSON.createObjectNode();
        sessionJson.put("planInputPath", "seed-plan.xlsx");
        DesktopSessionState session = DesktopSessionStateStore.parseSessionFragment(sessionJson);

        FactorySiteWorkspaceMigrator.migrateIfNeeded(
                "砂田", FactorySite.KONAN, List.of(), session, Map.of());

        Path marker = AppPaths.operatorLocalMigrationMarkerPath("砂田");
        assertTrue(Files.isRegularFile(marker));
        assertTrue(FactorySiteWorkspaceStore.load("砂田", FactorySite.KONAN).isEmpty());
    }
}
