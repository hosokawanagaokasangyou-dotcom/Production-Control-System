package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppPathsMasterDispatchSheetsTest {

    @Test
    void defaultSharedPath_isCurrentFactoryDataDirOnly() {
        Path konan = AppPaths.masterDispatchSheetsDefaultSharedPath(FactorySite.KONAN);
        Path kokubu = AppPaths.masterDispatchSheetsDefaultSharedPath(FactorySite.KOKUBU);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR, AppPaths.MASTER_DISPATCH_SHEETS_JSON_FILENAME),
                konan);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR, AppPaths.MASTER_DISPATCH_SHEETS_JSON_FILENAME),
                kokubu);
        assertFalse(konan.equals(kokubu));
    }

    @Test
    void preferredPath_usesExplicitEnvWithoutTouchingOtherFactory() {
        Path custom = Path.of("C:\\tmp\\master-dispatch-sheets.json");
        Path got =
                AppPaths.masterDispatchSheetsPreferredPath(
                        Map.of(
                                AppPaths.KEY_PM_AI_FACTORY_SITE,
                                "KONAN",
                                AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON,
                                custom.toString()));
        assertEquals(custom.toAbsolutePath().normalize(), got);
    }

    @Test
    void preferredPath_withoutEnv_isCurrentFactorySharedDefault() {
        Path got =
                AppPaths.masterDispatchSheetsPreferredPath(
                        Map.of(AppPaths.KEY_PM_AI_FACTORY_SITE, "KOKUBU"));
        assertEquals(AppPaths.masterDispatchSheetsDefaultSharedPath(FactorySite.KOKUBU), got);
    }

    @Test
    void jsonPath_fallsBackToLocalWhenPreferredParentMissing(@TempDir Path tmp) {
        Path missingParent = tmp.resolve("no-such-dir").resolve("master-dispatch-sheets.json");
        Path got =
                AppPaths.masterDispatchSheetsJsonPath(
                        Map.of(
                                AppPaths.KEY_PM_AI_FACTORY_SITE,
                                "KONAN",
                                AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON,
                                missingParent.toString()));
        assertEquals(AppPaths.masterDispatchSheetsLocalFallbackPath(FactorySite.KONAN), got);
        assertTrue(got.startsWith(AppPaths.resolveDesktopAppHomeDir()));
    }

    @Test
    void jsonPath_usesPreferredWhenParentWritable(@TempDir Path tmp) throws Exception {
        Path json = tmp.resolve("master-dispatch-sheets.json");
        Files.createDirectories(json.getParent());
        Path got =
                AppPaths.masterDispatchSheetsJsonPath(
                        Map.of(
                                AppPaths.KEY_PM_AI_FACTORY_SITE,
                                "KONAN",
                                AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON,
                                json.toString()));
        assertEquals(json.toAbsolutePath().normalize(), got);
    }

    @Test
    void sourceWorkbookPath_isFactoryDefaultNotMasterEnvOverride() {
        Path got =
                AppPaths.masterDispatchSheetsSourceWorkbookPath(
                        Map.of(
                                AppPaths.KEY_PM_AI_FACTORY_SITE,
                                "KONAN",
                                AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                                "C:\\local\\master.xlsm"));
        assertEquals(Path.of(AppPaths.DEFAULT_PM_AI_MASTER_WORKBOOK_KONAN), got);
    }

    @Test
    void overlayFactorySite_setsCurrentFactorySharedJsonPathOnly() {
        java.util.LinkedHashMap<String, String> map = new java.util.LinkedHashMap<>();
        map.put(AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON, "old");
        AppPaths.overlayFactorySiteMasterDispatchSheetsPath(map, FactorySite.KOKUBU);
        assertEquals(
                AppPaths.masterDispatchSheetsDefaultSharedPath(FactorySite.KOKUBU).toString(),
                map.get(AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON));
    }

    @Test
    void jsonFilePicker_includesMasterDispatchSheetsKey() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON));
        assertTrue(AppPaths.isJsonFilePathEnvKey(AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_MASTER_DISPATCH_SHEETS_JSON));
    }
}
