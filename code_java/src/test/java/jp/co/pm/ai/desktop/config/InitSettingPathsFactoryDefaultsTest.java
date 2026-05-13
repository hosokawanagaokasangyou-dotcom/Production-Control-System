package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class InitSettingPathsFactoryDefaultsTest {

    @Test
    void sessionAndTableDefaultsFileNames_useLowercaseEnumName() {
        assertEquals("session_defaults_konan.json", InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KONAN));
        assertEquals(
                "session_defaults_kokubu.json", InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KOKUBU));
        assertEquals(
                "table_column_defaults_konan.json",
                InitSettingPaths.tableColumnDefaultsFileForFactory(FactorySite.KONAN));
        assertEquals(
                "table_column_defaults_kokubu.json",
                InitSettingPaths.tableColumnDefaultsFileForFactory(FactorySite.KOKUBU));
    }
}
