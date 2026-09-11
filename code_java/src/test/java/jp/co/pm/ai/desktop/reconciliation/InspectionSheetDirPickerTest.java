package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.FactorySite;

class InspectionSheetDirPickerTest {

    @Test
    void matchesKensa_trimLowercaseOnly() {
        assertTrue(InspectionSheetDirPicker.matchesKensa("kensa"));
        assertTrue(InspectionSheetDirPicker.matchesKensa(" kensa "));
        assertFalse(InspectionSheetDirPicker.matchesKensa("KENSA"));
        assertFalse(InspectionSheetDirPicker.matchesKensa(""));
        assertFalse(InspectionSheetDirPicker.matchesKensa(null));
        assertFalse(InspectionSheetDirPicker.matchesKensa("kens"));
    }

    @Test
    void initialDirectoryCandidates_factoryLeafFirst() {
        var konan = InspectionSheetDirPicker.initialDirectoryCandidates(FactorySite.KONAN);
        assertTrue(konan.get(0).endsWith(Path.of("後加工検査表", "湖南工場")));
        var kokubu = InspectionSheetDirPicker.initialDirectoryCandidates(FactorySite.KOKUBU);
        assertTrue(kokubu.get(0).endsWith(Path.of("後加工検査表", "国分工場")));
    }
}
