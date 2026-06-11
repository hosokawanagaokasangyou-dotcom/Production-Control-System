package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class UncPathSegmentRepairTest {

    @Test
    void repair_fixesCollapsedKonan002Segment() {
        String wrong = "\\\\192.168.0.101\\share\\002 加工G\\scenarios\\a.ardrpa";
        String correct = "\\\\192.168.0.101\\share\\002  加工G\\scenarios\\a.ardrpa";
        assertEquals(correct, UncPathSegmentRepair.repair(wrong));
    }

    @Test
    void repair_preservesCorrectKonan002Segment() {
        String path = "\\\\192.168.0.101\\share\\002  加工G\\app.exe";
        assertEquals(path, UncPathSegmentRepair.repair(path));
    }
}
