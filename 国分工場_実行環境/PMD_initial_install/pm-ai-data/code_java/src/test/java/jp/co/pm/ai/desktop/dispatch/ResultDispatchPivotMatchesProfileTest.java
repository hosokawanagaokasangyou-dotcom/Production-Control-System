package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class ResultDispatchPivotMatchesProfileTest {

    @Test
    void matchesTaskProfile_includesTrialOrderInComparison() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        Map<String, String> profile = new LinkedHashMap<>();
        for (String c : cols) {
            if (!ResultDispatchSchema.COL_DISPATCH_DATE.equals(c)
                    && !ResultDispatchSchema.COL_DISPATCH_QTY.equals(c)) {
                profile.put(c, "");
            }
        }
        profile.put(ResultDispatchSchema.COL_PROCESS, "P");
        profile.put(ResultDispatchSchema.COL_MACHINE, "M");
        profile.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "1");

        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            row.put(c, "");
        }
        row.put(ResultDispatchSchema.COL_PROCESS, "P");
        row.put(ResultDispatchSchema.COL_MACHINE, "M");
        row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "1");
        row.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-10");
        row.put(ResultDispatchSchema.COL_DISPATCH_QTY, "10");

        assertTrue(ResultDispatchPivot.matchesTaskProfile(cols, profile, row));

        row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "2");
        assertFalse(ResultDispatchPivot.matchesTaskProfile(cols, profile, row));
    }
}
