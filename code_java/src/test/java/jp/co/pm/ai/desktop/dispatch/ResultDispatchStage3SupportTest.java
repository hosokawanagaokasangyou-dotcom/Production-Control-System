package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class ResultDispatchStage3SupportTest {

    @Test
    void applyStage3DisplayQuantities_copiesActualToPlanColumn() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            row.put(c, "");
        }
        row.put(ResultDispatchSchema.COL_DISPATCH_QTY, "4400");
        row.put(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL, "3000");
        rows.add(row);

        ResultDispatchStage3Support.applyStage3DisplayQuantities(cols, rows);

        assertEquals("3000", rows.get(0).get(ResultDispatchSchema.COL_DISPATCH_QTY));
    }

    @Test
    void removeRedundantActualColumn_dropsColumnFromMaps() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            row.put(c, "1");
        }
        rows.add(row);

        ResultDispatchStage3Support.removeRedundantActualColumnFromMaps(cols, rows);

        assertFalse(cols.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL));
        assertFalse(rows.get(0).containsKey(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL));
    }

    @Test
    void hasStage3ActualColumn_falseWhenAbsent() {
        assertFalse(
                ResultDispatchStage3Support.hasStage3ActualColumn(
                        ResultDispatchSchema.canonicalColumnOrder()));
    }

    @Test
    void hasStage3ActualColumn_trueWhenPresent() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        cols.add(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
        assertTrue(ResultDispatchStage3Support.hasStage3ActualColumn(cols));
    }
}
