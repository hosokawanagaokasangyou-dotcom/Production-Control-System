package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;

/** 検査表索引から依頼NOで照合する。 */
public final class InspectionSheetLookup {

    private InspectionSheetLookup() {}

    public static List<InspectionSheetIndexStore.Row> find(
            List<InspectionSheetIndexStore.Row> rows, String iraiNo) {
        if (rows == null || iraiNo == null || InspectionSheetIraiNo.normalize(iraiNo).isEmpty()) {
            return List.of();
        }
        List<InspectionSheetIndexStore.Row> hits = new ArrayList<>();
        for (InspectionSheetIndexStore.Row row : rows) {
            if (row != null && InspectionSheetIraiNo.matches(iraiNo, row.iraiNo())) {
                hits.add(row);
            }
        }
        return List.copyOf(hits);
    }
}
