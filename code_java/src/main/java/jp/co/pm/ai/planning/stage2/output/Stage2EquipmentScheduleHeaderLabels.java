package jp.co.pm.ai.planning.stage2.output;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Python {@code planning_core._core._equipment_schedule_header_labels} と同趣旨の表示ラベル。
 */
public final class Stage2EquipmentScheduleHeaderLabels {

    private Stage2EquipmentScheduleHeaderLabels() {}

    public static List<String> fromEquipmentCombos(List<String> equipmentList) {
        List<String> raw = new ArrayList<>();
        for (String eq : equipmentList) {
            String s = eq == null ? "" : eq.strip();
            if (s.isEmpty()) {
                raw.add("");
                continue;
            }
            if (s.contains("+")) {
                int p = s.indexOf('+');
                String mpart = s.substring(p + 1).strip();
                raw.add(!mpart.isEmpty() ? mpart : s);
            } else {
                raw.add(s);
            }
        }
        Map<String, Integer> counts = new HashMap<>();
        for (String r : raw) {
            counts.merge(r, 1, Integer::sum);
        }
        List<String> out = new ArrayList<>();
        for (int i = 0; i < equipmentList.size(); i++) {
            String eq = equipmentList.get(i);
            String r = raw.get(i);
            if (counts.getOrDefault(r, 0) > 1) {
                String s = eq == null ? "" : eq.strip();
                if (s.contains("+")) {
                    int p = s.indexOf('+');
                    String proc = s.substring(0, p).strip();
                    out.add(!proc.isEmpty() ? r + "（" + proc + "）" : r);
                } else {
                    out.add(r);
                }
            } else {
                out.add(r);
            }
        }
        return out;
    }
}
