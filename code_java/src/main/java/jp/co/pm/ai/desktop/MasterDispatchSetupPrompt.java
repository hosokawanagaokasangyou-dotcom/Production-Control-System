package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsDocument;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSeeder;
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.EquipmentRef;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.Evaluation;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

/**
 * 計画タスク上の工程+機械に対する配台マスタ（skills/need/組合せ/speed）完了状況を集計する。
 */
public final class MasterDispatchSetupPrompt {

    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_TASK = "依頼NO";
    private static final String COL_EXCLUDE = "配台不要";

    public record SheetBundle(
            List<List<String>> skills,
            List<List<String>> need,
            List<List<String>> combinations,
            List<List<String>> speed) {}

    private MasterDispatchSetupPrompt() {}

    public static List<EquipmentRef> collectPlanEquipment(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path plan = PlanTasksMissingSkillsColumnPrompt.resolvePlanInputPath(u);
        if (!Files.isRegularFile(plan)) {
            return List.of();
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        List<List<String>> rows = tr.tabular().rows();
        if (headers == null || headers.isEmpty() || rows == null) {
            return List.of();
        }
        int iProc = headers.indexOf(COL_PROCESS);
        int iMach = headers.indexOf(COL_MACHINE);
        int iTask = headers.indexOf(COL_TASK);
        int iEx = headers.indexOf(COL_EXCLUDE);
        if (iProc < 0 || iMach < 0) {
            return List.of();
        }
        LinkedHashMap<String, EquipmentRef> byKey = new LinkedHashMap<>();
        for (List<String> row : rows) {
            if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cell(row, iEx))) {
                continue;
            }
            String proc = cell(row, iProc);
            String mach = cell(row, iMach);
            if (proc.isBlank() || mach.isBlank()) {
                continue;
            }
            String nk = MasterTeamCombinationTableReader.normalizedComboKey(proc, mach);
            if (nk.isEmpty() || byKey.containsKey(nk)) {
                continue;
            }
            byKey.put(
                    nk,
                    new EquipmentRef(
                            proc.strip(),
                            mach.strip(),
                            iTask >= 0 ? cell(row, iTask) : ""));
        }
        return List.copyOf(byKey.values());
    }

    public static SheetBundle loadSheetBundle(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path json = AppPaths.masterDispatchSheetsJsonPath(u);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(u);
        String site =
                AppPaths.currentDispatchFactorySite(u) != null
                        ? AppPaths.currentDispatchFactorySite(u).name()
                        : "";
        MasterDispatchSheetsDocument doc =
                MasterDispatchSheetsSeeder.loadOrImport(json, source, site, false).document();
        return fromDocument(doc);
    }

    public static SheetBundle fromDocument(MasterDispatchSheetsDocument doc) {
        MasterDispatchSheetsDocument d =
                doc != null ? doc : MasterDispatchSheetsDocument.empty("");
        return new SheetBundle(
                rowsOf(d, MasterDispatchSheetsDocument.KEY_SKILLS),
                rowsOf(d, MasterDispatchSheetsDocument.KEY_NEED),
                rowsOf(d, MasterDispatchSheetsDocument.KEY_TEAM_COMBINATIONS),
                rowsOf(d, MasterDispatchSheetsDocument.KEY_SPEED));
    }

    public static Evaluation evaluate(Map<String, String> ui) throws IOException {
        return evaluate(collectPlanEquipment(ui), loadSheetBundle(ui));
    }

    public static Evaluation evaluate(List<EquipmentRef> equipment, SheetBundle sheets) {
        SheetBundle s =
                sheets != null
                        ? sheets
                        : new SheetBundle(List.of(), List.of(), List.of(), List.of());
        return MasterDispatchSetupCompleteness.evaluate(
                equipment, s.skills(), s.need(), s.combinations(), s.speed());
    }

    private static List<List<String>> rowsOf(MasterDispatchSheetsDocument doc, String key) {
        MasterDispatchSheetsDocument.SheetGrid g = doc.sheet(key);
        List<List<String>> rows = g != null ? g.rows() : null;
        return rows != null ? rows : List.of();
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }
}
