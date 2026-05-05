package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * 結果_タスク一覧 JSON から「ステータス=配台不可」の行を抽出する。
 *
 * <p>パス解決は {@code plan_workbook_sidecar.result_task_sidecar_path} と同じ規則（
 * production_plan のベース名 + {@code _結果_タスク一覧.json}）。{@code PM_AI_PLAN_RESULT_TASK_JSON_PATH}
 * が有効なファイルを指すときはそちらを優先する。
 */
public final class ResultTaskUnassignedLoader {

    private static final String COL_STATUS = "ステータス";
    /** 結果シートは {@code 配台不可} または {@code 配台不可（納期見直し必須）} 等 */
    private static final String STATUS_UNASSIGNED_PREFIX = "配台不可";
    private static final String COL_TASK_ID = "タスクID";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_MEMO = "配台状況メモ";

    private ResultTaskUnassignedLoader() {}

    /**
     * @param situationMemo 結果シート「配台状況メモ」と不足JSONヒントを統合した説明文
     */
    public record UnassignedRow(String taskId, String processName, String machineName, String situationMemo) {}

    public static List<UnassignedRow> loadUnassigned(Map<String, String> uiEnv, String productionPlanPathStr)
            throws IOException {
        return loadUnassigned(uiEnv, productionPlanPathStr, List.of());
    }

    /**
     * @param shortageHints {@code dispatch_trial_shortages.json} の op/as 不足（タスクID一致で補足文に結合）
     */
    public static List<UnassignedRow> loadUnassigned(
            Map<String, String> uiEnv,
            String productionPlanPathStr,
            List<DispatchTrialShortages.ShortageHint> shortageHints)
            throws IOException {
        String p = productionPlanPathStr != null ? productionPlanPathStr.trim() : "";
        if (p.isEmpty()) {
            return List.of();
        }
        Path sidecar = resolveResultTaskSidecarJson(uiEnv, Path.of(p));
        if (!Files.isRegularFile(sidecar)) {
            return List.of();
        }
        JsonTableIo.SheetTable table = JsonTableIo.loadFlatTable(sidecar);
        List<UnassignedRow> out = new ArrayList<>();
        for (Map<String, String> row : table.rows()) {
            String st = nz(row.get(COL_STATUS));
            if (!st.startsWith(STATUS_UNASSIGNED_PREFIX)) {
                continue;
            }
            String tid = nz(row.get(COL_TASK_ID));
            String memo = nz(row.get(COL_MEMO));
            if (memo.isEmpty() && shortageHints != null && !shortageHints.isEmpty()) {
                memo = DispatchTrialShortages.mergeHintsForTask(shortageHints, tid);
            }
            if (memo.isEmpty()) {
                memo = "（配台状況メモなし。計画上は割当履歴なし・残ありの組合せです）";
            }
            out.add(new UnassignedRow(tid, nz(row.get(COL_PROCESS)), nz(row.get(COL_MACHINE)), memo));
        }
        return List.copyOf(out);
    }

    static Path resolveResultTaskSidecarJson(Map<String, String> uiEnv, Path productionPlanPath) {
        if (uiEnv != null) {
            String override = nz(uiEnv.get(AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH)).trim();
            if (!override.isEmpty()) {
                Path o = Path.of(override);
                if (Files.isRegularFile(o)) {
                    return o;
                }
            }
        }
        Path parent = productionPlanPath.getParent();
        if (parent == null) {
            return Path.of(filenameStem(productionPlanPath.getFileName().toString()) + "_結果_タスク一覧.json");
        }
        String stem = filenameStem(productionPlanPath.getFileName().toString());
        return parent.resolve(stem + "_結果_タスク一覧.json");
    }

    static String filenameStem(String fileName) {
        if (fileName == null || fileName.isEmpty()) {
            return "";
        }
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    private static String nz(String s) {
        return s != null ? s.trim() : "";
    }
}
