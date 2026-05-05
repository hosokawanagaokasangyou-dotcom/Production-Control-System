package jp.co.pm.ai.desktop.config;

import java.util.List;

/**
 * Paths and fields restored on startup from {@link DesktopSessionStateStore}.
 *
 * @param planInputPath path field on 配台計画_タスク入力 tab
 * @param planInputSheet sheet name on the same tab
 * @param stage1PreviewPath Stage1 preview file path
 * @param stage1PreviewSheet Stage1 preview sheet name
 * @param excludeRulesPath PM_AI_EXCLUDE_RULES_JSON path (editor tab)
 * @param mainRunWorkbook task-input workbook field on run tab
 * @param mainRunPythonExe Python executable field
 * @param mainRunScriptDir code/python directory field
 * @param windowWidth last main window width ({@code 0} if unknown / use default scene size)
 * @param windowHeight last main window height ({@code 0} if unknown)
 * @param windowX last window X ({@link Double#NaN} if unknown / keep toolkit placement)
 * @param windowY last window Y ({@link Double#NaN} if unknown)
 * @param uiTheme persisted UI theme id ({@link DesktopTheme#storedId()}, empty defaults to light)
 * @param logFontFamily run-tab log font family name; empty means default family
 * @param logFontSize run-tab log size in points; {@code 0} means default size
 * @param mainRunLogFilter persisted run-tab log filter enum name ({@code ALL}, {@code ERRORS_ONLY}, ...); empty means ALL
 * @param mainRunLogLines last run-tab log lines (capped when saving)
 * @param mainRunLogScroll vertical scroll position as 0..1 proportion of the scroll bar; {@link Double#NaN} if unknown
 * @param mainRunStage2ProductionPlan last shown stage-2 production_plan xlsx path on run tab (empty if none)
 * @param mainRunStage2MemberSchedule last shown stage-2 member_schedule xlsx path on run tab (empty if none)
 * @param mainRunStage2WriteExcel whether stage-2 writes xlsx deliverables; when false only JSON (run tab)
 * @param mainRunStage2ResultBookFont stage-2 result Excel font family; empty with system default in UI means Python
 *     built-in default
 * @param uiEnvRows persisted 環境変数 tab rows (empty uses bootstrap defaults only)
 * @param mainShellTabOrder ordered {@link jp.co.pm.ai.desktop.MainShellTabId#key()} values for the main window
 *     tab strip; empty restores default FXML order
 * @param equipmentGanttGraphicZoomPercent 設備ガント・グラフィックタブの表示倍率（50〜200、0 は未保存として既定 100）
 * @param equipmentGanttMachineColWidth 同タブ左・機械名列の幅（px、0 は未保存として既定幅を使用）
 * @param equipmentGanttProcessColWidth 同タブ左・工程名列の幅（px、0 は未保存として既定幅を使用）
 * @param equipmentGanttBarFontFamily 同タブタイムライン・バー内ラベル用フォントファミリ（空はシステム既定）
 */
public record DesktopSessionState(
        String planInputPath,
        String planInputSheet,
        String stage1PreviewPath,
        String stage1PreviewSheet,
        String excludeRulesPath,
        String mainRunWorkbook,
        String mainRunPythonExe,
        String mainRunScriptDir,
        double windowWidth,
        double windowHeight,
        double windowX,
        double windowY,
        String uiTheme,
        String logFontFamily,
        double logFontSize,
        String mainRunLogFilter,
        List<String> mainRunLogLines,
        double mainRunLogScroll,
        String mainRunStage2ProductionPlan,
        String mainRunStage2MemberSchedule,
        boolean mainRunStage2WriteExcel,
        String mainRunStage2ResultBookFont,
        List<UiEnvRowSnapshot> uiEnvRows,
        List<String> mainShellTabOrder,
        double equipmentGanttGraphicZoomPercent,
        double equipmentGanttMachineColWidth,
        double equipmentGanttProcessColWidth,
        String equipmentGanttBarFontFamily) {

    public static DesktopSessionState empty() {
        return new DesktopSessionState(
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                0d,
                0d,
                Double.NaN,
                Double.NaN,
                "",
                "",
                0d,
                "",
                List.of(),
                Double.NaN,
                "",
                "",
                true,
                "",
                List.of(),
                List.of(),
                0d,
                0d,
                0d,
                "");
    }
}
