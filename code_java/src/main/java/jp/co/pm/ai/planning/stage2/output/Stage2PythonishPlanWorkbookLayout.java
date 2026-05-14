package jp.co.pm.ai.planning.stage2.output;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * Python 段階2の計画ブック（{@code _stage2_tabular_sheet_order}）と同じシート名・並びで xlsx を組み立てる。
 * セル内容はプレースホルダ（配台未実行）だが、UI／JSON ミラーがシート有無で分岐しないようにする。
 */
public final class Stage2PythonishPlanWorkbookLayout {

    /** Python {@code _core._stage2_tabular_sheet_order} と同一順（左から）。 */
    public static final List<String> PLAN_SHEET_ORDER =
            List.of(
                    "結果_設備毎の時間割",
                    "結果_設備毎の時間割_機械名毎",
                    "結果_カレンダー(出勤簿)",
                    "結果_メンバー別作業割合",
                    "列設定_結果_タスク一覧",
                    "結果_タスク一覧",
                    "結果_配台表",
                    "結果_AIログ");

    /** Python {@code RESULT_DISPATCH_TABLE_STATIC_HEADERS} + 動的配台列。 */
    private static final List<String> RESULT_DISPATCH_HEADERS =
            List.of(
                    "配台試行順番",
                    "工程名",
                    "機械名",
                    "受注日",
                    "受注NO",
                    "依頼NO",
                    "品名(原反)",
                    "使用原反",
                    "原反数",
                    "品名(製品)",
                    "製品名",
                    "換算数量",
                    "実加工数",
                    "加工内容",
                    "在庫場所",
                    "原反投入日",
                    "指定納期",
                    "回答納期",
                    "加工完了日",
                    "加工完了区分",
                    "実出来高",
                    "計画合計",
                    "原反投入場所",
                    "加工開始日時",
                    "加工終了日時",
                    "メンバー名",
                    "配台日",
                    "当日配台数量");

    private Stage2PythonishPlanWorkbookLayout() {}

    public static void write(
            Path path,
            PlanInputTabularIo.TabularSheet tasks,
            List<String> equipmentProcPlusMachine,
            List<String> memberDisplayNames)
            throws IOException {
        write(path, tasks, equipmentProcPlusMachine, memberDisplayNames, null);
    }

    /**
     * @param rollTableRepoRoot {@code PM_AI_REPO_ROOT} 相当。非 null なら {@code code/} 下のロール単位長さテーブルを読み、
     *     Python {@code _roll_unit_m_estimate_from_plan_row} に近い推定で未加工≤0 時の残量を算定する。
     */
    public static void write(
            Path path,
            PlanInputTabularIo.TabularSheet tasks,
            List<String> equipmentProcPlusMachine,
            List<String> memberDisplayNames,
            Path rollTableRepoRoot)
            throws IOException {
        Files.createDirectories(path.getParent());
        List<String> eq =
                equipmentProcPlusMachine != null && !equipmentProcPlusMachine.isEmpty()
                        ? equipmentProcPlusMachine
                        : List.of("未設定+プレースホルダ");
        List<String> schedLabels = Stage2EquipmentScheduleHeaderLabels.fromEquipmentCombos(eq);
        List<String> members =
                memberDisplayNames != null && !memberDisplayNames.isEmpty()
                        ? memberDisplayNames
                        : List.of("（メンバー未設定）");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            for (String sheetName : PLAN_SHEET_ORDER) {
                switch (sheetName) {
                    case "結果_設備毎の時間割" -> writeEquipmentScheduleSheet(wb, sheetName, eq, schedLabels);
                    case "結果_設備毎の時間割_機械名毎" ->
                            writeEquipmentByMachineSheet(wb, sheetName, schedLabels);
                    case "結果_カレンダー(出勤簿)" -> writeCalendarHeaderOnly(wb, sheetName);
                    case "結果_メンバー別作業割合" -> writeMemberUtilHeader(wb, sheetName, members);
                    case "列設定_結果_タスク一覧" -> writeColumnConfigSheet(wb, sheetName, tasks.headers());
                    case "結果_タスク一覧" -> writeResultTaskListSheet(wb, sheetName, tasks, rollTableRepoRoot);
                    case "結果_配台表" -> writeDispatchHeaderOnly(wb, sheetName);
                    case "結果_AIログ" -> writeAiLogPlaceholder(wb, sheetName);
                    default -> wb.createSheet(safeSheetName(sheetName));
                }
            }
            try (OutputStream os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }

    private static String safeSheetName(String name) {
        String n = name == null ? "sheet" : name.strip();
        if (n.length() > 31) {
            n = n.substring(0, 31);
        }
        return n.isEmpty() ? "sheet" : n;
    }

    private static void writeEquipmentScheduleSheet(
            XSSFWorkbook wb, String sheetName, List<String> eqCombos, List<String> schedLabels) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        List<String> headers = new ArrayList<>();
        headers.add("日時帯");
        for (int i = 0; i < eqCombos.size(); i++) {
            String lab = i < schedLabels.size() ? schedLabels.get(i) : eqCombos.get(i);
            headers.add(lab);
            headers.add(lab + "進度");
        }
        writeHeaderRow(sh, headers);
    }

    private static void writeEquipmentByMachineSheet(
            XSSFWorkbook wb, String sheetName, List<String> schedLabels) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        LinkedHashSet<String> uniq = new LinkedHashSet<>();
        List<String> headers = new ArrayList<>();
        headers.add("日時帯");
        for (String lab : schedLabels) {
            if (uniq.add(lab)) {
                headers.add(lab);
            }
        }
        writeHeaderRow(sh, headers);
    }

    private static void writeCalendarHeaderOnly(XSSFWorkbook wb, String sheetName) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        writeHeaderRow(sh, List.of("日付", "メンバー", "出勤", "退勤", "効率", "備考"));
    }

    private static void writeMemberUtilHeader(XSSFWorkbook wb, String sheetName, List<String> members) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        List<String> headers = new ArrayList<>();
        headers.add("年月日");
        headers.addAll(members);
        writeHeaderRow(sh, headers);
    }

    private static void writeColumnConfigSheet(XSSFWorkbook wb, String sheetName, List<String> taskHeaders) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        Row hr = sh.createRow(0);
        hr.createCell(0).setCellValue("列名");
        hr.createCell(1).setCellValue("表示");
        int r = 1;
        for (String col : taskHeaders) {
            Row rr = sh.createRow(r++);
            rr.createCell(0).setCellValue(col != null ? col : "");
            rr.createCell(1).setCellValue("True");
        }
    }

    /**
     * Python {@code default_result_task_sheet_column_order(0)} と同じ 28 列見出しで、計画入力の列を同名または
     * 依頼NO→タスクID にマッピングして転写する（配台未実行のため他列は空）。
     */
    private static void writeResultTaskListSheet(
            XSSFWorkbook wb, String sheetName, PlanInputTabularIo.TabularSheet tasks, Path rollTableRepoRoot)
            throws IOException {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        List<String> canon = Stage2ResultTaskListCanonicalHeaders.DEFAULT_ORDER_NO_HISTORY;
        writeTabular(sh, canon, expandResultTaskRows(tasks, canon, rollTableRepoRoot));
    }

    private static List<List<String>> expandResultTaskRows(
            PlanInputTabularIo.TabularSheet tasks, List<String> outCols, Path rollTableRepoRoot)
            throws IOException {
        Stage2RollUnitLengthTables tables =
                rollTableRepoRoot != null
                        ? Stage2RollUnitLengthTables.load(rollTableRepoRoot)
                        : Stage2RollUnitLengthTables.empty();
        List<String> inHeaders = tasks.headers();
        Map<String, Integer> inIndex = new HashMap<>();
        for (int i = 0; i < inHeaders.size(); i++) {
            String k = inHeaders.get(i) == null ? "" : inHeaders.get(i).strip();
            inIndex.put(k, i);
        }
        Map<String, String> planKeyToResultCol = new HashMap<>();
        planKeyToResultCol.put("依頼NO", "タスクID");
        String[] sameName =
                new String[] {
                    "工程名",
                    "機械名",
                    "回答納期",
                    "指定納期",
                    "原反投入日",
                    "加工速度",
                    "優先度",
                    "担当OP指定",
                };
        for (String s : sameName) {
            planKeyToResultCol.put(s, s);
        }
        Map<String, Integer> outIndex = new HashMap<>();
        for (int i = 0; i < outCols.size(); i++) {
            outIndex.put(outCols.get(i), i);
        }

        List<List<String>> out = new ArrayList<>();
        for (List<String> row : tasks.rows()) {
            List<String> cells = new ArrayList<>(Collections.nCopies(outCols.size(), ""));
            for (Map.Entry<String, String> e : planKeyToResultCol.entrySet()) {
                Integer si = inIndex.get(e.getKey());
                Integer ti = outIndex.get(e.getValue());
                if (si == null || ti == null || si < 0 || ti < 0) {
                    continue;
                }
                if (si < row.size()) {
                    String v = row.get(si);
                    cells.set(ti, v != null ? v : "");
                }
            }
            applyDispatchQtyMetricsToCells(cells, outIndex, inHeaders, row, tables);
            out.add(cells);
        }
        return out;
    }

    private static void applyDispatchQtyMetricsToCells(
            List<String> cells,
            Map<String, Integer> outIndex,
            List<String> inHeaders,
            List<String> row,
            Stage2RollUnitLengthTables tables) {
        Map<String, String> rowMap = new HashMap<>();
        for (int c = 0; c < inHeaders.size(); c++) {
            String hk = inHeaders.get(c) == null ? "" : inHeaders.get(c).strip();
            String val = "";
            if (c < row.size() && row.get(c) != null) {
                val = row.get(c).strip();
            }
            rowMap.put(hk, val);
        }
        Stage2PlanRowDispatchQtyMetrics.compute(rowMap, tables)
                .flatMap(Stage2PlanRowDispatchQtyMetrics::toResultSheetStrings)
                .ifPresent(
                        str -> {
                            putAt(outIndex, cells, "残加工量", str.remainingM());
                            putAt(outIndex, cells, "累計加工量", str.cumulativeDoneM());
                            putAt(outIndex, cells, "完了率(実行時点)", str.completionPct());
                        });
    }

    private static void putAt(Map<String, Integer> outIndex, List<String> cells, String col, String value) {
        Integer i = outIndex.get(col);
        if (i != null && i >= 0 && i < cells.size() && value != null && !value.isEmpty()) {
            cells.set(i, value);
        }
    }

    private static void writeDispatchHeaderOnly(XSSFWorkbook wb, String sheetName) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        writeHeaderRow(sh, RESULT_DISPATCH_HEADERS);
    }

    private static void writeAiLogPlaceholder(XSSFWorkbook wb, String sheetName) {
        Sheet sh = wb.createSheet(safeSheetName(sheetName));
        Row h = sh.createRow(0);
        h.createCell(0).setCellValue("項目");
        h.createCell(1).setCellValue("内容");
        Row d = sh.createRow(1);
        d.createCell(0).setCellValue("Java段階2");
        d.createCell(1).setCellValue("プレースホルダ（配台結果なし。本番の計画表は Python 段階2の出力を使用）");
    }

    private static void writeHeaderRow(Sheet sh, List<String> headers) {
        Row hr = sh.createRow(0);
        for (int c = 0; c < headers.size(); c++) {
            Cell cell = hr.createCell(c);
            String v = headers.get(c);
            cell.setCellValue(v != null ? v : "");
        }
    }

    private static void writeTabular(Sheet sh, List<String> headers, List<List<String>> rows) {
        Row hr = sh.createRow(0);
        for (int c = 0; c < headers.size(); c++) {
            Cell cell = hr.createCell(c);
            String v = headers.get(c);
            cell.setCellValue(v != null ? v : "");
        }
        int r = 1;
        if (rows != null) {
            for (List<String> rowVals : rows) {
                Row rr = sh.createRow(r++);
                for (int c = 0; c < headers.size(); c++) {
                    Cell cell = rr.createCell(c);
                    String v =
                            c < rowVals.size() && rowVals.get(c) != null ? rowVals.get(c) : "";
                    cell.setCellValue(v);
                }
            }
        }
    }
}
