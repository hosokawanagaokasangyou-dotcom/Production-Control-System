package jp.co.pm.ai.planning.stage2.dispatch;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.planning.stage2.Stage2ExitCodes;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;
import jp.co.pm.ai.planning.stage2.output.Stage2WorkbookJsonWriter;

/**
 * 日内配台のプレースホルダ: 計画入力を「結果_タスク一覧」シートに転写し、メンバーごとの空スケジュール表を載せた成果物を出力する。
 */
public final class Stage2PassThroughPlanner {

    private Stage2PassThroughPlanner() {}

    public static int run(Stage2RunContext ctx, Stage2InputSnapshot snap, Path outputDir) throws IOException {
        LocalDateTime runNow = LocalDateTime.now();
        String stamp = Stage2OutputNaming.formatStage2Stamp(runNow, runNow);
        String planBase = Stage2OutputNaming.PLAN_PREFIX + stamp;
        String memberBase = Stage2OutputNaming.MEMBER_PREFIX + stamp;
        Path planXlsx = outputDir.resolve(planBase + ".xlsx");
        Path memberXlsx = outputDir.resolve(memberBase + ".xlsx");
        Path planJson = outputDir.resolve(planBase + ".json");
        Path memberJson = outputDir.resolve(memberBase + ".json");

        writePlanWorkbook(planXlsx, snap.planningTasksSheet());
        writeMemberWorkbook(memberXlsx, snap.memberDisplayNames());

        if (ctx.stage2WriteExcel()) {
            ctx.log("[stage2-java] 成果 xlsx: " + planXlsx + " | " + memberXlsx);
        } else {
            ctx.log("[stage2-java] PM_AI_STAGE2_WRITE_EXCEL=0 — 一時 xlsx から JSON のみ確定します。");
        }

        if (ctx.mirrorPlanWorkbookJson()) {
            if (ctx.stage2WriteExcel()) {
                Stage2WorkbookJsonWriter.writeFromXlsx(planXlsx, planJson, Map.of());
            } else {
                Map<String, Object> payload =
                        Stage2WorkbookJsonWriter.buildPayloadFromXlsx(
                                planXlsx, Map.of("engine_note", "java_stage2_pass_through"));
                Stage2WorkbookJsonWriter.writePayload(planJson, payload);
                Files.deleteIfExists(planXlsx);
            }
            ctx.log("[stage2-java] 計画ブック JSON: " + planJson);
        } else {
            ctx.log("[stage2-java] PM_AI_PLAN_WORKBOOK_JSON 無効 — 計画 JSON はスキップ。");
            if (!ctx.stage2WriteExcel()) {
                Files.deleteIfExists(planXlsx);
            }
        }

        if (ctx.mirrorMemberScheduleJson()) {
            Map<String, Object> meta = new LinkedHashMap<>();
            meta.put("workbook_kind", "member_schedule");
            if (ctx.stage2WriteExcel()) {
                Stage2WorkbookJsonWriter.writeFromXlsx(memberXlsx, memberJson, meta);
            } else {
                Map<String, Object> payload = Stage2WorkbookJsonWriter.buildPayloadFromXlsx(memberXlsx, meta);
                Stage2WorkbookJsonWriter.writePayload(memberJson, payload);
                Files.deleteIfExists(memberXlsx);
            }
            ctx.log("[stage2-java] メンバー JSON: " + memberJson);
        } else {
            ctx.log("[stage2-java] PM_AI_MEMBER_SCHEDULE_JSON 無効 — メンバー JSON はスキップ。");
            if (!ctx.stage2WriteExcel()) {
                Files.deleteIfExists(memberXlsx);
            }
        }

        if (!ctx.stage2WriteExcel()) {
            Files.deleteIfExists(planXlsx);
            Files.deleteIfExists(memberXlsx);
        }

        ctx.log("[stage2-java] 完了（stamp=" + stamp + "）。本経路は配台アルゴリズムの段階移植用の足場です。");
        return Stage2ExitCodes.OK;
    }

    private static void writePlanWorkbook(Path path, PlanInputTabularIo.TabularSheet tasks) throws IOException {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet("結果_タスク一覧");
            writeTabular(sh, tasks.headers(), tasks.rows());
            Sheet meta = wb.createSheet("配台計画_タスク入力");
            writeTabular(meta, tasks.headers(), tasks.rows());
            try (var os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }

    private static void writeMemberWorkbook(Path path, List<String> members) throws IOException {
        Files.createDirectories(path.getParent());
        List<String> headers = List.of("日付", "内容");
        Set<String> used = new HashSet<>();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            for (String m : members) {
                String base = sanitizeSheetName(m);
                String name = base;
                int i = 2;
                while (used.contains(name)) {
                    name = base + "_" + i++;
                }
                used.add(name);
                Sheet sh = wb.createSheet(name);
                writeTabular(sh, headers, List.of());
            }
            try (var os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }

    private static String sanitizeSheetName(String name) {
        String n = name == null ? "member" : name.strip();
        if (n.isEmpty()) {
            n = "member";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < n.length() && sb.length() < 28; i++) {
            char c = n.charAt(i);
            if (c == '[' || c == ']' || c == '*' || c == '/' || c == '\\' || c == '?') {
                sb.append('_');
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    private static void writeTabular(Sheet sh, List<String> headers, List<List<String>> rows) {
        Row hr = sh.createRow(0);
        for (int c = 0; c < headers.size(); c++) {
            Cell cell = hr.createCell(c);
            String v = headers.get(c);
            cell.setCellValue(v != null ? v : "");
        }
        int r = 1;
        for (List<String> rowVals : rows) {
            Row rr = sh.createRow(r++);
            for (int c = 0; c < headers.size(); c++) {
                Cell cell = rr.createCell(c);
                String v = c < rowVals.size() && rowVals.get(c) != null ? rowVals.get(c) : "";
                cell.setCellValue(v);
            }
        }
    }
}
