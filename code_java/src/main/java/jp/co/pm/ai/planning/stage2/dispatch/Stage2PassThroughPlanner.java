package jp.co.pm.ai.planning.stage2.dispatch;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.SkillsSheetEquipmentListReader;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.planning.stage2.Stage2ExitCodes;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;
import jp.co.pm.ai.planning.stage2.output.Stage2EquipmentGanttContractPlaceholder;
import jp.co.pm.ai.planning.stage2.output.Stage2PythonishMemberWorkbookLayout;
import jp.co.pm.ai.planning.stage2.output.Stage2PythonishPlanWorkbookLayout;
import jp.co.pm.ai.planning.stage2.output.Stage2WorkbookJsonWriter;

/**
 * 日内配台のプレースホルダ: 計画入力を「結果_タスク一覧」に転写し、Python 段階2と同一シート名・並びの計画ブック xlsx を出力する。
 *
 * <p>本経路は Python {@code _generate_plan_impl} 相当の完全置換ではない。計画 xlsx は Python 段階2と<strong>同一のシート名・並び</strong>で
 * プレースホルダを出す（中身は未配台）。計画ブック JSON をミラーする場合のみ、設備ガント契約（計画 JSON と兄弟の「設」付き JSON）を
 * Python と同型の<strong>空イベント</strong>プレースホルダとして出力する。実イベント・本番ガントは {@code PM_AI_STAGE2_ENGINE=python} を推奨。
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

        ctx.log(
                "[stage2-java] 注意: 配台コアは未移植。計画ブックは Python と同じ 8 シート構成のプレースホルダです。"
                        + " 計画 JSON ミラー時は設備ガント契約 JSON（空イベント）も出します。本番は PM_AI_STAGE2_ENGINE=python を推奨。");
        PlanInputTabularIo.TabularSheet tasks = snap.planningTasksSheet();
        List<String> equipmentCombos =
                SkillsSheetEquipmentListReader.readEquipmentProcPlusMachineCombos(snap.masterPath());
        Stage2PythonishPlanWorkbookLayout.write(
                planXlsx, tasks, equipmentCombos, snap.memberDisplayNames());
        Stage2PythonishMemberWorkbookLayout.write(
                memberXlsx,
                snap.memberDisplayNames(),
                runNow.toLocalDate(),
                snap.factoryStart(),
                snap.factoryEnd());

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
            Path contractJson = outputDir.resolve(planBase + "設.json");
            Stage2EquipmentGanttContractPlaceholder.write(contractJson, runNow.toLocalDate());
            ctx.log("[stage2-java] 設備ガント契約 JSON（プレースホルダ）: " + contractJson);
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
}
