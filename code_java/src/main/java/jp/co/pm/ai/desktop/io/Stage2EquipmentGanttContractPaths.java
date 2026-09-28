package jp.co.pm.ai.desktop.io;

import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 段階2計画成果物に付随する設備ガント描画契約 JSON（{@code …設.json} 等）のパス解決。
 */
public final class Stage2EquipmentGanttContractPaths {

    private Stage2EquipmentGanttContractPaths() {}

    /**
     * 計画 xlsx/json パスから兄弟の設備ガント契約 JSON を探す。見つからなければ null。
     */
    public static Path resolveEquipmentContractSibling(Path planArtifactPath) {
        ContractCandidates c = candidates(planArtifactPath);
        if (c == null) {
            return null;
        }
        if (c.selfIfContract != null) {
            return Files.isRegularFile(c.selfIfContract) ? c.selfIfContract : null;
        }
        if (c.modern != null && Files.isRegularFile(c.modern)) {
            return c.modern;
        }
        if (c.legacy != null && Files.isRegularFile(c.legacy)) {
            return c.legacy;
        }
        return null;
    }

    /**
     * スナップショットへコピーする契約パス。ファイルが無いときは期待する {@code …設.json} を返す
     * （呼び出し側が missing に記録できる）。計画以外は null。
     */
    public static Path preferredContractSiblingForCopy(Path planArtifactPath) {
        ContractCandidates c = candidates(planArtifactPath);
        if (c == null) {
            return null;
        }
        if (c.selfIfContract != null) {
            return Files.isRegularFile(c.selfIfContract) ? c.selfIfContract : null;
        }
        if (c.modern != null && Files.isRegularFile(c.modern)) {
            return c.modern;
        }
        if (c.legacy != null && Files.isRegularFile(c.legacy)) {
            return c.legacy;
        }
        return c.modern;
    }

    private record ContractCandidates(Path modern, Path legacy, Path selfIfContract) {}

    private static ContractCandidates candidates(Path planArtifactPath) {
        if (planArtifactPath == null) {
            return null;
        }
        Path fn = planArtifactPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        boolean xlsx = name.endsWith(".xlsx");
        boolean json = name.endsWith(".json");
        if (!xlsx && !json) {
            return null;
        }
        String stem = name.substring(0, name.length() - 5);
        if (json && (stem.endsWith("_equipment_gantt_contract") || stem.endsWith("設"))) {
            return new ContractCandidates(null, null, planArtifactPath);
        }
        String baseStem = stripStage2PlanStemVariants(stem);
        Path modern = planArtifactPath.resolveSibling(baseStem + "設.json");
        Path legacy = planArtifactPath.resolveSibling(baseStem + "_equipment_gantt_contract.json");
        return new ContractCandidates(modern, legacy, null);
    }

    /** {@code 結果_配台表.json} 近傍から設備ガント契約を解決する（shortages の production_plan を優先）。 */
    public static Path resolveNearResultDispatchJson(Path resultDispatchJson) {
        if (resultDispatchJson == null) {
            return null;
        }
        Path shortagePath = resultDispatchJson.resolveSibling("dispatch_trial_shortages.json");
        if (Files.isRegularFile(shortagePath)) {
            try {
                var paths = jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages.read(shortagePath);
                String plan = paths.productionPlan();
                if (plan != null && !plan.isBlank()) {
                    Path fromPlan = resolveEquipmentContractSibling(Path.of(plan));
                    if (fromPlan != null) {
                        return fromPlan;
                    }
                }
            } catch (Exception ignored) {
                // fall through
            }
        }
        Path dir = resultDispatchJson.getParent();
        if (dir != null && Files.isDirectory(dir)) {
            try {
                Path newest =
                        Stage2OutputNaming.newestMatching(
                                dir,
                                p -> {
                                    String n = p.getFileName().toString();
                                    return n.endsWith("設.json")
                                            || n.endsWith("_equipment_gantt_contract.json");
                                });
                if (newest != null) {
                    return newest;
                }
            } catch (Exception ignored) {
                // fall through
            }
        }
        return null;
    }

    static String stripStage2PlanStemVariants(String stem) {
        String s = stem;
        while (true) {
            boolean changed = false;
            if (s.endsWith("_equipment_gantt_contract")) {
                s = s.substring(0, s.length() - "_equipment_gantt_contract".length());
                changed = true;
            } else if (s.endsWith("_logical_view")) {
                s = s.substring(0, s.length() - "_logical_view".length());
                changed = true;
            } else if (s.endsWith("_tabular_source")) {
                s = s.substring(0, s.length() - "_tabular_source".length());
                changed = true;
            } else if (s.endsWith("_actual_detail_gantt_contract")) {
                s = s.substring(0, s.length() - "_actual_detail_gantt_contract".length());
                changed = true;
            } else if (s.endsWith("_結果_タスク一覧")) {
                s = s.substring(0, s.length() - "_結果_タスク一覧".length());
                changed = true;
            } else if (s.endsWith("一覧")) {
                s = s.substring(0, s.length() - 2);
                changed = true;
            } else if (s.endsWith("表") || s.endsWith("論") || s.endsWith("設") || s.endsWith("実")) {
                s = s.substring(0, s.length() - 1);
                changed = true;
            }
            if (!changed) {
                break;
            }
        }
        return s;
    }
}
