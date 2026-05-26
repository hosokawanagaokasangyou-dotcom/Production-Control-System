package jp.co.pm.ai.desktop.dispatch.rules.validation;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEntry;

/** Static conflict detection (Python parity subset). */
public final class DispatchRuleConflictChecker {

    public record ConflictItem(
            String kind, String severity, List<String> ruleIds, String message) {}

    public record ConflictReport(List<ConflictItem> conflicts) {
        public long errorCount() {
            return conflicts.stream().filter(c -> "error".equals(c.severity())).count();
        }

        public long warningCount() {
            return conflicts.stream().filter(c -> "warning".equals(c.severity())).count();
        }
    }

    private DispatchRuleConflictChecker() {}

    public static ConflictReport check(DispatchRuleDocument doc) {
        List<ConflictItem> items = new ArrayList<>();
        Map<Integer, List<String>> byOrder = new HashMap<>();
        List<DispatchRuleEntry> enabled = new ArrayList<>();
        for (DispatchRuleEntry r : doc.rules) {
            if (r.enabled) {
                enabled.add(r);
                byOrder.computeIfAbsent(r.applyOrder, k -> new ArrayList<>()).add(r.id);
            }
        }
        byOrder.forEach(
                (order, ids) -> {
                    if (ids.size() > 1) {
                        items.add(
                                new ConflictItem(
                                        "apply_order_tie",
                                        "warning",
                                        List.copyOf(ids),
                                        "applyOrder=" + order + " が重複"));
                    }
                });
        boolean l10 = enabled.stream().anyMatch(r -> "L10".equals(r.id));
        boolean l13 = enabled.stream().anyMatch(r -> "L13".equals(r.id));
        if (l10 && l13) {
            items.add(
                    new ConflictItem(
                            "pipeline_incompatible",
                            "warning",
                            List.of("L10", "L13"),
                            "L10 と L13 が同時有効 — SEC ゲート経路に注意"));
        }
        return new ConflictReport(items);
    }
}
