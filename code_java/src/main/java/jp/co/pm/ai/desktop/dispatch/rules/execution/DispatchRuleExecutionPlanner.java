package jp.co.pm.ai.desktop.dispatch.rules.execution;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEntry;

/** Resolves enabled rules in applyOrder (Python execution_planner parity subset). */
public final class DispatchRuleExecutionPlanner {

    public enum Source {
        DSL,
        LEGACY,
        SKIP
    }

    public record PlannedEntry(DispatchRuleEntry rule, Source source) {}

    private DispatchRuleExecutionPlanner() {}

    public static List<PlannedEntry> plan(DispatchRuleDocument doc, boolean engineGloballyEnabled) {
        List<DispatchRuleEntry> enabled = new ArrayList<>();
        for (DispatchRuleEntry r : doc.rules) {
            if (r.enabled) {
                enabled.add(r);
            }
        }
        enabled.sort(Comparator.comparingInt(r -> r.applyOrder));
        List<PlannedEntry> out = new ArrayList<>();
        for (DispatchRuleEntry r : enabled) {
            out.add(new PlannedEntry(r, resolveSource(r, engineGloballyEnabled)));
        }
        return out;
    }

    public static Source resolveSource(DispatchRuleEntry rule, boolean engineGloballyEnabled) {
        if (!rule.enabled) {
            return Source.SKIP;
        }
        String mode = rule.executionMode != null ? rule.executionMode.strip().toLowerCase() : "auto";
        return switch (mode) {
            case "legacy" -> Source.LEGACY;
            case "dsl" -> Source.DSL;
            default -> engineGloballyEnabled ? Source.DSL : Source.LEGACY;
        };
    }
}
