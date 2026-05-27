package jp.co.pm.ai.desktop.dispatch.rules.simulation;

import java.util.List;

/** Simulation result for test lab. */
public record DispatchRuleSimulationResult(
        boolean finalBlocked,
        String summaryJa,
        List<DispatchRuleSimulationStep> steps,
        int rollTotal,
        int blockedAtRoll) {

    public DispatchRuleSimulationResult {
        steps = steps != null ? List.copyOf(steps) : List.of();
    }

    public DispatchRuleSimulationResult(boolean finalBlocked, String summaryJa, List<DispatchRuleSimulationStep> steps) {
        this(finalBlocked, summaryJa, steps, 0, 0);
    }

    public static DispatchRuleSimulationResult empty(String message) {
        return new DispatchRuleSimulationResult(false, message, List.of());
    }
}
