package jp.co.pm.ai.desktop.dispatch.rules.simulation;

/** One simulation step from Python {@code simulate_task}. */
public record DispatchRuleSimulationStep(
        int sequence,
        String phase,
        String ruleId,
        String nodeId,
        String nodeType,
        String edgeFrom,
        String edgeTo,
        String effect,
        String summaryJa,
        int rollIndex,
        int rollTotal,
        double wipCount,
        String animationKind,
        int preInputRawRolls,
        int connectionRolls,
        int secBeforeWipRolls,
        int secCompleteRolls,
        String flowPhase) {

    public boolean rollAccumulateStep() {
        return "roll_accumulate".equals(animationKind);
    }
}
