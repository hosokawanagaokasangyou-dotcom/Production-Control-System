package jp.co.pm.ai.desktop.dispatch.rules.execution;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEntry;

class DispatchRuleExecutionPlannerTest {

    @Test
    void legacyModeWhenEngineOff() {
        DispatchRuleEntry e = new DispatchRuleEntry();
        e.enabled = true;
        e.executionMode = "auto";
        assertEquals(
                DispatchRuleExecutionPlanner.Source.LEGACY,
                DispatchRuleExecutionPlanner.resolveSource(e, false));
    }

    @Test
    void dslModeWhenForced() {
        DispatchRuleEntry e = new DispatchRuleEntry();
        e.enabled = true;
        e.executionMode = "dsl";
        assertEquals(
                DispatchRuleExecutionPlanner.Source.DSL,
                DispatchRuleExecutionPlanner.resolveSource(e, false));
    }
}
