package jp.co.pm.ai.desktop.dispatch.rules.validation;

import static org.junit.jupiter.api.Assertions.assertFalse;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEntry;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleGraph;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleNode;

class DispatchRuleConflictCheckerTest {

    @Test
    void detectsDuplicateSpeedRules() {
        DispatchRuleDocument doc = new DispatchRuleDocument();
        doc.rules.add(speedRule("L4", 20));
        doc.rules.add(speedRule("L6", 30));
        var report = DispatchRuleConflictChecker.check(doc);
        assertFalse(report.conflicts().isEmpty());
    }

    private static DispatchRuleEntry speedRule(String id, int order) {
        DispatchRuleEntry e = new DispatchRuleEntry();
        e.id = id;
        e.name = id;
        e.enabled = true;
        e.applyOrder = order;
        e.graph = new DispatchRuleGraph();
        DispatchRuleNode n = new DispatchRuleNode();
        n.type = "action.set_speed_mpm";
        e.graph.nodes.add(n);
        return e;
    }
}
