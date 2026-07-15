package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.MainShellTabId;

class Stage3MainShellLayoutTest {

    @Test
    void mergeHiddenTabPreservesItsNestedPositionWhileKeepingVisibleReorder() {
        var stage3 = tab(MainShellTabId.PLAN_INPUT_STAGE3.key());
        var complete =
                List.of(
                        MainShellTabLayoutNode.groupNode(
                                "入力",
                                "",
                                List.of(tab("run"), stage3, tab("planInput"))),
                        tab("env"));
        var visible =
                List.of(
                        MainShellTabLayoutNode.groupNode(
                                "入力", "", List.of(tab("planInput"), tab("run"))),
                        tab("env"));

        var merged = Stage3MainShellLayout.mergeHiddenTab(visible, complete);

        assertEquals(
                List.of("planInput", MainShellTabId.PLAN_INPUT_STAGE3.key(), "run"),
                merged.getFirst().children().stream().map(MainShellTabLayoutNode::id).toList());
    }

    @Test
    void removeAndMergeRoundTripKeepsCompleteLayout() {
        var complete =
                List.of(
                        tab("run"),
                        tab(MainShellTabId.PLAN_INPUT_STAGE3.key()),
                        tab("planInput"));

        assertEquals(
                complete,
                Stage3MainShellLayout.mergeHiddenTab(
                        Stage3MainShellLayout.withoutStage3(complete), complete));
    }

    private static MainShellTabLayoutNode tab(String id) {
        return MainShellTabLayoutNode.tabNode(id, "");
    }
}
