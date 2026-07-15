package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.Stage3UiVisibility;

class Stage3UiPresentationTest {

    @Test
    void hidesOnlyStage3MainTabAndTimingKindsWhenDisabled() {
        Map<String, String> hidden = Map.of(AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE, "0");
        Map<String, String> visible = Map.of(AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE, "1");

        assertFalse(Stage3UiVisibility.isMainShellTabVisible(MainShellTabId.PLAN_INPUT_STAGE3, hidden));
        assertTrue(Stage3UiVisibility.isMainShellTabVisible(MainShellTabId.PLAN_INPUT_STAGE3, visible));
        assertTrue(Stage3UiVisibility.isMainShellTabVisible(MainShellTabId.PLAN_INPUT, hidden));

        assertFalse(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE3_0, hidden));
        assertFalse(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE3_1, hidden));
        assertFalse(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE3_2, hidden));
        assertFalse(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE3, hidden));
        assertTrue(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE2_1, hidden));
        assertTrue(Stage3UiVisibility.isTimingKindVisible(PipelineExecutionTimingKind.STAGE3, visible));
    }

    @Test
    void fxmlExposesStage3OnlyContainersWithoutRevivingLegacyTrialButton() throws IOException {
        String dispatch = resource("DispatchInteractiveTab.fxml");
        assertTrue(dispatch.contains("fx:id=\"stage3QtyControls\""));
        assertTrue(dispatch.contains("fx:id=\"stage3ActionControls\""));
        assertTrue(dispatch.contains("fx:id=\"dispatchPlanningStageBadgeLabel\""));
        assertTrue(
                dispatch.contains(
                        "fx:id=\"dispatchTrialButton\" minWidth=\"-1.0\" mnemonicParsing=\"false\"\n"
                                + "                    text=\"段階3\" visible=\"false\" managed=\"false\""));

        String run = resource("MainRunTab.fxml");
        assertTrue(run.contains("fx:id=\"pipelineTimingStage3Rows\""));

        String pushButtons = resource("PushButtonDesignTab.fxml");
        assertTrue(pushButtons.contains("fx:id=\"stage3PreviewButton\""));
        assertTrue(pushButtons.contains("fx:id=\"stage3DesignControls\""));
    }

    private static String resource(String name) throws IOException {
        try (InputStream in =
                Stage3UiPresentationTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/" + name)) {
            if (in == null) {
                throw new IOException("resource not found: " + name);
            }
            return new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
    }
}
