package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import org.junit.jupiter.api.Test;

class MainRunStage2ProgressTest {

    @Test
    void excelGenerationStateHasVisibleRunningPresentation() {
        assertEquals(
                "アラジン入力用Excelを生成中…",
                MainRunStage2Progress.State.EXCEL_GENERATING.message());
        assertEquals(
                "pm-stage2-progress-running",
                MainRunStage2Progress.State.EXCEL_GENERATING.styleClass());
    }

    @Test
    void fxmlExposesStage2ProgressAccordionControls() throws IOException {
        try (InputStream in =
                MainRunStage2ProgressTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/MainRunTab.fxml")) {
            assertTrue(in != null);
            String fxml = new String(in.readAllBytes(), StandardCharsets.UTF_8);
            assertTrue(fxml.contains("fx:id=\"stage2ProgressAccordion\""));
            assertTrue(fxml.contains("fx:id=\"stage2ProgressPane\""));
            assertTrue(fxml.contains("fx:id=\"stage2ProgressLabel\""));
        }
    }
}
