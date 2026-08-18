package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Region;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class RunLogListViewSupportTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void longLineConstrainedNarrow_staysSingleLineHeight() {
        Font font = Font.font(16);
        String item =
                "[request-form-input] "
                        + "C:\\very\\long\\path\\".repeat(20)
                        + "file.xlsx";
        Node graphic =
                RunLogListViewSupport.buildLineGraphic(item, font, Color.WHITE, "");
        assertInstanceOf(HBox.class, graphic);
        assertFalse(graphic instanceof TextFlow);
        Region region = (Region) graphic;
        double maxSingleLine = RunLogListViewSupport.measureLineHeightPx(font) * 1.6;
        assertTrue(
                region.prefHeight(220) <= maxSingleLine,
                () -> "expected single-line pref height but was " + region.prefHeight(220));
        for (Node child : region.getChildrenUnmodifiable()) {
            if (child instanceof Text text) {
                assertEquals(0, text.getWrappingWidth(), 0.001);
            }
        }
    }

    @Test
    void searchHits_staySingleLineAndKeepHitClass() {
        Font font = Font.font(14);
        Node graphic =
                RunLogListViewSupport.buildLineGraphic(
                        "[env] PM_AI_TASK_INPUT_SOURCE_DIR=\\\\server\\share",
                        font,
                        Color.WHITE,
                        "TASK");
        Region region = (Region) graphic;
        double maxSingleLine = RunLogListViewSupport.measureLineHeightPx(font) * 1.6;
        assertTrue(region.prefHeight(180) <= maxSingleLine);
        long hits =
                region.getChildrenUnmodifiable().stream()
                        .filter(n -> n instanceof Text t && t.getStyleClass().contains("pm-log-search-hit"))
                        .count();
        assertEquals(1, hits);
    }

    @Test
    void fixedCellSize_coversMeasuredLineAndCellPadding() {
        Font font = Font.font(16);
        double line = RunLogListViewSupport.measureLineHeightPx(font);
        double cell = RunLogListViewSupport.fixedCellSizePx(font);
        assertTrue(cell >= line + RunLogListViewSupport.CELL_VERTICAL_PADDING_PX);
        assertTrue(cell <= 72.0);
    }
}
