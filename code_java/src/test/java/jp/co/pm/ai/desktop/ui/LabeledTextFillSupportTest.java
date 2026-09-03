package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import javafx.application.Platform;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.text.Text;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class LabeledTextFillSupportTest {

    @BeforeAll
    static void startFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            /* already started */
        }
    }

    @Test
    void applyCssColorToken_setsInlineFillOnTextAndLabeled() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        Text text = new Text("x");
                        Label label = new Label("y");
                        label.setGraphic(text);
                        LabeledTextFillSupport.applyCssColorToken(
                                label, LabeledTextFillSupport.THEME_MID);
                        assertTrue(label.getStyle().contains("-fx-mid-text-color"));
                        assertTrue(text.getStyle().contains("-fx-fill: -fx-mid-text-color"));
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
    }

    @Test
    void applyToButton_setsFillOnGraphicTextWithoutReplacingButtonStyle() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        Button button = new Button("保存");
                        button.setStyle("-fx-border-color: #38bdf8;");
                        Text ghost = new Text("保存");
                        button.setGraphic(ghost);
                        LabeledTextFillSupport.applyToButton(
                                button, LabeledTextFillSupport.THEME_MID);
                        assertTrue(button.getStyle().contains("-fx-border-color: #38bdf8"));
                        assertTrue(ghost.getStyle().contains("-fx-fill: -fx-mid-text-color"));
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
    }
}
