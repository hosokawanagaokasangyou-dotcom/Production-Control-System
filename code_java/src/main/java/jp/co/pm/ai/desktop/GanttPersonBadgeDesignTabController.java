package jp.co.pm.ai.desktop;

import javafx.animation.PauseTransition;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

/** 設備ガント・担当バッジのデザイン編集タブ。 */
public final class GanttPersonBadgeDesignTabController {

    @FXML
    private Accordion designAccordion;

    @FXML
    private HBox badgePreviewBox;

    @FXML
    private TextField badgeFontField;

    @FXML
    private Slider badgeFontPctSlider;

    @FXML
    private Label badgeFontPctLabel;

    @FXML
    private TextField badgeFillField;

    @FXML
    private TextField badgeTextField;

    @FXML
    private TextField badgeStrokeField;

    @FXML
    private Slider badgeStrokeSlider;

    @FXML
    private Label badgeStrokeLabel;

    @FXML
    private Slider badgeCornerSlider;

    @FXML
    private Label badgeCornerLabel;

    @FXML
    private CheckBox badgePillCheck;

    @FXML
    private TextField badgeGlowColorField;

    @FXML
    private Slider badgeGlowRadiusSlider;

    @FXML
    private Label badgeGlowRadiusLabel;

    @FXML
    private Slider badgeGlowSpreadSlider;

    @FXML
    private Label badgeGlowSpreadLabel;

    @FXML
    private Button badgeResetDefaultsButton;

    private MainShellController shell;

    private PauseTransition persistDelay;

    private boolean suppress;

    @FXML
    private void initialize() {
        if (designAccordion != null && designAccordion.getPanes().size() > 0) {
            TitledPane p = designAccordion.getPanes().get(0);
            designAccordion.setExpandedPane(p);
        }
        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> saveAndRefresh());

        attachListeners();
        refreshPreview();
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
    }

    void applyPersonBadgeDesignSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppress = true;
        try {
            PersonBadgeStyle st = s.resolvedPersonBadgeStyle();
            if (badgeFontField != null) {
                badgeFontField.setText(st.fontFamily() != null ? st.fontFamily() : "");
            }
            if (badgeFontPctSlider != null) {
                badgeFontPctSlider.setValue(Math.clamp(st.fontPercent(), 40, 160));
            }
            if (badgeFillField != null) {
                badgeFillField.setText(st.fillHex());
            }
            if (badgeTextField != null) {
                badgeTextField.setText(st.textHex());
            }
            if (badgeStrokeField != null) {
                badgeStrokeField.setText(st.strokeHex());
            }
            if (badgeStrokeSlider != null) {
                badgeStrokeSlider.setValue(Math.clamp(st.strokeWidth(), 0, 6));
            }
            if (badgeCornerSlider != null) {
                badgeCornerSlider.setValue(Math.clamp(st.cornerRadius(), 0, 24));
            }
            if (badgePillCheck != null) {
                badgePillCheck.setSelected(st.pill());
            }
            if (badgeGlowColorField != null) {
                badgeGlowColorField.setText(st.glowColorHex());
            }
            if (badgeGlowRadiusSlider != null) {
                badgeGlowRadiusSlider.setValue(Math.clamp(st.glowRadius(), 0, 40));
            }
            if (badgeGlowSpreadSlider != null) {
                badgeGlowSpreadSlider.setValue(Math.clamp(st.glowSpread(), 0, 1));
            }
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
    }

    private void attachListeners() {
        Runnable r = this::schedulePersist;
        if (badgeFontField != null) {
            badgeFontField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeFontPctSlider, badgeFontPctLabel, "%.0f%%", r);
        if (badgeFillField != null) {
            badgeFillField.textProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeTextField != null) {
            badgeTextField.textProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeStrokeField != null) {
            badgeStrokeField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeStrokeSlider, badgeStrokeLabel, "%.1f", r);
        wireSlider(badgeCornerSlider, badgeCornerLabel, "%.0f", r);
        if (badgePillCheck != null) {
            badgePillCheck.selectedProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeGlowColorField != null) {
            badgeGlowColorField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeGlowRadiusSlider, badgeGlowRadiusLabel, "%.0f", r);
        wireSlider(badgeGlowSpreadSlider, badgeGlowSpreadLabel, "%.2f", r);
    }

    private static void wireSlider(Slider sl, Label lb, String fmt, Runnable onChange) {
        if (sl == null) {
            return;
        }
        sl.valueProperty()
                .addListener(
                        (o, a, b) -> {
                            if (lb != null) {
                                lb.setText(String.format(fmt, sl.getValue()));
                            }
                            onChange.run();
                        });
    }

    private void syncLabelsFromSliders() {
        if (badgeFontPctSlider != null && badgeFontPctLabel != null) {
            badgeFontPctLabel.setText(String.format("%.0f%%", badgeFontPctSlider.getValue()));
        }
        if (badgeStrokeSlider != null && badgeStrokeLabel != null) {
            badgeStrokeLabel.setText(String.format("%.1f", badgeStrokeSlider.getValue()));
        }
        if (badgeCornerSlider != null && badgeCornerLabel != null) {
            badgeCornerLabel.setText(String.format("%.0f", badgeCornerSlider.getValue()));
        }
        if (badgeGlowRadiusSlider != null && badgeGlowRadiusLabel != null) {
            badgeGlowRadiusLabel.setText(String.format("%.0f", badgeGlowRadiusSlider.getValue()));
        }
        if (badgeGlowSpreadSlider != null && badgeGlowSpreadLabel != null) {
            badgeGlowSpreadLabel.setText(String.format("%.2f", badgeGlowSpreadSlider.getValue()));
        }
    }

    private void schedulePersist() {
        if (suppress) {
            return;
        }
        refreshPreview();
        persistDelay.stop();
        persistDelay.playFromStart();
    }

    private void saveAndRefresh() {
        if (suppress) {
            return;
        }
        if (shell != null) {
            shell.persistDesktopSessionNow();
            shell.refreshEquipmentGanttGraphicForBadgeChange();
        }
    }

    PersonBadgeStyle previewStyleForGantt() {
        return previewStyle();
    }

    String snapshotPersonBadgeFontFamily() {
        return nz(badgeFontField);
    }

    double snapshotPersonBadgeFontPercent() {
        return badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : 85d;
    }

    String snapshotPersonBadgeFillHex() {
        return nz(badgeFillField);
    }

    String snapshotPersonBadgeTextHex() {
        return nz(badgeTextField);
    }

    String snapshotPersonBadgeStrokeHex() {
        return nz(badgeStrokeField);
    }

    double snapshotPersonBadgeStrokeWidth() {
        return badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : 1d;
    }

    double snapshotPersonBadgeCornerRadius() {
        return badgeCornerSlider != null ? badgeCornerSlider.getValue() : 6d;
    }

    boolean snapshotPersonBadgePill() {
        return badgePillCheck != null && badgePillCheck.isSelected();
    }

    String snapshotPersonBadgeGlowColorHex() {
        return nz(badgeGlowColorField);
    }

    double snapshotPersonBadgeGlowRadius() {
        return badgeGlowRadiusSlider != null ? badgeGlowRadiusSlider.getValue() : 14d;
    }

    double snapshotPersonBadgeGlowSpread() {
        return badgeGlowSpreadSlider != null ? badgeGlowSpreadSlider.getValue() : 0.28d;
    }

    private static String nz(TextField f) {
        return f != null && f.getText() != null ? f.getText().strip() : "";
    }

    @FXML
    private void onBadgeResetDefaultsAction() {
        suppress = true;
        try {
            PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
            if (badgeFontField != null) {
                badgeFontField.setText("");
            }
            if (badgeFontPctSlider != null) {
                badgeFontPctSlider.setValue(d.fontPercent());
            }
            if (badgeFillField != null) {
                badgeFillField.setText(d.fillHex());
            }
            if (badgeTextField != null) {
                badgeTextField.setText(d.textHex());
            }
            if (badgeStrokeField != null) {
                badgeStrokeField.setText(d.strokeHex());
            }
            if (badgeStrokeSlider != null) {
                badgeStrokeSlider.setValue(d.strokeWidth());
            }
            if (badgeCornerSlider != null) {
                badgeCornerSlider.setValue(d.cornerRadius());
            }
            if (badgePillCheck != null) {
                badgePillCheck.setSelected(d.pill());
            }
            if (badgeGlowColorField != null) {
                badgeGlowColorField.setText(d.glowColorHex());
            }
            if (badgeGlowRadiusSlider != null) {
                badgeGlowRadiusSlider.setValue(d.glowRadius());
            }
            if (badgeGlowSpreadSlider != null) {
                badgeGlowSpreadSlider.setValue(d.glowSpread());
            }
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
        saveAndRefresh();
    }

    private void refreshPreview() {
        if (badgePreviewBox == null) {
            return;
        }
        badgePreviewBox.getChildren().clear();
        PersonBadgeStyle st = previewStyle();
        double zoom = 1.0;
        double rowPx = 13;
        badgePreviewBox
                .getChildren()
                .add(PersonBadgeNodeFactory.createBadge("山田", st, zoom, rowPx));
        badgePreviewBox
                .getChildren()
                .add(PersonBadgeNodeFactory.createBadge("佐藤", st, zoom, rowPx));
    }

    private PersonBadgeStyle previewStyle() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        return new PersonBadgeStyle(
                nz(badgeFontField),
                badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : d.fontPercent(),
                nzOr(badgeFillField, d.fillHex()),
                nzOr(badgeTextField, d.textHex()),
                nzOr(badgeStrokeField, d.strokeHex()),
                badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : d.strokeWidth(),
                badgeCornerSlider != null ? badgeCornerSlider.getValue() : d.cornerRadius(),
                badgePillCheck != null && badgePillCheck.isSelected(),
                nzOr(badgeGlowColorField, d.glowColorHex()),
                badgeGlowRadiusSlider != null ? badgeGlowRadiusSlider.getValue() : d.glowRadius(),
                badgeGlowSpreadSlider != null ? badgeGlowSpreadSlider.getValue() : d.glowSpread());
    }

    private static String nzOr(TextField f, String fallback) {
        String s = nz(f);
        return s.isEmpty() ? fallback : s;
    }
}
