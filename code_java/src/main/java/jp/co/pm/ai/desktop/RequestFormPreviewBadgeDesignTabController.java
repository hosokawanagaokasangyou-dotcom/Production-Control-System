package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;

import javafx.animation.PauseTransition;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPreviewBadgeConfig;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;

/** 依頼書プレビュー上部・原本更新バッジのデザイン編集タブ。 */
public final class RequestFormPreviewBadgeDesignTabController {

    private static final String FONT_COMBO_DEFAULT_LABEL = "(既定)";

    @FXML private Accordion designAccordion;
    @FXML private TextField badgeTextField;
    @FXML private HBox badgePreviewBox;
    @FXML private ComboBox<String> badgeFontCombo;
    @FXML private Slider badgeFontPctSlider;
    @FXML private Label badgeFontPctLabel;
    @FXML private ColorPicker badgeFillPicker;
    @FXML private ColorPicker badgeTextPicker;
    @FXML private ColorPicker badgeStrokePicker;
    @FXML private Slider badgeStrokeSlider;
    @FXML private Label badgeStrokeLabel;
    @FXML private Slider badgeCornerSlider;
    @FXML private Label badgeCornerLabel;
    @FXML private CheckBox badgePillCheck;
    @FXML private ColorPicker badgeGlowColorPicker;
    @FXML private Slider badgeGlowRadiusPctSlider;
    @FXML private Label badgeGlowRadiusPctLabel;
    @FXML private Slider badgeGlowSpreadPctSlider;
    @FXML private Label badgeGlowSpreadPctLabel;

    private MainShellController shell;
    private PauseTransition persistDelay;
    private boolean suppress;
    private PersonBadgeStyle badgeStyle = PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();

    @FXML
    private void initialize() {
        if (designAccordion != null && !designAccordion.getPanes().isEmpty()) {
            designAccordion.setExpandedPane(designAccordion.getPanes().get(0));
        }
        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> persistNow());
        populateFontCombo();
        attachListeners();
        suppress = true;
        try {
            if (badgeTextField != null) {
                badgeTextField.setText("更新");
            }
            pushStyleToUi(badgeStyle);
            syncLabelsFromSliders();
            refreshPreview();
        } finally {
            suppress = false;
        }
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
    }

    void applyRequestFormPreviewBadgeSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppress = true;
        try {
            String lbl = s.requestFormPreviewUpdateBadgeLabel();
            if (badgeTextField != null) {
                badgeTextField.setText(lbl != null && !lbl.isBlank() ? lbl.strip() : "更新");
            }
            PersonBadgeStyle st = s.requestFormPreviewUpdateBadgeStyle();
            badgeStyle = st != null ? st : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
            pushStyleToUi(badgeStyle);
            syncLabelsFromSliders();
            refreshPreview();
        } finally {
            suppress = false;
        }
    }

    void flushEditsBeforeSnapshot() {
        commitStyleFromUi();
    }

    RequestFormPreviewBadgeConfig snapshotPreviewBadgeConfig() {
        commitStyleFromUi();
        String label =
                badgeTextField != null && badgeTextField.getText() != null
                        ? badgeTextField.getText().strip()
                        : "更新";
        if (label.isEmpty()) {
            label = "更新";
        }
        PersonBadgeStyle st =
                badgeStyle != null
                        ? badgeStyle
                        : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        return new RequestFormPreviewBadgeConfig(label, st);
    }

    String snapshotPreviewBadgeLabel() {
        return snapshotPreviewBadgeConfig().label();
    }

    PersonBadgeStyle snapshotPreviewBadgeStyle() {
        return snapshotPreviewBadgeConfig().style();
    }

    private void populateFontCombo() {
        if (badgeFontCombo == null) {
            return;
        }
        ObservableList<String> items = FXCollections.observableArrayList(FONT_COMBO_DEFAULT_LABEL);
        List<String> sorted = new ArrayList<>(Font.getFamilies());
        sorted.sort(String.CASE_INSENSITIVE_ORDER);
        items.addAll(sorted);
        badgeFontCombo.setItems(items);
    }

    private void attachListeners() {
        Runnable r = this::schedulePersist;
        if (badgeFontCombo != null) {
            badgeFontCombo.valueProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeTextField != null) {
            badgeTextField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeFontPctSlider, badgeFontPctLabel, "%.0f%%", r);
        addColorPickerListener(badgeFillPicker, r);
        addColorPickerListener(badgeTextPicker, r);
        addColorPickerListener(badgeStrokePicker, r);
        wireSlider(badgeStrokeSlider, badgeStrokeLabel, "%.1f", r);
        wireSlider(badgeCornerSlider, badgeCornerLabel, "%.0f", r);
        if (badgePillCheck != null) {
            badgePillCheck.selectedProperty().addListener((o, a, b) -> r.run());
        }
        addColorPickerListener(badgeGlowColorPicker, r);
        wireSlider(badgeGlowRadiusPctSlider, badgeGlowRadiusPctLabel, "%.0f%%", r);
        wireSlider(badgeGlowSpreadPctSlider, badgeGlowSpreadPctLabel, "%.0f%%", r);
    }

    private static void addColorPickerListener(ColorPicker cp, Runnable r) {
        if (cp != null) {
            cp.valueProperty().addListener((o, a, b) -> r.run());
        }
    }

    private static void wireSlider(Slider sl, Label lb, String fmt, Runnable onChange) {
        if (sl == null) {
            return;
        }
        SliderCommittedChangeSupport.install(
                sl,
                () -> {
                    if (lb != null) {
                        lb.setText(String.format(fmt, sl.getValue()));
                    }
                },
                onChange);
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
        if (badgeGlowRadiusPctSlider != null && badgeGlowRadiusPctLabel != null) {
            badgeGlowRadiusPctLabel.setText(String.format("%.0f%%", badgeGlowRadiusPctSlider.getValue()));
        }
        if (badgeGlowSpreadPctSlider != null && badgeGlowSpreadPctLabel != null) {
            badgeGlowSpreadPctLabel.setText(String.format("%.0f%%", badgeGlowSpreadPctSlider.getValue()));
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

    private void persistNow() {
        if (suppress) {
            return;
        }
        if (shell != null) {
            shell.persistDesktopSessionNow();
            shell.refreshRequestFormPreviewBadgeAppearance();
        }
    }

    private void commitStyleFromUi() {
        if (suppress) {
            return;
        }
        badgeStyle = buildStyleFromUiFields();
    }

    private PersonBadgeStyle buildStyleFromUiFields() {
        PersonBadgeStyle d = PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        String fontFam = "";
        if (badgeFontCombo != null && badgeFontCombo.getValue() != null) {
            String v = badgeFontCombo.getValue().strip();
            if (!v.isEmpty() && !FONT_COMBO_DEFAULT_LABEL.equals(v)) {
                fontFam = v;
            }
        }
        double baseR = PersonBadgeStyle.defaultStyle().glowRadius();
        double baseS = PersonBadgeStyle.defaultStyle().glowSpread();
        double rPct = badgeGlowRadiusPctSlider != null ? badgeGlowRadiusPctSlider.getValue() : 100.0;
        double sPct = badgeGlowSpreadPctSlider != null ? badgeGlowSpreadPctSlider.getValue() : 100.0;
        double glowR = baseR * (rPct / 100.0);
        double glowS = Math.min(1.0, Math.max(0.0, baseS * (sPct / 100.0)));
        return new PersonBadgeStyle(
                fontFam,
                badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : d.fontPercent(),
                colorToHex(badgeFillPicker, d.fillHex()),
                colorToHex(badgeTextPicker, d.textHex()),
                colorToHex(badgeStrokePicker, d.strokeHex()),
                badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : d.strokeWidth(),
                badgeCornerSlider != null ? badgeCornerSlider.getValue() : d.cornerRadius(),
                badgePillCheck != null && badgePillCheck.isSelected(),
                colorToHex(badgeGlowColorPicker, d.glowColorHex()),
                glowR,
                glowS,
                d.opacity());
    }

    private void pushStyleToUi(PersonBadgeStyle st) {
        PersonBadgeStyle x = st != null ? st : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        PersonBadgeStyle d = PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        applyFontFamilyToCombo(x.fontFamily());
        if (badgeFontPctSlider != null) {
            badgeFontPctSlider.setValue(Math.clamp(x.fontPercent(), 40, 160));
        }
        if (badgeFillPicker != null) {
            badgeFillPicker.setValue(parseHexToColor(x.fillHex(), Color.web(d.fillHex())));
        }
        if (badgeTextPicker != null) {
            badgeTextPicker.setValue(parseHexToColor(x.textHex(), Color.web(d.textHex())));
        }
        if (badgeStrokePicker != null) {
            badgeStrokePicker.setValue(parseHexToColor(x.strokeHex(), Color.web(d.strokeHex())));
        }
        if (badgeStrokeSlider != null) {
            badgeStrokeSlider.setValue(Math.clamp(x.strokeWidth(), 0, 6));
        }
        if (badgeCornerSlider != null) {
            badgeCornerSlider.setValue(Math.clamp(x.cornerRadius(), 0, 24));
        }
        if (badgePillCheck != null) {
            badgePillCheck.setSelected(x.pill());
        }
        if (badgeGlowColorPicker != null) {
            badgeGlowColorPicker.setValue(parseHexToColor(x.glowColorHex(), Color.web(d.glowColorHex())));
        }
        double baseR = PersonBadgeStyle.defaultStyle().glowRadius();
        double baseS = PersonBadgeStyle.defaultStyle().glowSpread();
        if (badgeGlowRadiusPctSlider != null) {
            double pct = baseR > 1e-9 ? (x.glowRadius() / baseR) * 100.0 : 0.0;
            badgeGlowRadiusPctSlider.setValue(Math.clamp(pct, 0, 400));
        }
        if (badgeGlowSpreadPctSlider != null) {
            double pct =
                    baseS > 1e-12
                            ? (x.glowSpread() / baseS) * 100.0
                            : (x.glowSpread() <= 1e-12 ? 0.0 : 100.0);
            badgeGlowSpreadPctSlider.setValue(Math.clamp(pct, 0, 400));
        }
    }

    private void applyFontFamilyToCombo(String fontFamily) {
        if (badgeFontCombo == null) {
            return;
        }
        if (fontFamily == null || fontFamily.isBlank()) {
            badgeFontCombo.getSelectionModel().select(0);
            return;
        }
        ObservableList<String> items = badgeFontCombo.getItems();
        for (int i = 0; i < items.size(); i++) {
            if (fontFamily.equals(items.get(i))) {
                badgeFontCombo.getSelectionModel().select(i);
                return;
            }
        }
        items.add(fontFamily);
        badgeFontCombo.getSelectionModel().select(fontFamily);
    }

    private static String colorToHex(ColorPicker cp, String fallbackHex) {
        if (cp == null || cp.getValue() == null) {
            return fallbackHex;
        }
        return colorToHex(cp.getValue());
    }

    private static String colorToHex(Color c) {
        int r = (int) Math.round(c.getRed() * 255.0);
        int g = (int) Math.round(c.getGreen() * 255.0);
        int b = (int) Math.round(c.getBlue() * 255.0);
        return String.format("#%02x%02x%02x", r, g, b);
    }

    private static Color parseHexToColor(String hex, Color fallback) {
        try {
            String h = hex != null ? hex.strip() : "";
            return h.isEmpty() ? fallback : Color.web(h);
        } catch (IllegalArgumentException e) {
            return fallback;
        }
    }

    private void refreshPreview() {
        if (badgePreviewBox == null) {
            return;
        }
        if (!suppress) {
            commitStyleFromUi();
        }
        badgePreviewBox.getChildren().clear();
        String txt =
                badgeTextField != null && badgeTextField.getText() != null
                        ? badgeTextField.getText().strip()
                        : "更新";
        if (txt.isEmpty()) {
            txt = "更新";
        }
        PersonBadgeStyle st =
                badgeStyle != null ? badgeStyle : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        badgePreviewBox.getChildren().add(PersonBadgeNodeFactory.createBadge(txt, st, 1.0, 13.0));
    }
}
