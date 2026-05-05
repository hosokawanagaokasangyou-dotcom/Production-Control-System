package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Accordion;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.Duration;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.config.UiBadgePurpose;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;

/**
 * 用途別 UI バッジのデザイン（実行タブのキャッシュ表示など）。
 */
public final class UiBadgeDesignTabController {

    private static final String FONT_COMBO_DEFAULT_LABEL = "(既定)";

    @FXML
    private Accordion designAccordion;

    @FXML
    private ComboBox<UiBadgePurpose> purposeCombo;

    @FXML
    private StackPane purposeStack;

    @FXML
    private VBox stage1EditorBox;

    @FXML
    private VBox foreignPurposeBox;

    @FXML
    private Label foreignPurposeHintLabel;

    @FXML
    private TextField stage1BadgeTextField;

    @FXML
    private HBox badgePreviewBox;

    @FXML
    private ComboBox<String> badgeFontCombo;

    @FXML
    private Slider badgeFontPctSlider;

    @FXML
    private Label badgeFontPctLabel;

    @FXML
    private ColorPicker badgeFillPicker;

    @FXML
    private ColorPicker badgeTextPicker;

    @FXML
    private ColorPicker badgeStrokePicker;

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
    private ColorPicker badgeGlowColorPicker;

    @FXML
    private Slider badgeGlowRadiusPctSlider;

    @FXML
    private Label badgeGlowRadiusPctLabel;

    @FXML
    private Slider badgeGlowSpreadPctSlider;

    @FXML
    private Label badgeGlowSpreadPctLabel;

    private MainShellController shell;

    private PauseTransition persistDelay;

    private boolean suppress;

    private PersonBadgeStyle stage1NetworkCacheStyle = PersonBadgeStyle.networkSourceCacheBadgeDefault();

    @FXML
    private void initialize() {
        if (designAccordion != null && !designAccordion.getPanes().isEmpty()) {
            TitledPane p = designAccordion.getPanes().get(0);
            designAccordion.setExpandedPane(p);
        }
        ObservableList<UiBadgePurpose> items = FXCollections.observableArrayList(UiBadgePurpose.values());
        purposeCombo.setItems(items);
        purposeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(UiBadgePurpose object) {
                        return object != null ? object.displayLabel() : "";
                    }

                    @Override
                    public UiBadgePurpose fromString(String string) {
                        return UiBadgePurpose.STAGE1_NETWORK_CACHE;
                    }
                });
        purposeCombo.setValue(UiBadgePurpose.STAGE1_NETWORK_CACHE);
        purposeCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            if (!suppress) {
                                refreshPurposePanels();
                            }
                        });

        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> persistNow());

        populateFontCombo();
        attachListeners();
        suppress = true;
        try {
            if (stage1BadgeTextField != null) {
                stage1BadgeTextField.setText("キャッシュ");
            }
            pushStyleToUi(stage1NetworkCacheStyle);
            syncLabelsFromSliders();
            refreshPurposePanels();
            refreshPreview();
        } finally {
            suppress = false;
        }
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
    }

    void applyUiBadgeSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppress = true;
        try {
            String lbl = s.stage1NetworkCacheBadgeLabel();
            if (stage1BadgeTextField != null) {
                stage1BadgeTextField.setText(
                        lbl != null && !lbl.isBlank() ? lbl.strip() : "キャッシュ");
            }
            PersonBadgeStyle st = s.stage1NetworkCacheBadgeStyle();
            stage1NetworkCacheStyle =
                    st != null ? st : PersonBadgeStyle.networkSourceCacheBadgeDefault();
            pushStyleToUi(stage1NetworkCacheStyle);
            syncLabelsFromSliders();
            refreshPreview();
        } finally {
            suppress = false;
        }
    }

    void flushEditsBeforeSnapshot() {
        commitStage1StyleFromUi();
    }

    String snapshotStage1NetworkCacheBadgeLabel() {
        commitStage1StyleFromUi();
        if (stage1BadgeTextField == null) {
            return "キャッシュ";
        }
        String t = stage1BadgeTextField.getText();
        return t != null && !t.isBlank() ? t.strip() : "キャッシュ";
    }

    PersonBadgeStyle snapshotStage1NetworkCacheBadgeStyle() {
        commitStage1StyleFromUi();
        return stage1NetworkCacheStyle != null
                ? stage1NetworkCacheStyle
                : PersonBadgeStyle.networkSourceCacheBadgeDefault();
    }

    private void refreshPurposePanels() {
        UiBadgePurpose p = purposeCombo != null ? purposeCombo.getValue() : null;
        boolean stage1 = p == UiBadgePurpose.STAGE1_NETWORK_CACHE;
        if (stage1EditorBox != null) {
            stage1EditorBox.setManaged(stage1);
            stage1EditorBox.setVisible(stage1);
        }
        if (foreignPurposeBox != null) {
            foreignPurposeBox.setManaged(!stage1);
            foreignPurposeBox.setVisible(!stage1);
        }
        if (foreignPurposeHintLabel != null && p != null) {
            foreignPurposeHintLabel.setText(hintForPurpose(p));
        }
        if (stage1) {
            refreshPreview();
        }
    }

    private static String hintForPurpose(UiBadgePurpose p) {
        return switch (p) {
            case EQUIPMENT_GANTT_PERSON ->
                    "設備ガントの担当者バッジの詳細（メンバー別スタイルなど）は、タブ「設備ガント・担当バッジ」で編集してください。"
                            + " 用途一覧には将来追加するバッジも並びます。";
            case RESERVED -> "将来の UI バッジ用の予約スロットです。";
            default -> "";
        };
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
        if (stage1BadgeTextField != null) {
            stage1BadgeTextField.textProperty().addListener((o, a, b) -> r.run());
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
        }
    }

    private void commitStage1StyleFromUi() {
        if (suppress) {
            return;
        }
        stage1NetworkCacheStyle = buildStyleFromUiFields();
    }

    private PersonBadgeStyle buildStyleFromUiFields() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        String fontFam = "";
        if (badgeFontCombo != null && badgeFontCombo.getValue() != null) {
            String v = badgeFontCombo.getValue().strip();
            if (!v.isEmpty() && !FONT_COMBO_DEFAULT_LABEL.equals(v)) {
                fontFam = v;
            }
        }
        double baseR = d.glowRadius();
        double baseS = d.glowSpread();
        double rPct = badgeGlowRadiusPctSlider != null ? badgeGlowRadiusPctSlider.getValue() : 100.0;
        double sPct = badgeGlowSpreadPctSlider != null ? badgeGlowSpreadPctSlider.getValue() : 100.0;
        double glowR = baseR * (rPct / 100.0);
        double glowS = Math.min(1.0, Math.max(0.0, baseS * (sPct / 100.0)));
        return new PersonBadgeStyle(
                fontFam,
                badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : d.fontPercent(),
                colorToHex(badgeFillPicker, PersonBadgeStyle.networkSourceCacheBadgeDefault().fillHex()),
                colorToHex(badgeTextPicker, PersonBadgeStyle.networkSourceCacheBadgeDefault().textHex()),
                colorToHex(badgeStrokePicker, PersonBadgeStyle.networkSourceCacheBadgeDefault().strokeHex()),
                badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : d.strokeWidth(),
                badgeCornerSlider != null ? badgeCornerSlider.getValue() : d.cornerRadius(),
                badgePillCheck != null && badgePillCheck.isSelected(),
                colorToHex(
                        badgeGlowColorPicker, PersonBadgeStyle.networkSourceCacheBadgeDefault().glowColorHex()),
                glowR,
                glowS);
    }

    private void pushStyleToUi(PersonBadgeStyle st) {
        PersonBadgeStyle x = st != null ? st : PersonBadgeStyle.networkSourceCacheBadgeDefault();
        PersonBadgeStyle d = PersonBadgeStyle.networkSourceCacheBadgeDefault();
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
            commitStage1StyleFromUi();
        }
        badgePreviewBox.getChildren().clear();
        String txt =
                stage1BadgeTextField != null && stage1BadgeTextField.getText() != null
                        ? stage1BadgeTextField.getText().strip()
                        : "キャッシュ";
        if (txt.isEmpty()) {
            txt = "キャッシュ";
        }
        PersonBadgeStyle st =
                stage1NetworkCacheStyle != null
                        ? stage1NetworkCacheStyle
                        : PersonBadgeStyle.networkSourceCacheBadgeDefault();
        badgePreviewBox
                .getChildren()
                .add(PersonBadgeNodeFactory.createBadge(txt, st, 1.0, 13.0));
    }
}
