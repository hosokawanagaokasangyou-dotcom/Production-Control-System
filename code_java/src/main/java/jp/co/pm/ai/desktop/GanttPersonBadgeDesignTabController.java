package jp.co.pm.ai.desktop;

import java.util.Collection;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import javafx.animation.PauseTransition;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

/** 設備ガント・担当バッジのデザイン編集タブ（担当者キーごとにスタイルを保持）。 */
public final class GanttPersonBadgeDesignTabController {

    /** セッションの「既定」―マップ未登録のバッジ表示に適用。 */
    public static final String GLOBAL_EDIT_LABEL = "（既定・未登録の担当者）";

    @FXML
    private Accordion designAccordion;

    @FXML
    private ComboBox<String> personBadgeEditTargetCombo;

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

    /** フラット項目に保存される「既定」スタイル。 */
    private PersonBadgeStyle globalStyle = PersonBadgeStyle.defaultStyle();

    private final LinkedHashMap<String, PersonBadgeStyle> perLabelStyles = new LinkedHashMap<>();
    private final LinkedHashSet<String> observedBadgeLabels = new LinkedHashSet<>();

    @FXML
    private void initialize() {
        if (designAccordion != null && designAccordion.getPanes().size() > 0) {
            TitledPane p = designAccordion.getPanes().get(0);
            designAccordion.setExpandedPane(p);
        }
        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> saveAndRefresh());

        if (personBadgeEditTargetCombo != null) {
            personBadgeEditTargetCombo.setEditable(true);
            personBadgeEditTargetCombo
                    .valueProperty()
                    .addListener(
                            (o, oldVal, newVal) -> {
                                if (!suppress) {
                                    onEditTargetChanged(oldVal, newVal);
                                }
                            });
        }

        attachListeners();
        rebuildComboItems();
        suppress = true;
        try {
            if (personBadgeEditTargetCombo != null) {
                personBadgeEditTargetCombo.getSelectionModel().select(GLOBAL_EDIT_LABEL);
            }
            pushStyleToUi(globalStyle);
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
    }

    /**
     * 設備ガントで検出したバッジ表示キーをコンボに追加する（選択は維持）。
     */
    void mergeObservedBadgeLabels(Collection<String> labels) {
        if (labels == null || labels.isEmpty()) {
            return;
        }
        boolean changed = false;
        for (String raw : labels) {
            String k = PersonBadgeStyle.normalizeLabelKey(raw);
            if (!k.isEmpty() && observedBadgeLabels.add(k)) {
                changed = true;
            }
        }
        if (!changed) {
            return;
        }
        suppress = true;
        try {
            rebuildComboItems();
        } finally {
            suppress = false;
        }
    }

    void applyPersonBadgeDesignSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppress = true;
        try {
            globalStyle = s.resolvedPersonBadgeStyle();
            perLabelStyles.clear();
            perLabelStyles.putAll(s.equipmentGanttPersonBadgeStylesByLabel());
            rebuildComboItems();
            if (personBadgeEditTargetCombo != null) {
                personBadgeEditTargetCombo.getSelectionModel().select(GLOBAL_EDIT_LABEL);
            }
            pushStyleToUi(globalStyle);
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
    }

    /** セッション保存直前に、現在の UI を global / per-label に反映する。 */
    void flushBadgeEditsBeforeSnapshot() {
        commitUiToTarget();
    }

    /** ガント描画用：バッジ文字列ごとの実効スタイル。 */
    public PersonBadgeStyle resolveStyleForBadgeLabel(String badgeLabel) {
        commitUiToTarget();
        return resolveStyleForBadgeLabelRaw(badgeLabel);
    }

    private PersonBadgeStyle resolveStyleForBadgeLabelRaw(String badgeLabel) {
        String k = PersonBadgeStyle.normalizeLabelKey(badgeLabel);
        if (k.isEmpty()) {
            return globalStyle;
        }
        PersonBadgeStyle o = perLabelStyles.get(k);
        return o != null ? o : globalStyle;
    }

    Map<String, PersonBadgeStyle> snapshotPersonBadgeStylesByLabel() {
        commitUiToTarget();
        return Map.copyOf(perLabelStyles);
    }

    PersonBadgeStyle previewStyleForGantt() {
        return globalStyle;
    }

    String snapshotPersonBadgeFontFamily() {
        return globalStyle.fontFamily() != null ? globalStyle.fontFamily().strip() : "";
    }

    double snapshotPersonBadgeFontPercent() {
        return globalStyle.fontPercent();
    }

    String snapshotPersonBadgeFillHex() {
        return globalStyle.fillHex() != null ? globalStyle.fillHex() : "";
    }

    String snapshotPersonBadgeTextHex() {
        return globalStyle.textHex() != null ? globalStyle.textHex() : "";
    }

    String snapshotPersonBadgeStrokeHex() {
        return globalStyle.strokeHex() != null ? globalStyle.strokeHex() : "";
    }

    double snapshotPersonBadgeStrokeWidth() {
        return globalStyle.strokeWidth();
    }

    double snapshotPersonBadgeCornerRadius() {
        return globalStyle.cornerRadius();
    }

    boolean snapshotPersonBadgePill() {
        return globalStyle.pill();
    }

    String snapshotPersonBadgeGlowColorHex() {
        return globalStyle.glowColorHex() != null ? globalStyle.glowColorHex() : "";
    }

    double snapshotPersonBadgeGlowRadius() {
        return globalStyle.glowRadius();
    }

    double snapshotPersonBadgeGlowSpread() {
        return globalStyle.glowSpread();
    }

    private void onEditTargetChanged(String oldVal, String newVal) {
        suppress = true;
        try {
            if (oldVal != null) {
                PersonBadgeStyle built = buildStyleFromUiFields();
                if (GLOBAL_EDIT_LABEL.equals(oldVal)) {
                    globalStyle = built;
                } else {
                    String k = PersonBadgeStyle.normalizeLabelKey(oldVal);
                    if (!k.isEmpty()) {
                        perLabelStyles.put(k, built);
                    }
                }
            }
            String nv = newVal != null ? newVal : GLOBAL_EDIT_LABEL;
            if (GLOBAL_EDIT_LABEL.equals(nv)) {
                pushStyleToUi(globalStyle);
            } else {
                String k = PersonBadgeStyle.normalizeLabelKey(nv);
                PersonBadgeStyle st = perLabelStyles.getOrDefault(k, globalStyle);
                pushStyleToUi(st);
            }
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
    }

    private void rebuildComboItems() {
        if (personBadgeEditTargetCombo == null) {
            return;
        }
        String prev = personBadgeEditTargetCombo.getValue();
        var items = FXCollections.<String>observableArrayList();
        items.add(GLOBAL_EDIT_LABEL);
        TreeSet<String> sorted = new TreeSet<>(Comparator.naturalOrder());
        sorted.addAll(perLabelStyles.keySet());
        sorted.addAll(observedBadgeLabels);
        for (String s : sorted) {
            if (s != null && !s.isBlank()) {
                items.add(s);
            }
        }
        personBadgeEditTargetCombo.setItems(items);
        if (prev != null && items.contains(prev)) {
            personBadgeEditTargetCombo.getSelectionModel().select(prev);
        } else if (GLOBAL_EDIT_LABEL.equals(prev) || prev == null) {
            personBadgeEditTargetCombo.getSelectionModel().select(GLOBAL_EDIT_LABEL);
        }
    }

    private void commitUiToTarget() {
        if (suppress || personBadgeEditTargetCombo == null) {
            return;
        }
        PersonBadgeStyle built = buildStyleFromUiFields();
        String raw = personBadgeEditTargetCombo.getValue();
        if (raw == null || GLOBAL_EDIT_LABEL.equals(raw)) {
            globalStyle = built;
        } else {
            String k = PersonBadgeStyle.normalizeLabelKey(raw);
            if (!k.isEmpty()) {
                perLabelStyles.put(k, built);
            }
        }
    }

    private void pushStyleToUi(PersonBadgeStyle st) {
        PersonBadgeStyle x = st != null ? st : PersonBadgeStyle.defaultStyle();
        if (badgeFontField != null) {
            badgeFontField.setText(x.fontFamily() != null ? x.fontFamily() : "");
        }
        if (badgeFontPctSlider != null) {
            badgeFontPctSlider.setValue(Math.clamp(x.fontPercent(), 40, 160));
        }
        if (badgeFillField != null) {
            badgeFillField.setText(x.fillHex());
        }
        if (badgeTextField != null) {
            badgeTextField.setText(x.textHex());
        }
        if (badgeStrokeField != null) {
            badgeStrokeField.setText(x.strokeHex());
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
        if (badgeGlowColorField != null) {
            badgeGlowColorField.setText(x.glowColorHex());
        }
        if (badgeGlowRadiusSlider != null) {
            badgeGlowRadiusSlider.setValue(Math.clamp(x.glowRadius(), 0, 40));
        }
        if (badgeGlowSpreadSlider != null) {
            badgeGlowSpreadSlider.setValue(Math.clamp(x.glowSpread(), 0, 1));
        }
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

    private static String nz(TextField f) {
        return f != null && f.getText() != null ? f.getText().strip() : "";
    }

    private static String nzOr(TextField f, String fallback) {
        String s = nz(f);
        return s.isEmpty() ? fallback : s;
    }

    private PersonBadgeStyle buildStyleFromUiFields() {
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

    @FXML
    private void onBadgeResetDefaultsAction() {
        suppress = true;
        try {
            String cur =
                    personBadgeEditTargetCombo != null ? personBadgeEditTargetCombo.getValue() : GLOBAL_EDIT_LABEL;
            PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
            if (cur == null || GLOBAL_EDIT_LABEL.equals(cur)) {
                globalStyle = d;
                pushStyleToUi(globalStyle);
            } else {
                String k = PersonBadgeStyle.normalizeLabelKey(cur);
                perLabelStyles.remove(k);
                pushStyleToUi(globalStyle);
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
        if (!suppress) {
            commitUiToTarget();
        }
        badgePreviewBox.getChildren().clear();
        double zoom = 1.0;
        double rowPx = 13;
        for (String sample : List.of("山田", "佐藤")) {
            PersonBadgeStyle st = resolveStyleForBadgeLabelRaw(sample);
            badgePreviewBox
                    .getChildren()
                    .add(PersonBadgeNodeFactory.createBadge(sample, st, zoom, rowPx));
        }
    }
}
