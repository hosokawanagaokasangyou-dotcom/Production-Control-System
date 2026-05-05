package jp.co.pm.ai.desktop;

import javafx.animation.PauseTransition;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.paint.Color;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PushButtonDesignPrefs;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;

/**
 * アプリ内プッシュボタン（実行タブの通常ボタン・段階実行ボタン）の見た目オーバーライド。
 */
public final class PushButtonDesignTabController {

    @FXML
    private Accordion designAccordion;

    @FXML
    private BorderPane previewRunPane;

    @FXML
    private CheckBox customGeneralCheck;

    @FXML
    private Slider generalRadiusSlider;

    @FXML
    private Label generalRadiusLabel;

    @FXML
    private Slider generalPadVSlider;

    @FXML
    private Label generalPadVLabel;

    @FXML
    private Slider generalPadHSlider;

    @FXML
    private Label generalPadHLabel;

    @FXML
    private Slider generalFontSlider;

    @FXML
    private Label generalFontLabel;

    @FXML
    private ColorPicker generalBgPicker;

    @FXML
    private ColorPicker generalBorderPicker;

    @FXML
    private ColorPicker generalTextPicker;

    @FXML
    private ColorPicker generalHoverPicker;

    @FXML
    private ColorPicker generalPressedPicker;

    @FXML
    private CheckBox customStageCheck;

    @FXML
    private Slider stageFontSlider;

    @FXML
    private Label stageFontLabel;

    @FXML
    private Slider stageMinWidthSlider;

    @FXML
    private Label stageMinWidthLabel;

    @FXML
    private Slider stageMinHeightSlider;

    @FXML
    private Label stageMinHeightLabel;

    @FXML
    private Slider stagePadVSlider;

    @FXML
    private Label stagePadVLabel;

    @FXML
    private Slider stagePadHSlider;

    @FXML
    private Label stagePadHLabel;

    @FXML
    private Slider stageRadiusSlider;

    @FXML
    private Label stageRadiusLabel;

    @FXML
    private ColorPicker stage1BgPicker;

    @FXML
    private ColorPicker stage1BorderPicker;

    @FXML
    private ColorPicker stage1HoverPicker;

    @FXML
    private ColorPicker stage1PressedPicker;

    @FXML
    private ColorPicker stage2BgPicker;

    @FXML
    private ColorPicker stage2BorderPicker;

    @FXML
    private ColorPicker stage2HoverPicker;

    @FXML
    private ColorPicker stage2PressedPicker;

    @FXML
    private ColorPicker stage3BgPicker;

    @FXML
    private ColorPicker stage3BorderPicker;

    @FXML
    private ColorPicker stage3HoverPicker;

    @FXML
    private ColorPicker stage3PressedPicker;

    private MainShellController shell;

    private PauseTransition persistDelay;

    private boolean suppress;

    @FXML
    private void initialize() {
        if (designAccordion != null && !designAccordion.getPanes().isEmpty()) {
            TitledPane p = designAccordion.getPanes().get(0);
            designAccordion.setExpandedPane(p);
        }
        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> persistNow());

        PushButtonDesignPrefs d = PushButtonDesignPrefs.builtInSnapshot();
        suppress = true;
        try {
            pushPrefsToUi(d);
            if (customGeneralCheck != null) {
                customGeneralCheck.setSelected(false);
            }
            if (customStageCheck != null) {
                customStageCheck.setSelected(false);
            }
            syncLabelsFromSliders();
            refreshPreviewDisabledMask();
        } finally {
            suppress = false;
        }

        attachListeners();
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
    }

    void applyPushButtonSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        PushButtonDesignPrefs p = s.pushButtonDesignPrefs();
        suppress = true;
        try {
            pushPrefsToUi(p != null ? p : PushButtonDesignPrefs.inactiveDefaults());
            syncLabelsFromSliders();
            refreshPreviewDisabledMask();
        } finally {
            suppress = false;
        }
    }

    void flushEditsBeforeSnapshot() {
        // UI が唯一のソース（隠しバッファなし）
    }

    PushButtonDesignPrefs snapshotPrefs() {
        PushButtonDesignPrefs d = PushButtonDesignPrefs.builtInSnapshot();
        boolean g = customGeneralCheck != null && customGeneralCheck.isSelected();
        boolean st = customStageCheck != null && customStageCheck.isSelected();
        return new PushButtonDesignPrefs(
                g,
                generalRadiusSlider != null ? generalRadiusSlider.getValue() : d.generalBorderRadius(),
                generalPadVSlider != null ? generalPadVSlider.getValue() : d.generalPaddingV(),
                generalPadHSlider != null ? generalPadHSlider.getValue() : d.generalPaddingH(),
                generalFontSlider != null ? generalFontSlider.getValue() : d.generalFontPx(),
                colorToHex(generalBgPicker, d.generalBgHex()),
                colorToHex(generalBorderPicker, d.generalBorderHex()),
                colorToHex(generalTextPicker, d.generalTextHex()),
                colorToHex(generalHoverPicker, d.generalHoverBgHex()),
                colorToHex(generalPressedPicker, d.generalPressedBgHex()),
                st,
                stageFontSlider != null ? stageFontSlider.getValue() : d.stageFontPx(),
                stageMinWidthSlider != null ? stageMinWidthSlider.getValue() : d.stageMinWidth(),
                stageMinHeightSlider != null ? stageMinHeightSlider.getValue() : d.stageMinHeight(),
                stagePadVSlider != null ? stagePadVSlider.getValue() : d.stagePaddingV(),
                stagePadHSlider != null ? stagePadHSlider.getValue() : d.stagePaddingH(),
                stageRadiusSlider != null ? stageRadiusSlider.getValue() : d.stageBorderRadius(),
                colorToHex(stage1BgPicker, d.stage1BgHex()),
                colorToHex(stage1BorderPicker, d.stage1BorderHex()),
                colorToHex(stage1HoverPicker, d.stage1HoverBgHex()),
                colorToHex(stage1PressedPicker, d.stage1PressedBgHex()),
                colorToHex(stage2BgPicker, d.stage2BgHex()),
                colorToHex(stage2BorderPicker, d.stage2BorderHex()),
                colorToHex(stage2HoverPicker, d.stage2HoverBgHex()),
                colorToHex(stage2PressedPicker, d.stage2PressedBgHex()),
                colorToHex(stage3BgPicker, d.stage3BgHex()),
                colorToHex(stage3BorderPicker, d.stage3BorderHex()),
                colorToHex(stage3HoverPicker, d.stage3HoverBgHex()),
                colorToHex(stage3PressedPicker, d.stage3PressedBgHex()));
    }

    /** メイン {@link javafx.scene.Scene} 準備後に呼び、保存済みプリファレンスを反映する。 */
    void installStylesheetWhenSceneReady() {
        refreshStylesheetOnShell();
    }

    @FXML
    private void onResetToBuiltInAction() {
        suppress = true;
        try {
            pushPrefsToUi(PushButtonDesignPrefs.builtInSnapshot());
            if (customGeneralCheck != null) {
                customGeneralCheck.setSelected(true);
            }
            if (customStageCheck != null) {
                customStageCheck.setSelected(true);
            }
            syncLabelsFromSliders();
            refreshPreviewDisabledMask();
        } finally {
            suppress = false;
        }
        refreshStylesheetOnShell();
        schedulePersist();
    }

    private void pushPrefsToUi(PushButtonDesignPrefs p) {
        PushButtonDesignPrefs x = p != null ? p : PushButtonDesignPrefs.inactiveDefaults();
        if (customGeneralCheck != null) {
            customGeneralCheck.setSelected(x.customizeGeneralRunTab());
        }
        if (customStageCheck != null) {
            customStageCheck.setSelected(x.customizeStageRunButtons());
        }
        if (generalRadiusSlider != null) {
            generalRadiusSlider.setValue(Math.clamp(x.generalBorderRadius(), 0, 24));
        }
        if (generalPadVSlider != null) {
            generalPadVSlider.setValue(Math.clamp(x.generalPaddingV(), 0, 32));
        }
        if (generalPadHSlider != null) {
            generalPadHSlider.setValue(Math.clamp(x.generalPaddingH(), 0, 40));
        }
        if (generalFontSlider != null) {
            generalFontSlider.setValue(Math.clamp(x.generalFontPx(), 9, 24));
        }
        setPicker(generalBgPicker, x.generalBgHex());
        setPicker(generalBorderPicker, x.generalBorderHex());
        setPicker(generalTextPicker, x.generalTextHex());
        setPicker(generalHoverPicker, x.generalHoverBgHex());
        setPicker(generalPressedPicker, x.generalPressedBgHex());

        if (stageFontSlider != null) {
            stageFontSlider.setValue(Math.clamp(x.stageFontPx(), 10, 28));
        }
        if (stageMinWidthSlider != null) {
            stageMinWidthSlider.setValue(Math.clamp(x.stageMinWidth(), 120, 400));
        }
        if (stageMinHeightSlider != null) {
            stageMinHeightSlider.setValue(Math.clamp(x.stageMinHeight(), 36, 72));
        }
        if (stagePadVSlider != null) {
            stagePadVSlider.setValue(Math.clamp(x.stagePaddingV(), 4, 28));
        }
        if (stagePadHSlider != null) {
            stagePadHSlider.setValue(Math.clamp(x.stagePaddingH(), 8, 48));
        }
        if (stageRadiusSlider != null) {
            stageRadiusSlider.setValue(Math.clamp(x.stageBorderRadius(), 0, 20));
        }
        setPicker(stage1BgPicker, x.stage1BgHex());
        setPicker(stage1BorderPicker, x.stage1BorderHex());
        setPicker(stage1HoverPicker, x.stage1HoverBgHex());
        setPicker(stage1PressedPicker, x.stage1PressedBgHex());
        setPicker(stage2BgPicker, x.stage2BgHex());
        setPicker(stage2BorderPicker, x.stage2BorderHex());
        setPicker(stage2HoverPicker, x.stage2HoverBgHex());
        setPicker(stage2PressedPicker, x.stage2PressedBgHex());
        setPicker(stage3BgPicker, x.stage3BgHex());
        setPicker(stage3BorderPicker, x.stage3BorderHex());
        setPicker(stage3HoverPicker, x.stage3HoverBgHex());
        setPicker(stage3PressedPicker, x.stage3PressedBgHex());
    }

    private static void setPicker(ColorPicker cp, String hex) {
        if (cp == null) {
            return;
        }
        try {
            String h = hex != null ? hex.strip() : "";
            cp.setValue(h.isEmpty() ? Color.GRAY : Color.web(h));
        } catch (IllegalArgumentException e) {
            cp.setValue(Color.GRAY);
        }
    }

    private void attachListeners() {
        Runnable refreshAndPersist =
                () -> {
                    refreshPreviewDisabledMask();
                    refreshStylesheetOnShell();
                    schedulePersist();
                };
        if (customGeneralCheck != null) {
            customGeneralCheck.selectedProperty().addListener((o, a, b) -> refreshAndPersist.run());
        }
        if (customStageCheck != null) {
            customStageCheck.selectedProperty().addListener((o, a, b) -> refreshAndPersist.run());
        }

        wireSlider(generalRadiusSlider, generalRadiusLabel, "%.0f", refreshAndPersist);
        wireSlider(generalPadVSlider, generalPadVLabel, "%.0f", refreshAndPersist);
        wireSlider(generalPadHSlider, generalPadHLabel, "%.0f", refreshAndPersist);
        wireSlider(generalFontSlider, generalFontLabel, "%.0f", refreshAndPersist);

        wireSlider(stageFontSlider, stageFontLabel, "%.0f", refreshAndPersist);
        wireSlider(stageMinWidthSlider, stageMinWidthLabel, "%.0f", refreshAndPersist);
        wireSlider(stageMinHeightSlider, stageMinHeightLabel, "%.0f", refreshAndPersist);
        wireSlider(stagePadVSlider, stagePadVLabel, "%.0f", refreshAndPersist);
        wireSlider(stagePadHSlider, stagePadHLabel, "%.0f", refreshAndPersist);
        wireSlider(stageRadiusSlider, stageRadiusLabel, "%.0f", refreshAndPersist);

        addPicker(generalBgPicker, refreshAndPersist);
        addPicker(generalBorderPicker, refreshAndPersist);
        addPicker(generalTextPicker, refreshAndPersist);
        addPicker(generalHoverPicker, refreshAndPersist);
        addPicker(generalPressedPicker, refreshAndPersist);
        addPicker(stage1BgPicker, refreshAndPersist);
        addPicker(stage1BorderPicker, refreshAndPersist);
        addPicker(stage1HoverPicker, refreshAndPersist);
        addPicker(stage1PressedPicker, refreshAndPersist);
        addPicker(stage2BgPicker, refreshAndPersist);
        addPicker(stage2BorderPicker, refreshAndPersist);
        addPicker(stage2HoverPicker, refreshAndPersist);
        addPicker(stage2PressedPicker, refreshAndPersist);
        addPicker(stage3BgPicker, refreshAndPersist);
        addPicker(stage3BorderPicker, refreshAndPersist);
        addPicker(stage3HoverPicker, refreshAndPersist);
        addPicker(stage3PressedPicker, refreshAndPersist);
    }

    private static void addPicker(ColorPicker cp, Runnable r) {
        if (cp != null) {
            cp.valueProperty().addListener((o, a, b) -> r.run());
        }
    }

    private void wireSlider(Slider sl, Label lb, String fmt, Runnable onChange) {
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
        if (generalRadiusSlider != null && generalRadiusLabel != null) {
            generalRadiusLabel.setText(String.format("%.0f", generalRadiusSlider.getValue()));
        }
        if (generalPadVSlider != null && generalPadVLabel != null) {
            generalPadVLabel.setText(String.format("%.0f", generalPadVSlider.getValue()));
        }
        if (generalPadHSlider != null && generalPadHLabel != null) {
            generalPadHLabel.setText(String.format("%.0f", generalPadHSlider.getValue()));
        }
        if (generalFontSlider != null && generalFontLabel != null) {
            generalFontLabel.setText(String.format("%.0f", generalFontSlider.getValue()));
        }
        if (stageFontSlider != null && stageFontLabel != null) {
            stageFontLabel.setText(String.format("%.0f", stageFontSlider.getValue()));
        }
        if (stageMinWidthSlider != null && stageMinWidthLabel != null) {
            stageMinWidthLabel.setText(String.format("%.0f", stageMinWidthSlider.getValue()));
        }
        if (stageMinHeightSlider != null && stageMinHeightLabel != null) {
            stageMinHeightLabel.setText(String.format("%.0f", stageMinHeightSlider.getValue()));
        }
        if (stagePadVSlider != null && stagePadVLabel != null) {
            stagePadVLabel.setText(String.format("%.0f", stagePadVSlider.getValue()));
        }
        if (stagePadHSlider != null && stagePadHLabel != null) {
            stagePadHLabel.setText(String.format("%.0f", stagePadHSlider.getValue()));
        }
        if (stageRadiusSlider != null && stageRadiusLabel != null) {
            stageRadiusLabel.setText(String.format("%.0f", stageRadiusSlider.getValue()));
        }
    }

    private void refreshPreviewDisabledMask() {
        boolean ge = customGeneralCheck == null || !customGeneralCheck.isSelected();
        boolean se = customStageCheck == null || !customStageCheck.isSelected();
        setDisableGeneral(ge);
        setDisableStage(se);
        if (previewRunPane != null) {
            previewRunPane.setOpacity(1.0);
        }
    }

    private void setDisableGeneral(boolean dis) {
        if (generalRadiusSlider != null) {
            generalRadiusSlider.setDisable(dis);
        }
        if (generalPadVSlider != null) {
            generalPadVSlider.setDisable(dis);
        }
        if (generalPadHSlider != null) {
            generalPadHSlider.setDisable(dis);
        }
        if (generalFontSlider != null) {
            generalFontSlider.setDisable(dis);
        }
        if (generalBgPicker != null) {
            generalBgPicker.setDisable(dis);
        }
        if (generalBorderPicker != null) {
            generalBorderPicker.setDisable(dis);
        }
        if (generalTextPicker != null) {
            generalTextPicker.setDisable(dis);
        }
        if (generalHoverPicker != null) {
            generalHoverPicker.setDisable(dis);
        }
        if (generalPressedPicker != null) {
            generalPressedPicker.setDisable(dis);
        }
    }

    private void setDisableStage(boolean dis) {
        if (stageFontSlider != null) {
            stageFontSlider.setDisable(dis);
        }
        if (stageMinWidthSlider != null) {
            stageMinWidthSlider.setDisable(dis);
        }
        if (stageMinHeightSlider != null) {
            stageMinHeightSlider.setDisable(dis);
        }
        if (stagePadVSlider != null) {
            stagePadVSlider.setDisable(dis);
        }
        if (stagePadHSlider != null) {
            stagePadHSlider.setDisable(dis);
        }
        if (stageRadiusSlider != null) {
            stageRadiusSlider.setDisable(dis);
        }
        if (stage1BgPicker != null) {
            stage1BgPicker.setDisable(dis);
        }
        if (stage1BorderPicker != null) {
            stage1BorderPicker.setDisable(dis);
        }
        if (stage1HoverPicker != null) {
            stage1HoverPicker.setDisable(dis);
        }
        if (stage1PressedPicker != null) {
            stage1PressedPicker.setDisable(dis);
        }
        if (stage2BgPicker != null) {
            stage2BgPicker.setDisable(dis);
        }
        if (stage2BorderPicker != null) {
            stage2BorderPicker.setDisable(dis);
        }
        if (stage2HoverPicker != null) {
            stage2HoverPicker.setDisable(dis);
        }
        if (stage2PressedPicker != null) {
            stage2PressedPicker.setDisable(dis);
        }
        if (stage3BgPicker != null) {
            stage3BgPicker.setDisable(dis);
        }
        if (stage3BorderPicker != null) {
            stage3BorderPicker.setDisable(dis);
        }
        if (stage3HoverPicker != null) {
            stage3HoverPicker.setDisable(dis);
        }
        if (stage3PressedPicker != null) {
            stage3PressedPicker.setDisable(dis);
        }
    }

    private void refreshStylesheetOnShell() {
        if (suppress || shell == null) {
            return;
        }
        shell.refreshPushButtonStylesheet();
    }

    private void schedulePersist() {
        if (suppress) {
            return;
        }
        persistDelay.stop();
        persistDelay.playFromStart();
    }

    private void persistNow() {
        if (suppress || shell == null) {
            return;
        }
        shell.persistDesktopSessionNow();
    }

    private static String colorToHex(ColorPicker cp, String fallbackHex) {
        if (cp == null || cp.getValue() == null) {
            return fallbackHex;
        }
        Color c = cp.getValue();
        int r = (int) Math.round(c.getRed() * 255.0);
        int g = (int) Math.round(c.getGreen() * 255.0);
        int b = (int) Math.round(c.getBlue() * 255.0);
        return String.format("#%02x%02x%02x", r, g, b);
    }
}
