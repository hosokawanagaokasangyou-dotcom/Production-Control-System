package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.function.Consumer;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;

/** ダッシュボード見た目設定 UI（タブ内 TitledPane 向け）。 */
public final class EquipmentStatusDashboardAppearancePanel {

    private final Consumer<EquipmentStatusDashboardAppearancePrefs> onChange;
    private EquipmentStatusDashboardAppearancePrefs prefs;

    private Spinner<Integer> columnSpinner;
    private Slider cardWidthSlider;
    private Slider fullscreenScaleSlider;
    private Slider paddingSlider;
    private Slider gapHSlider;
    private Slider gapVSlider;
    private Slider radiusSlider;
    private ComboBox<String> shadowCombo;
    private ComboBox<String> fontCombo;
    private Slider machineFontSlider;
    private Slider metaFontSlider;
    private Slider planFontSlider;
    private Slider pctFontSlider;
    private Slider chartSizeSlider;
    private ColorPicker doneColorPicker;
    private ColorPicker remainColorPicker;
    private ComboBox<String> chartStyleCombo;
    private CheckBox chartShadowCheckBox;
    private Label cardWidthValue;
    private Label chartSizeValue;

    private boolean suppress;

    public EquipmentStatusDashboardAppearancePanel(
            EquipmentStatusDashboardAppearancePrefs initial,
            Consumer<EquipmentStatusDashboardAppearancePrefs> onChange) {
        this.prefs = initial != null ? initial : EquipmentStatusDashboardAppearancePrefs.defaults();
        this.onChange = onChange != null ? onChange : __ -> {};
    }

    public VBox buildRoot() {
        GridPane grid = new GridPane();
        grid.setHgap(12);
        grid.setVgap(8);
        grid.setPadding(new Insets(4, 0, 4, 0));
        int row = 0;

        columnSpinner = new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 12, prefs.columnCount()));
        columnSpinner.setEditable(true);
        columnSpinner.setPrefWidth(80);
        addRow(grid, row++, "表示列数", columnSpinner, "0=自動");

        cardWidthSlider = slider(160, 520, prefs.cardWidth(), 10);
        cardWidthValue = valueLabel();
        addRow(grid, row++, "カード幅 (px)", labeledSlider(cardWidthSlider, cardWidthValue), null);

        fullscreenScaleSlider = slider(80, 200, prefs.fullscreenCardWidthPercent(), 1);
        addRow(grid, row++, "全画面カード幅 (%)", labeledSlider(fullscreenScaleSlider, valueLabel()), null);

        paddingSlider = slider(4, 32, prefs.cardPadding(), 1);
        addRow(grid, row++, "カード内余白 (px)", slider(paddingSlider), null);

        gapHSlider = slider(0, 48, prefs.cardGapH(), 1);
        gapVSlider = slider(0, 48, prefs.cardGapV(), 1);
        addRow(grid, row++, "横間隔 (px)", slider(gapHSlider), null);
        addRow(grid, row++, "縦間隔 (px)", slider(gapVSlider), null);

        radiusSlider = slider(0, 32, prefs.cardBorderRadius(), 1);
        addRow(grid, row++, "角丸 (px)", slider(radiusSlider), null);

        shadowCombo = new ComboBox<>();
        shadowCombo.getItems().addAll("なし", "弱", "中", "強");
        shadowCombo.setValue(shadowLabel(prefs.cardShadowStyle()));
        shadowCombo.setMaxWidth(Double.MAX_VALUE);
        addRow(grid, row++, "カードの影", shadowCombo, null);

        fontCombo = new ComboBox<>();
        List<String> families = new ArrayList<>();
        families.add("（システム既定）");
        families.addAll(Font.getFamilies());
        fontCombo.getItems().setAll(families);
        if (prefs.fontFamily().isBlank()) {
            fontCombo.setValue("（システム既定）");
        } else {
            fontCombo.setValue(prefs.fontFamily());
        }
        fontCombo.setMaxWidth(Double.MAX_VALUE);
        addRow(grid, row++, "フォント", fontCombo, null);

        machineFontSlider = slider(9, 28, prefs.machineFontPx(), 1);
        metaFontSlider = slider(8, 20, prefs.metaFontPx(), 1);
        planFontSlider = slider(8, 18, prefs.planFontPx(), 1);
        pctFontSlider = slider(10, 32, prefs.pctFontPx(), 1);
        addRow(grid, row++, "機械名 (px)", slider(machineFontSlider), null);
        addRow(grid, row++, "依頼・工程 (px)", slider(metaFontSlider), null);
        addRow(grid, row++, "予定行 (px)", slider(planFontSlider), null);
        addRow(grid, row++, "完了率 (px)", slider(pctFontSlider), null);

        chartSizeSlider = slider(40, 240, prefs.chartSizePx(), 4);
        chartSizeValue = valueLabel();
        addRow(grid, row++, "円グラフ (px)", labeledSlider(chartSizeSlider, chartSizeValue), null);

        doneColorPicker = colorPicker(prefs.chartDoneColorHex());
        remainColorPicker = colorPicker(prefs.chartRemainColorHex());
        HBox colors = new HBox(8, new Label("完了"), doneColorPicker, new Label("残"), remainColorPicker);
        colors.setAlignment(Pos.CENTER_LEFT);
        addRow(grid, row++, "円グラフ色", colors, null);

        chartStyleCombo = new ComboBox<>();
        chartStyleCombo.getItems().addAll("平面", "立体風");
        chartStyleCombo.setValue(
                EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(prefs.chartStyle())
                        ? "立体風"
                        : "平面");
        chartStyleCombo.setMaxWidth(Double.MAX_VALUE);
        addRow(grid, row++, "円グラフ", chartStyleCombo, null);

        chartShadowCheckBox = new CheckBox("円グラフに影");
        chartShadowCheckBox.setSelected(prefs.chartShadowEnabled());
        addRow(grid, row++, "", chartShadowCheckBox, null);

        Button reset = new Button("既定に戻す");
        reset.setOnAction(e -> applyToControls(EquipmentStatusDashboardAppearancePrefs.defaults()));

        VBox root = new VBox(8, grid, reset);
        wireListeners();
        updateValueLabels();
        return root;
    }

    public EquipmentStatusDashboardAppearancePrefs snapshot() {
        return prefs;
    }

    public void applyPrefs(EquipmentStatusDashboardAppearancePrefs p) {
        prefs = p != null ? p : EquipmentStatusDashboardAppearancePrefs.defaults();
        applyToControls(prefs);
    }

    private void applyToControls(EquipmentStatusDashboardAppearancePrefs p) {
        suppress = true;
        try {
            columnSpinner.getValueFactory().setValue(p.columnCount());
            cardWidthSlider.setValue(p.cardWidth());
            fullscreenScaleSlider.setValue(p.fullscreenCardWidthPercent());
            paddingSlider.setValue(p.cardPadding());
            gapHSlider.setValue(p.cardGapH());
            gapVSlider.setValue(p.cardGapV());
            radiusSlider.setValue(p.cardBorderRadius());
            shadowCombo.setValue(shadowLabel(p.cardShadowStyle()));
            if (p.fontFamily().isBlank()) {
                fontCombo.setValue("（システム既定）");
            } else {
                fontCombo.setValue(p.fontFamily());
            }
            machineFontSlider.setValue(p.machineFontPx());
            metaFontSlider.setValue(p.metaFontPx());
            planFontSlider.setValue(p.planFontPx());
            pctFontSlider.setValue(p.pctFontPx());
            chartSizeSlider.setValue(p.chartSizePx());
            doneColorPicker.setValue(Color.web(p.chartDoneColorHex()));
            remainColorPicker.setValue(Color.web(p.chartRemainColorHex()));
            chartStyleCombo.setValue(
                    EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(p.chartStyle())
                            ? "立体風"
                            : "平面");
            chartShadowCheckBox.setSelected(p.chartShadowEnabled());
            updateValueLabels();
        } finally {
            suppress = false;
        }
        prefs = p;
    }

    private void wireListeners() {
        Runnable fire =
                () -> {
                    if (suppress) {
                        return;
                    }
                    prefs = readFromControls();
                    updateValueLabels();
                    onChange.accept(prefs);
                };
        columnSpinner.valueProperty().addListener((o, a, b) -> fire.run());
        for (Slider s :
                List.of(
                        cardWidthSlider,
                        fullscreenScaleSlider,
                        paddingSlider,
                        gapHSlider,
                        gapVSlider,
                        radiusSlider,
                        machineFontSlider,
                        metaFontSlider,
                        planFontSlider,
                        pctFontSlider,
                        chartSizeSlider)) {
            s.valueProperty().addListener((o, a, b) -> fire.run());
        }
        shadowCombo.valueProperty().addListener((o, a, b) -> fire.run());
        fontCombo.valueProperty().addListener((o, a, b) -> fire.run());
        doneColorPicker.valueProperty().addListener((o, a, b) -> fire.run());
        remainColorPicker.valueProperty().addListener((o, a, b) -> fire.run());
        chartStyleCombo.valueProperty().addListener((o, a, b) -> fire.run());
        chartShadowCheckBox.selectedProperty().addListener((o, a, b) -> fire.run());
    }

    private EquipmentStatusDashboardAppearancePrefs readFromControls() {
        String font =
                fontCombo.getValue() == null || "（システム既定）".equals(fontCombo.getValue())
                        ? ""
                        : fontCombo.getValue().strip();
        return new EquipmentStatusDashboardAppearancePrefs(
                columnSpinner.getValue() != null ? columnSpinner.getValue() : 0,
                cardWidthSlider.getValue(),
                fullscreenScaleSlider.getValue(),
                paddingSlider.getValue(),
                gapHSlider.getValue(),
                gapVSlider.getValue(),
                radiusSlider.getValue(),
                shadowKey(shadowCombo.getValue()),
                font,
                machineFontSlider.getValue(),
                metaFontSlider.getValue(),
                planFontSlider.getValue(),
                pctFontSlider.getValue(),
                chartSizeSlider.getValue(),
                toHex(doneColorPicker.getValue()),
                toHex(remainColorPicker.getValue()),
                "立体風".equals(chartStyleCombo.getValue())
                        ? EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH
                        : EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                chartShadowCheckBox.isSelected());
    }

    private void updateValueLabels() {
        if (cardWidthValue != null) {
            cardWidthValue.setText(String.format(Locale.ROOT, "%.0f", cardWidthSlider.getValue()));
        }
        if (chartSizeValue != null) {
            chartSizeValue.setText(String.format(Locale.ROOT, "%.0f", chartSizeSlider.getValue()));
        }
    }

    private static Slider slider(double min, double max, double val, double block) {
        Slider s = new Slider(min, max, val);
        s.setBlockIncrement(block);
        s.setMajorTickUnit(block * 5);
        s.setShowTickMarks(false);
        s.setMaxWidth(Double.MAX_VALUE);
        return s;
    }

    private static HBox labeledSlider(Slider slider, Label value) {
        HBox box = new HBox(8, slider, value);
        HBox.setHgrow(slider, Priority.ALWAYS);
        slider.valueProperty().addListener((o, a, b) -> value.setText(String.format(Locale.ROOT, "%.0f", b.doubleValue())));
        return box;
    }

    private static Slider slider(Slider s) {
        HBox.setHgrow(s, Priority.ALWAYS);
        return s;
    }

    private static Label valueLabel() {
        Label l = new Label();
        l.setMinWidth(36);
        return l;
    }

    private static ColorPicker colorPicker(String hex) {
        ColorPicker p = new ColorPicker();
        try {
            p.setValue(Color.web(hex));
        } catch (Exception ex) {
            p.setValue(Color.web("#0d9488"));
        }
        return p;
    }

    private static void addRow(GridPane grid, int row, String label, javafx.scene.Node control, String hint) {
        Label l = new Label(label);
        grid.add(l, 0, row);
        if (control instanceof Region r) {
            GridPane.setHgrow(r, Priority.ALWAYS);
        }
        grid.add(control, 1, row);
        if (hint != null) {
            Label h = new Label(hint);
            h.getStyleClass().add("pm-equipment-status-appearance-hint");
            grid.add(h, 2, row);
        }
    }

    private static String shadowLabel(String key) {
        return switch (key) {
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_NONE -> "なし";
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_MEDIUM -> "中";
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_STRONG -> "強";
            default -> "弱";
        };
    }

    private static String shadowKey(String label) {
        if (label == null) {
            return EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE;
        }
        return switch (label) {
            case "なし" -> EquipmentStatusDashboardAppearancePrefs.SHADOW_NONE;
            case "中" -> EquipmentStatusDashboardAppearancePrefs.SHADOW_MEDIUM;
            case "強" -> EquipmentStatusDashboardAppearancePrefs.SHADOW_STRONG;
            default -> EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE;
        };
    }

    private static String toHex(Color c) {
        if (c == null) {
            return "#000000";
        }
        return String.format(
                Locale.ROOT,
                "#%02x%02x%02x",
                (int) Math.round(c.getRed() * 255),
                (int) Math.round(c.getGreen() * 255),
                (int) Math.round(c.getBlue() * 255));
    }
}
