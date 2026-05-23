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
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.ColumnConstraints;
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
    private Spinner<Integer> cardWidthSpinner;
    private Spinner<Integer> fullscreenScaleSpinner;
    private Spinner<Integer> paddingSpinner;
    private Spinner<Integer> gapHSpinner;
    private Spinner<Integer> gapVSpinner;
    private Spinner<Integer> radiusSpinner;
    private ComboBox<String> shadowCombo;
    private ComboBox<String> fontCombo;
    private Spinner<Integer> machineFontSpinner;
    private Spinner<Integer> metaFontSpinner;
    private Spinner<Integer> planFontSpinner;
    private Spinner<Integer> pctFontSpinner;
    private Spinner<Integer> chartSizeSpinner;
    private ColorPicker doneColorPicker;
    private ColorPicker remainColorPicker;
    private ComboBox<String> chartStyleCombo;
    private ComboBox<String> fullscreenThemeCombo;
    private CheckBox chartShadowCheckBox;

    private final List<Spinner<Integer>> numericSpinners = new ArrayList<>();

    private boolean suppress;

    public EquipmentStatusDashboardAppearancePanel(
            EquipmentStatusDashboardAppearancePrefs initial,
            Consumer<EquipmentStatusDashboardAppearancePrefs> onChange) {
        this.prefs = initial != null ? initial : EquipmentStatusDashboardAppearancePrefs.defaults();
        this.onChange = onChange != null ? onChange : __ -> {};
    }

    public VBox buildRoot() {
        GridPane grid = new GridPane();
        grid.setHgap(10);
        grid.setVgap(6);
        grid.setPadding(new Insets(4, 0, 4, 0));
        ColumnConstraints labelCol = new ColumnConstraints();
        labelCol.setMinWidth(Region.USE_PREF_SIZE);
        ColumnConstraints fieldCol = new ColumnConstraints();
        fieldCol.setMinWidth(96);
        fieldCol.setMaxWidth(120);
        ColumnConstraints hintCol = new ColumnConstraints();
        hintCol.setHgrow(Priority.ALWAYS);
        grid.getColumnConstraints().addAll(labelCol, fieldCol, hintCol);

        int row = 0;
        row = addSection(grid, row, "レイアウト");

        columnSpinner = intSpinner(0, 12, prefs.columnCount(), 1);
        addRow(grid, row++, "表示列数", columnSpinner, "0=自動");

        cardWidthSpinner = intSpinner(160, 520, (int) prefs.cardWidth(), 10);
        addRow(grid, row++, "カード幅 (px)", cardWidthSpinner, null);

        fullscreenScaleSpinner = intSpinner(80, 200, (int) prefs.fullscreenCardWidthPercent(), 1);
        addRow(grid, row++, "全画面幅 (%)", fullscreenScaleSpinner, null);

        paddingSpinner = intSpinner(4, 32, (int) prefs.cardPadding(), 1);
        addRow(grid, row++, "内余白 (px)", paddingSpinner, null);

        gapHSpinner = intSpinner(0, 48, (int) prefs.cardGapH(), 1);
        addRow(grid, row++, "横間隔 (px)", gapHSpinner, null);

        gapVSpinner = intSpinner(0, 48, (int) prefs.cardGapV(), 1);
        addRow(grid, row++, "縦間隔 (px)", gapVSpinner, null);

        radiusSpinner = intSpinner(0, 32, (int) prefs.cardBorderRadius(), 1);
        addRow(grid, row++, "角丸 (px)", radiusSpinner, null);

        shadowCombo = new ComboBox<>();
        shadowCombo.getItems().addAll("なし", "弱", "中", "強");
        shadowCombo.setValue(shadowLabel(prefs.cardShadowStyle()));
        shadowCombo.setMaxWidth(120);
        addRow(grid, row++, "カードの影", shadowCombo, null);

        row = addSection(grid, row, "全画面");

        fullscreenThemeCombo = new ComboBox<>();
        fullscreenThemeCombo.getItems().addAll("ダーク", "ライト", "壁面（高コントラスト）");
        fullscreenThemeCombo.setValue(fullscreenThemeLabel(prefs.fullscreenTheme()));
        fullscreenThemeCombo.setMaxWidth(200);
        addRow(grid, row++, "テーマ", fullscreenThemeCombo, null);

        row = addSection(grid, row, "フォント");

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
        fontCombo.setMaxWidth(200);
        addRow(grid, row++, "種類", fontCombo, null);

        machineFontSpinner = intSpinner(9, 28, (int) prefs.machineFontPx(), 1);
        addRow(grid, row++, "機械名 (px)", machineFontSpinner, null);

        metaFontSpinner = intSpinner(8, 20, (int) prefs.metaFontPx(), 1);
        addRow(grid, row++, "依頼・工程 (px)", metaFontSpinner, null);

        planFontSpinner = intSpinner(8, 18, (int) prefs.planFontPx(), 1);
        addRow(grid, row++, "予定行 (px)", planFontSpinner, null);

        pctFontSpinner = intSpinner(10, 32, (int) prefs.pctFontPx(), 1);
        addRow(grid, row++, "完了率 (px)", pctFontSpinner, null);

        row = addSection(grid, row, "円グラフ");

        chartSizeSpinner = intSpinner(40, 240, (int) prefs.chartSizePx(), 4);
        addRow(grid, row++, "サイズ (px)", chartSizeSpinner, null);

        doneColorPicker = colorPicker(prefs.chartDoneColorHex());
        remainColorPicker = colorPicker(prefs.chartRemainColorHex());
        HBox colors = new HBox(8, new Label("完了"), doneColorPicker, new Label("残"), remainColorPicker);
        colors.setAlignment(Pos.CENTER_LEFT);
        addRow(grid, row++, "色", colors, null);

        chartStyleCombo = new ComboBox<>();
        chartStyleCombo.getItems().addAll("平面", "立体風");
        chartStyleCombo.setValue(
                EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(prefs.chartStyle())
                        ? "立体風"
                        : "平面");
        chartStyleCombo.setMaxWidth(120);
        addRow(grid, row++, "スタイル", chartStyleCombo, null);

        chartShadowCheckBox = new CheckBox("影を付ける");
        chartShadowCheckBox.setSelected(prefs.chartShadowEnabled());
        addRow(grid, row++, "", chartShadowCheckBox, null);

        Button reset = new Button("既定に戻す");
        reset.setOnAction(e -> applyToControls(EquipmentStatusDashboardAppearancePrefs.defaults()));

        VBox root = new VBox(8, grid, reset);
        wireListeners();
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
            cardWidthSpinner.getValueFactory().setValue((int) p.cardWidth());
            fullscreenScaleSpinner.getValueFactory().setValue((int) p.fullscreenCardWidthPercent());
            paddingSpinner.getValueFactory().setValue((int) p.cardPadding());
            gapHSpinner.getValueFactory().setValue((int) p.cardGapH());
            gapVSpinner.getValueFactory().setValue((int) p.cardGapV());
            radiusSpinner.getValueFactory().setValue((int) p.cardBorderRadius());
            shadowCombo.setValue(shadowLabel(p.cardShadowStyle()));
            fullscreenThemeCombo.setValue(fullscreenThemeLabel(p.fullscreenTheme()));
            if (p.fontFamily().isBlank()) {
                fontCombo.setValue("（システム既定）");
            } else {
                fontCombo.setValue(p.fontFamily());
            }
            machineFontSpinner.getValueFactory().setValue((int) p.machineFontPx());
            metaFontSpinner.getValueFactory().setValue((int) p.metaFontPx());
            planFontSpinner.getValueFactory().setValue((int) p.planFontPx());
            pctFontSpinner.getValueFactory().setValue((int) p.pctFontPx());
            chartSizeSpinner.getValueFactory().setValue((int) p.chartSizePx());
            doneColorPicker.setValue(Color.web(p.chartDoneColorHex()));
            remainColorPicker.setValue(Color.web(p.chartRemainColorHex()));
            chartStyleCombo.setValue(
                    EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(p.chartStyle())
                            ? "立体風"
                            : "平面");
            chartShadowCheckBox.setSelected(p.chartShadowEnabled());
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
                    onChange.accept(prefs);
                };
        for (Spinner<Integer> s : numericSpinners) {
            s.valueProperty().addListener((o, a, b) -> fire.run());
        }
        shadowCombo.valueProperty().addListener((o, a, b) -> fire.run());
        fontCombo.valueProperty().addListener((o, a, b) -> fire.run());
        doneColorPicker.valueProperty().addListener((o, a, b) -> fire.run());
        remainColorPicker.valueProperty().addListener((o, a, b) -> fire.run());
        chartStyleCombo.valueProperty().addListener((o, a, b) -> fire.run());
        fullscreenThemeCombo.valueProperty().addListener((o, a, b) -> fire.run());
        chartShadowCheckBox.selectedProperty().addListener((o, a, b) -> fire.run());
    }

    private EquipmentStatusDashboardAppearancePrefs readFromControls() {
        String font =
                fontCombo.getValue() == null || "（システム既定）".equals(fontCombo.getValue())
                        ? ""
                        : fontCombo.getValue().strip();
        return new EquipmentStatusDashboardAppearancePrefs(
                spinnerIntValue(columnSpinner),
                spinnerInt(cardWidthSpinner),
                spinnerInt(fullscreenScaleSpinner),
                spinnerInt(paddingSpinner),
                spinnerInt(gapHSpinner),
                spinnerInt(gapVSpinner),
                spinnerInt(radiusSpinner),
                shadowKey(shadowCombo.getValue()),
                font,
                spinnerInt(machineFontSpinner),
                spinnerInt(metaFontSpinner),
                spinnerInt(planFontSpinner),
                spinnerInt(pctFontSpinner),
                spinnerInt(chartSizeSpinner),
                toHex(doneColorPicker.getValue()),
                toHex(remainColorPicker.getValue()),
                "立体風".equals(chartStyleCombo.getValue())
                        ? EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH
                        : EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                chartShadowCheckBox.isSelected(),
                fullscreenThemeKey(fullscreenThemeCombo.getValue()));
    }

    private Spinner<Integer> intSpinner(int min, int max, int value, int step) {
        Spinner<Integer> s =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(min, max, value, step));
        s.setEditable(true);
        s.setPrefWidth(96);
        numericSpinners.add(s);
        return s;
    }

    private static double spinnerInt(Spinner<Integer> s) {
        return spinnerIntValue(s);
    }

    private static int spinnerIntValue(Spinner<Integer> s) {
        Integer v = s != null ? s.getValue() : null;
        return v != null ? v : 0;
    }

    private static int addSection(GridPane grid, int row, String title) {
        Label head = new Label(title);
        head.getStyleClass().add("pm-equipment-status-appearance-section");
        GridPane.setColumnSpan(head, 3);
        grid.add(head, 0, row);
        return row + 1;
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
        grid.add(control, 1, row);
        if (hint != null) {
            Label h = new Label(hint);
            h.getStyleClass().add("pm-equipment-status-appearance-hint");
            grid.add(h, 2, row);
        }
    }

    private static String fullscreenThemeLabel(String key) {
        return switch (key != null ? key : EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK) {
            case EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_LIGHT -> "ライト";
            case EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_WALL -> "壁面（高コントラスト）";
            default -> "ダーク";
        };
    }

    private static String fullscreenThemeKey(String label) {
        if (label == null) {
            return EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK;
        }
        return switch (label) {
            case "ライト" -> EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_LIGHT;
            case "壁面（高コントラスト）" -> EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_WALL;
            default -> EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK;
        };
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
