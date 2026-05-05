package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.input.ScrollEvent;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.text.Font;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;
import jp.co.pm.ai.desktop.ui.GanttSheetKind;

/**
 * 「結果_設備ガント」等の時刻軸シートを plan JSON から読み、グラフィック表示する独立タブ。
 */
public final class EquipmentGanttGraphicTabController {

    private static final String DEFAULT_SHEET = "結果_設備ガント";

    private static final String HINT =
            "計画結果ビューアと同じ production_plan_multi_day*.json を指定します。"
                    + " 設備タイムライン（時刻列 HH:MM）と判定されるシートのうち、既定では「結果_設備ガント」を表示します。"
                    + " グラフィック表示はメインのこのタブから利用してください。";

    @FXML
    private Button reloadButton;

    @FXML
    private Button syncLatestButton;

    @FXML
    private TextField planJsonField;

    @FXML
    private Button browsePlanButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private Accordion sourceAccordion;

    @FXML
    private TitledPane sourceTitledPane;

    @FXML
    private BorderPane contentPane;

    @FXML
    private VBox graphicToolbarHost;

    private MainShellController shell;

    private Stage ownerStage;

    private String lastLoadedPlanPath = "";

    /** 再描画用に保持する最新の選択シート（ズーム・テーマ変更時）。 */
    private JsonTableIo.SheetTable lastGraphicSheet;

    private BorderPane graphicRootWrapper;

    private Slider graphicZoomSlider;

    private Label graphicZoomPercentLabel;

    private Slider graphicRowHeightSlider;

    private Label graphicRowHeightPctLabel;

    private Slider graphicSlotWidthSlider;

    private Label graphicSlotWidthPctLabel;

    private Slider graphicHeaderHeightSlider;

    private Label graphicHeaderHeightPctLabel;

    private ComboBox<String> equipmentGraphicBarFontCombo;

    private Slider graphicBarFontPctSlider;

    private Label graphicBarFontPctLabel;

    private Slider graphicDateColSlider;

    private Label graphicDateColWidthLabel;

    private Slider graphicMachColSlider;

    private Label graphicMachColWidthLabel;

    private Slider graphicProcColSlider;

    private Label graphicProcColWidthLabel;

    private Slider graphicShiftWheelHSlider;

    private Label graphicShiftWheelHLabel;

    /** 日付列幅スライダー上限（px）。0 は自動計測 */
    private static final double DATE_COL_WIDTH_SLIDER_MAX = 220;

    /** 機械名・工程名列幅スライダー上限（px）。0 は自動計測 */
    private static final double SIDE_COL_WIDTH_SLIDER_MAX = 800;

    /** Shift+ホイール横スクロール感度（％）。100＝従来のステップ相当 */
    private static final double SHIFT_WHEEL_H_SCROLL_MIN = 50;

    private static final double SHIFT_WHEEL_H_SCROLL_MAX = 400;

    private PauseTransition equipmentGraphicPersistDelay;

    private boolean graphicWheelHookInstalled;

    /**
     * 設備グラフィックの {@link EquipmentGraphicGanttPane#build} は重いため、スライダー連続変更では
     * この間隔（ms）より頻繁には再構築しない。ドラッグ終了時は {@link #flushGraphicRebuildNow()} で必ず反映する。
     */
    private static final long GRAPHIC_REBUILD_MIN_GAP_MS = 33L;

    private long graphicRebuildLastEmitMs;

    private PauseTransition graphicRebuildTrailing;

    /** {@link #applyEquipmentGanttSession} 等で複数スライダーを一度に動かすときの再構築抑制 */
    private boolean suppressGraphicRebuild;

    @FXML
    private void initialize() {
        if (hintLabel != null) {
            hintLabel.setText(HINT);
        }
        if (sourceAccordion != null && sourceTitledPane != null) {
            sourceAccordion.setExpandedPane(sourceTitledPane);
            /* ツールバー（列幅・Shift横スクロール等）を見つけやすくする */
            sourceTitledPane.setExpanded(true);
        }
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder("JSON を指定して再読みしてください。"));
        }
        graphicZoomSlider = new Slider(50, 200, 100);
        graphicZoomSlider.setPrefWidth(220);
        graphicZoomSlider.setShowTickLabels(true);
        graphicZoomSlider.setShowTickMarks(true);
        graphicZoomSlider.setMajorTickUnit(25);
        graphicZoomSlider.setMinorTickCount(4);
        graphicZoomPercentLabel = new Label("100%");
        graphicZoomSlider
                .valueProperty()
                .addListener(
                        (obs, oldV, v) -> {
                            graphicZoomPercentLabel.setText(String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicZoomSlider);

        graphicRowHeightSlider = new Slider(50, 200, 100);
        graphicRowHeightSlider.setPrefWidth(140);
        graphicRowHeightSlider.setShowTickLabels(false);
        graphicRowHeightPctLabel = new Label("100%");
        graphicRowHeightSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicRowHeightPctLabel.setText(String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicRowHeightSlider);

        graphicSlotWidthSlider = new Slider(50, 500, 100);
        graphicSlotWidthSlider.setPrefWidth(140);
        graphicSlotWidthSlider.setShowTickLabels(false);
        graphicSlotWidthPctLabel = new Label("100%");
        graphicSlotWidthSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicSlotWidthPctLabel.setText(String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicSlotWidthSlider);

        graphicHeaderHeightSlider = new Slider(50, 200, 100);
        graphicHeaderHeightSlider.setPrefWidth(140);
        graphicHeaderHeightSlider.setShowTickLabels(false);
        graphicHeaderHeightPctLabel = new Label("100%");
        graphicHeaderHeightSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicHeaderHeightPctLabel.setText(
                                    String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicHeaderHeightSlider);

        equipmentGraphicBarFontCombo = new ComboBox<>();
        List<String> families = new ArrayList<>(Font.getFamilies());
        Collections.sort(families);
        equipmentGraphicBarFontCombo.getItems().add("");
        equipmentGraphicBarFontCombo.getItems().addAll(families);
        equipmentGraphicBarFontCombo.setEditable(false);
        equipmentGraphicBarFontCombo.setPrefWidth(220);
        equipmentGraphicBarFontCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            flushGraphicRebuildNow();
                            scheduleEquipmentGraphicPersist();
                        });

        graphicBarFontPctSlider = new Slider(50, 200, 100);
        graphicBarFontPctSlider.setPrefWidth(140);
        graphicBarFontPctSlider.setShowTickLabels(false);
        graphicBarFontPctLabel = new Label("100%");
        graphicBarFontPctSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicBarFontPctLabel.setText(String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicBarFontPctSlider);

        graphicDateColSlider = new Slider(0, DATE_COL_WIDTH_SLIDER_MAX, 0);
        graphicDateColSlider.setPrefWidth(112);
        graphicDateColWidthLabel = new Label(formatLeftColWidthLabel(0));
        graphicDateColSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicDateColWidthLabel.setText(
                                    formatLeftColWidthLabel(v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicDateColSlider);

        graphicMachColSlider = new Slider(0, SIDE_COL_WIDTH_SLIDER_MAX, 0);
        graphicMachColSlider.setPrefWidth(112);
        graphicMachColWidthLabel = new Label(formatLeftColWidthLabel(0));
        graphicMachColSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicMachColWidthLabel.setText(
                                    formatLeftColWidthLabel(v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicMachColSlider);

        graphicProcColSlider = new Slider(0, SIDE_COL_WIDTH_SLIDER_MAX, 0);
        graphicProcColSlider.setPrefWidth(112);
        graphicProcColWidthLabel = new Label(formatLeftColWidthLabel(0));
        graphicProcColSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicProcColWidthLabel.setText(
                                    formatLeftColWidthLabel(v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicProcColSlider);

        graphicShiftWheelHSlider =
                new Slider(SHIFT_WHEEL_H_SCROLL_MIN, SHIFT_WHEEL_H_SCROLL_MAX, 200);
        graphicShiftWheelHSlider.setPrefWidth(130);
        graphicShiftWheelHSlider.setShowTickLabels(false);
        graphicShiftWheelHLabel = new Label("200%");
        graphicShiftWheelHSlider
                .valueProperty()
                .addListener(
                        (o, a, v) -> {
                            graphicShiftWheelHLabel.setText(
                                    String.format("%.0f%%", v.doubleValue()));
                            requestThrottledGraphicRebuild();
                            scheduleEquipmentGraphicPersist();
                        });
        wireGraphicSliderFlushOnDragEnd(graphicShiftWheelHSlider);

        if (graphicToolbarHost != null) {
            graphicToolbarHost.getChildren().setAll(buildGraphicToolbar());
        }
    }

    private static String formatLeftColWidthLabel(double px) {
        if (px <= 0.5) {
            return "自動";
        }
        return String.format("%.0fpx", px);
    }

    void applyEquipmentGanttSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppressGraphicRebuild = true;
        try {
            applyEquipmentGanttSessionBody(s);
        } finally {
            suppressGraphicRebuild = false;
            flushGraphicRebuildNow();
        }
    }

    private void applyEquipmentGanttSessionBody(DesktopSessionState s) {
        double z = s.equipmentGanttGraphicZoomPercent();
        if (graphicZoomSlider != null && Double.isFinite(z) && z >= 50 && z <= 200) {
            graphicZoomSlider.setValue(z);
        }
        double rh = s.equipmentGanttRowHeightPercent();
        if (graphicRowHeightSlider != null && Double.isFinite(rh) && rh >= 50 && rh <= 200) {
            graphicRowHeightSlider.setValue(rh);
            graphicRowHeightPctLabel.setText(String.format("%.0f%%", rh));
        }
        double hh = s.equipmentGanttHeaderHeightPercent();
        if (graphicHeaderHeightSlider != null && Double.isFinite(hh) && hh >= 50 && hh <= 200) {
            graphicHeaderHeightSlider.setValue(hh);
            graphicHeaderHeightPctLabel.setText(String.format("%.0f%%", hh));
        }
        double sw = s.equipmentGanttSlotWidthPercent();
        if (graphicSlotWidthSlider != null && Double.isFinite(sw) && sw >= 50 && sw <= 500) {
            graphicSlotWidthSlider.setValue(sw);
            graphicSlotWidthPctLabel.setText(String.format("%.0f%%", sw));
        }
        double bfp = s.equipmentGanttBarFontPercent();
        if (graphicBarFontPctSlider != null && Double.isFinite(bfp) && bfp >= 50 && bfp <= 200) {
            graphicBarFontPctSlider.setValue(bfp);
            graphicBarFontPctLabel.setText(String.format("%.0f%%", bfp));
        }
        String f = s.equipmentGanttBarFontFamily();
        if (equipmentGraphicBarFontCombo != null) {
            if (f == null || f.isBlank()) {
                equipmentGraphicBarFontCombo.setValue("");
            } else {
                String fs = f.strip();
                if (!equipmentGraphicBarFontCombo.getItems().contains(fs)) {
                    equipmentGraphicBarFontCombo.getItems().add(1, fs);
                }
                equipmentGraphicBarFontCombo.setValue(fs);
            }
        }
        double dwc = s.equipmentGanttDateColWidth();
        if (graphicDateColSlider != null
                && Double.isFinite(dwc)
                && dwc >= 0
                && dwc <= DATE_COL_WIDTH_SLIDER_MAX) {
            graphicDateColSlider.setValue(dwc);
            graphicDateColWidthLabel.setText(formatLeftColWidthLabel(dwc));
        }
        double mwc = s.equipmentGanttMachineColWidth();
        if (graphicMachColSlider != null
                && Double.isFinite(mwc)
                && mwc >= 0
                && mwc <= SIDE_COL_WIDTH_SLIDER_MAX) {
            graphicMachColSlider.setValue(mwc);
            graphicMachColWidthLabel.setText(formatLeftColWidthLabel(mwc));
        }
        double pwc = s.equipmentGanttProcessColWidth();
        if (graphicProcColSlider != null
                && Double.isFinite(pwc)
                && pwc >= 0
                && pwc <= SIDE_COL_WIDTH_SLIDER_MAX) {
            graphicProcColSlider.setValue(pwc);
            graphicProcColWidthLabel.setText(formatLeftColWidthLabel(pwc));
        }
        double swh = s.equipmentGanttShiftWheelHScrollPercent();
        if (graphicShiftWheelHSlider != null
                && Double.isFinite(swh)
                && swh >= SHIFT_WHEEL_H_SCROLL_MIN
                && swh <= SHIFT_WHEEL_H_SCROLL_MAX) {
            graphicShiftWheelHSlider.setValue(swh);
            graphicShiftWheelHLabel.setText(String.format("%.0f%%", swh));
        }
    }

    private void wireGraphicSliderFlushOnDragEnd(Slider slider) {
        if (slider == null) {
            return;
        }
        slider.valueChangingProperty()
                .addListener(
                        (obs, wasChanging, changing) -> {
                            if (!changing) {
                                flushGraphicRebuildNow();
                            }
                        });
    }

    private void requestThrottledGraphicRebuild() {
        if (suppressGraphicRebuild) {
            return;
        }
        long now = System.currentTimeMillis();
        long gap = GRAPHIC_REBUILD_MIN_GAP_MS;
        long elapsed = now - graphicRebuildLastEmitMs;
        if (elapsed >= gap || graphicRebuildLastEmitMs == 0L) {
            graphicRebuildLastEmitMs = now;
            rebuildGraphicView();
            return;
        }
        if (graphicRebuildTrailing == null) {
            graphicRebuildTrailing = new PauseTransition();
            graphicRebuildTrailing.setOnFinished(
                    e -> {
                        graphicRebuildLastEmitMs = System.currentTimeMillis();
                        if (!suppressGraphicRebuild) {
                            rebuildGraphicView();
                        }
                    });
        }
        graphicRebuildTrailing.stop();
        graphicRebuildTrailing.setDuration(Duration.millis(Math.max(1, gap - elapsed)));
        graphicRebuildTrailing.playFromStart();
    }

    private void flushGraphicRebuildNow() {
        if (suppressGraphicRebuild) {
            return;
        }
        if (graphicRebuildTrailing != null) {
            graphicRebuildTrailing.stop();
        }
        graphicRebuildLastEmitMs = System.currentTimeMillis();
        rebuildGraphicView();
    }

    double snapshotEquipmentGanttZoomPercent() {
        return graphicZoomSlider != null ? graphicZoomSlider.getValue() : 100d;
    }

    double snapshotEquipmentGanttDateColWidth() {
        return graphicDateColSlider != null ? graphicDateColSlider.getValue() : 0d;
    }

    double snapshotEquipmentGanttMachineColWidth() {
        return graphicMachColSlider != null ? graphicMachColSlider.getValue() : 0d;
    }

    double snapshotEquipmentGanttProcessColWidth() {
        return graphicProcColSlider != null ? graphicProcColSlider.getValue() : 0d;
    }

    double snapshotEquipmentGanttRowHeightPercent() {
        return graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d;
    }

    double snapshotEquipmentGanttHeaderHeightPercent() {
        return graphicHeaderHeightSlider != null ? graphicHeaderHeightSlider.getValue() : 100d;
    }

    double snapshotEquipmentGanttSlotWidthPercent() {
        return graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d;
    }

    String snapshotEquipmentGanttBarFontFamily() {
        if (equipmentGraphicBarFontCombo != null) {
            String v = equipmentGraphicBarFontCombo.getValue();
            return v != null ? v.strip() : "";
        }
        return "";
    }

    double snapshotEquipmentGanttBarFontPercent() {
        return graphicBarFontPctSlider != null ? graphicBarFontPctSlider.getValue() : 100d;
    }

    double snapshotEquipmentGanttShiftWheelHScrollPercent() {
        return graphicShiftWheelHSlider != null ? graphicShiftWheelHSlider.getValue() : 200d;
    }

    private void scheduleEquipmentGraphicPersist() {
        if (equipmentGraphicPersistDelay == null) {
            equipmentGraphicPersistDelay = new PauseTransition(Duration.millis(450));
            equipmentGraphicPersistDelay.setOnFinished(
                    e -> {
                        if (shell != null) {
                            shell.persistDesktopSessionNow();
                        }
                    });
        }
        equipmentGraphicPersistDelay.stop();
        equipmentGraphicPersistDelay.playFromStart();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        Platform.runLater(this::reloadFromFields);
    }

    /**
     * 実行タブの計画ブックパスと同じステムの .json があればフィールドに反映し再読み。
     */
    void tryAutofillJsonFromStage2Xlsx(String productionPlanPath, String memberSchedulePath) {
        if (planJsonField == null) {
            return;
        }
        String p = productionPlanPath != null ? productionPlanPath.strip() : "";
        if (p.isEmpty()) {
            return;
        }
        Path json = siblingJson(Path.of(p));
        if (json != null && Files.isRegularFile(json)) {
            planJsonField.setText(json.toString());
            Platform.runLater(this::reloadFromFields);
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromFields();
    }

    @FXML
    private void onSyncLatestButtonAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path plan = newestMatching(dir, "production_plan_multi_day_*.json");
            if (plan != null) {
                planJsonField.setText(plan.toString());
            }
            if (plan == null) {
                statusLabel.setText("このフォルダに production_plan_multi_day*.json がありません: " + dir);
                return;
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        reloadFromFields();
    }

    @FXML
    private void onBrowsePlanJsonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle("JSON");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("JSON", "*.json"));
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("All", "*.*"));
        if (shell != null) {
            try {
                Map<String, String> ui = shell.snapshotUiEnv();
                Path dir = AppPaths.defaultPlanningOutputDir(ui);
                if (Files.isDirectory(dir)) {
                    ch.setInitialDirectory(dir.toFile());
                }
            } catch (Exception ignored) {
                // ignore
            }
        }
        java.io.File picked = ch.showOpenDialog(ownerStage);
        if (picked != null) {
            planJsonField.setText(picked.getAbsolutePath());
            reloadFromFields();
        }
    }

    private void reloadFromFields() {
        if (contentPane == null) {
            return;
        }
        reloadButton.setDisable(true);
        syncLatestButton.setDisable(true);
        try {
            String ps = planJsonField != null ? planJsonField.getText().strip() : "";
            Path planPath = ps.isEmpty() ? null : Path.of(ps);
            if (planPath == null || !Files.isRegularFile(planPath)) {
                resetGraphicState("ファイルが指定されていないか、見つかりません。");
                statusLabel.setText("読み込み対象なし");
                return;
            }

            SheetLoad loaded = loadWorkbookSheetsForGraphic(planPath);
            Map<String, JsonTableIo.SheetTable> sheets = loaded.sheets();
            lastLoadedPlanPath = planPath.toString();

            Map<String, JsonTableIo.SheetTable> eligible = filterEquipmentTimelineSheets(sheets);
            if (eligible.isEmpty()) {
                resetGraphicState(
                        "設備タイムライン形式のシートが見つかりません（時刻列 HH:MM のシート）。");
                statusLabel.setText("対象シートなし: " + planPath.getFileName());
                return;
            }

            List<String> names = eligible.keySet().stream().sorted().toList();
            applySelectedSheetFromMap(eligible);
            String sheetUsed =
                    eligible.containsKey(DEFAULT_SHEET) ? DEFAULT_SHEET : names.get(0);
            statusLabel.setText(
                    "読み込み: "
                            + planPath.getFileName()
                            + " → "
                            + loaded.description()
                            + " / 表示シート="
                            + sheetUsed
                            + " / 対象シート数="
                            + names.size());
        } catch (Exception ex) {
            resetGraphicState("エラー");
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] " + ex.getMessage());
            }
        } finally {
            reloadButton.setDisable(false);
            syncLatestButton.setDisable(false);
        }
    }

    private void applySelectedSheetFromMap(Map<String, JsonTableIo.SheetTable> eligible) {
        if (eligible == null || eligible.isEmpty()) {
            return;
        }
        String name =
                eligible.containsKey(DEFAULT_SHEET)
                        ? DEFAULT_SHEET
                        : eligible.keySet().stream().sorted().findFirst().orElse("");
        if (name.isBlank()) {
            return;
        }
        JsonTableIo.SheetTable st = eligible.get(name);
        if (st == null) {
            return;
        }
        lastGraphicSheet = st;
        applyGraphicCenter(st);
    }

    private void resetGraphicState(String placeholderMsg) {
        lastGraphicSheet = null;
        graphicRootWrapper = null;
        graphicWheelHookInstalled = false;
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder(placeholderMsg));
        }
    }

    /**
     * 1 行の {@link HBox} に詰めると横幅不足で右側のコントロールが画面外に押し出されるため、複数行に分割する。
     */
    private VBox buildGraphicToolbar() {
        Label z = new Label("表示倍率");
        Label rhLab = new Label("行の高さ");
        Label hhLab = new Label("見出し行の高さ");
        Label swLab = new Label("時刻列幅");
        Label shiftHWLab = new Label("Shift横スクロール");
        Label dateColLab = new Label("日付列幅");
        Label machColLab = new Label("機械名列幅");
        Label procColLab = new Label("工程名列幅");
        Label fontHint = new Label("バー文字フォント");
        Label bfLab = new Label("バー文字サイズ");

        HBox rowMain =
                new HBox(
                        8,
                        z,
                        graphicZoomSlider,
                        graphicZoomPercentLabel,
                        rhLab,
                        graphicRowHeightSlider,
                        graphicRowHeightPctLabel,
                        hhLab,
                        graphicHeaderHeightSlider,
                        graphicHeaderHeightPctLabel,
                        swLab,
                        graphicSlotWidthSlider,
                        graphicSlotWidthPctLabel);
        rowMain.setAlignment(Pos.CENTER_LEFT);

        HBox rowCols =
                new HBox(
                        8,
                        shiftHWLab,
                        graphicShiftWheelHSlider,
                        graphicShiftWheelHLabel,
                        dateColLab,
                        graphicDateColSlider,
                        graphicDateColWidthLabel,
                        machColLab,
                        graphicMachColSlider,
                        graphicMachColWidthLabel,
                        procColLab,
                        graphicProcColSlider,
                        graphicProcColWidthLabel);
        rowCols.setAlignment(Pos.CENTER_LEFT);

        HBox rowFont =
                new HBox(8, fontHint, equipmentGraphicBarFontCombo, bfLab, graphicBarFontPctSlider, graphicBarFontPctLabel);
        rowFont.setAlignment(Pos.CENTER_LEFT);

        VBox stack = new VBox(6, rowMain, rowCols, rowFont);
        stack.setPadding(new Insets(0, 0, 4, 0));
        stack.setFillWidth(true);
        return stack;
    }

    private void applyGraphicCenter(JsonTableIo.SheetTable st) {
        if (contentPane == null || st == null) {
            return;
        }
        double zoom = graphicZoomSlider != null ? graphicZoomSlider.getValue() / 100.0 : 1.0;
        double rowPct = graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d;
        double slotPct = graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d;
        double headerPct =
                graphicHeaderHeightSlider != null ? graphicHeaderHeightSlider.getValue() : 100d;
        double barFp =
                graphicBarFontPctSlider != null ? graphicBarFontPctSlider.getValue() : 100d;
        DesktopTheme theme =
                shell != null ? shell.currentDesktopTheme() : DesktopTheme.LIGHT;
        ObservableList<ObservableList<String>> rows = toObservableRows(st);
        BorderPane gantt =
                EquipmentGraphicGanttPane.build(
                        st.columns(),
                        rows,
                        theme,
                        zoom,
                        rowPct,
                        slotPct,
                        snapshotEquipmentGanttBarFontFamily(),
                        barFp,
                        headerPct,
                        snapshotEquipmentGanttDateColWidth(),
                        snapshotEquipmentGanttMachineColWidth(),
                        snapshotEquipmentGanttProcessColWidth(),
                        snapshotEquipmentGanttShiftWheelHScrollPercent());
        if (graphicRootWrapper == null) {
            graphicRootWrapper = new BorderPane();
            contentPane.setCenter(graphicRootWrapper);
        }
        graphicRootWrapper.setCenter(gantt);
        installGraphicWheelZoomIfNeeded();
    }

    private void installGraphicWheelZoomIfNeeded() {
        if (graphicRootWrapper == null || graphicWheelHookInstalled) {
            return;
        }
        graphicWheelHookInstalled = true;
        graphicRootWrapper.addEventFilter(
                ScrollEvent.SCROLL,
                e -> {
                    if (!e.isControlDown()) {
                        return;
                    }
                    e.consume();
                    if (graphicZoomSlider == null) {
                        return;
                    }
                    double delta = e.getDeltaY() > 0 ? 5 : -5;
                    double v = Math.clamp(graphicZoomSlider.getValue() + delta, 50, 200);
                    graphicZoomSlider.setValue(v);
                    scheduleEquipmentGraphicPersist();
                });
    }

    private void rebuildGraphicView() {
        if (lastGraphicSheet == null || contentPane == null) {
            return;
        }
        applyGraphicCenter(lastGraphicSheet);
    }

    /** メインの {@link DesktopTheme} 変更時に Canvas 帯の配色を合わせて再描画する。 */
    void refreshGraphicForTheme() {
        flushGraphicRebuildNow();
    }

    /**
     * {@link GanttScheduleStyle#resolveKind} と同趣旨（設備ガント・グラフィック専用タブのみで使用し、
     * GanttScheduleStyle の Spreadsheet API 版とシグネチャ競合させない）。
     */
    private static GanttSheetKind resolveEquipmentGraphicSheetKind(
            String sheetName, List<String> columns) {
        if (columns != null && !columns.isEmpty() && "日時帯".equals(columns.get(0))) {
            return GanttSheetKind.EQUIPMENT_TIMELINE;
        }
        if (sheetName != null) {
            if (sheetName.contains("設備")
                    && (sheetName.contains("ガント") || sheetName.contains("時間割"))) {
                return GanttSheetKind.EQUIPMENT_TIMELINE;
            }
        }
        return GanttSheetKind.DEFAULT;
    }

    private static Map<String, JsonTableIo.SheetTable> filterEquipmentTimelineSheets(
            Map<String, JsonTableIo.SheetTable> sheets) {
        Map<String, JsonTableIo.SheetTable> out = new LinkedHashMap<>();
        for (Map.Entry<String, JsonTableIo.SheetTable> e : sheets.entrySet()) {
            GanttSheetKind k = resolveEquipmentGraphicSheetKind(e.getKey(), e.getValue().columns());
            if (k == GanttSheetKind.EQUIPMENT_TIMELINE) {
                out.put(e.getKey(), e.getValue());
            }
        }
        return out;
    }

    private static ObservableList<ObservableList<String>> toObservableRows(JsonTableIo.SheetTable t) {
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        List<String> cols = t.columns();
        for (Map<String, String> map : t.rows()) {
            ObservableList<String> line = FXCollections.observableArrayList();
            for (String h : cols) {
                line.add(map != null ? map.getOrDefault(h, "") : "");
            }
            rows.add(line);
        }
        return rows;
    }

    private StackPane emptyPlaceholder(String msg) {
        StackPane p = new StackPane(new Label(msg));
        StackPane.setAlignment(p.getChildren().get(0), Pos.CENTER);
        return p;
    }

    private record SheetLoad(Map<String, JsonTableIo.SheetTable> sheets, String description) {}

    /**
     * ブック JSON は論理ビューがあればそれを読む（他シートの結合セル展開用）。
     * 「結果_設備ガント」のタイムセルは xlsx 由来 JSON では欠損しがち（シェイプ描画のため）なので、
     * 兄弟の {@code *_equipment_gantt_contract.json} があればそのシートだけ契約から組み立てた表で上書きする。
     */
    private static SheetLoad loadWorkbookSheetsForGraphic(Path planJsonFromField)
            throws IOException {
        Path fn0 = planJsonFromField.getFileName();
        if (fn0 != null && fn0.toString().endsWith(".json")) {
            String stem0 = fn0.toString().substring(0, fn0.toString().length() - 5);
            if (stem0.endsWith("_equipment_gantt_contract")) {
                JsonTableIo.SheetTable ganttOnly =
                        EquipmentGanttContractSheetTableBuilder.buildFromContractPath(
                                planJsonFromField);
                Map<String, JsonTableIo.SheetTable> m = new LinkedHashMap<>();
                m.put(DEFAULT_SHEET, ganttOnly);
                return new SheetLoad(
                        m, fn0.toString() + " (設備ガント契約・直接)");
            }
        }
        Path logical = resolveLogicalViewPath(planJsonFromField);
        Path workbookJson =
                logical != null && Files.isRegularFile(logical) ? logical : planJsonFromField;
        Map<String, JsonTableIo.SheetTable> sheets =
                new LinkedHashMap<>(JsonTableIo.loadSheetsWorkbook(workbookJson));

        Path contract = resolveEquipmentContractSibling(planJsonFromField);
        String desc;
        if (logical != null && workbookJson.equals(logical)) {
            desc = logical.getFileName().toString() + " (論理ビュー)";
        } else {
            desc = planJsonFromField.getFileName().toString();
        }
        if (contract != null && Files.isRegularFile(contract)) {
            JsonTableIo.SheetTable gantt =
                    EquipmentGanttContractSheetTableBuilder.buildFromContractPath(contract);
            sheets.put(DEFAULT_SHEET, gantt);
            desc = desc + " / " + contract.getFileName() + " (設備ガント帯)";
        }
        return new SheetLoad(sheets, desc);
    }

    /** 論理ビュー JSON 本体のパス（直接指定または sibling）。無ければ null。 */
    private static Path resolveLogicalViewPath(Path planJsonFromField) {
        if (planJsonFromField == null || !Files.isRegularFile(planJsonFromField)) {
            return null;
        }
        Path fn = planJsonFromField.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        if (!name.endsWith(".json")) {
            return null;
        }
        String stem = name.substring(0, name.length() - 5);
        if (stem.endsWith("_logical_view")) {
            return planJsonFromField;
        }
        Path sibling = planJsonFromField.resolveSibling(stem + "_logical_view.json");
        return Files.isRegularFile(sibling) ? sibling : null;
    }

    /**
     * {@code production_plan_multi_day_xxx.json} と並ぶ {@code …_equipment_gantt_contract.json}。
     * {@code *_logical_view.json} を開いているときは stem から {@code _logical_view} を除いて兄弟を解決する。
     */
    private static Path resolveEquipmentContractSibling(Path planJsonFromField) {
        if (planJsonFromField == null) {
            return null;
        }
        Path fn = planJsonFromField.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        if (!name.endsWith(".json")) {
            return null;
        }
        String stem = name.substring(0, name.length() - 5);
        if (stem.endsWith("_equipment_gantt_contract")) {
            return null;
        }
        if (stem.endsWith("_logical_view")) {
            stem = stem.substring(0, stem.length() - "_logical_view".length());
        }
        return planJsonFromField.resolveSibling(stem + "_equipment_gantt_contract.json");
    }

    private static Path siblingJson(Path workbookPath) {
        if (workbookPath == null) {
            return null;
        }
        Path fn = workbookPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        String stem;
        if (name.endsWith(".xlsx")) {
            stem = name.substring(0, name.length() - 5);
        } else if (name.endsWith(".xlsm")) {
            stem = name.substring(0, name.length() - 5);
        } else {
            return null;
        }
        return workbookPath.resolveSibling(stem + ".json");
    }

    private static Path newestMatching(Path dir, String glob) throws IOException {
        if (!Files.isDirectory(dir)) {
            return null;
        }
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, glob)) {
            for (Path p : stream) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t >= bestTime) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
    }
}
