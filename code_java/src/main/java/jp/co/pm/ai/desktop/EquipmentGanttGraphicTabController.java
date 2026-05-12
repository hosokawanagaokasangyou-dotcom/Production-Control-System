package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Pattern;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.print.PageLayout;
import javafx.print.PageOrientation;
import javafx.print.Paper;
import javafx.print.Printer;
import javafx.print.PrinterJob;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.input.ScrollEvent;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Slider;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.Parent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.Duration;
import javafx.util.StringConverter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.EquipmentGanttBadgeDragDelta;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttSheetBundle;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintDaySlices;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintPageWrapper;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;
import jp.co.pm.ai.desktop.ui.EquipmentGanttPersonBadgeWireDashStyle;
import jp.co.pm.ai.desktop.ui.GanttSheetKind;

/**
 * 「結果_設備ガント」等の時刻軸シートを plan JSON から読み、グラフィック表示する独立タブ。
 * グラフィック調整ツールバーのレイアウトは {@code EquipmentGanttGraphicTab.fxml} の FlowPane で定義する。
 */
public final class EquipmentGanttGraphicTabController {

    private static final String DEFAULT_SHEET = "結果_設備ガント";

    /** {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} の時刻列検出と整合する見出し判定（ログ・検証用） */
    private static final Pattern HH_MM_COLUMN_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    /**
     * 非空タイムラインが皆無の「大表」で {@link EquipmentGraphicGanttPane#build}（行ごと Canvas）を
     * 起動しない行数の下限。配台が空のときのメモリ急増を抑える。
     */
    private static final int HEAVY_EMPTY_GRID_MIN_ROWS = 300;

    /**
     * 上記判定で先頭から走査する最大行数（全行が空の想定のとき概ね足りるが、上限で打ち切る）。
     */
    private static final int HEAVY_EMPTY_GRID_SCAN_MAX_ROWS = 2000;

    private static final ObjectMapper GANTT_CONTRACT_PEEK_OM = new ObjectMapper();

    private static final String HINT =
            "計画結果ビューアと同じ 計画*.json（または旧 production_plan_multi_day*.json）を指定します。"
                    + " 設備タイムライン（時刻列 HH:MM）と判定されるシートのうち、既定では「結果_設備ガント」を表示します。"
                    + " グラフィック表示はメインのこのタブから利用してください。";

    @FXML
    private Button reloadButton;

    @FXML
    private Button syncLatestButton;

    @FXML
    private Button printGanttButton;

    @FXML
    private Button exportGanttPdfButton;

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

    private MainShellController shell;

    private Stage ownerStage;

    private String lastLoadedPlanPath = "";

    /** 再描画用に保持する最新の選択シート（ズーム・テーマ変更時）。 */
    private JsonTableIo.SheetTable lastGraphicSheet;

    /** 契約 JSON から得たバッジグリッド（{@link #DEFAULT_SHEET} 表示時のみ使用）。 */
    private List<List<String>> loadedContractBadgeRows;

    /** {@link #applyGraphicCenter} に渡す現在のバッジ行（シートに応じて null）。 */
    private List<List<String>> badgeRowsForCurrentGraphic;

    /** {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane#computeDataFingerprint} と同一キーでデータ同一判定に使う。 */
    private String equipmentGanttGraphicDataFingerprint = "";

    /** データ同一時のみ有効なバッジドラッグずれ（セッションと同期）。 */
    private final Map<String, EquipmentGanttBadgeDragDelta> equipmentGanttBadgeDragDeltas =
            new LinkedHashMap<>();

    private BorderPane graphicRootWrapper;

    @FXML
    private CheckBox personBadgeShowCheckBox;

    @FXML
    private CheckBox personBadgeDragAdjustCheckBox;

    @FXML
    private CheckBox personBadgeWireShowCheckBox;

    @FXML
    private Slider graphicZoomSlider;

    @FXML
    private Label graphicZoomPercentLabel;

    @FXML
    private Slider graphicRowHeightSlider;

    @FXML
    private Label graphicRowHeightPctLabel;

    @FXML
    private Slider graphicSlotWidthSlider;

    @FXML
    private Label graphicSlotWidthPctLabel;

    @FXML
    private Slider graphicHeaderHeightSlider;

    @FXML
    private Label graphicHeaderHeightPctLabel;

    @FXML
    private ComboBox<String> equipmentGraphicBarFontCombo;

    @FXML
    private Slider graphicBarFontPctSlider;

    @FXML
    private Label graphicBarFontPctLabel;

    @FXML
    private Slider graphicDateColSlider;

    @FXML
    private Label graphicDateColWidthLabel;

    @FXML
    private Slider graphicMachColSlider;

    @FXML
    private Label graphicMachColWidthLabel;

    @FXML
    private Slider graphicProcColSlider;

    @FXML
    private Label graphicProcColWidthLabel;

    @FXML
    private Slider graphicShiftWheelHSlider;

    @FXML
    private Label graphicShiftWheelHLabel;

    @FXML
    private Slider graphicPersonBadgeGapSlider;

    @FXML
    private Label graphicPersonBadgeGapLabel;

    @FXML
    private Slider graphicPersonBadgeBandVerticalOffsetSlider;

    @FXML
    private Label graphicPersonBadgeBandVerticalOffsetLabel;

    @FXML
    private CheckBox graphicPersonBadgeWireThemeColorCheckBox;

    @FXML
    private ColorPicker graphicPersonBadgeWireColorPicker;

    @FXML
    private Slider graphicPersonBadgeWireWidthSlider;

    @FXML
    private Label graphicPersonBadgeWireWidthLabel;

    @FXML
    private ComboBox<EquipmentGanttPersonBadgeWireDashStyle> graphicPersonBadgeWireDashCombo;

    @FXML
    private Slider graphicPersonBadgeWireMaxLengthSlider;

    @FXML
    private Label graphicPersonBadgeWireMaxLengthLabel;

    /** 日付列幅スライダー上限（px）。0 は自動計測 */
    private static final double DATE_COL_WIDTH_SLIDER_MAX = 220;

    /** 機械名・工程名列幅スライダー上限（px）。0 は自動計測 */
    private static final double SIDE_COL_WIDTH_SLIDER_MAX = 800;

    /** Shift+ホイール横スクロール感度（％）。100＝従来のステップ相当 */
    private static final double SHIFT_WHEEL_H_SCROLL_MIN = 50;

    /** ワイヤー太さスライダー上限（px）。{@link DesktopSessionState#MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX} と整合 */
    private static final double WIRE_WIDTH_SLIDER_MAX =
            DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX;

    /** ワイヤー長上限スライダー（px）。{@link DesktopSessionState#MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX} と整合 */
    private static final double WIRE_MAX_LENGTH_SLIDER_MAX =
            DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX;

    private static final double SHIFT_WHEEL_H_SCROLL_MAX = 1000;

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

    /**
     * Ctrl+ホイールで拡大率変更した直後の再構築のみ、横スクロールをマウス位置基準で復元する。
     * {@link #applyGraphicCenter} で消費して null に戻す。
     */
    private EquipmentGraphicGanttPane.HorizontalZoomAnchor pendingHorizontalZoomAnchor;

    @FXML
    private void initialize() {
        if (hintLabel != null) {
            hintLabel.setText(HINT);
        }
        // FXML の expanded="false" を尊重する。起動時に setExpandedPane すると開アニメが走り、
        // 直後の reload 成功時の setExpanded(false) と競合して見た目だけ中途半端に残ることがある。
        if (sourceTitledPane != null) {
            sourceTitledPane
                    .expandedProperty()
                    .addListener(
                            (o, a, b) -> {
                                scheduleEquipmentGraphicPersist();
                                requestThrottledGraphicRebuild();
                            });
        }
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder("JSON を指定して再読みしてください。"));
        }
        populateEquipmentGraphicBarFontComboItems();
        attachGraphicToolbarListeners();
    }

    /** FXML で定義済みのコントロールへ値変更リスナーを付ける（ノード生成は FXML 側）。 */
    private void attachGraphicToolbarListeners() {
        if (graphicZoomSlider == null) {
            return;
        }
        Runnable graphicCommitted =
                () -> {
                    flushGraphicRebuildNow();
                    scheduleEquipmentGraphicPersist();
                };

        SliderCommittedChangeSupport.install(
                graphicZoomSlider,
                () ->
                        graphicZoomPercentLabel.setText(
                                String.format("%.0f%%", graphicZoomSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicRowHeightSlider,
                () ->
                        graphicRowHeightPctLabel.setText(
                                String.format("%.0f%%", graphicRowHeightSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicSlotWidthSlider,
                () ->
                        graphicSlotWidthPctLabel.setText(
                                String.format("%.0f%%", graphicSlotWidthSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicHeaderHeightSlider,
                () ->
                        graphicHeaderHeightPctLabel.setText(
                                String.format("%.0f%%", graphicHeaderHeightSlider.getValue())),
                graphicCommitted);

        equipmentGraphicBarFontCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            flushGraphicRebuildNow();
                            scheduleEquipmentGraphicPersist();
                        });

        SliderCommittedChangeSupport.install(
                graphicBarFontPctSlider,
                () ->
                        graphicBarFontPctLabel.setText(
                                String.format("%.0f%%", graphicBarFontPctSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicDateColSlider,
                () ->
                        graphicDateColWidthLabel.setText(
                                formatLeftColWidthLabel(graphicDateColSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicMachColSlider,
                () ->
                        graphicMachColWidthLabel.setText(
                                formatLeftColWidthLabel(graphicMachColSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicProcColSlider,
                () ->
                        graphicProcColWidthLabel.setText(
                                formatLeftColWidthLabel(graphicProcColSlider.getValue())),
                graphicCommitted);

        SliderCommittedChangeSupport.install(
                graphicShiftWheelHSlider,
                () ->
                        graphicShiftWheelHLabel.setText(
                                String.format("%.0f%%", graphicShiftWheelHSlider.getValue())),
                graphicCommitted);

        if (graphicPersonBadgeGapSlider != null) {
            SliderCommittedChangeSupport.install(
                    graphicPersonBadgeGapSlider,
                    () -> {
                        if (graphicPersonBadgeGapLabel != null) {
                            graphicPersonBadgeGapLabel.setText(
                                    String.format(
                                            "%.0fpx",
                                            graphicPersonBadgeGapSlider.getValue()));
                        }
                    },
                    graphicCommitted);
        }
        if (graphicPersonBadgeBandVerticalOffsetSlider != null) {
            SliderCommittedChangeSupport.install(
                    graphicPersonBadgeBandVerticalOffsetSlider,
                    () -> {
                        if (graphicPersonBadgeBandVerticalOffsetLabel != null) {
                            graphicPersonBadgeBandVerticalOffsetLabel.setText(
                                    String.format(
                                            "%+.0fpx",
                                            graphicPersonBadgeBandVerticalOffsetSlider.getValue()));
                        }
                    },
                    graphicCommitted);
        }
        if (personBadgeDragAdjustCheckBox != null) {
            personBadgeDragAdjustCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                flushGraphicRebuildNow();
                                scheduleEquipmentGraphicPersist();
                            });
        }
        if (personBadgeWireShowCheckBox != null) {
            personBadgeWireShowCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                flushGraphicRebuildNow();
                                scheduleEquipmentGraphicPersist();
                            });
        }
        configurePersonBadgeWireToolbar(graphicCommitted);
    }

    private void configurePersonBadgeWireToolbar(Runnable graphicCommitted) {
        if (graphicPersonBadgeWireDashCombo != null) {
            graphicPersonBadgeWireDashCombo
                    .getItems()
                    .setAll(EquipmentGanttPersonBadgeWireDashStyle.values());
            graphicPersonBadgeWireDashCombo.setConverter(
                    new StringConverter<>() {
                        @Override
                        public String toString(EquipmentGanttPersonBadgeWireDashStyle o) {
                            return o != null ? o.labelJa() : "";
                        }

                        @Override
                        public EquipmentGanttPersonBadgeWireDashStyle fromString(String s) {
                            return EquipmentGanttPersonBadgeWireDashStyle.SOLID;
                        }
                    });
            graphicPersonBadgeWireDashCombo
                    .getSelectionModel()
                    .select(EquipmentGanttPersonBadgeWireDashStyle.SOLID);
            graphicPersonBadgeWireDashCombo
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener((o, a, b) -> graphicCommitted.run());
        }
        if (graphicPersonBadgeWireWidthSlider != null) {
            graphicPersonBadgeWireWidthSlider.setMax(WIRE_WIDTH_SLIDER_MAX);
            SliderCommittedChangeSupport.install(
                    graphicPersonBadgeWireWidthSlider,
                    () -> {
                        if (graphicPersonBadgeWireWidthLabel != null) {
                            graphicPersonBadgeWireWidthLabel.setText(
                                    formatPersonBadgeWireWidthLabel(
                                            graphicPersonBadgeWireWidthSlider.getValue()));
                        }
                    },
                    graphicCommitted);
        }
        if (graphicPersonBadgeWireMaxLengthSlider != null) {
            graphicPersonBadgeWireMaxLengthSlider.setMax(WIRE_MAX_LENGTH_SLIDER_MAX);
            SliderCommittedChangeSupport.install(
                    graphicPersonBadgeWireMaxLengthSlider,
                    () -> {
                        if (graphicPersonBadgeWireMaxLengthLabel != null) {
                            graphicPersonBadgeWireMaxLengthLabel.setText(
                                    formatPersonBadgeWireMaxLengthLabel(
                                            graphicPersonBadgeWireMaxLengthSlider.getValue()));
                        }
                    },
                    graphicCommitted);
        }
        if (graphicPersonBadgeWireThemeColorCheckBox != null) {
            graphicPersonBadgeWireThemeColorCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                updateWireColorPickerDisabledFromThemeToggle();
                                if (!suppressGraphicRebuild) {
                                    graphicCommitted.run();
                                }
                            });
        }
        if (graphicPersonBadgeWireColorPicker != null) {
            graphicPersonBadgeWireColorPicker
                    .valueProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (!suppressGraphicRebuild) {
                                    graphicCommitted.run();
                                }
                            });
        }
        updateWireColorPickerDisabledFromThemeToggle();
    }

    /** テーマ色モードのときはパレットを無効化（保存は空文字＝レンダラ側でテーマ色）。 */
    private void updateWireColorPickerDisabledFromThemeToggle() {
        if (graphicPersonBadgeWireColorPicker == null || graphicPersonBadgeWireThemeColorCheckBox == null) {
            return;
        }
        graphicPersonBadgeWireColorPicker.setDisable(
                graphicPersonBadgeWireThemeColorCheckBox.isSelected());
    }

    private static String formatPersonBadgeWireWidthLabel(double sliderValue) {
        if (!Double.isFinite(sliderValue) || sliderValue < 0.05) {
            return "自動";
        }
        return String.format("%.2fpx", sliderValue);
    }

    private static String formatPersonBadgeWireMaxLengthLabel(double sliderValue) {
        if (!Double.isFinite(sliderValue) || sliderValue < 0.5) {
            return "無制限";
        }
        return String.format("%.0fpx", sliderValue);
    }

    private void populateEquipmentGraphicBarFontComboItems() {
        if (equipmentGraphicBarFontCombo == null) {
            return;
        }
        List<String> families = new ArrayList<>(Font.getFamilies());
        Collections.sort(families);
        equipmentGraphicBarFontCombo.getItems().clear();
        equipmentGraphicBarFontCombo.getItems().add("");
        equipmentGraphicBarFontCombo.getItems().addAll(families);
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
        if (planJsonField != null) {
            String pp = s.equipmentGanttPlanJsonPath();
            if (pp != null && !pp.isBlank()) {
                planJsonField.setText(pp.strip());
            }
        }
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
        if (personBadgeShowCheckBox != null) {
            personBadgeShowCheckBox.setSelected(s.equipmentGanttPersonBadgeEnabled());
        }
        if (personBadgeWireShowCheckBox != null) {
            personBadgeWireShowCheckBox.setSelected(s.equipmentGanttPersonBadgeWireEnabled());
        }
        String wireHex = s.equipmentGanttPersonBadgeWireStrokeHex();
        boolean wireTheme = wireHex == null || wireHex.isBlank();
        if (graphicPersonBadgeWireThemeColorCheckBox != null) {
            graphicPersonBadgeWireThemeColorCheckBox.setSelected(wireTheme);
        }
        if (graphicPersonBadgeWireColorPicker != null) {
            if (!wireTheme) {
                try {
                    graphicPersonBadgeWireColorPicker.setValue(Color.web(wireHex.strip()));
                } catch (IllegalArgumentException | NullPointerException ignored) {
                    graphicPersonBadgeWireColorPicker.setValue(Color.web("#64748b"));
                }
            } else {
                graphicPersonBadgeWireColorPicker.setValue(Color.web("#64748b"));
            }
        }
        updateWireColorPickerDisabledFromThemeToggle();
        if (graphicPersonBadgeWireWidthSlider != null) {
            double ww = s.equipmentGanttPersonBadgeWireWidthPx();
            if (!Double.isFinite(ww) || ww <= 1e-6) {
                graphicPersonBadgeWireWidthSlider.setValue(0d);
            } else {
                graphicPersonBadgeWireWidthSlider.setValue(
                        Math.min(ww, WIRE_WIDTH_SLIDER_MAX));
            }
            if (graphicPersonBadgeWireWidthLabel != null) {
                graphicPersonBadgeWireWidthLabel.setText(
                        formatPersonBadgeWireWidthLabel(graphicPersonBadgeWireWidthSlider.getValue()));
            }
        }
        if (graphicPersonBadgeWireDashCombo != null) {
            EquipmentGanttPersonBadgeWireDashStyle d =
                    EquipmentGanttPersonBadgeWireDashStyle.fromStored(
                            s.equipmentGanttPersonBadgeWireDashStyleKey());
            graphicPersonBadgeWireDashCombo.getSelectionModel().select(d);
        }
        if (graphicPersonBadgeWireMaxLengthSlider != null) {
            double wMax = s.equipmentGanttPersonBadgeWireMaxLengthPx();
            if (!Double.isFinite(wMax) || wMax < 0) {
                graphicPersonBadgeWireMaxLengthSlider.setValue(0d);
            } else {
                graphicPersonBadgeWireMaxLengthSlider.setValue(
                        Math.min(wMax, WIRE_MAX_LENGTH_SLIDER_MAX));
            }
            if (graphicPersonBadgeWireMaxLengthLabel != null) {
                graphicPersonBadgeWireMaxLengthLabel.setText(
                        formatPersonBadgeWireMaxLengthLabel(
                                graphicPersonBadgeWireMaxLengthSlider.getValue()));
            }
        }
        double bgap = s.equipmentGanttPersonBadgeGapPx();
        if (graphicPersonBadgeGapSlider != null) {
            double mx = DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX;
            if (Double.isFinite(bgap) && bgap >= 0 && bgap <= mx) {
                graphicPersonBadgeGapSlider.setValue(bgap);
            } else {
                graphicPersonBadgeGapSlider.setValue(
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX);
            }
            if (graphicPersonBadgeGapLabel != null) {
                graphicPersonBadgeGapLabel.setText(
                        String.format("%.0fpx", graphicPersonBadgeGapSlider.getValue()));
            }
        }
        double bvOff = s.equipmentGanttPersonBadgeBandVerticalOffsetPx();
        if (graphicPersonBadgeBandVerticalOffsetSlider != null) {
            double mn = DesktopSessionState.MIN_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX;
            double mx = DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX;
            if (Double.isFinite(bvOff) && bvOff >= mn && bvOff <= mx) {
                graphicPersonBadgeBandVerticalOffsetSlider.setValue(bvOff);
            } else {
                graphicPersonBadgeBandVerticalOffsetSlider.setValue(
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX);
            }
            if (graphicPersonBadgeBandVerticalOffsetLabel != null) {
                graphicPersonBadgeBandVerticalOffsetLabel.setText(
                        String.format(
                                "%+.0fpx",
                                graphicPersonBadgeBandVerticalOffsetSlider.getValue()));
            }
        }
        equipmentGanttGraphicDataFingerprint =
                s.equipmentGanttGraphicDataFingerprint() != null
                        ? s.equipmentGanttGraphicDataFingerprint()
                        : "";
        equipmentGanttBadgeDragDeltas.clear();
        if (s.equipmentGanttBadgeDragDeltas() != null) {
            equipmentGanttBadgeDragDeltas.putAll(s.equipmentGanttBadgeDragDeltas());
        }
        if (personBadgeDragAdjustCheckBox != null) {
            personBadgeDragAdjustCheckBox.setSelected(s.equipmentGanttPersonBadgeDragAdjustEnabled());
        }
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

    boolean snapshotEquipmentGanttPersonBadgeEnabled() {
        return personBadgeShowCheckBox == null || personBadgeShowCheckBox.isSelected();
    }

    boolean snapshotEquipmentGanttPersonBadgeWireEnabled() {
        return personBadgeWireShowCheckBox == null
                || personBadgeWireShowCheckBox.isSelected();
    }

    String snapshotEquipmentGanttPersonBadgeWireStrokeHex() {
        if (graphicPersonBadgeWireThemeColorCheckBox != null
                && graphicPersonBadgeWireThemeColorCheckBox.isSelected()) {
            return "";
        }
        if (graphicPersonBadgeWireColorPicker == null) {
            return "";
        }
        Color c = graphicPersonBadgeWireColorPicker.getValue();
        return c != null ? formatWireColorToWebHex(c) : "";
    }

    /** ColorPicker の値をセッション用の #RRGGBB / #RRGGBBAA に変換する。 */
    private static String formatWireColorToWebHex(Color c) {
        int r = (int) Math.round(c.getRed() * 255);
        int g = (int) Math.round(c.getGreen() * 255);
        int b = (int) Math.round(c.getBlue() * 255);
        double op = c.getOpacity();
        if (op >= 1.0 - 1e-3) {
            return String.format("#%02x%02x%02x", r, g, b);
        }
        int a = (int) Math.round(op * 255);
        return String.format("#%02x%02x%02x%02x", r, g, b, a);
    }

    double snapshotEquipmentGanttPersonBadgeWireWidthPx() {
        if (graphicPersonBadgeWireWidthSlider == null) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX;
        }
        double v = graphicPersonBadgeWireWidthSlider.getValue();
        return !Double.isFinite(v) || v < 0.05
                ? DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX
                : Math.min(v, WIRE_WIDTH_SLIDER_MAX);
    }

    String snapshotEquipmentGanttPersonBadgeWireDashStyleKey() {
        if (graphicPersonBadgeWireDashCombo == null) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_DASH_STYLE_KEY;
        }
        EquipmentGanttPersonBadgeWireDashStyle sel =
                graphicPersonBadgeWireDashCombo.getSelectionModel().getSelectedItem();
        return sel != null
                ? sel.storedKey()
                : DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_DASH_STYLE_KEY;
    }

    double snapshotEquipmentGanttPersonBadgeWireMaxLengthPx() {
        if (graphicPersonBadgeWireMaxLengthSlider == null) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX;
        }
        double v = graphicPersonBadgeWireMaxLengthSlider.getValue();
        if (!Double.isFinite(v) || v < 0) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX;
        }
        return Math.min(v, WIRE_MAX_LENGTH_SLIDER_MAX);
    }

    double snapshotEquipmentGanttPersonBadgeGapPx() {
        return graphicPersonBadgeGapSlider != null
                ? graphicPersonBadgeGapSlider.getValue()
                : DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX;
    }

    double snapshotEquipmentGanttPersonBadgeBandVerticalOffsetPx() {
        return graphicPersonBadgeBandVerticalOffsetSlider != null
                ? graphicPersonBadgeBandVerticalOffsetSlider.getValue()
                : DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX;
    }

    String snapshotEquipmentGanttGraphicDataFingerprint() {
        return equipmentGanttGraphicDataFingerprint != null ? equipmentGanttGraphicDataFingerprint : "";
    }

    Map<String, EquipmentGanttBadgeDragDelta> snapshotEquipmentGanttBadgeDragDeltas() {
        return Map.copyOf(equipmentGanttBadgeDragDeltas);
    }

    boolean snapshotEquipmentGanttPersonBadgeDragAdjustEnabled() {
        return personBadgeDragAdjustCheckBox != null && personBadgeDragAdjustCheckBox.isSelected();
    }

    /**
     * アコーディオン「閲覧モード」ではドラッグを無効にする。チェックボックス ON でも閉じている間は操作しない。
     */
    private boolean effectivePersonBadgeDragAdjustEnabled() {
        if (personBadgeDragAdjustCheckBox == null || !personBadgeDragAdjustCheckBox.isSelected()) {
            return false;
        }
        return sourceTitledPane == null || sourceTitledPane.isExpanded();
    }

    String snapshotEquipmentGanttPlanJsonPath() {
        return planJsonField != null ? planJsonField.getText().strip() : "";
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
    private void onPersonBadgeShowAction() {
        scheduleEquipmentGraphicPersist();
        requestThrottledGraphicRebuild();
    }

    @FXML
    private void onSyncLatestButtonAction() {
        syncLatestPlanJsonFromOutputDirAndReload();
    }

    /**
     * 既定出力フォルダの最新 {@code 計画*.json} を選択してグラフィックを再構築する。「同期して最新を表示」と同一。
     * 段階2・配台試行（段階3）完了後にメインシェルから呼び、ディスク上の最新計画を確実に反映する。
     */
    void syncLatestPlanJsonFromOutputDirAndReload() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path plan = Stage2OutputNaming.newestPrimaryPlanJson(dir);
            if (plan != null) {
                planJsonField.setText(plan.toString());
            }
            if (plan == null) {
                statusLabel.setText("このフォルダに 計画*.json（または旧 production_plan_multi_day*.json）がありません: " + dir);
                if (shell != null) {
                    shell.appendLog(
                            "[equipment-gantt-graphic] ガント同期をスキップ: 出力フォルダに計画JSONがありません（"
                                    + dir
                                    + "）");
                }
                return;
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            if (shell != null) {
                shell.appendLog(
                        "[equipment-gantt-graphic] ガント同期エラー: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex));
            }
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

    /**
     * JSON パス欄が空または無効なとき、実行タブの段階2計画ブックに対応する .json、または既定出力の最新計画 JSON
     * でフィールドを埋める。
     */
    private void tryAutofillPlanJsonIfEmpty() {
        if (planJsonField == null) {
            return;
        }
        String cur = planJsonField.getText().strip();
        if (!cur.isEmpty() && Files.isRegularFile(Path.of(cur))) {
            return;
        }
        if (shell != null) {
            String stage2 = shell.mainRunStage2ProductionPlanPathOrEmpty();
            if (!stage2.isEmpty()) {
                tryAutofillJsonFromStage2Xlsx(stage2, "");
            }
        }
        String again = planJsonField.getText().strip();
        if (!again.isEmpty() && Files.isRegularFile(Path.of(again))) {
            return;
        }
        if (shell == null) {
            return;
        }
        try {
            java.util.Map<String, String> ui = shell.snapshotUiEnv();
            Path dir = AppPaths.defaultPlanningOutputDir(ui);
            Path newest = Stage2OutputNaming.newestPrimaryPlanJson(dir);
            if (newest != null && Files.isRegularFile(newest)) {
                planJsonField.setText(newest.toString());
            }
        } catch (Exception ignored) {
            // ignore
        }
    }

    private void reloadFromFields() {
        if (contentPane == null) {
            return;
        }
        reloadButton.setDisable(true);
        syncLatestButton.setDisable(true);
        try {
            tryAutofillPlanJsonIfEmpty();
            String ps = planJsonField != null ? planJsonField.getText().strip() : "";
            Path planPath = ps.isEmpty() ? null : Path.of(ps);
            if (planPath == null || !Files.isRegularFile(planPath)) {
                resetGraphicState(
                        "ファイルが指定されていないか、見つかりません。", false);
                statusLabel.setText("読み込み対象なし");
                return;
            }

            SheetLoad loaded = loadWorkbookSheetsForGraphic(planPath);
            Map<String, JsonTableIo.SheetTable> sheets = loaded.sheets();
            loadedContractBadgeRows = loaded.contractBadgeSlotRows();
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
            if (shell != null) {
                shell.appendLog(
                        "[equipment-gantt-graphic] ガントを更新: "
                                + planPath.getFileName()
                                + "（対象シート数="
                                + names.size()
                                + ", 表示="
                                + sheetUsed
                                + "）");
            }
            collapseSourceAccordionAfterSuccessfulLoad();
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
        if (loadedContractBadgeRows != null && DEFAULT_SHEET.equals(name)) {
            badgeRowsForCurrentGraphic = loadedContractBadgeRows;
        } else {
            badgeRowsForCurrentGraphic = null;
        }
        applyGraphicCenter(st);
    }

    /** 再読み／最新JSON検索が成功したあとアコーディオンを閉じて表示領域を広げる */
    private void collapseSourceAccordionAfterSuccessfulLoad() {
        if (sourceTitledPane != null) {
            boolean anim = sourceTitledPane.isAnimated();
            sourceTitledPane.setAnimated(false);
            sourceTitledPane.setExpanded(false);
            if (sourceAccordion != null) {
                sourceAccordion.setExpandedPane(null);
            }
            sourceTitledPane.setAnimated(anim);
        } else if (sourceAccordion != null) {
            sourceAccordion.setExpandedPane(null);
        }
    }

    /** パス未指定・読み込み失敗時は設定パネルを開く */
    private void expandSourceAccordionForAttention() {
        if (sourceAccordion != null && sourceTitledPane != null) {
            sourceAccordion.setExpandedPane(sourceTitledPane);
            sourceTitledPane.setExpanded(true);
        }
    }

    private void resetGraphicState(String placeholderMsg) {
        resetGraphicState(placeholderMsg, true);
    }

    /**
     * @param clearBadgeSessionData false のとき、セッション復元済みのフィンガープリント・バッジドラッグずれは消さない（パス未指定時など）。
     */
    private void resetGraphicState(String placeholderMsg, boolean clearBadgeSessionData) {
        lastGraphicSheet = null;
        loadedContractBadgeRows = null;
        badgeRowsForCurrentGraphic = null;
        if (clearBadgeSessionData) {
            equipmentGanttGraphicDataFingerprint = "";
            equipmentGanttBadgeDragDeltas.clear();
        }
        graphicRootWrapper = null;
        graphicWheelHookInstalled = false;
        pendingHorizontalZoomAnchor = null;
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder(placeholderMsg));
        }
        expandSourceAccordionForAttention();
    }

    private void applyGraphicCenter(JsonTableIo.SheetTable st) {
        if (contentPane == null || st == null) {
            return;
        }
        Path planPath =
                lastLoadedPlanPath != null && !lastLoadedPlanPath.isBlank()
                        ? Path.of(lastLoadedPlanPath)
                        : null;
        String skipReason = skipHeavyGraphicReason(planPath, st);
        if (skipReason != null) {
            if (shell != null) {
                shell.refreshEquipmentGanttObservedBadgeLabels(List.of());
            }
            if (graphicRootWrapper == null) {
                graphicRootWrapper = new BorderPane();
                contentPane.setCenter(graphicRootWrapper);
            }
            graphicRootWrapper.setCenter(
                    emptyPlaceholder(
                            "設備ガント（グラフィック）の Canvas 生成をスキップしました（メモリ負荷の抑制）。\n"
                                    + skipReason
                                    + "\n正しい master.xlsm と段階2出力を確認してください。"));
            installGraphicWheelZoomIfNeeded();
            return;
        }

        BorderPane oldGantt =
                graphicRootWrapper != null && graphicRootWrapper.getCenter() instanceof BorderPane ob
                        ? ob
                        : null;
        EquipmentGraphicGanttPane.EquipmentGanttScrollState scrollSnap =
                EquipmentGraphicGanttPane.snapshotScroll(oldGantt);
        EquipmentGraphicGanttPane.HorizontalZoomAnchor zoomAnchor = pendingHorizontalZoomAnchor;
        pendingHorizontalZoomAnchor = null;

        ObservableList<ObservableList<String>> rows = toObservableRows(st);

        String fpNow =
                EquipmentGraphicGanttPane.computeDataFingerprint(
                        st.columns(), rows, badgeRowsForCurrentGraphic);
        if (!fpNow.equals(equipmentGanttGraphicDataFingerprint)) {
            equipmentGanttBadgeDragDeltas.clear();
        }
        equipmentGanttGraphicDataFingerprint = fpNow;

        long buildT0 = System.nanoTime();
        BorderPane gantt =
                buildEquipmentGanttBorderPane(
                        st.columns(),
                        rows,
                        badgeRowsForCurrentGraphic,
                        effectivePersonBadgeDragAdjustEnabled());
        if (Boolean.getBoolean("pm.ai.gantt.profile")) {
            long ms = (System.nanoTime() - buildT0) / 1_000_000L;
            if (shell != null) {
                shell.appendLog("[gantt-profile] EquipmentGraphicGanttPane.build ms=" + ms);
            }
        }
        if (shell != null) {
            shell.refreshEquipmentGanttObservedBadgeLabels(distinctBadgeLabelsFromGrid(badgeRowsForCurrentGraphic));
        }
        if (graphicRootWrapper == null) {
            graphicRootWrapper = new BorderPane();
            contentPane.setCenter(graphicRootWrapper);
        }
        graphicRootWrapper.setCenter(gantt);
        EquipmentGraphicGanttPane.restoreScrollAfterRebuild(gantt, scrollSnap, zoomAnchor);
        installGraphicWheelZoomIfNeeded();
        Object ud = gantt.getUserData();
        if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h
                && h.scheduleViewportRepaint() != null) {
            Platform.runLater(h.scheduleViewportRepaint());
        }
    }

    /**
     * 設備ガント（グラフィック）の {@link EquipmentGraphicGanttPane#build} を、現在のツールバー設定で行う。
     *
     * @param interactiveDragBadges false のとき印刷用（ドラッグずれの保存を行わない）
     */
    private BorderPane buildEquipmentGanttBorderPane(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsSlice,
            boolean interactiveDragBadges) {
        double zoom = graphicZoomSlider != null ? graphicZoomSlider.getValue() / 100.0 : 1.0;
        double rowPct = graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d;
        double slotPct = graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d;
        double headerPct =
                graphicHeaderHeightSlider != null ? graphicHeaderHeightSlider.getValue() : 100d;
        double barFp =
                graphicBarFontPctSlider != null ? graphicBarFontPctSlider.getValue() : 100d;
        DesktopTheme theme =
                shell != null ? shell.currentDesktopTheme() : DesktopTheme.LIGHT;
        java.util.function.Function<String, PersonBadgeStyle> badgeResolver =
                shell != null
                        ? shell.personBadgeStyleResolverForGantt()
                        : (String __) -> PersonBadgeStyle.defaultStyle();
        boolean showBadges = snapshotEquipmentGanttPersonBadgeEnabled();

        boolean dragEffective =
                interactiveDragBadges && effectivePersonBadgeDragAdjustEnabled();
        java.util.function.BiConsumer<String, EquipmentGanttBadgeDragDelta> dragSink =
                dragEffective
                        ? (k, d) -> {
                            if (Math.abs(d.dx()) < 1e-6 && Math.abs(d.dy()) < 1e-6) {
                                equipmentGanttBadgeDragDeltas.remove(k);
                            } else {
                                equipmentGanttBadgeDragDeltas.put(k, d);
                            }
                            scheduleEquipmentGraphicPersist();
                        }
                        : null;
        java.util.Map<String, EquipmentGanttBadgeDragDelta> dragMap =
                dragEffective ? equipmentGanttBadgeDragDeltas : java.util.Map.of();

        return EquipmentGraphicGanttPane.build(
                columns,
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
                snapshotEquipmentGanttShiftWheelHScrollPercent(),
                badgeSlotRowsSlice,
                showBadges,
                badgeResolver,
                snapshotEquipmentGanttPersonBadgeGapPx(),
                snapshotEquipmentGanttPersonBadgeBandVerticalOffsetPx(),
                dragEffective,
                dragMap,
                dragSink,
                snapshotEquipmentGanttPersonBadgeWireStrokeHex(),
                snapshotEquipmentGanttPersonBadgeWireWidthPx(),
                snapshotEquipmentGanttPersonBadgeWireDashStyleKey(),
                snapshotEquipmentGanttPersonBadgeWireMaxLengthPx(),
                snapshotEquipmentGanttPersonBadgeWireEnabled());
    }

    @FXML
    private void onPrintGanttAction() {
        runEquipmentGanttPrintJob(false);
    }

    @FXML
    private void onExportGanttPdfAction() {
        runEquipmentGanttPrintJob(true);
    }

    private void runEquipmentGanttPrintJob(boolean preselectPdfPrinter) {
        if (lastGraphicSheet == null) {
            if (statusLabel != null) {
                statusLabel.setText("先に JSON を読み込んでください。");
            }
            return;
        }
        Stage stage = ownerStage != null ? ownerStage : (shell != null ? shell.getPrimaryStage() : null);
        if (stage == null) {
            return;
        }
        ObservableList<ObservableList<String>> fullRows = toObservableRows(lastGraphicSheet);
        List<String> cols = lastGraphicSheet.columns();
        List<List<Integer>> groups =
                EquipmentGanttPrintDaySlices.rowIndexGroupsOnePagePerDay(cols, fullRows);
        if (groups.isEmpty()) {
            if (statusLabel != null) {
                statusLabel.setText("印刷する行がありません。");
            }
            return;
        }
        int slotCols = EquipmentGraphicGanttPane.countTimeSlotHeadersInColumns(cols);
        if (slotCols <= 0) {
            if (statusLabel != null) {
                statusLabel.setText("時刻列（HH:MM）が無いため印刷できません。");
            }
            return;
        }

        PrinterJob job = PrinterJob.createPrinterJob();
        if (job == null) {
            if (statusLabel != null) {
                statusLabel.setText("印刷ジョブを作成できませんでした。");
            }
            return;
        }
        if (preselectPdfPrinter) {
            Printer pdf = findLikelyPdfPrinter();
            if (pdf != null) {
                job.setPrinter(pdf);
            } else {
                Alert info = new Alert(AlertType.INFORMATION);
                info.setTitle("PDF へ出力");
                info.setHeaderText(null);
                info.setContentText(
                        "一覧から「Microsoft Print to PDF」など PDF 用プリンターを選んでください。\n"
                                + "（環境によっては名前が異なります）");
                info.initOwner(stage);
                info.showAndWait();
            }
        }
        if (!job.showPrintDialog(stage)) {
            if (statusLabel != null) {
                statusLabel.setText("印刷をキャンセルしました。");
            }
            return;
        }
        Printer printer = job.getPrinter();
        if (printer == null) {
            if (statusLabel != null) {
                statusLabel.setText("プリンターが選択されていません。");
            }
            return;
        }
        PageLayout layout =
                printer.createPageLayout(
                        Paper.A3, PageOrientation.LANDSCAPE, Printer.MarginType.DEFAULT);

        int okPages = 0;
        try {
            for (List<Integer> idxGroup : groups) {
                ObservableList<ObservableList<String>> slice =
                        EquipmentGanttPrintDaySlices.sliceRowsByIndices(fullRows, idxGroup);
                List<List<String>> badgeSlice =
                        EquipmentGanttPrintDaySlices.sliceBadgeRowsAligned(
                                badgeRowsForCurrentGraphic, idxGroup, slotCols);
                BorderPane page =
                        buildEquipmentGanttBorderPane(cols, slice, badgeSlice, false);
                Parent printRoot =
                        EquipmentGanttPrintPageWrapper.fitGanttToSinglePrintablePage(page, layout);
                if (!job.printPage(layout, printRoot)) {
                    if (shell != null) {
                        shell.appendLog(
                                "[equipment-gantt-graphic] printPage が false を返しました（"
                                        + (okPages + 1)
                                        + " ページ目）");
                    }
                    break;
                }
                okPages++;
            }
        } catch (Exception ex) {
            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            if (statusLabel != null) {
                statusLabel.setText("印刷エラー: " + msg);
            }
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] print: " + msg);
            }
            return;
        } finally {
            job.endJob();
        }
        if (statusLabel != null) {
            statusLabel.setText(
                    "印刷ジョブを送信しました（A3 横向き・"
                            + okPages
                            + " ページ。各ページは暦日単位で分割し、用紙 1 枚に収まるよう縮小します）。");
        }
    }

    private static Printer findLikelyPdfPrinter() {
        Printer fallback = null;
        for (Printer p : Printer.getAllPrinters()) {
            if (p == null) {
                continue;
            }
            String n = p.getName();
            if (n == null) {
                continue;
            }
            String low = n.toLowerCase(Locale.ROOT);
            if (low.contains("pdf") || low.contains("pdfwriter")) {
                if (low.contains("microsoft") || low.contains("windows")) {
                    return p;
                }
                if (fallback == null) {
                    fallback = p;
                }
            }
        }
        return fallback;
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
                    double cur = graphicZoomSlider.getValue();
                    double delta = e.getDeltaY() > 0 ? 5 : -5;
                    double next = Math.clamp(cur + delta, 50, 200);
                    if (next == cur) {
                        return;
                    }
                    BorderPane oldGantt =
                            graphicRootWrapper.getCenter() instanceof BorderPane ob ? ob : null;
                    ScrollPane sp = null;
                    if (oldGantt != null
                            && oldGantt.getUserData()
                                    instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h) {
                        sp = h.timelineScroll();
                    }
                    pendingHorizontalZoomAnchor =
                            sp != null
                                    ? EquipmentGraphicGanttPane.computeHorizontalZoomAnchor(sp, e)
                                    : null;
                    /*
                     * 拡大率の数値・スライダーを即時反映し、重い build は次フレームへ遅延する
                     *（同一イベント内で Canvas 再構築するとツールバーの追従が遅く見える）。
                     */
                    suppressGraphicRebuild = true;
                    try {
                        graphicZoomSlider.setValue(next);
                        scheduleEquipmentGraphicPersist();
                    } finally {
                        suppressGraphicRebuild = false;
                    }
                    Platform.runLater(
                            () -> {
                                flushGraphicRebuildNow();
                            });
                });
    }

    private void rebuildGraphicView() {
        if (lastGraphicSheet == null || contentPane == null) {
            pendingHorizontalZoomAnchor = null;
            return;
        }
        applyGraphicCenter(lastGraphicSheet);
    }

    /** メインの {@link DesktopTheme} 変更時に Canvas 帯の配色を合わせて再描画する。 */
    void refreshGraphicForTheme() {
        flushGraphicRebuildNow();
    }

    /** 担当バッジデザイン変更時に設備ガントを即時再描画する。 */
    void refreshGraphicForPersonBadge() {
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

    /**
     * 計画 JSON パスから、設備ガント契約 JSON（…設.json 等）のパスを返す。計画パス自体が契約のときはそのまま。
     */
    private static Path resolveContractJsonForPlanPath(Path planJsonPath) {
        Path fn = planJsonPath.getFileName();
        if (fn == null) {
            return resolveEquipmentContractSibling(planJsonPath);
        }
        String n = fn.toString();
        if (!n.endsWith(".json")) {
            return resolveEquipmentContractSibling(planJsonPath);
        }
        String stem = n.substring(0, n.length() - 5);
        if (stem.endsWith("設") || stem.endsWith("_equipment_gantt_contract")) {
            return planJsonPath;
        }
        return resolveEquipmentContractSibling(planJsonPath);
    }

    /**
     * 契約 JSON の {@code kwargs_packed.timeline_events} が無い・空配列のとき true（配台イベントなし）。
     */
    private static boolean isContractTimelineEventsEmpty(Path contractJsonPath) {
        try {
            JsonNode root =
                    GANTT_CONTRACT_PEEK_OM.readTree(
                            Files.readString(contractJsonPath, StandardCharsets.UTF_8));
            JsonNode packed = root.get("kwargs_packed");
            if (packed == null || !packed.isObject()) {
                return false;
            }
            JsonNode te = packed.get("timeline_events");
            if (te == null || !te.isArray()) {
                return true;
            }
            return te.isEmpty();
        } catch (IOException | RuntimeException ignored) {
            return false;
        }
    }

    /**
     * 重いグラフィックビルドを省略すべき理由を返す（無ければ null）。
     * 契約の空 timeline_events、または大規模表でタイムライン非空セルが皆無のとき。
     */
    private static String skipHeavyGraphicReason(Path planJsonPath, JsonTableIo.SheetTable st) {
        if (planJsonPath != null && Files.isRegularFile(planJsonPath)) {
            Path contract = resolveContractJsonForPlanPath(planJsonPath);
            if (contract != null && Files.isRegularFile(contract)) {
                if (isContractTimelineEventsEmpty(contract)) {
                    return "設備ガント契約の kwargs_packed.timeline_events にイベントがありません（配台結果なしと同等）。";
                }
            }
        }
        if (st != null) {
            int n = st.rows().size();
            if (n >= HEAVY_EMPTY_GRID_MIN_ROWS) {
                int scan = Math.min(n, HEAVY_EMPTY_GRID_SCAN_MAX_ROWS);
                if (countNonEmptyTimelineSlotCells(st, scan) == 0) {
                    return "タイムライン（HH:MM）列に非空セルがありません（空の大表のため描画を省略しました）。";
                }
            }
        }
        return null;
    }

    /**
     * HH:MM 列の非空白セル数（先頭 {@code maxRows} 行まで）。契約マージ後の表にタイムライン文言があるかの簡易検証用。
     */
    private static int countNonEmptyTimelineSlotCells(JsonTableIo.SheetTable st, int maxRows) {
        List<String> cols = st.columns();
        List<String> slotHeaders = new ArrayList<>();
        for (String h : cols) {
            if (h != null && HH_MM_COLUMN_HEADER.matcher(h).matches()) {
                slotHeaders.add(h);
            }
        }
        if (slotHeaders.isEmpty()) {
            return 0;
        }
        int nonEmpty = 0;
        int rowsSeen = 0;
        for (Map<String, String> row : st.rows()) {
            if (rowsSeen >= maxRows) {
                break;
            }
            rowsSeen++;
            if (row == null) {
                continue;
            }
            for (String h : slotHeaders) {
                String v = row.get(h);
                if (v != null && !v.isBlank()) {
                    nonEmpty++;
                }
            }
        }
        return nonEmpty;
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

    private record SheetLoad(
            Map<String, JsonTableIo.SheetTable> sheets,
            String description,
            List<List<String>> contractBadgeSlotRows) {}

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
            if (stem0.endsWith("_equipment_gantt_contract") || stem0.endsWith("設")) {
                EquipmentGanttSheetBundle bundle =
                        EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(
                                planJsonFromField);
                Map<String, JsonTableIo.SheetTable> m = new LinkedHashMap<>();
                m.put(DEFAULT_SHEET, bundle.table());
                return new SheetLoad(
                        m, fn0.toString() + " (設備ガント契約・直接)", bundle.badgeSlotRows());
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
        List<List<String>> badgeRows = null;
        if (contract != null && Files.isRegularFile(contract)) {
            EquipmentGanttSheetBundle bundle =
                    EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);
            sheets.put(DEFAULT_SHEET, bundle.table());
            badgeRows = bundle.badgeSlotRows();
            desc = desc + " / " + contract.getFileName() + " (設備ガント帯)";
        }
        return new SheetLoad(sheets, desc, badgeRows);
    }

    /**
     * 計画ブック JSON のファイル名 stem（拡張子除く）から、サイドカー接尾辞・旧英語接尾辞を繰り返し除去する。
     */
    private static String stripStage2PlanJsonStemVariants(String stem) {
        String s = stem;
        while (true) {
            boolean changed = false;
            if (s.endsWith("_equipment_gantt_contract")) {
                s = s.substring(0, s.length() - "_equipment_gantt_contract".length());
                changed = true;
            } else if (s.endsWith("_logical_view")) {
                s = s.substring(0, s.length() - "_logical_view".length());
                changed = true;
            } else if (s.endsWith("_tabular_source")) {
                s = s.substring(0, s.length() - "_tabular_source".length());
                changed = true;
            } else if (s.endsWith("_actual_detail_gantt_contract")) {
                s = s.substring(0, s.length() - "_actual_detail_gantt_contract".length());
                changed = true;
            } else if (s.endsWith("_結果_タスク一覧")) {
                s = s.substring(0, s.length() - "_結果_タスク一覧".length());
                changed = true;
            } else if (s.endsWith("一覧")) {
                s = s.substring(0, s.length() - 2);
                changed = true;
            } else if (s.endsWith("表") || s.endsWith("論") || s.endsWith("設") || s.endsWith("実")) {
                s = s.substring(0, s.length() - 1);
                changed = true;
            }
            if (!changed) {
                break;
            }
        }
        return s;
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
        if (stem.endsWith("_logical_view") || stem.endsWith("論")) {
            return planJsonFromField;
        }
        String baseStem = stripStage2PlanJsonStemVariants(stem);
        Path sibling = planJsonFromField.resolveSibling(baseStem + "論.json");
        if (Files.isRegularFile(sibling)) {
            return sibling;
        }
        Path legacy = planJsonFromField.resolveSibling(baseStem + "_logical_view.json");
        return Files.isRegularFile(legacy) ? legacy : null;
    }

    /**
     * 計画ブック本体 JSON と並ぶ設備ガント契約 JSON（{@code …設.json}、または旧 {@code …_equipment_gantt_contract.json}）。
     * サイドカー JSON を開いているときは stem から接尾辞を除いて兄弟を解決する。
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
        if (stem.endsWith("_equipment_gantt_contract") || stem.endsWith("設")) {
            return null;
        }
        String baseStem = stripStage2PlanJsonStemVariants(stem);
        Path modern = planJsonFromField.resolveSibling(baseStem + "設.json");
        if (Files.isRegularFile(modern)) {
            return modern;
        }
        Path legacy = planJsonFromField.resolveSibling(baseStem + "_equipment_gantt_contract.json");
        return Files.isRegularFile(legacy) ? legacy : null;
    }

    /** バッジグリッドから表示キー（姓2文字等）を重複除去して列挙する。 */
    private static List<String> distinctBadgeLabelsFromGrid(List<List<String>> grid) {
        if (grid == null || grid.isEmpty()) {
            return List.of();
        }
        LinkedHashSet<String> out = new LinkedHashSet<>();
        for (List<String> row : grid) {
            if (row == null) {
                continue;
            }
            for (String cell : row) {
                if (cell == null || cell.isBlank()) {
                    continue;
                }
                for (String part : PersonNameBadgeText.splitBadgeCell(cell)) {
                    String k = PersonBadgeStyle.normalizeLabelKey(part);
                    if (!k.isEmpty()) {
                        out.add(k);
                    }
                }
            }
        }
        return List.copyOf(out);
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
}
