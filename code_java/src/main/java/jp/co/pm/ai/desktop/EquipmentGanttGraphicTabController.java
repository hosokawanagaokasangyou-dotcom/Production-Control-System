package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
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
import javafx.scene.input.ScrollEvent;
import javafx.scene.control.Accordion;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.Slider;
import javafx.scene.control.DatePicker;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.ToggleGroup;
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
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.Stage3UiVisibility;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage3Support;
import jp.co.pm.ai.desktop.config.EquipmentGanttBadgeDragDelta;
import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentBadgeGridUpdater;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentBarContext;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentDragPayload;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentDropTarget;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditActions;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditModel;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentInteraction;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentMetadata;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentOpSubCodec;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentSyncPython;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentSyncResult;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentPerson;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentRole;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttSheetBundle;
import jp.co.pm.ai.desktop.io.gantt.GanttContractValueDecoder;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintCompositor;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintDaySlices;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintPageSpec;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintTableData;
import jp.co.pm.ai.desktop.print.EquipmentGanttPdfExporter;
import jp.co.pm.ai.desktop.print.EquipmentGanttPrintTimelineColumnDensifier;
import jp.co.pm.ai.desktop.ui.EquipmentGanttAssignmentMemberPicker;
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
    private Label dataStageBadgeLabel;

    @FXML
    private Button syncLatestButton;

    @FXML
    private Button printGanttButton;

    @FXML
    private Button exportGanttPdfButton;

    @FXML
    private Button openGanttPdfButton;

    /** 直近に作成した設備ガント PDF（サマリ Excel と同一フォルダ）。 */
    private Path lastExportedGanttPdfPath;

    @FXML
    private RadioButton printTimeModeRegularRadio;

    @FXML
    private RadioButton printTimeModeRangeRadio;

    @FXML
    private ToggleGroup printTimeModeToggleGroup;

    @FXML
    private TextField printTimeRangeStartField;

    @FXML
    private TextField printTimeRangeEndField;

    @FXML
    private Label printTimeRegularHintLabel;

    /** 契約 JSON の {@code regular_shift_times}（定常 A15/B15 相当）。 */
    private LocalTime loadedRegularShiftStart;

    private LocalTime loadedRegularShiftEnd;

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
    private DatePicker ganttJumpDatePicker;

    @FXML
    private Button ganttJumpButton;

    /** 読込直後の DatePicker 初期化中はジャンプを抑止する。 */
    private boolean suppressGanttJumpDatePicker;

    private MainShellController shell;

    private Stage ownerStage;

    private String lastLoadedPlanPath = "";

    /** 再描画用に保持する最新の選択シート（ズーム・テーマ変更時）。 */
    private JsonTableIo.SheetTable lastGraphicSheet;

    /** 契約 JSON から得たバッジグリッド（{@link #DEFAULT_SHEET} 表示時のみ使用）。 */
    private List<List<String>> loadedContractBadgeRows;

    private EquipmentGanttAssignmentMetadata loadedAssignmentMetadata;

    private EquipmentGanttAssignmentEditModel assignmentEditModel;

    private boolean assignmentDirty;

    private final EquipmentGanttAssignmentEditActions assignmentEditActions =
            new EquipmentGanttAssignmentEditActions() {
                @Override
                public void onAddPersonRequested(String barId, double screenX, double screenY) {
                    promptAddAssignmentPerson(barId, screenX, screenY);
                }

                @Override
                public void onRemovePersonRequested(
                        String barId, String memberKey, double screenX, double screenY) {
                    removeAssignmentPerson(barId, memberKey);
                }
            };

    /** {@link #applyGraphicCenter} に渡す現在のバッジ行（シートに応じて null）。 */
    private List<List<String>> badgeRowsForCurrentGraphic;

    /** {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane#computeDataFingerprint} と同一キーでデータ同一判定に使う。 */
    private String equipmentGanttGraphicDataFingerprint = "";

    /**
     * 表データ＋表示パラメータの合成キー。直前と同一なら {@link EquipmentGraphicGanttPane#build} を省略する。
     */
    private String equipmentGanttGraphicRenderKey = "";

    /** データ同一時のみ有効なバッジドラッグずれ（セッションと同期）。 */
    private final Map<String, EquipmentGanttBadgeDragDelta> equipmentGanttBadgeDragDeltas =
            new LinkedHashMap<>();

    private BorderPane graphicRootWrapper;

    @FXML
    private CheckBox personBadgeShowCheckBox;

    @FXML
    private RadioButton personBadgeInteractionNoneRadio;

    @FXML
    private ToggleGroup personBadgeInteractionModeToggleGroup;

    @FXML
    private RadioButton personBadgeDragAdjustRadio;

    @FXML
    private RadioButton personBadgeAssignmentEditRadio;

    @FXML
    private Button undoAssignmentButton;

    @FXML
    private Button saveAssignmentButton;

    private boolean assignmentSaveInProgress;

    @FXML
    private CheckBox personBadgeWireShowCheckBox;

    @FXML
    private CheckBox prepTimeLabelsShowCheckBox;

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

    /** 担当割当編集と位置ドラッグの排他同期中（再入防止）。 */

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

    /** 起動時はタブが遅延プレースホルダのため {@link #flushGraphicRebuildNow} を後回しにする。 */
    private boolean pendingGraphicRebuildAfterSessionApply;

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
                                if (!b) {
                                    resetPersonBadgeInteractionToViewMode();
                                } else if (isPersonBadgeEditModeRadioSelected()) {
                                    flushGraphicRebuildNow();
                                }
                                scheduleEquipmentGraphicPersist();
                            });
        }
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder("JSON を指定して再読みしてください。"));
        }
        populateEquipmentGraphicBarFontComboItems();
        attachGraphicToolbarListeners();
        attachPrintTimeRangeListeners();
        attachGanttJumpDatePickerListeners();
        refreshPrintTimeRegularHint();
    }

    private void attachGanttJumpDatePickerListeners() {
        if (ganttJumpDatePicker != null) {
            ganttJumpDatePicker
                    .valueProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (suppressGanttJumpDatePicker || b == null) {
                                    return;
                                }
                                performGanttDateJump(b);
                            });
        }
    }

    @FXML
    private void onGanttJumpButtonAction() {
        if (ganttJumpDatePicker == null) {
            return;
        }
        LocalDate date = ganttJumpDatePicker.getValue();
        if (date != null) {
            performGanttDateJump(date);
        }
    }

    private void performGanttDateJump(LocalDate date) {
        BorderPane gantt = currentEquipmentGanttBorderPane();
        if (gantt == null) {
            setGanttJumpStatus("ガントが未読込のためジャンプできません。");
            return;
        }
        EquipmentGraphicGanttPane.GanttDateJumpResult result =
                EquipmentGraphicGanttPane.scrollToDate(gantt, date);
        switch (result) {
            case SUCCESS -> setGanttJumpStatus(
                    String.format("日付 %s の行へジャンプしました。", formatGanttJumpDate(date)));
            case NO_SUCH_DATE -> setGanttJumpStatus(
                    String.format("日付 %s はこのガントに含まれません。", formatGanttJumpDate(date)));
            case NO_GANTT -> setGanttJumpStatus("ガントが未読込のためジャンプできません。");
            case SCROLL_UNAVAILABLE -> setGanttJumpStatus("スクロール位置を変更できませんでした。");
            default -> setGanttJumpStatus("");
        }
    }

    private BorderPane currentEquipmentGanttBorderPane() {
        if (graphicRootWrapper == null || !(graphicRootWrapper.getCenter() instanceof BorderPane bp)) {
            return null;
        }
        return bp;
    }

    private void refreshGanttJumpDatePicker() {
        if (ganttJumpDatePicker == null) {
            return;
        }
        BorderPane gantt = currentEquipmentGanttBorderPane();
        List<LocalDate> dates =
                gantt != null
                        ? EquipmentGraphicGanttPane.listAvailableDates(gantt)
                        : List.of();
        boolean enabled = !dates.isEmpty();
        ganttJumpDatePicker.setDisable(!enabled);
        if (ganttJumpButton != null) {
            ganttJumpButton.setDisable(!enabled);
        }
        if (!enabled) {
            suppressGanttJumpDatePicker = true;
            try {
                ganttJumpDatePicker.setValue(null);
            } finally {
                suppressGanttJumpDatePicker = false;
            }
            return;
        }
        suppressGanttJumpDatePicker = true;
        try {
            LocalDate current = ganttJumpDatePicker.getValue();
            if (current == null || !dates.contains(current)) {
                ganttJumpDatePicker.setValue(dates.get(0));
            }
        } finally {
            suppressGanttJumpDatePicker = false;
        }
    }

    private static String formatGanttJumpDate(LocalDate date) {
        return date != null ? date.toString() : "—";
    }

    private void setGanttJumpStatus(String message) {
        if (statusLabel != null && message != null && !message.isBlank()) {
            statusLabel.setText(message);
        }
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
        if (personBadgeInteractionModeToggleGroup != null) {
            personBadgeInteractionModeToggleGroup
                    .selectedToggleProperty()
                    .addListener(
                            (o, a, b) -> {
                                flushGraphicRebuildNow();
                                scheduleEquipmentGraphicPersist();
                                refreshUndoAssignmentButtonState();
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
        if (prepTimeLabelsShowCheckBox != null) {
            prepTimeLabelsShowCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                flushGraphicRebuildNow();
                                scheduleEquipmentGraphicPersist();
                            });
        }
        configurePersonBadgeWireToolbar(graphicCommitted);
    }

    private void attachPrintTimeRangeListeners() {
        if (printTimeModeToggleGroup != null) {
            printTimeModeToggleGroup
                    .selectedToggleProperty()
                    .addListener((o, a, b) -> updatePrintTimeRangeFieldsEnabled());
        }
        updatePrintTimeRangeFieldsEnabled();
    }

    private void updatePrintTimeRangeFieldsEnabled() {
        boolean custom = printTimeModeRangeRadio != null && printTimeModeRangeRadio.isSelected();
        if (printTimeRangeStartField != null) {
            printTimeRangeStartField.setDisable(!custom);
        }
        if (printTimeRangeEndField != null) {
            printTimeRangeEndField.setDisable(!custom);
        }
        if (printTimeRegularHintLabel != null) {
            printTimeRegularHintLabel.setManaged(!custom);
            printTimeRegularHintLabel.setVisible(!custom);
        }
    }

    private void refreshPrintTimeRegularHint() {
        if (printTimeRegularHintLabel == null) {
            return;
        }
        if (loadedRegularShiftStart != null && loadedRegularShiftEnd != null) {
            printTimeRegularHintLabel.setText(
                    "（定常 "
                            + formatHm(loadedRegularShiftStart)
                            + "～"
                            + formatHm(loadedRegularShiftEnd)
                            + "）");
        } else {
            printTimeRegularHintLabel.setText("（定常時刻は契約 JSON に未設定）");
        }
    }

    private void applyRegularShiftTimesToPrintFields() {
        if (loadedRegularShiftStart == null || loadedRegularShiftEnd == null) {
            return;
        }
        if (printTimeRangeStartField != null) {
            printTimeRangeStartField.setText(formatHm(loadedRegularShiftStart));
        }
        if (printTimeRangeEndField != null) {
            printTimeRangeEndField.setText(formatHm(loadedRegularShiftEnd));
        }
    }

    private static String formatHm(LocalTime t) {
        return t != null ? String.format("%d:%02d", t.getHour(), t.getMinute()) : "";
    }

    private record ResolvedPrintTimeRange(LocalTime startInclusive, LocalTime endExclusive) {}

    /**
     * 印刷用タイムライン列の半開区間。解決できないときは {@code null}（呼び出し側でメッセージ表示）。
     */
    private ResolvedPrintTimeRange resolvePrintTimeRangeForJob() {
        if (printTimeModeRangeRadio != null && printTimeModeRangeRadio.isSelected()) {
            LocalTime start = parseHmField(printTimeRangeStartField);
            LocalTime end = parseHmField(printTimeRangeEndField);
            if (start == null || end == null) {
                return null;
            }
            if (!start.isBefore(end)) {
                return null;
            }
            return new ResolvedPrintTimeRange(start, end);
        }
        if (loadedRegularShiftStart == null || loadedRegularShiftEnd == null) {
            return null;
        }
        if (!loadedRegularShiftStart.isBefore(loadedRegularShiftEnd)) {
            return null;
        }
        return new ResolvedPrintTimeRange(loadedRegularShiftStart, loadedRegularShiftEnd);
    }

    private static LocalTime parseHmField(TextField field) {
        if (field == null) {
            return null;
        }
        return parseHmText(field.getText());
    }

    private static LocalTime parseHmText(String raw) {
        if (raw == null) {
            return null;
        }
        String t = raw.strip();
        var m = HH_MM_COLUMN_HEADER.matcher(t);
        if (!m.matches()) {
            return null;
        }
        try {
            return LocalTime.of(Integer.parseInt(m.group(1)), Integer.parseInt(m.group(2)));
        } catch (Exception e) {
            return null;
        }
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
            if (contentPane != null && contentPane.getScene() != null) {
                flushGraphicRebuildNow();
            } else {
                pendingGraphicRebuildAfterSessionApply = true;
            }
        }
    }

    /** メインシェルで設備ガントタブが実体化されたあと、セッション適用分の再構築を行う。 */
    void flushPendingGraphicRebuildAfterSessionApply() {
        if (!pendingGraphicRebuildAfterSessionApply) {
            return;
        }
        pendingGraphicRebuildAfterSessionApply = false;
        flushGraphicRebuildNow();
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
        if (prepTimeLabelsShowCheckBox != null) {
            prepTimeLabelsShowCheckBox.setSelected(s.equipmentGanttPrepTimeLabelsEnabled());
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
        if (personBadgeDragAdjustRadio != null) {
            if (s.equipmentGanttPersonBadgeDragAdjustEnabled()) {
                personBadgeDragAdjustRadio.setSelected(true);
            } else if (personBadgeInteractionNoneRadio != null) {
                personBadgeInteractionNoneRadio.setSelected(true);
            }
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

    boolean snapshotEquipmentGanttPrepTimeLabelsEnabled() {
        return prepTimeLabelsShowCheckBox != null && prepTimeLabelsShowCheckBox.isSelected();
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
        return personBadgeDragAdjustRadio != null && personBadgeDragAdjustRadio.isSelected();
    }

    private boolean isPersonBadgeEditModeRadioSelected() {
        return (personBadgeDragAdjustRadio != null && personBadgeDragAdjustRadio.isSelected())
                || (personBadgeAssignmentEditRadio != null
                        && personBadgeAssignmentEditRadio.isSelected());
    }

    /**
     * アコーディオン閉じ（閲覧モード）時に編集ラジオを通常表示へ戻す。
     * ラジオ選択のままだと再構築後に effective* が false となり操作不能になる。
     */
    private void resetPersonBadgeInteractionToViewMode() {
        if (!isPersonBadgeEditModeRadioSelected()) {
            return;
        }
        if (personBadgeInteractionNoneRadio != null) {
            personBadgeInteractionNoneRadio.setSelected(true);
            return;
        }
        flushGraphicRebuildNow();
    }

    /**
     * アコーディオン「閲覧モード」ではドラッグを無効にする。ラジオ ON でも閉じている間は操作しない。
     */
    private boolean effectivePersonBadgeDragAdjustEnabled() {
        if (personBadgeDragAdjustRadio == null || !personBadgeDragAdjustRadio.isSelected()) {
            return false;
        }
        return sourceTitledPane == null || sourceTitledPane.isExpanded();
    }

    /**
     * アコーディオン「閲覧モード」では担当割当編集を無効にする。位置調整と同様。
     */
    private boolean effectivePersonBadgeAssignmentEditEnabled() {
        if (personBadgeAssignmentEditRadio == null
                || !personBadgeAssignmentEditRadio.isSelected()) {
            return false;
        }
        if (loadedAssignmentMetadata == null || loadedAssignmentMetadata.barUnits().isEmpty()) {
            return false;
        }
        return sourceTitledPane == null || sourceTitledPane.isExpanded();
    }

    @FXML
    private void onUndoAssignmentChangesAction() {
        undoAssignmentChanges();
    }

    @FXML
    private void onSaveAssignmentChangesAction() {
        beginSaveAssignmentChanges();
    }

    private void beginSaveAssignmentChanges() {
        if (shell == null
                || assignmentSaveInProgress
                || !assignmentDirty
                || assignmentEditModel == null
                || loadedAssignmentMetadata == null) {
            return;
        }
        String planPath = lastLoadedPlanPath != null ? lastLoadedPlanPath.strip() : "";
        if (planPath.isEmpty()) {
            shell.showWarningDialog("担当割当の保存", "計画 JSON が読み込まれていません。");
            return;
        }
        Path planJson = Path.of(planPath);
        Path contract = resolveEquipmentContractSibling(planJson);
        if (contract == null || !Files.isRegularFile(contract)) {
            shell.showWarningDialog(
                    "担当割当の保存", "設備ガント契約 JSON が見つかりません。JSON を再読み込みしてください。");
            return;
        }
        Path updatesJson;
        try {
            updatesJson = writeAssignmentUpdatesJson();
        } catch (IOException ex) {
            shell.showErrorDialog("担当割当の保存", "更新内容の書き出しに失敗しました: " + ex.getMessage());
            return;
        }
        Path planXlsx = resolvePlanXlsxSibling(planJson);
        Map<String, String> ui = shell.snapshotUiEnv();
        Path pythonExe = StagePythonExecutable.resolve(ui);
        Path scriptDir = EquipmentGanttAssignmentSyncPython.defaultScriptDir(ui);
        assignmentSaveInProgress = true;
        refreshAssignmentActionButtonsState();
        shell.appendLog("[equipment-gantt-assignment] save: validate start");
        EquipmentGanttAssignmentSyncPython.Request validateReq =
                new EquipmentGanttAssignmentSyncPython.Request(
                        contract,
                        updatesJson,
                        planXlsx,
                        null,
                        null,
                        false,
                        true);
        EquipmentGanttAssignmentSyncPython.runAsync(
                        pythonExe, scriptDir, ui, validateReq, shell::appendLog)
                .whenComplete(
                        (preview, err) ->
                                Platform.runLater(
                                        () ->
                                                onAssignmentValidateComplete(
                                                        preview,
                                                        err,
                                                        contract,
                                                        updatesJson,
                                                        planXlsx,
                                                        pythonExe,
                                                        scriptDir,
                                                        ui)));
    }

    private void onAssignmentValidateComplete(
            EquipmentGanttAssignmentSyncResult preview,
            Throwable err,
            Path contract,
            Path updatesJson,
            Path planXlsx,
            Path pythonExe,
            Path scriptDir,
            Map<String, String> ui) {
        if (err != null) {
            assignmentSaveInProgress = false;
            refreshAssignmentActionButtonsState();
            shell.showErrorDialog(
                    "担当割当の保存",
                    err.getCause() != null ? err.getCause().getMessage() : err.getMessage());
            return;
        }
        if (preview == null) {
            assignmentSaveInProgress = false;
            refreshAssignmentActionButtonsState();
            shell.showErrorDialog("担当割当の保存", "検証応答が空です。");
            return;
        }
        if (!preview.ok() && preview.hasWarnings()) {
            String msg =
                    "次の警告があります。保存を続行しますか？\n\n"
                            + preview.formatIssuesForDialog();
            if (!confirmAssignmentSaveWithWarnings(msg)) {
                assignmentSaveInProgress = false;
                refreshAssignmentActionButtonsState();
                return;
            }
            applyAssignmentChangesAfterConfirm(
                    contract,
                    updatesJson,
                    planXlsx,
                    pythonExe,
                    scriptDir,
                    ui,
                    preview.timelineHash(),
                    preview.confirmToken());
            return;
        }
        if (!preview.ok()) {
            assignmentSaveInProgress = false;
            refreshAssignmentActionButtonsState();
            shell.showErrorDialog("担当割当の保存", preview.formatIssuesForDialog());
            return;
        }
        applyAssignmentChangesAfterConfirm(
                contract,
                updatesJson,
                planXlsx,
                pythonExe,
                scriptDir,
                ui,
                preview.timelineHash(),
                null);
    }

    private void applyAssignmentChangesAfterConfirm(
            Path contract,
            Path updatesJson,
            Path planXlsx,
            Path pythonExe,
            Path scriptDir,
            Map<String, String> ui,
            String expectedTimelineHash,
            String confirmToken) {
        boolean forceWarnings = confirmToken != null && !confirmToken.isBlank();
        EquipmentGanttAssignmentSyncPython.Request applyReq =
                new EquipmentGanttAssignmentSyncPython.Request(
                        contract,
                        updatesJson,
                        planXlsx,
                        expectedTimelineHash,
                        confirmToken,
                        forceWarnings,
                        false);
        EquipmentGanttAssignmentSyncPython.runAsync(
                        pythonExe, scriptDir, ui, applyReq, shell::appendLog)
                .whenComplete(
                        (applied, err) ->
                                Platform.runLater(
                                        () ->
                                                onAssignmentApplyComplete(applied, err)));
    }

    private void onAssignmentApplyComplete(
            EquipmentGanttAssignmentSyncResult applied, Throwable err) {
        assignmentSaveInProgress = false;
        refreshAssignmentActionButtonsState();
        if (err != null) {
            shell.showErrorDialog(
                    "担当割当の保存",
                    err.getCause() != null ? err.getCause().getMessage() : err.getMessage());
            return;
        }
        if (applied == null || !applied.ok()) {
            String msg =
                    applied != null ? applied.formatIssuesForDialog() : "保存応答が空です。";
            shell.showErrorDialog("担当割当の保存", msg);
            return;
        }
        shell.appendLog("[equipment-gantt-assignment] save: applied " + applied.detail());
        shell.showInformationDialog("担当割当の保存", "担当割当を保存しました。");
        shell.refreshOperatorCardAfterPipelineRun();
        reloadFromFields(false);
    }

    private boolean confirmAssignmentSaveWithWarnings(String message) {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        Stage owner = ownerStage != null ? ownerStage : shell.getPrimaryStage();
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.setTitle("担当割当の保存");
        alert.setHeaderText(null);
        alert.setContentText(message);
        return alert.showAndWait().filter(r -> r == ButtonType.OK).isPresent();
    }

    private Path writeAssignmentUpdatesJson() throws IOException {
        Map<Integer, EquipmentGanttAssignmentOpSubCodec.OpSubPair> updates =
                EquipmentGanttAssignmentOpSubCodec.eventUpdates(
                        loadedAssignmentMetadata,
                        assignmentEditModel.snapshotPersonsByBarId());
        ObjectMapper mapper = new ObjectMapper();
        ObjectNode root = mapper.createObjectNode();
        ObjectNode updatesNode = mapper.createObjectNode();
        for (Map.Entry<Integer, EquipmentGanttAssignmentOpSubCodec.OpSubPair> e :
                updates.entrySet()) {
            ObjectNode pair = mapper.createObjectNode();
            pair.put("op", e.getValue().op());
            pair.put("sub", e.getValue().sub());
            updatesNode.set(String.valueOf(e.getKey()), pair);
        }
        root.set("updates", updatesNode);
        Path tmp = Files.createTempFile("pm_ai_gantt_assignment_updates_", ".json");
        tmp.toFile().deleteOnExit();
        Files.writeString(
                tmp,
                mapper.writerWithDefaultPrettyPrinter().writeValueAsString(root),
                StandardCharsets.UTF_8);
        return tmp;
    }

    private static Path resolvePlanXlsxSibling(Path planJsonFromField) {
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
        String baseStem = stripStage2PlanJsonStemVariants(stem);
        Path xlsx = planJsonFromField.resolveSibling(baseStem + ".xlsx");
        return Files.isRegularFile(xlsx) ? xlsx : null;
    }

    private void undoAssignmentChanges() {
        resetAssignmentEditState();
        if (loadedContractBadgeRows != null) {
            badgeRowsForCurrentGraphic = deepCopyBadgeRows(loadedContractBadgeRows);
        }
        refreshAssignmentActionButtonsState();
        flushGraphicRebuildNow();
    }

    private void resetAssignmentEditState() {
        assignmentEditModel =
                loadedAssignmentMetadata != null
                                && !loadedAssignmentMetadata.barUnits().isEmpty()
                        ? new EquipmentGanttAssignmentEditModel(loadedAssignmentMetadata)
                        : null;
        assignmentDirty = false;
    }

    private void refreshAssignmentBadgeGrid() {
        if (badgeRowsForCurrentGraphic == null
                || assignmentEditModel == null
                || loadedAssignmentMetadata == null) {
            return;
        }
        EquipmentGanttAssignmentBadgeGridUpdater.applyToBadgeRows(
                badgeRowsForCurrentGraphic,
                loadedAssignmentMetadata,
                assignmentEditModel.snapshotPersonsByBarId());
    }

    private EquipmentGanttAssignmentInteraction buildAssignmentInteraction() {
        boolean effective = effectivePersonBadgeAssignmentEditEnabled();
        boolean modelPresent = assignmentEditModel != null;
        EquipmentGanttAssignmentInteraction built;
        if (!effective || !modelPresent) {
            built = EquipmentGanttAssignmentInteraction.disabled();
        } else {
            built =
                    new EquipmentGanttAssignmentInteraction(
                            true,
                            loadedAssignmentMetadata,
                            assignmentEditModel.snapshotPersonsByBarId(),
                            this::handleAssignmentDrop,
                            assignmentEditActions);
        }
        return built;
    }

    private void promptAddAssignmentPerson(String barId, double screenX, double screenY) {
        if (assignmentEditModel == null || !effectivePersonBadgeAssignmentEditEnabled()) {
            return;
        }
        if (shell == null) {
            return;
        }
        Optional<EquipmentGanttAssignmentBarContext.ProcessMachine> ctx =
                EquipmentGanttAssignmentBarContext.resolve(
                        loadedAssignmentMetadata, lastGraphicSheet, barId);
        if (ctx.isEmpty()) {
            shell.showWarningDialog("担当を追加", "対象バーの工程・機械名を解決できません。");
            return;
        }
        Stage owner = ownerStage != null ? ownerStage : shell.getPrimaryStage();
        EquipmentGanttAssignmentMemberPicker.pickSingleMemberAsync(
                shell,
                owner,
                ctx.get().processName(),
                ctx.get().machineName(),
                screenX,
                screenY,
                fullName -> {
                    EquipmentGanttAssignmentPerson person =
                            EquipmentGanttAssignmentPerson.fromRawName(
                                    fullName, EquipmentGanttAssignmentRole.SUB);
                    applyAssignmentMutation(assignmentEditModel.addPerson(barId, person));
                });
    }

    private void removeAssignmentPerson(String barId, String memberKey) {
        if (assignmentEditModel == null || !effectivePersonBadgeAssignmentEditEnabled()) {
            return;
        }
        applyAssignmentMutation(assignmentEditModel.removePerson(barId, memberKey));
    }

    private boolean applyAssignmentMutation(
            Optional<EquipmentGanttAssignmentEditModel.Failure> failure) {
        if (failure.isPresent()) {
            EquipmentGanttAssignmentEditModel.Failure f = failure.get();
            if (f != EquipmentGanttAssignmentEditModel.Failure.SAME_BAR_NOOP && shell != null) {
                shell.appendLog("[equipment-gantt-assignment] rejected: " + f);
                String message = assignmentFailureMessage(f);
                if (!message.isBlank()) {
                    shell.showWarningDialog("担当割当", message);
                }
            }
            return false;
        }
        assignmentDirty = true;
        refreshAssignmentBadgeGrid();
        refreshAssignmentActionButtonsState();
        flushGraphicRebuildNow();
        return true;
    }

    private void refreshAssignmentActionButtonsState() {
        refreshUndoAssignmentButtonState();
        refreshSaveAssignmentButtonState();
    }

    private static String assignmentFailureMessage(
            EquipmentGanttAssignmentEditModel.Failure failure) {
        return switch (failure) {
            case UNKNOWN_BAR -> "対象バーが見つかりません。";
            case UNKNOWN_PERSON -> "対象担当者が見つかりません。";
            case DUPLICATE_PERSON -> "同じ担当者は既に割り当て済みです。";
            case EMPTY_BAR_FORBIDDEN -> "バーを0名にすることはできません。";
            case SAME_BAR_NOOP -> "";
        };
    }

    private boolean handleAssignmentDrop(
            EquipmentGanttAssignmentDragPayload source,
            EquipmentGanttAssignmentDropTarget target) {
        if (assignmentEditModel == null || !effectivePersonBadgeAssignmentEditEnabled()) {
            return false;
        }
        if (source == null || target == null) {
            return false;
        }
        Optional<EquipmentGanttAssignmentEditModel.Failure> failure;
        if (target.memberKey() == null || target.memberKey().isBlank()) {
            failure =
                    assignmentEditModel.movePerson(
                            source.barId(), target.barId(), source.memberKey());
        } else {
            failure =
                    assignmentEditModel.swapPerson(
                            source.barId(),
                            target.barId(),
                            source.memberKey(),
                            target.memberKey());
        }
        return applyAssignmentMutation(failure);
    }

    private void refreshUndoAssignmentButtonState() {
        if (undoAssignmentButton != null) {
            undoAssignmentButton.setDisable(!assignmentDirty || assignmentSaveInProgress);
        }
    }

    private void refreshSaveAssignmentButtonState() {
        if (saveAssignmentButton != null) {
            saveAssignmentButton.setDisable(
                    !assignmentDirty || assignmentSaveInProgress || assignmentEditModel == null);
        }
    }

    /** 契約メタデータが無いときは担当割当編集ラジオを選べないようにする。 */
    private void refreshAssignmentEditRadioState() {
        if (personBadgeAssignmentEditRadio == null) {
            return;
        }
        boolean hasMeta =
                loadedAssignmentMetadata != null
                        && !loadedAssignmentMetadata.barUnits().isEmpty();
        personBadgeAssignmentEditRadio.setDisable(!hasMeta);
        if (!hasMeta
                && personBadgeAssignmentEditRadio.isSelected()
                && personBadgeInteractionNoneRadio != null) {
            personBadgeInteractionNoneRadio.setSelected(true);
        }
    }

    private static List<List<String>> deepCopyBadgeRows(List<List<String>> src) {
        if (src == null) {
            return null;
        }
        List<List<String>> out = new ArrayList<>(src.size());
        for (List<String> row : src) {
            out.add(row != null ? new ArrayList<>(row) : new ArrayList<>());
        }
        return out;
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
        refreshExistingGanttPdfPathFromEnv();
        Platform.runLater(() -> reloadFromFields(false));
    }

    void applyStage3UiVisibility(boolean visible) {
        refreshPlanningStageBadgeFromDispatchJson();
        if (!visible
                && dataStageBadgeLabel != null
                && dataStageBadgeLabel.getText() != null
                && dataStageBadgeLabel.getText().startsWith("段階3")) {
            Stage3UiVisibility.apply(dataStageBadgeLabel, false);
        }
    }

    private void refreshExistingGanttPdfPathFromEnv() {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path existing = AppPaths.equipmentGanttPdfPath(ui);
        if (Files.isRegularFile(existing)) {
            lastExportedGanttPdfPath = existing;
        }
        refreshOpenGanttPdfButtonState();
    }

    private void refreshOpenGanttPdfButtonState() {
        if (openGanttPdfButton == null) {
            return;
        }
        boolean ok =
                lastExportedGanttPdfPath != null
                        && Files.isRegularFile(lastExportedGanttPdfPath);
        openGanttPdfButton.setDisable(!ok);
    }

    /**
     * 段階2.1 正本反映後: 指定した計画 JSON でガントを再読込する。無効なら出力フォルダ最新にフォールバック。
     */
    void syncPlanJsonPathAndReload(Path planJson, boolean userCompletionDialog) {
        if (shell == null) {
            return;
        }
        try {
            if (planJson != null && Files.isRegularFile(planJson) && planJsonField != null) {
                planJsonField.setText(planJson.toAbsolutePath().normalize().toString());
                reloadFromFields(userCompletionDialog);
                return;
            }
        } catch (Exception ex) {
            if (shell != null) {
                shell.appendLog(
                        "[equipment-gantt-graphic] 段階2.1 計画 JSON 同期エラー: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex));
            }
        }
        syncLatestPlanJsonFromOutputDirAndReload(userCompletionDialog);
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
            Platform.runLater(() -> reloadFromFields(false));
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromFields(true);
    }

    @FXML
    private void onPersonBadgeShowAction() {
        scheduleEquipmentGraphicPersist();
        requestThrottledGraphicRebuild();
    }

    @FXML
    private void onSyncLatestButtonAction() {
        syncLatestPlanJsonFromOutputDirAndReload(true);
    }

    /**
     * 既定出力フォルダの最新 {@code 計画*.json} を選択してグラフィックを再構築する。「同期して最新を表示」と同一。
     * 段階2・配台試行（段階3）完了後にメインシェルから呼び、ディスク上の最新計画を確実に反映する。
     */
    void syncLatestPlanJsonFromOutputDirAndReload(boolean userCompletionDialog) {
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
                if (userCompletionDialog) {
                    shell.showWarningDialog(
                            "同期",
                            "出力フォルダに計画 JSON がありません。\n" + dir);
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
            if (userCompletionDialog) {
                shell.showErrorDialog(
                        "同期エラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
            return;
        }
        reloadFromFields(userCompletionDialog);
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
            reloadFromFields(true);
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

    private void reloadFromFields(boolean userCompletionDialog) {
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
                if (userCompletionDialog && shell != null) {
                    shell.showWarningDialog(
                            "再読み",
                            "計画 JSON が指定されていないか、見つかりません。\n"
                                    + (planPath != null ? planPath : "（パス空）"));
                }
                return;
            }

            SheetLoad loaded = loadWorkbookSheetsForGraphic(planPath);
            Map<String, JsonTableIo.SheetTable> sheets = loaded.sheets();
            loadedContractBadgeRows = loaded.contractBadgeSlotRows();
            loadedAssignmentMetadata = loaded.assignmentMetadata();
            resetAssignmentEditState();
            refreshAssignmentActionButtonsState();
            refreshAssignmentEditRadioState();
            loadRegularShiftTimesFromPlan(planPath);
            lastLoadedPlanPath = planPath.toString();

            Map<String, JsonTableIo.SheetTable> eligible = filterEquipmentTimelineSheets(sheets);
            if (eligible.isEmpty()) {
                resetGraphicState(
                        "設備タイムライン形式のシートが見つかりません（時刻列 HH:MM のシート）。");
                statusLabel.setText("対象シートなし: " + planPath.getFileName());
                if (userCompletionDialog && shell != null) {
                    shell.showWarningDialog(
                            "再読み",
                            "設備タイムライン形式のシートがありません。\n" + planPath);
                }
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
            refreshPlanningStageBadgeFromDispatchJson();
            if (userCompletionDialog && shell != null) {
                shell.showInformationDialog(
                        "再読み完了",
                        "設備ガント（グラフィック）を更新しました。\n"
                                + planPath
                                + "\n表示シート: "
                                + sheetUsed);
            }
        } catch (Exception ex) {
            resetGraphicState("エラー");
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] " + ex.getMessage());
            }
            if (userCompletionDialog && shell != null) {
                shell.showErrorDialog(
                        "再読みエラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        } finally {
            refreshPlanningStageBadgeFromDispatchJson();
            reloadButton.setDisable(false);
            syncLatestButton.setDisable(false);
        }
    }

    private void refreshPlanningStageBadgeFromDispatchJson() {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(ui);
        ResultDispatchStage3Support.applyPlanningStageBadgeFromDispatchJson(
                dataStageBadgeLabel, jsonPath);
        Stage3UiVisibility.applyPlanningStageBadgePolicy(dataStageBadgeLabel, ui);
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
            badgeRowsForCurrentGraphic = deepCopyBadgeRows(loadedContractBadgeRows);
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
        loadedAssignmentMetadata = null;
        assignmentEditModel = null;
        assignmentDirty = false;
        refreshAssignmentActionButtonsState();
        refreshAssignmentEditRadioState();
        loadedRegularShiftStart = null;
        loadedRegularShiftEnd = null;
        refreshPrintTimeRegularHint();
        badgeRowsForCurrentGraphic = null;
        equipmentGanttGraphicRenderKey = "";
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
        refreshGanttJumpDatePicker();
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
            equipmentGanttGraphicRenderKey = "";
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
            refreshGanttJumpDatePicker();
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

        String renderKeyNow = computeGraphicRenderKey(st.columns(), rows, fpNow);
        if (oldGantt != null && renderKeyNow.equals(equipmentGanttGraphicRenderKey)) {
            EquipmentGraphicGanttPane.restoreScrollAfterRebuild(oldGantt, scrollSnap, zoomAnchor);
            installGraphicWheelZoomIfNeeded();
            Object ud = oldGantt.getUserData();
            if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h
                    && h.scheduleViewportRepaint() != null) {
                Platform.runLater(h.scheduleViewportRepaint());
            }
            refreshGanttJumpDatePicker();
            return;
        }
        equipmentGanttGraphicRenderKey = renderKeyNow;

        long buildT0 = System.nanoTime();
        BorderPane gantt =
                buildEquipmentGanttBorderPane(
                        st.columns(),
                        rows,
                        badgeRowsForCurrentGraphic,
                        true);
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
        refreshGanttJumpDatePicker();
    }

    /**
     * 設備ガント（グラフィック）の {@link EquipmentGraphicGanttPane#build} を、現在のツールバー設定で行う。
     *
     * @param interactiveDragBadges false のとき印刷用（ドラッグずれマップを渡さない）。画面表示は true。
     */
    private BorderPane buildEquipmentGanttBorderPane(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsSlice,
            boolean interactiveDragBadges) {
        return buildEquipmentGanttBorderPane(
                columns, rows, badgeSlotRowsSlice, interactiveDragBadges, null, false);
    }

    /**
     * @param slotWidthPercentOverride 非 null のときツールバー「時刻列幅」％の代わりに使用（印刷の用紙幅合わせ）
     * @param highQualityPrint 印刷向け高解像度 Canvas（画面表示では false）
     */
    private BorderPane buildEquipmentGanttBorderPane(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsSlice,
            boolean interactiveDragBadges,
            Double slotWidthPercentOverride,
            boolean highQualityPrint) {
        double zoom = graphicZoomSlider != null ? graphicZoomSlider.getValue() / 100.0 : 1.0;
        double rowPct = graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d;
        double slotPct =
                slotWidthPercentOverride != null
                        ? slotWidthPercentOverride
                        : (graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d);
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

        boolean dragAdjustActive =
                interactiveDragBadges && effectivePersonBadgeDragAdjustEnabled();
        java.util.function.BiConsumer<String, EquipmentGanttBadgeDragDelta> dragSink =
                dragAdjustActive
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
                interactiveDragBadges ? equipmentGanttBadgeDragDeltas : java.util.Map.of();

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
                dragAdjustActive,
                dragMap,
                dragSink,
                snapshotEquipmentGanttPersonBadgeWireStrokeHex(),
                snapshotEquipmentGanttPersonBadgeWireWidthPx(),
                snapshotEquipmentGanttPersonBadgeWireDashStyleKey(),
                snapshotEquipmentGanttPersonBadgeWireMaxLengthPx(),
                snapshotEquipmentGanttPersonBadgeWireEnabled(),
                snapshotEquipmentGanttPrepTimeLabelsEnabled(),
                highQualityPrint,
                buildAssignmentInteraction());
    }

    @FXML
    private void onPrintGanttAction() {
        runEquipmentGanttPrintJob();
    }

    @FXML
    private void onExportGanttPdfAction() {
        runEquipmentGanttPdfExport();
    }

    @FXML
    private void onOpenGanttPdfAction() {
        if (lastExportedGanttPdfPath == null || !Files.isRegularFile(lastExportedGanttPdfPath)) {
            if (statusLabel != null) {
                statusLabel.setText("PDF が未作成です。先に PDF を作成してください。");
            }
            refreshOpenGanttPdfButtonState();
            return;
        }
        try {
            DesktopFileOpener.openFile(lastExportedGanttPdfPath);
            if (statusLabel != null) {
                statusLabel.setText("PDF を開きました: " + lastExportedGanttPdfPath);
            }
        } catch (IOException ex) {
            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            if (statusLabel != null) {
                statusLabel.setText("PDF を開けません: " + msg);
            }
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] open pdf: " + msg);
            }
        }
    }

    private record PreparedGanttPrintJob(
            List<String> printCols,
            ObservableList<ObservableList<String>> printRows,
            EquipmentGanttPrintTableData printTable,
            List<List<Integer>> groups,
            int slotCols,
            ResolvedPrintTimeRange printRange) {}

    private PreparedGanttPrintJob prepareGanttPrintJob() {
        if (lastGraphicSheet == null) {
            if (statusLabel != null) {
                statusLabel.setText("先に JSON を読み込んでください。");
            }
            return null;
        }
        ResolvedPrintTimeRange printRange = resolvePrintTimeRangeForJob();
        if (printRange == null) {
            if (statusLabel != null) {
                if (printTimeModeRangeRadio != null && printTimeModeRangeRadio.isSelected()) {
                    statusLabel.setText(
                            "印刷の時刻範囲を HH:MM 形式で指定してください（開始＜終了）。");
                } else {
                    statusLabel.setText(
                            "定常時刻が契約 JSON に無いか不正です。範囲指定に切り替えるか、段階2を再実行してください。");
                }
            }
            return null;
        }
        ObservableList<ObservableList<String>> fullRows = toObservableRows(lastGraphicSheet);
        List<String> cols = lastGraphicSheet.columns();
        EquipmentGanttPrintTableData printTable =
                EquipmentGanttPrintTimelineColumnDensifier.densify(
                        cols,
                        fullRows,
                        badgeRowsForCurrentGraphic,
                        printRange.startInclusive(),
                        printRange.endExclusive());
        List<String> printCols = printTable.columns();
        ObservableList<ObservableList<String>> printRows = printTable.rows();
        List<List<Integer>> groups =
                EquipmentGanttPrintDaySlices.rowIndexGroupsOnePagePerDay(printCols, printRows);
        if (groups.isEmpty()) {
            if (statusLabel != null) {
                statusLabel.setText("印刷する行がありません。");
            }
            return null;
        }
        int slotCols = EquipmentGraphicGanttPane.countTimeSlotHeadersInColumns(printCols);
        if (slotCols <= 0) {
            if (statusLabel != null) {
                statusLabel.setText(
                        "指定した時刻範囲（"
                                + formatHm(printRange.startInclusive())
                                + "～"
                                + formatHm(printRange.endExclusive())
                                + "）に該当する時刻列がありません。");
            }
            return null;
        }
        return new PreparedGanttPrintJob(
                printCols, printRows, printTable, groups, slotCols, printRange);
    }

    private List<Parent> composeGanttPrintPages(PreparedGanttPrintJob prepared, PageLayout layout) {
        double paperW = layout.getPrintableWidth();
        double paperH = layout.getPrintableHeight();
        List<Parent> pages = new ArrayList<>();
        for (List<Integer> idxGroup : prepared.groups()) {
            ObservableList<ObservableList<String>> slice =
                    EquipmentGanttPrintDaySlices.sliceRowsByIndices(
                            prepared.printRows(), idxGroup);
            List<List<String>> badgeSlice =
                    EquipmentGanttPrintDaySlices.sliceBadgeRowsAligned(
                            prepared.printTable().badgeSlotRows(),
                            idxGroup,
                            prepared.slotCols());
            EquipmentGanttPrintPageSpec daySpec =
                    equipmentGanttPrintPageSpec(
                            prepared.printCols(),
                            slice,
                            badgeSlice,
                            prepared.printRange());
            List<List<Integer>> localChunks =
                    EquipmentGraphicGanttPane.splitLocalRowIndicesToFitPaper(
                            daySpec, paperW, paperH);
            for (List<Integer> localIdx : localChunks) {
                ObservableList<ObservableList<String>> pageRows =
                        EquipmentGanttPrintDaySlices.sliceRowsByIndices(slice, localIdx);
                List<List<String>> pageBadges =
                        EquipmentGanttPrintDaySlices.sliceBadgeRowsAligned(
                                badgeSlice, localIdx, prepared.slotCols());
                EquipmentGanttPrintPageSpec printSpec =
                        equipmentGanttPrintPageSpec(
                                prepared.printCols(),
                                pageRows,
                                pageBadges,
                                prepared.printRange());
                pages.add(EquipmentGanttPrintCompositor.composePage(printSpec, layout));
            }
        }
        return pages;
    }

    private void runEquipmentGanttPrintJob() {
        PreparedGanttPrintJob prepared = prepareGanttPrintJob();
        if (prepared == null) {
            return;
        }
        Stage stage = ownerStage != null ? ownerStage : (shell != null ? shell.getPrimaryStage() : null);
        if (stage == null) {
            return;
        }
        PrinterJob job = PrinterJob.createPrinterJob();
        if (job == null) {
            if (statusLabel != null) {
                statusLabel.setText("印刷ジョブを作成できませんでした。");
            }
            return;
        }
        applyDefaultEquipmentGanttPrintPageLayout(job);
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
        PageLayout layout = job.getJobSettings().getPageLayout();
        if (layout == null) {
            layout = defaultEquipmentGanttPrintPageLayout(printer);
            job.getJobSettings().setPageLayout(layout);
        }
        List<Parent> pages = composeGanttPrintPages(prepared, layout);
        int okPages = 0;
        try {
            for (Parent printRoot : pages) {
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
            job.cancelJob();
            return;
        }
        if (okPages > 0) {
            job.endJob();
        } else {
            job.cancelJob();
            if (statusLabel != null) {
                statusLabel.setText("印刷ページを生成できませんでした。");
            }
            return;
        }
        if (statusLabel != null) {
            statusLabel.setText(
                    "印刷ジョブを送信しました（"
                            + okPages
                            + " ページ・時刻 "
                            + formatHm(prepared.printRange().startInclusive())
                            + "～"
                            + formatHm(prepared.printRange().endExclusive())
                            + "・A3 横）。");
        }
    }

    private void runEquipmentGanttPdfExport() {
        if (shell != null && shell.isPlanningPipelineStageRunning()) {
            if (statusLabel != null) {
                statusLabel.setText("段階1～段階3.5 実行中のため PDF 作成を一時停止しています。");
            }
            shell.appendLog(
                    "[equipment-gantt-graphic] pdf export suppressed: planning pipeline stage running");
            return;
        }
        PreparedGanttPrintJob prepared = prepareGanttPrintJob();
        if (prepared == null) {
            return;
        }
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path outputPath = EquipmentGanttPdfExporter.resolveOutputPath(ui);
        PageLayout layout;
        try {
            layout = EquipmentGanttPdfExporter.defaultPageLayout();
        } catch (IOException ex) {
            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            if (statusLabel != null) {
                statusLabel.setText("PDF ページサイズの決定に失敗: " + msg);
            }
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] pdf layout: " + msg);
            }
            return;
        }
        List<Parent> pages = composeGanttPrintPages(prepared, layout);
        try {
            EquipmentGanttPdfExporter.export(outputPath, layout, pages);
        } catch (Exception ex) {
            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            if (statusLabel != null) {
                statusLabel.setText("PDF 出力エラー: " + msg);
            }
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] pdf export: " + msg);
            }
            return;
        }
        lastExportedGanttPdfPath = outputPath;
        refreshOpenGanttPdfButtonState();
        try {
            DesktopFileOpener.openFile(outputPath);
        } catch (IOException ex) {
            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
            if (statusLabel != null) {
                statusLabel.setText(
                        "PDF を出力しましたが自動で開けません: "
                                + msg
                                + " （"
                                + outputPath
                                + "）");
            }
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] pdf auto-open: " + msg);
            }
            return;
        }
        if (statusLabel != null) {
            statusLabel.setText(
                    "PDF を出力しました（"
                            + pages.size()
                            + " ページ・"
                            + outputPath.getFileName()
                            + "）。自動で開きました。");
        }
        if (shell != null) {
            shell.appendLog("[equipment-gantt-graphic] pdf: " + outputPath);
        }
    }

    /** 設備ガント印刷の既定：A3 横向き・余白はプリンター最小。 */
    private static void applyDefaultEquipmentGanttPrintPageLayout(PrinterJob job) {
        if (job == null) {
            return;
        }
        Printer printer = job.getPrinter();
        if (printer == null) {
            printer = Printer.getDefaultPrinter();
        }
        if (printer == null) {
            return;
        }
        job.getJobSettings().setPageLayout(defaultEquipmentGanttPrintPageLayout(printer));
    }

    private static PageLayout defaultEquipmentGanttPrintPageLayout(Printer printer) {
        return printer.createPageLayout(
                Paper.A3,
                PageOrientation.LANDSCAPE,
                Printer.MarginType.HARDWARE_MINIMUM);
    }

  /** {@link EquipmentGanttPrintCompositor} 向けにツールバー設定を束ねた印刷仕様。 */
    private EquipmentGanttPrintPageSpec equipmentGanttPrintPageSpec(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsSlice,
            ResolvedPrintTimeRange printRange) {
        java.util.function.Function<String, PersonBadgeStyle> badgeResolver =
                shell != null
                        ? shell.personBadgeStyleResolverForGantt()
                        : (String __) -> PersonBadgeStyle.defaultStyle();
        return new EquipmentGanttPrintPageSpec(
                columns,
                rows,
                badgeSlotRowsSlice,
                graphicZoomSlider != null ? graphicZoomSlider.getValue() / 100.0 : 1.0,
                graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d,
                graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d,
                snapshotEquipmentGanttBarFontFamily(),
                graphicBarFontPctSlider != null ? graphicBarFontPctSlider.getValue() : 100d,
                graphicHeaderHeightSlider != null ? graphicHeaderHeightSlider.getValue() : 100d,
                snapshotEquipmentGanttDateColWidth(),
                snapshotEquipmentGanttMachineColWidth(),
                snapshotEquipmentGanttProcessColWidth(),
                snapshotEquipmentGanttPersonBadgeEnabled(),
                badgeResolver,
                snapshotEquipmentGanttPersonBadgeGapPx(),
                snapshotEquipmentGanttPersonBadgeBandVerticalOffsetPx(),
                equipmentGanttBadgeDragDeltas,
                snapshotEquipmentGanttPersonBadgeWireStrokeHex(),
                snapshotEquipmentGanttPersonBadgeWireWidthPx(),
                snapshotEquipmentGanttPersonBadgeWireDashStyleKey(),
                snapshotEquipmentGanttPersonBadgeWireMaxLengthPx(),
                snapshotEquipmentGanttPersonBadgeWireEnabled(),
                snapshotEquipmentGanttPrepTimeLabelsEnabled(),
                printRange != null ? printRange.startInclusive() : null,
                printRange != null ? printRange.endExclusive() : null);
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

    /**
     * {@link #buildEquipmentGanttBorderPane} に渡す表示設定と表データ指紋を連結したキー。
     * 印刷専用 build は対象外。
     */
    private String computeGraphicRenderKey(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            String dataFingerprint) {
        StringBuilder sb = new StringBuilder(384);
        sb.append(dataFingerprint != null ? dataFingerprint : "").append('\u0001');
        DesktopTheme theme =
                shell != null ? shell.currentDesktopTheme() : DesktopTheme.LIGHT;
        sb.append(theme.name()).append('\u0001');
        sb.append(graphicZoomSlider != null ? graphicZoomSlider.getValue() : 100d)
                .append('\u0001');
        sb.append(graphicRowHeightSlider != null ? graphicRowHeightSlider.getValue() : 100d)
                .append('\u0001');
        sb.append(graphicSlotWidthSlider != null ? graphicSlotWidthSlider.getValue() : 100d)
                .append('\u0001');
        sb.append(graphicHeaderHeightSlider != null ? graphicHeaderHeightSlider.getValue() : 100d)
                .append('\u0001');
        sb.append(graphicBarFontPctSlider != null ? graphicBarFontPctSlider.getValue() : 100d)
                .append('\u0001');
        sb.append(snapshotEquipmentGanttBarFontFamily()).append('\u0001');
        sb.append(snapshotEquipmentGanttDateColWidth()).append('\u0001');
        sb.append(snapshotEquipmentGanttMachineColWidth()).append('\u0001');
        sb.append(snapshotEquipmentGanttProcessColWidth()).append('\u0001');
        sb.append(snapshotEquipmentGanttShiftWheelHScrollPercent()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeEnabled()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeGapPx()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeBandVerticalOffsetPx()).append('\u0001');
        sb.append(effectivePersonBadgeDragAdjustEnabled()).append('\u0001');
        sb.append(effectivePersonBadgeAssignmentEditEnabled()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeWireEnabled()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeWireStrokeHex()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeWireWidthPx()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeWireDashStyleKey()).append('\u0001');
        sb.append(snapshotEquipmentGanttPersonBadgeWireMaxLengthPx()).append('\u0001');
        sb.append(snapshotEquipmentGanttPrepTimeLabelsEnabled()).append('\u0001');
        sb.append(columns != null ? columns.size() : 0).append('\u0001');
        sb.append(rows != null ? rows.size() : 0);
        if (!equipmentGanttBadgeDragDeltas.isEmpty()) {
            equipmentGanttBadgeDragDeltas.entrySet().stream()
                    .sorted(Map.Entry.comparingByKey())
                    .forEach(
                            e -> {
                                EquipmentGanttBadgeDragDelta d = e.getValue();
                                sb.append('\u0001')
                                        .append(e.getKey())
                                        .append(':')
                                        .append(d != null ? d.dx() : 0)
                                        .append(',')
                                        .append(d != null ? d.dy() : 0);
                            });
        }
        if (assignmentEditModel != null && assignmentDirty) {
            assignmentEditModel.snapshotPersonsByBarId().entrySet().stream()
                    .sorted(Map.Entry.comparingByKey())
                    .forEach(
                            e -> {
                                sb.append('\u0001').append(e.getKey()).append('=');
                                if (e.getValue() != null) {
                                    e.getValue()
                                            .forEach(
                                                    p ->
                                                            sb.append(
                                                                            p != null
                                                                                    ? p.memberKey()
                                                                                    : "")
                                                                    .append(','));
                                }
                            });
        }
        return sb.toString();
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

    /** 契約 JSON の {@code kwargs_packed.regular_shift_times} を読み、印刷「定常時刻」に使う。 */
    private void loadRegularShiftTimesFromPlan(Path planJsonPath) {
        loadedRegularShiftStart = null;
        loadedRegularShiftEnd = null;
        if (planJsonPath == null || !Files.isRegularFile(planJsonPath)) {
            refreshPrintTimeRegularHint();
            return;
        }
        Path contract = resolveContractJsonForPlanPath(planJsonPath);
        if (contract == null || !Files.isRegularFile(contract)) {
            refreshPrintTimeRegularHint();
            return;
        }
        try {
            JsonNode root =
                    GANTT_CONTRACT_PEEK_OM.readTree(
                            Files.readString(contract, StandardCharsets.UTF_8));
            JsonNode packed = root.get("kwargs_packed");
            if (packed == null || !packed.isObject()) {
                refreshPrintTimeRegularHint();
                return;
            }
            Object decoded = GanttContractValueDecoder.decodeValue(packed.get("regular_shift_times"));
            if (decoded instanceof List<?> list && list.size() >= 2) {
                loadedRegularShiftStart = GanttContractValueDecoder.toLocalTime(list.get(0));
                loadedRegularShiftEnd = GanttContractValueDecoder.toLocalTime(list.get(1));
            }
        } catch (IOException | RuntimeException ignored) {
            loadedRegularShiftStart = null;
            loadedRegularShiftEnd = null;
        }
        applyRegularShiftTimesToPrintFields();
        refreshPrintTimeRegularHint();
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
            List<List<String>> contractBadgeSlotRows,
            EquipmentGanttAssignmentMetadata assignmentMetadata) {}

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
                        m,
                        fn0.toString() + " (設備ガント契約・直接)",
                        bundle.badgeSlotRows(),
                        bundle.assignmentMetadata());
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
        EquipmentGanttAssignmentMetadata assignmentMetadata = EquipmentGanttAssignmentMetadata.empty();
        if (contract != null && Files.isRegularFile(contract)) {
            EquipmentGanttSheetBundle bundle =
                    EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(contract);
            sheets.put(DEFAULT_SHEET, bundle.table());
            badgeRows = bundle.badgeSlotRows();
            assignmentMetadata = bundle.assignmentMetadata();
            desc = desc + " / " + contract.getFileName() + " (設備ガント帯)";
        }
        return new SheetLoad(sheets, desc, badgeRows, assignmentMetadata);
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
