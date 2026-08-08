package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.CompanyCalendarDayVisual;
import jp.co.pm.ai.desktop.ui.EditableMachineCalendarGridPane;
import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;
import jp.co.pm.ai.desktop.ui.InlineMonthCalendarPane;
import jp.co.pm.ai.desktop.ui.MachineCalendarCellValues;

/** 機械カレンダー（JSON 正本）編集タブ。 */
public class MachineCalendarTabController {

    public enum UnsavedPromptResult {
        SAVED,
        DISCARDED,
        CANCELLED
    }

    private static final String CALENDAR_XLSX_LABEL = "勤怠・機械カレンダー.xlsx";

    private static final ObjectMapper JSON = new ObjectMapper();

    @FXML private VBox gridHost;
    @FXML private VBox monthCalendarHost;
    @FXML private Label statusLabel;
    @FXML private Button saveButton;
    @FXML private Button createInitialValuesButton;
    @FXML private Button restoreButton;
    @FXML private Button openCalendarButton;
    @FXML private Button refreshButton;
    @FXML private Button fillAllButton;
    @FXML private Button clearAllButton;
    @FXML private Button invertAllButton;
    @FXML private Button undoButton;
    @FXML private ToggleButton paintOccupiedButton;
    @FXML private ToggleButton paintClearButton;
    @FXML private Spinner<Integer> cellSizeSpinner;
    @FXML private Spinner<Integer> columnWidthSpinner;
    @FXML private Spinner<Integer> columnGapSpinner;

    private int machineCalendarColumnWidthPx =
            AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_PX;
    private int machineCalendarColumnGapPx =
            AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_GAP_PX;

    private MainShellController shell;
    private EditableMachineCalendarGridPane gridPane;
    private InlineMonthCalendarPane monthCalendar;
    private ButtonAttentionGlow saveButtonGlow;
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private final AtomicLong companyCalendarLoadGeneration = new AtomicLong(0);
    private boolean suppressDateGuard = false;
    private boolean gridDirty = false;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        LocalDate today = LocalDate.now();
        if (gridHost != null && gridPane == null) {
            gridPane = new EditableMachineCalendarGridPane();
            gridPane.setDirtyListener(this::applyGridDirtyState);
            gridPane.setUndoStateListener(this::applyUndoButtonState);
            gridPane.setCommentDialogOwner(shell.primaryStageForDialogs());
            VBox.setVgrow(gridPane, Priority.ALWAYS);
            gridPane.setMaxHeight(Double.MAX_VALUE);
            gridHost.setMaxHeight(Double.MAX_VALUE);
            gridHost.getChildren().add(gridPane);
        }
        installPaintModeToggles();
        if (undoButton != null) {
            undoButton.setDisable(true);
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        installCellSizeSpinner();
        installColumnWidthSpinner();
        installColumnGapSpinner();
        applyGridCellSize(shell.attendanceGridCellSizePx());
        applyColumnWidth(machineCalendarColumnWidthPx);
        applyColumnGap(machineCalendarColumnGapPx);
        installMonthCalendar(today);
        loadGridFromPython();
    }

    private void installCellSizeSpinner() {
        if (cellSizeSpinner == null) {
            return;
        }
        cellSizeSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AttendanceGridCellSizing.MIN_PX,
                        AttendanceGridCellSizing.MAX_PX,
                        AttendanceGridCellSizing.DEFAULT_CELL_PX,
                        2));
        cellSizeSpinner.valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null && shell != null) {
                                shell.setAttendanceGridCellSizePx(n);
                            }
                        });
    }

    private void installColumnWidthSpinner() {
        if (columnWidthSpinner == null) {
            return;
        }
        columnWidthSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AttendanceGridCellSizing.MACHINE_CALENDAR_COLUMN_MIN_PX,
                        AttendanceGridCellSizing.MACHINE_CALENDAR_COLUMN_MAX_PX,
                        AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_PX,
                        4));
        columnWidthSpinner.valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null) {
                                applyColumnWidth(n);
                            }
                        });
    }

    public void applyColumnWidth(int px) {
        int clamped = AttendanceGridCellSizing.clampMachineCalendarColumnWidth(px);
        machineCalendarColumnWidthPx = clamped;
        if (gridPane != null) {
            gridPane.setColumnWidthPx(clamped);
        }
        if (columnWidthSpinner != null
                && columnWidthSpinner.getValueFactory().getValue() != clamped) {
            columnWidthSpinner.getValueFactory().setValue(clamped);
        }
    }

    private void installColumnGapSpinner() {
        if (columnGapSpinner == null) {
            return;
        }
        columnGapSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AttendanceGridCellSizing.MACHINE_CALENDAR_COLUMN_GAP_MIN_PX,
                        AttendanceGridCellSizing.MACHINE_CALENDAR_COLUMN_GAP_MAX_PX,
                        AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_GAP_PX,
                        1));
        columnGapSpinner.valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null) {
                                applyColumnGap(n);
                            }
                        });
    }

    public void applyColumnGap(int px) {
        int clamped = AttendanceGridCellSizing.clampMachineCalendarColumnGap(px);
        machineCalendarColumnGapPx = clamped;
        if (gridPane != null) {
            gridPane.setColumnGapPx(clamped);
        }
        if (columnGapSpinner != null
                && columnGapSpinner.getValueFactory().getValue() != clamped) {
            columnGapSpinner.getValueFactory().setValue(clamped);
        }
    }

    public void syncGridCellSizeSpinner(int px) {
        if (cellSizeSpinner != null) {
            cellSizeSpinner.getValueFactory().setValue(AttendanceGridCellSizing.clamp(px));
        }
    }

    public void applyGridCellSize(int px) {
        if (gridPane != null) {
            gridPane.setCellSizePx(px);
        }
    }

    private void installMonthCalendar(LocalDate today) {
        if (monthCalendarHost == null || monthCalendar != null) {
            return;
        }
        monthCalendar = new InlineMonthCalendarPane(false);
        monthCalendar.setCompanyCalendarMode(true);
        monthCalendar.setSelectedDate(today);
        monthCalendar.displayedMonthProperty()
                .addListener((obs, o, n) -> {
                    if (n != null) {
                        loadCompanyCalendarForMiniCalendar();
                    }
                });
        monthCalendar.selectedDateProperty().addListener((obs, o, n) -> {
            if (n == null || suppressDateGuard) {
                return;
            }
            LocalDate oldDate = o != null ? o : today;
            handleUnsavedThen(
                    "日付を変える",
                    () -> loadGridFromPython(),
                    () -> {
                        suppressDateGuard = true;
                        monthCalendar.setSelectedDate(oldDate);
                        suppressDateGuard = false;
                    });
        });
        monthCalendarHost.getChildren().add(monthCalendar);
        loadCompanyCalendarForMiniCalendar();
    }

    private void loadCompanyCalendarForMiniCalendar() {
        if (shell == null || monthCalendar == null) {
            return;
        }
        LocalDate ref = selectedDate();
        FiscalYearPeriod period = shell.attendanceFiscalPeriod();
        int fiscalYear = FiscalYearPeriod.fiscalYearLabelFor(ref, period);
        long gen = companyCalendarLoadGeneration.incrementAndGet();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "company_calendar",
                        Integer.toString(fiscalYear),
                        Integer.toString(period.startMonth()),
                        Integer.toString(period.startDay())),
                node -> {
                    if (gen != companyCalendarLoadGeneration.get()) {
                        return;
                    }
                    Map<LocalDate, CompanyCalendarDayVisual.DayInfo> days =
                            CompanyCalendarDayVisual.parseDays(node.path("days"));
                    monthCalendar.setCompanyCalendarDays(days);
                },
                null);
    }

    private LocalDate selectedDate() {
        return monthCalendar != null && monthCalendar.getSelectedDate() != null
                ? monthCalendar.getSelectedDate()
                : LocalDate.now();
    }

    public boolean hasUnsavedEdits() {
        return gridPane != null && gridPane.hasUnsavedEdits();
    }

    private void handleUnsavedThen(
            String actionDescription, Runnable onProceed, Runnable onCancel) {
        if (!hasUnsavedEdits()) {
            onProceed.run();
            return;
        }
        UnsavedPromptResult result = promptUnsavedChanges(actionDescription);
        if (result == UnsavedPromptResult.CANCELLED) {
            onCancel.run();
            return;
        }
        if (result == UnsavedPromptResult.DISCARDED) {
            onProceed.run();
            return;
        }
        saveEditsAsync(
                saved -> {
                    if (saved) {
                        onProceed.run();
                    } else {
                        onCancel.run();
                    }
                });
    }

    private UnsavedPromptResult promptUnsavedChanges(String actionDescription) {
        if (!hasUnsavedEdits()) {
            return UnsavedPromptResult.DISCARDED;
        }
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        if (shell != null) {
            alert.initOwner(shell.primaryStageForDialogs());
            shell.applyAlertStylesheets(alert);
        }
        alert.setTitle("未保存の変更");
        alert.setHeaderText(null);
        alert.setContentText(
                "機械カレンダーに未保存の変更があります。"
                        + actionDescription
                        + "前に保存しますか？");
        ButtonType save = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        ButtonType discard = new ButtonType("保存しない", ButtonBar.ButtonData.NO);
        alert.getButtonTypes().setAll(save, discard, ButtonType.CANCEL);
        Optional<ButtonType> ans = alert.showAndWait();
        if (ans.isEmpty() || ans.get() == ButtonType.CANCEL) {
            return UnsavedPromptResult.CANCELLED;
        }
        if (ans.get() == discard) {
            return UnsavedPromptResult.DISCARDED;
        }
        return UnsavedPromptResult.SAVED;
    }

    private void installPaintModeToggles() {
        if (paintOccupiedButton == null || paintClearButton == null) {
            return;
        }
        ToggleGroup group = new ToggleGroup();
        paintOccupiedButton.setToggleGroup(group);
        paintClearButton.setToggleGroup(group);
        paintOccupiedButton.setSelected(true);
        group.selectedToggleProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (gridPane == null || n == null) {
                                return;
                            }
                            if (n == paintOccupiedButton) {
                                gridPane.setPaintMode(
                                        MachineCalendarCellValues.OccupancyMode.OCCUPIED);
                            } else if (n == paintClearButton) {
                                gridPane.setPaintMode(
                                        MachineCalendarCellValues.OccupancyMode.AVAILABLE);
                            }
                        });
    }

    private void applyUndoButtonState(boolean canUndo) {
        if (undoButton != null) {
            undoButton.setDisable(!canUndo);
        }
    }

    @FXML
    private void onFillAllOccupied() {
        if (gridPane != null) {
            gridPane.fillAllOccupied();
        }
    }

    @FXML
    private void onClearAll() {
        if (gridPane != null) {
            gridPane.clearAll();
        }
    }

    @FXML
    private void onInvertAll() {
        if (gridPane != null) {
            gridPane.invertAll();
        }
    }

    @FXML
    private void onUndo() {
        if (gridPane != null) {
            gridPane.undo();
        }
    }

    @FXML
    private void onSave() {
        saveEditsAsync(null);
    }

    private void saveEditsAsync(Consumer<Boolean> onComplete) {
        if (shell == null || gridPane == null) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        if (!FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "機械カレンダー保存",
                "編集内容を machine-calendar-data.json（正本）に保存し、"
                        + CALENDAR_XLSX_LABEL
                        + " へ出力します。",
                "保存")) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        gridPane.setGridLoading(true, "JSON 正本へ保存中…");
        setToolbarBusy(true);
        try {
            Map<String, Object> patch = gridPane.exportPatchJson();
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-machine-calendar-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildMachineCalendarIoRequest("merge", "--patch-file", tmp.toString()),
                    mergeNode -> {
                        if (statusLabel != null) {
                            statusLabel.setText(CALENDAR_XLSX_LABEL + " を出力中…");
                        }
                        gridPane.setGridLoadingMessage(CALENDAR_XLSX_LABEL + " を出力中…");
                        runAsync(
                                shell.buildAttendanceDataIoRequest("export_calendar_xlsx"),
                                exportNode -> {
                                    gridPane.captureSavedBaseline();
                                    applyGridDirtyState(false);
                                    endGridLoading();
                                    if (statusLabel != null) {
                                        statusLabel.setText(
                                                "保存・Excel 出力完了: "
                                                        + mergeNode.path("json_path").asText("")
                                                        + " / "
                                                        + exportNode.path("calendar_xlsx_path")
                                                                .asText(""));
                                    }
                                    if (onComplete != null) {
                                        onComplete.accept(true);
                                    }
                                },
                                tmp);
                    },
                    tmp);
        } catch (Exception e) {
            endGridLoading();
            if (statusLabel != null) {
                statusLabel.setText("エラー: " + e.getMessage());
            }
            if (onComplete != null) {
                onComplete.accept(false);
            }
        }
    }

    @FXML
    private void onCreateInitialValues() {
        if (shell == null) {
            return;
        }
        int fiscalYear = shell.attendanceFiscalYearLabel();
        var period = shell.attendanceFiscalPeriod();
        String range = period.rangeLabel(fiscalYear);
        String message =
                "会計年度（"
                        + range
                        + "）の機械カレンダーを初期化します。"
                        + "列は need シートの工程×機械、稼働枠は 8:00〜19:00。"
                        + "土日・祭日も時刻範囲内は空（稼働可能）で初期化します（会社カレンダーとは連動しません）。"
                        + "配台では人の勤怠が先に効きます。既存の占有データは当該年度で上書きされます。";
        if (!FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "機械カレンダー初期値作成",
                message,
                "初期値を作る")) {
            return;
        }
        runAsync(
                shell.buildMachineCalendarIoRequest(
                        "initialize_defaults",
                        Integer.toString(fiscalYear),
                        Integer.toString(period.startMonth()),
                        Integer.toString(period.startDay())),
                node -> {
                    if (statusLabel != null) {
                        statusLabel.setText("勤怠カレンダー.xlsx を出力中…");
                    }
                    runAsync(
                            shell.buildAttendanceDataIoRequest("export_calendar_xlsx"),
                            exportNode -> {
                                if (statusLabel != null) {
                                    statusLabel.setText(
                                            "初期値作成: 列="
                                                    + node.path("columns").asInt(0)
                                                    + " / Excel: "
                                                    + exportNode.path("calendar_xlsx_path")
                                                            .asText(""));
                                }
                                loadGridFromPython();
                            },
                            null);
                },
                null);
    }

    @FXML
    private void onRestoreFromHistory() {
        if (shell == null) {
            return;
        }
        MachineCalendarJsonHistoryDialog.show(shell, this::loadGridFromPython);
    }

    @FXML
    private void onOpenAttendanceCalendar() {
        if (shell == null) {
            return;
        }
        Path path = AppPaths.attendanceCalendarXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(path)) {
            shell.showWarningDialog(
                    "勤怠・機械カレンダーを開く",
                    "ファイルが見つかりません。\n"
                            + path
                            + "\n「保存」で "
                            + CALENDAR_XLSX_LABEL
                            + " を出力してから開いてください。");
            return;
        }
        if (!shell.openAttendanceCalendarXlsxInDesktop("[machine-calendar]")) {
            shell.showErrorDialog(
                    "勤怠・機械カレンダーを開く",
                    "ファイルを開けませんでした。\n" + path);
        }
    }

    @FXML
    private void onRefresh() {
        handleUnsavedThen(
                "再読込",
                () -> {
                    loadCompanyCalendarForMiniCalendar();
                    loadGridFromPython();
                },
                () -> {});
    }

    private void loadGridFromPython() {
        if (shell == null || gridPane == null) {
            return;
        }
        LocalDate d = selectedDate();
        long gen = loadGeneration.incrementAndGet();
        runAsync(
                shell.buildMachineCalendarIoRequest("day_grid", d.toString()),
                node -> {
                    if (gen != loadGeneration.get()) {
                        return;
                    }
                    gridPane.loadFromDayGridJson(node);
                    if (statusLabel != null) {
                        statusLabel.setText(
                                "読込 "
                                        + d
                                        + " 列="
                                        + node.path("columns").size()
                                        + " 行="
                                        + node.path("rows").size());
                    }
                },
                null);
    }

    private void applyGridDirtyState(boolean dirty) {
        gridDirty = dirty;
        if (saveButtonGlow != null) {
            if (dirty) {
                saveButtonGlow.ensureActive();
            } else {
                saveButtonGlow.stop();
            }
        }
        if (shell != null) {
            shell.onMachineCalendarDirtyChanged(dirty);
        }
    }

    private void endGridLoading() {
        if (gridPane != null) {
            gridPane.setGridLoading(false);
        }
        setToolbarBusy(false);
        applyGridDirtyState(gridDirty);
    }

    private void setToolbarBusy(boolean busy) {
        if (saveButton != null) {
            saveButton.setDisable(busy);
        }
        if (createInitialValuesButton != null) {
            createInitialValuesButton.setDisable(busy);
        }
        if (restoreButton != null) {
            restoreButton.setDisable(busy);
        }
        if (openCalendarButton != null) {
            openCalendarButton.setDisable(busy);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(busy);
        }
        if (fillAllButton != null) {
            fillAllButton.setDisable(busy);
        }
        if (clearAllButton != null) {
            clearAllButton.setDisable(busy);
        }
        if (invertAllButton != null) {
            invertAllButton.setDisable(busy);
        }
        if (undoButton != null) {
            undoButton.setDisable(
                    busy || gridPane == null || !gridPane.canUndo());
        }
        if (paintOccupiedButton != null) {
            paintOccupiedButton.setDisable(busy);
        }
        if (paintClearButton != null) {
            paintClearButton.setDisable(busy);
        }
        if (cellSizeSpinner != null) {
            cellSizeSpinner.setDisable(busy);
        }
        if (columnWidthSpinner != null) {
            columnWidthSpinner.setDisable(busy);
        }
        if (columnGapSpinner != null) {
            columnGapSpinner.setDisable(busy);
        }
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            Consumer<JsonNode> onOk,
            Path tempPatchFile) {
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            if (tempPatchFile != null) {
                                                try {
                                                    Files.deleteIfExists(tempPatchFile);
                                                } catch (Exception ignored) {
                                                    // ignore
                                                }
                                            }
                                            if (err != null) {
                                                endGridLoading();
                                                if (statusLabel != null) {
                                                    statusLabel.setText(
                                                            "エラー: " + err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                endGridLoading();
                                                if (statusLabel != null) {
                                                    statusLabel.setText("失敗");
                                                }
                                                return;
                                            }
                                            try {
                                                JsonNode node =
                                                        JSON.readTree(
                                                                AttendanceOvertimePreview
                                                                        .MasterReadSummaryJson
                                                                        .extractLastJsonLine(
                                                                                cap.stdout()));
                                                if (!node.path("ok").asBoolean(false)) {
                                                    endGridLoading();
                                                    if (statusLabel != null) {
                                                        statusLabel.setText(
                                                                "エラー: "
                                                                        + node.path("error")
                                                                                .asText("失敗"));
                                                    }
                                                    return;
                                                }
                                                onOk.accept(node);
                                            } catch (Exception e) {
                                                endGridLoading();
                                                if (statusLabel != null) {
                                                    statusLabel.setText(e.getMessage());
                                                }
                                            }
                                        }));
    }
}
