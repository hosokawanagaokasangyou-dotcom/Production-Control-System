package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.animation.KeyFrame;
import javafx.animation.PauseTransition;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.AttendanceSyncStatusPane;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.EditableCompanyCalendarPane;
import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;

/** 会社カレンダー編集タブ。 */
public class CompanyCalendarTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    public enum UnsavedPromptResult {
        CANCELLED,
        DISCARDED,
        SAVED
    }

    @FXML
    private VBox calendarHost;

    @FXML
    private VBox statusHost;

    @FXML
    private Spinner<Integer> fiscalYearSpinner;

    @FXML
    private Spinner<Integer> fiscalStartMonthSpinner;

    @FXML
    private Spinner<Integer> fiscalStartDaySpinner;

    @FXML
    private Label statusLabel;

    @FXML
    private Button saveButton;

    @FXML
    private Button initializeButton;

    @FXML
    private Button setupButton;

    @FXML
    private Button restoreButton;

    @FXML
    private Button refreshButton;

    @FXML
    private Button openCalendarButton;

    @FXML
    private Spinner<Integer> cellSizeSpinner;

    private MainShellController shell;
    private EditableCompanyCalendarPane calendarPane;
    private AttendanceSyncStatusPane syncStatusPane;
    private ButtonAttentionGlow saveButtonGlow;
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private final PauseTransition fiscalDebounce = new PauseTransition(Duration.millis(350));
    private boolean attendanceLoadEnabled = false;
    private boolean suppressFiscalSpinner = false;
    private String pendingStatusOverride = null;
    private int tabProcessingDepth = 0;
    private int setupWizardGridOverlayDepth = 0;
    private String activeLoadingMessage = "処理中";
    private ProgressIndicator statusProgress;
    private Timeline statusActivityTick;
    private long statusActivityStartMs = 0L;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        if (statusHost != null && syncStatusPane == null) {
            syncStatusPane = new AttendanceSyncStatusPane();
            statusHost.getChildren().add(syncStatusPane);
        }
        if (calendarHost != null && calendarPane == null) {
            calendarPane = new EditableCompanyCalendarPane();
            calendarPane.setDirtyListener(this::applyGridDirtyState);
            calendarHost.getChildren().add(calendarPane);
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        installGridCellSizeSpinner();
        applyGridCellSizeToPane(shell.attendanceGridCellSizePx());
        LocalDate today = LocalDate.now();
        int defaultFiscalYear =
                FiscalYearPeriod.fiscalYearLabelFor(
                        today, FiscalYearPeriod.DEFAULT_APRIL_MARCH);
        if (fiscalYearSpinner != null) {
            fiscalYearSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(
                            2020, 2040, defaultFiscalYear));
            fiscalYearSpinner
                    .valueProperty()
                    .addListener(
                            (obs, o, n) ->
                                    onFiscalSpinnerChanged(
                                            o, n, fiscalYearSpinner, "会計年度を変える"));
        }
        if (fiscalStartMonthSpinner != null) {
            fiscalStartMonthSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 12, 4));
            fiscalStartMonthSpinner
                    .valueProperty()
                    .addListener((obs, o, n) -> updateFiscalStartDaySpinnerMax());
            fiscalStartMonthSpinner
                    .valueProperty()
                    .addListener(
                            (obs, o, n) ->
                                    onFiscalSpinnerChanged(
                                            o,
                                            n,
                                            fiscalStartMonthSpinner,
                                            "期間開始を変える"));
        }
        if (fiscalStartDaySpinner != null) {
            fiscalStartDaySpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 31, 1));
            fiscalStartDaySpinner
                    .valueProperty()
                    .addListener(
                            (obs, o, n) ->
                                    onFiscalSpinnerChanged(
                                            o,
                                            n,
                                            fiscalStartDaySpinner,
                                            "期間開始を変える"));
        }
        installToolbarTooltips();
        installStatusActivityRow();
        updateFiscalStartDaySpinnerMax();
        fiscalDebounce.setOnFinished(
                e -> {
                    applyFiscalSettingsToPane();
                    if (attendanceLoadEnabled) {
                        refreshFromPython();
                    }
                });
        applyFiscalSettingsToPane();
    }

    private void installStatusActivityRow() {
        if (statusLabel == null || statusLabel.getParent() == null) {
            return;
        }
        statusProgress = new ProgressIndicator();
        statusProgress.setPrefSize(16, 16);
        statusProgress.setMaxSize(16, 16);
        statusProgress.setVisible(false);
        HBox row = new HBox(8, statusProgress, statusLabel);
        row.setAlignment(Pos.CENTER_LEFT);
        if (statusLabel.getParent() instanceof VBox parent) {
            int idx = parent.getChildren().indexOf(statusLabel);
            if (idx >= 0) {
                parent.getChildren().set(idx, row);
            }
        }
        statusActivityTick =
                new Timeline(
                        new KeyFrame(
                                Duration.millis(400),
                                e -> updateStatusActivityLabel()));
        statusActivityTick.setCycleCount(Timeline.INDEFINITE);
    }

    private void installGridCellSizeSpinner() {
        if (cellSizeSpinner == null || shell == null) {
            return;
        }
        cellSizeSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AttendanceGridCellSizing.MIN_PX,
                        AttendanceGridCellSizing.MAX_PX,
                        shell.attendanceGridCellSizePx()));
        cellSizeSpinner
                .valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null && shell != null) {
                                shell.setAttendanceGridCellSizePx(n);
                            }
                        });
    }

    public void applyGridCellSizeToPane(int px) {
        if (calendarPane != null) {
            calendarPane.setCellSizePx(px);
        }
    }

    public void syncGridCellSizeSpinner(int px) {
        if (cellSizeSpinner == null) {
            return;
        }
        SpinnerValueFactory<Integer> vf = cellSizeSpinner.getValueFactory();
        if (vf instanceof SpinnerValueFactory.IntegerSpinnerValueFactory intVf
                && intVf.getValue() != px) {
            intVf.setValue(px);
        }
    }

    public boolean hasUnsavedEdits() {
        return calendarPane != null && calendarPane.hasUnsavedEdits();
    }

    public void discardUnsavedEdits() {
        refreshFromPython();
    }

    public void clearUnsavedWithoutReload() {
        if (calendarPane != null) {
            calendarPane.clearUnsavedEditFlags();
            applyGridDirtyState(false);
        }
    }

    public UnsavedPromptResult promptUnsavedChanges(String actionDescription) {
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
                "会社カレンダーに未保存の変更があります。"
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

    public void saveEditsAsync(Consumer<Boolean> onComplete) {
        if (shell == null || calendarPane == null) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        if (!confirmSaveWithFourDigit()) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        try {
            int fiscalYear = currentFiscalYearLabel();
            FiscalYearPeriod period = currentFiscalPeriod();
            Map<String, Object> patch = new HashMap<>();
            patch.put("year", fiscalYear);
            patch.put("fiscal_start_month", period.startMonth());
            patch.put("fiscal_start_day", period.startDay());
            patch.put("days", calendarPane.exportDaysJsonForFiscalYear(fiscalYear, period));
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-attendance-patch-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildAttendanceDataIoRequest(
                            "merge_company_calendar", "--patch-file", tmp.toString()),
                    mergeNode -> {
                        setActiveLoadingMessage("勤怠カレンダー.xlsx を出力中");
                        runAsync(
                                shell.buildAttendanceDataIoRequest("export_calendar_xlsx"),
                                exportNode -> {
                                    calendarPane.clearUnsavedEditFlags();
                                    applyGridDirtyState(false);
                                    if (statusLabel != null) {
                                        statusLabel.setText(
                                                "保存・勤怠カレンダー.xlsx 出力完了: "
                                                        + mergeNode.path("json_path").asText("")
                                                        + " / "
                                                        + exportNode.path("calendar_xlsx_path").asText("")
                                                        + " / シート "
                                                        + exportNode.path("sheets_updated")
                                                                .toString());
                                    }
                                    shell.refreshAttendanceReadiness();
                                    refreshLocalReadiness();
                                },
                                false,
                                null,
                                null,
                                exportSuccess -> {
                                    if (!exportSuccess && statusLabel != null) {
                                        statusLabel.setText(
                                                "JSON 保存済みだが勤怠カレンダー.xlsx の出力に失敗しました。"
                                                        + mergeNode.path("json_path").asText(""));
                                    }
                                    if (onComplete != null) {
                                        onComplete.accept(exportSuccess);
                                    }
                                });
                    },
                    false,
                    tmp,
                    null,
                    mergeSuccess -> {
                        if (!mergeSuccess && onComplete != null) {
                            onComplete.accept(false);
                        }
                    },
                    "会社カレンダーを JSON に保存中");
        } catch (Exception e) {
            statusLabel.setText("エラー: " + e.getMessage());
            if (onComplete != null) {
                onComplete.accept(false);
            }
        }
    }

    private void applyGridDirtyState(boolean dirty) {
        if (saveButtonGlow != null) {
            if (dirty) {
                saveButtonGlow.startIfIdle();
            } else {
                saveButtonGlow.stop();
            }
        }
        if (shell != null) {
            shell.onCompanyCalendarDirtyChanged(dirty);
        }
    }

    /** セッション・環境変数復元後に MainShell から呼ぶ。JSON 正本を読み込む。 */
    public void enableAttendanceLoadAndRefresh() {
        if (attendanceLoadEnabled) {
            refreshFromPython();
            refreshLocalReadiness();
            return;
        }
        attendanceLoadEnabled = true;
        applyFiscalSettingsToPane();
        refreshFromPython();
        refreshLocalReadiness();
    }

    /** 環境変数・パス確定後の再読込（起動時・工場ワークスペース復元後）。 */
    public void reloadAttendanceDataFromJson() {
        if (!attendanceLoadEnabled) {
            enableAttendanceLoadAndRefresh();
            return;
        }
        refreshFromPython();
        refreshLocalReadiness();
    }

    public int getFiscalYearLabel() {
        return currentFiscalYearLabel();
    }

    public FiscalYearPeriod getFiscalPeriod() {
        return currentFiscalPeriod();
    }

    private void onFiscalSpinnerChanged(
            Integer oldVal, Integer newVal, Spinner<Integer> spinner, String actionDescription) {
        if (suppressFiscalSpinner || newVal == null) {
            return;
        }
        if (oldVal != null && oldVal.equals(newVal)) {
            return;
        }
        if (!attendanceLoadEnabled) {
            applyFiscalSettingsToPane();
            return;
        }
        handleUnsavedThen(
                actionDescription,
                () -> scheduleFiscalReload(),
                () -> revertSpinner(spinner, oldVal));
    }

    private void revertSpinner(Spinner<Integer> spinner, Integer oldVal) {
        if (spinner == null || oldVal == null) {
            return;
        }
        suppressFiscalSpinner = true;
        SpinnerValueFactory<Integer> vf = spinner.getValueFactory();
        if (vf instanceof SpinnerValueFactory.IntegerSpinnerValueFactory intVf) {
            intVf.setValue(oldVal);
        }
        suppressFiscalSpinner = false;
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

    private void installToolbarTooltips() {
        installTooltip(saveButton, "編集内容を JSON に保存し 勤怠カレンダー.xlsx を出力します");
        installTooltip(
                setupButton,
                "祝日・週末公休の取得とメンバー勤怠の同期（初回／再取得）");
        installTooltip(initializeButton, "表示中会計年度の手動設定を削除し平日／週末既定に戻します");
        installTooltip(restoreButton, "attendance-data.json の過去リビジョンから復元します");
        installTooltip(
                openCalendarButton,
                "勤怠カレンダー.xlsx を Excel で読み取り専用で開きます（未出力の場合は先に保存してください）");
        installTooltip(
                refreshButton,
                "JSON 正本から再読込します（未保存の変更がある場合は確認します）");
        if (cellSizeSpinner != null) {
            installTooltip(cellSizeSpinner, "グリッドセルサイズ（メンバー勤怠と共通）");
        }
    }

    private static void installTooltip(javafx.scene.control.Control control, String text) {
        if (control != null) {
            control.setTooltip(new Tooltip(text));
        }
    }

    private void scheduleFiscalReload() {
        fiscalDebounce.playFromStart();
    }

    private void updateFiscalStartDaySpinnerMax() {
        if (fiscalStartDaySpinner == null || fiscalStartMonthSpinner == null) {
            return;
        }
        int month = fiscalStartMonthSpinner.getValue();
        int year =
                fiscalYearSpinner != null
                        ? fiscalYearSpinner.getValue()
                        : LocalDate.now().getYear();
        int max = java.time.YearMonth.of(year, month).lengthOfMonth();
        int cur = fiscalStartDaySpinner.getValue();
        SpinnerValueFactory<Integer> vf =
                new SpinnerValueFactory.IntegerSpinnerValueFactory(1, max, Math.min(cur, max));
        fiscalStartDaySpinner.setValueFactory(vf);
    }

    private FiscalYearPeriod currentFiscalPeriod() {
        int month =
                fiscalStartMonthSpinner != null
                        ? fiscalStartMonthSpinner.getValue()
                        : 4;
        int day =
                fiscalStartDaySpinner != null ? fiscalStartDaySpinner.getValue() : 1;
        return new FiscalYearPeriod(month, day);
    }

    private int currentFiscalYearLabel() {
        return fiscalYearSpinner != null
                ? fiscalYearSpinner.getValue()
                : FiscalYearPeriod.fiscalYearLabelFor(
                        LocalDate.now(), FiscalYearPeriod.DEFAULT_APRIL_MARCH);
    }

    private boolean confirmSaveWithFourDigit() {
        if (shell == null) {
            return false;
        }
        return FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "会社カレンダー保存",
                "編集内容を attendance-data.json（正本）と 勤怠カレンダー.xlsx に保存します。",
                "保存");
    }

    private void applyFiscalSettingsToPane() {
        if (calendarPane != null) {
            calendarPane.setFiscalYear(currentFiscalYearLabel(), currentFiscalPeriod());
        }
    }

    @FXML
    private void onSetupWizard() {
        if (shell != null) {
            beginSetupWizardGridOverlay();
            LocalDate today = LocalDate.now();
            AttendanceSetupWizard.show(
                    shell,
                    currentFiscalYearLabel(),
                    currentFiscalPeriod(),
                    today.getYear(),
                    today.getMonthValue(),
                    ok -> {
                        endSetupWizardGridOverlay();
                        if (ok) {
                            refreshFromPython();
                            shell.refreshAttendanceReadiness();
                        }
                    });
            endSetupWizardGridOverlay();
        }
    }

    @FXML
    private void onRestoreFromHistory() {
        if (shell == null) {
            return;
        }
        AttendanceJsonHistoryDialog.show(
                shell,
                () -> {
                    refreshFromPython();
                    refreshLocalReadiness();
                });
    }

    /** タブ内ステータスパネル用（グローバル段階2ブロックは MainShell が担当）。 */
    private void refreshLocalReadiness() {
        if (shell == null) {
            return;
        }
        LocalDate today = LocalDate.now();
        int year = today.getYear();
        int month = today.getMonthValue();
        shell.runAttendanceDataIoAsync(
                shell.buildAttendanceDataIoRequest(
                        "readiness", Integer.toString(year), Integer.toString(month)),
                node -> {
                    if (syncStatusPane != null) {
                        syncStatusPane.updateFromReadiness(node);
                    }
                    if (calendarPane != null) {
                        calendarPane.setGridNeedsAttention(
                                node.path("needs_setup").asBoolean(false));
                    }
                },
                err -> {
                    if (statusLabel != null) {
                        statusLabel.setText("状態取得失敗: " + err);
                    }
                });
    }

    @FXML
    private void onSave() {
        saveEditsAsync(null);
    }

    @FXML
    private void onInitializeCompanyCalendar() {
        if (shell == null) {
            return;
        }
        if (hasUnsavedEdits()) {
            UnsavedPromptResult unsaved = promptUnsavedChanges("初期化");
            if (unsaved == UnsavedPromptResult.CANCELLED) {
                return;
            }
            if (unsaved == UnsavedPromptResult.SAVED) {
                saveEditsAsync(
                        saved -> {
                            if (saved) {
                                runInitializeCompanyCalendarAfterConfirm();
                            }
                        });
                return;
            }
        }
        runInitializeCompanyCalendarAfterConfirm();
    }

    private void runInitializeCompanyCalendarAfterConfirm() {
        int fiscalYear = currentFiscalYearLabel();
        FiscalYearPeriod period = currentFiscalPeriod();
        String range = period.rangeLabel(fiscalYear);
        String message =
                "表示中の会計年度（"
                        + range
                        + "）の会社カレンダーを初期化します。"
                        + "手動設定した休日・特別休暇は削除され、平日／週末の既定表示に戻ります。"
                        + "メンバー勤怠は変更しません。";
        if (!FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "会社カレンダー初期化",
                message,
                "初期化")) {
            return;
        }
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "initialize_company_calendar",
                        Integer.toString(fiscalYear),
                        Integer.toString(period.startMonth()),
                        Integer.toString(period.startDay())),
                node -> {
                    if (calendarPane != null) {
                        calendarPane.clearUnsavedEditFlags();
                        calendarPane.setGridNeedsAttention(false);
                        applyGridDirtyState(false);
                    }
                    pendingStatusOverride =
                            "初期化完了: "
                                    + range
                                    + " — 手動設定 "
                                    + node.path("removed").asInt(0)
                                    + " 日を削除";
                    shell.refreshAttendanceReadiness();
                    refreshLocalReadiness();
                    refreshFromPython();
                },
                false,
                null);
    }

    @FXML
    private void onRefresh() {
        handleUnsavedThen(
                "再読込",
                () -> {
                    refreshFromPython();
                    refreshLocalReadiness();
                },
                () -> {});
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
                            + "\n「保存」で 勤怠カレンダー.xlsx を出力してから開いてください。");
            return;
        }
        if (!shell.openAttendanceCalendarXlsxInDesktop("[company-calendar]")) {
            shell.showErrorDialog(
                    "勤怠・機械カレンダーを開く",
                    "ファイルを開けませんでした。\n" + path);
        }
    }

    private void refreshFromPython() {
        if (shell == null) {
            return;
        }
        long gen = loadGeneration.incrementAndGet();
        int fiscalYear = currentFiscalYearLabel();
        FiscalYearPeriod period = currentFiscalPeriod();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "company_calendar",
                        Integer.toString(fiscalYear),
                        Integer.toString(period.startMonth()),
                        Integer.toString(period.startDay())),
                node -> {
                    if (gen != loadGeneration.get()) {
                        return;
                    }
                    if (calendarPane != null && node.path("days").isObject()) {
                        Map<String, Map<String, Object>> days = new HashMap<>();
                        node.path("days")
                                .fields()
                                .forEachRemaining(
                                        e -> {
                                            Map<String, Object> row = new HashMap<>();
                                            JsonNode v = e.getValue();
                                            row.put("kind", v.path("kind").asText(""));
                                            row.put("label", v.path("label").asText(""));
                                            if (v.has("source")) {
                                                row.put("source", v.path("source").asText(""));
                                            }
                                            if (v.has("manual_edit")) {
                                                row.put(
                                                        "manual_edit",
                                                        v.path("manual_edit").asBoolean(false));
                                            }
                                            days.put(e.getKey(), row);
                                        });
                        calendarPane.setFiscalYearAndDays(fiscalYear, period, days);
                    }
                    statusLabel.setText(
                            pendingStatusOverride != null
                                    ? pendingStatusOverride
                                    : "読込 "
                                            + period.rangeLabel(fiscalYear)
                                            + " revision="
                                            + node.path("revision").asInt(0));
                    pendingStatusOverride = null;
                },
                false,
                null,
                null,
                null,
                "会社カレンダーを読込中");
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshAfter,
            Path tempPatchFile) {
        runAsync(req, onOk, refreshAfter, tempPatchFile, null, null, null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshAfter,
            Path tempPatchFile,
            Long gridLoadGen,
            Consumer<Boolean> onFinished) {
        runAsync(req, onOk, refreshAfter, tempPatchFile, gridLoadGen, onFinished, null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshAfter,
            Path tempPatchFile,
            Long gridLoadGen,
            Consumer<Boolean> onFinished,
            String loadingMessage) {
        setActiveLoadingMessage(loadingMessage);
        pushTabProcessing();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            try {
                                                if (tempPatchFile != null) {
                                                    try {
                                                        Files.deleteIfExists(tempPatchFile);
                                                    } catch (Exception ignored) {
                                                        // ignore
                                                    }
                                                }
                                                if (err != null) {
                                                    statusLabel.setText("エラー: " + err.getMessage());
                                                    shell.appendLog("[company-calendar] " + err);
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                    return;
                                                }
                                                if (cap == null) {
                                                    statusLabel.setText("失敗");
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
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
                                                        statusLabel.setText(
                                                                "エラー: "
                                                                        + node.path("error")
                                                                                .asText("失敗"));
                                                        shell.appendLog(
                                                                "[company-calendar] exit="
                                                                        + cap.exitCode()
                                                                        + " "
                                                                        + cap.stdout());
                                                        if (onFinished != null) {
                                                            onFinished.accept(false);
                                                        }
                                                        return;
                                                    }
                                                    if (cap.exitCode() != 0) {
                                                        statusLabel.setText(
                                                                "失敗 exit=" + cap.exitCode());
                                                        shell.appendLog(
                                                                "[company-calendar] "
                                                                        + cap.stdout());
                                                        if (onFinished != null) {
                                                            onFinished.accept(false);
                                                        }
                                                        return;
                                                    }
                                                    onOk.accept(node);
                                                    if (onFinished != null) {
                                                        onFinished.accept(true);
                                                    }
                                                    if (refreshAfter) {
                                                        refreshFromPython();
                                                    }
                                                } catch (Exception e) {
                                                    statusLabel.setText("JSON 解析失敗: " + e.getMessage());
                                                    shell.appendLog("[company-calendar] " + e);
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                }
                                            } finally {
                                                popTabProcessing();
                                            }
                                        }));
    }

    private void pushTabProcessing() {
        tabProcessingDepth++;
        if (tabProcessingDepth == 1) {
            setToolbarBusy(true);
            beginStatusActivity();
        }
    }

    private void popTabProcessing() {
        if (tabProcessingDepth > 0) {
            tabProcessingDepth--;
        }
        if (tabProcessingDepth == 0) {
            setToolbarBusy(false);
            endStatusActivity();
        }
    }

    private void setActiveLoadingMessage(String message) {
        activeLoadingMessage =
                message != null && !message.isBlank() ? message.strip() : "処理中";
        updateGridLoadingOverlay();
        if (tabProcessingDepth > 0) {
            updateStatusActivityLabel();
        }
    }

    private void beginStatusActivity() {
        statusActivityStartMs = System.currentTimeMillis();
        if (statusProgress != null) {
            statusProgress.setVisible(true);
        }
        updateStatusActivityLabel();
        if (statusActivityTick != null) {
            statusActivityTick.play();
        }
    }

    private void endStatusActivity() {
        if (statusActivityTick != null) {
            statusActivityTick.stop();
        }
        if (statusProgress != null) {
            statusProgress.setVisible(false);
        }
    }

    private void updateStatusActivityLabel() {
        if (statusLabel == null) {
            return;
        }
        double sec = (System.currentTimeMillis() - statusActivityStartMs) / 1000.0;
        statusLabel.setText(
                String.format("%s…（経過 %.1f 秒）", activeLoadingMessage, sec));
    }

    private void beginSetupWizardGridOverlay() {
        setupWizardGridOverlayDepth++;
        updateGridLoadingOverlay();
    }

    private void endSetupWizardGridOverlay() {
        if (setupWizardGridOverlayDepth > 0) {
            setupWizardGridOverlayDepth--;
        }
        updateGridLoadingOverlay();
    }

    private void updateGridLoadingOverlay() {
        if (calendarPane == null) {
            return;
        }
        boolean loading = setupWizardGridOverlayDepth > 0 || tabProcessingDepth > 0;
        String message =
                loading
                        ? (setupWizardGridOverlayDepth > 0
                                ? "セットアップ準備中"
                                : activeLoadingMessage)
                        : null;
        calendarPane.setGridLoading(loading, message);
    }

    private void setToolbarBusy(boolean busy) {
        if (saveButton != null) {
            saveButton.setDisable(busy);
        }
        if (initializeButton != null) {
            initializeButton.setDisable(busy);
        }
        if (setupButton != null) {
            setupButton.setDisable(busy);
        }
        if (restoreButton != null) {
            restoreButton.setDisable(busy);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(busy);
        }
        if (openCalendarButton != null) {
            openCalendarButton.setDisable(busy);
        }
        if (fiscalYearSpinner != null) {
            fiscalYearSpinner.setDisable(busy);
        }
        if (fiscalStartMonthSpinner != null) {
            fiscalStartMonthSpinner.setDisable(busy);
        }
        if (fiscalStartDaySpinner != null) {
            fiscalStartDaySpinner.setDisable(busy);
        }
        if (cellSizeSpinner != null) {
            cellSizeSpinner.setDisable(busy);
        }
        updateGridLoadingOverlay();
    }
}
