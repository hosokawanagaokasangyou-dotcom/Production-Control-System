package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.AttendanceSyncStatusPane;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.EditableMemberAttendanceGridPane;
import jp.co.pm.ai.desktop.ui.InlineMonthCalendarPane;
import jp.co.pm.ai.desktop.ui.MemberHourlyAttendanceDialog;

/** メンバー勤怠（カレンダー方式）編集タブ。 */
public class MemberAttendanceTabController {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final String DEBUG_SESSION = "871314";

    @FXML
    private VBox gridHost;

    @FXML
    private VBox statusHost;

    @FXML
    private VBox monthCalendarHost;

    @FXML
    private Label statusLabel;

    @FXML
    private Button syncButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button exportMasterButton;

    @FXML
    private Button refreshButton;

    @FXML
    private Spinner<Integer> cellSizeSpinner;

    private MainShellController shell;
    private EditableMemberAttendanceGridPane gridPane;
    private AttendanceSyncStatusPane syncStatusPane;
    private InlineMonthCalendarPane monthCalendar;
    private ButtonAttentionGlow saveButtonGlow;
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private final PauseTransition gridReloadDebounce = new PauseTransition(Duration.millis(350));
    private boolean attendanceLoadEnabled = false;

    public enum UnsavedPromptResult {
        SAVED,
        DISCARDED,
        CANCELLED
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        LocalDate today = LocalDate.now();
        if (statusHost != null && syncStatusPane == null) {
            syncStatusPane = new AttendanceSyncStatusPane();
            statusHost.getChildren().add(syncStatusPane);
        }
        if (gridHost != null && gridPane == null) {
            gridPane = new EditableMemberAttendanceGridPane();
            gridPane.setCellDetailHandler(
                    req -> {
                        Stage owner =
                                shell != null ? shell.primaryStageForDialogs() : null;
                        MemberHourlyAttendanceDialog.show(
                                owner,
                                req.member(),
                                req.date().toString(),
                                req.state().hourly(),
                                hourly ->
                                        gridPane.applyHourlyEdit(
                                                req.date(),
                                                req.member(),
                                                hourly,
                                                req.state().dayPreset()));
                    });
            gridPane.setDirtyListener(this::applyGridDirtyState);
            gridHost.getChildren().add(gridPane);
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        installGridCellSizeSpinner();
        applyGridCellSizeToPane(shell.attendanceGridCellSizePx());
        installMonthCalendar(today);
    }

    private void installMonthCalendar(LocalDate today) {
        if (monthCalendarHost == null || monthCalendar != null) {
            return;
        }
        monthCalendar = new InlineMonthCalendarPane(true);
        monthCalendar.setSelectedDate(today);
        monthCalendarHost.getChildren().add(monthCalendar);
        gridReloadDebounce.setOnFinished(
                e -> {
                    if (attendanceLoadEnabled) {
                        loadGridFromPython();
                    }
                });
        monthCalendar
                .selectedDateProperty()
                .addListener(
                        (obs, oldDate, newDate) -> {
                            if (!attendanceLoadEnabled || newDate == null) {
                                return;
                            }
                            debugCalendarLog(
                                    "H1",
                                    "monthCalendar",
                                    Map.of(
                                            "oldDate",
                                            oldDate != null ? oldDate.toString() : null,
                                            "newDate",
                                            newDate.toString(),
                                            "displayedMonth",
                                            monthCalendar.getDisplayedMonth() != null
                                                    ? monthCalendar.getDisplayedMonth().toString()
                                                    : null));
                            if (oldDate != null
                                    && YearMonth.from(oldDate).equals(YearMonth.from(newDate))) {
                                return;
                            }
                            scheduleGridReload();
                        });
    }

    private LocalDate selectedCalendarDate() {
        if (monthCalendar != null && monthCalendar.getSelectedDate() != null) {
            return monthCalendar.getSelectedDate();
        }
        return LocalDate.now();
    }

    public boolean hasUnsavedEdits() {
        return gridPane != null && gridPane.hasUnsavedEdits();
    }

    public void discardUnsavedEdits() {
        loadGridFromPython();
    }

    public void clearUnsavedWithoutReload() {
        if (gridPane != null) {
            gridPane.clearUnsavedEditFlags();
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
                "メンバー勤怠に未保存の変更があります。"
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
        if (shell == null || gridPane == null) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        try {
            Map<String, Object> patch = gridPane.exportPatchJson();
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-member-attendance-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildAttendanceDataIoRequest(
                            "merge_member_attendance", "--patch-file", tmp.toString()),
                    node -> {
                        gridPane.clearUnsavedEditFlags();
                        applyGridDirtyState(false);
                        statusLabel.setText(
                                "保存完了: "
                                        + node.path("applied").asInt(0)
                                        + " セル → "
                                        + node.path("json_path").asText(""));
                        shell.refreshAttendanceReadiness();
                        refreshLocalReadiness();
                    },
                    false,
                    tmp,
                    null,
                    success -> {
                        if (onComplete != null) {
                            onComplete.accept(success);
                        }
                    });
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
            shell.onMemberAttendanceDirtyChanged(dirty);
        }
    }

    private void scheduleGridReload() {
        gridReloadDebounce.playFromStart();
    }

    private void debugCalendarLog(String hypothesisId, String location, Map<String, Object> data) {
        if (shell == null) {
            return;
        }
        // #region agent log
        Map<String, Object> payload = new LinkedHashMap<>(data);
        payload.put("toolbarBusy", syncButton != null && syncButton.isDisabled());
        AgentDebugLog.appendStructured(
                shell.snapshotUiEnv(),
                DEBUG_SESSION,
                hypothesisId,
                "MemberAttendanceTabController:" + location,
                location,
                payload);
        // #endregion
    }

    /** セッション・環境変数復元後に MainShell から呼ぶ。JSON 正本を読み込む。 */
    public void enableAttendanceLoadAndRefresh() {
        if (attendanceLoadEnabled) {
            loadGridFromPython();
            refreshLocalReadiness();
            return;
        }
        attendanceLoadEnabled = true;
        loadGridFromPython();
        refreshLocalReadiness();
    }

    /** 環境変数・パス確定後の再読込（起動時・工場ワークスペース復元後）。 */
    public void reloadAttendanceDataFromJson() {
        if (!attendanceLoadEnabled) {
            enableAttendanceLoadAndRefresh();
            return;
        }
        loadGridFromPython();
        refreshLocalReadiness();
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
        if (gridPane != null) {
            gridPane.setCellSizePx(px);
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

    @FXML
    private void onSetupWizard() {
        if (shell != null) {
            AttendanceSetupWizard.show(
                    shell,
                    ok -> {
                        if (ok) {
                            loadGridFromPython();
                            shell.refreshAttendanceReadiness();
                            refreshLocalReadiness();
                        }
                    });
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
                    loadGridFromPython();
                    refreshLocalReadiness();
                });
    }

    private void refreshLocalReadiness() {
        if (shell == null) {
            return;
        }
        LocalDate today = LocalDate.now();
        shell.runAttendanceDataIoAsync(
                shell.buildAttendanceDataIoRequest(
                        "readiness",
                        Integer.toString(today.getYear()),
                        Integer.toString(today.getMonthValue())),
                node -> {
                    if (syncStatusPane != null) {
                        syncStatusPane.updateFromReadiness(node);
                    }
                },
                err -> {
                    if (statusLabel != null) {
                        statusLabel.setText("状態取得失敗: " + err);
                    }
                });
    }

    @FXML
    private void onSyncFromCompanyCalendar() {
        if (shell == null) {
            return;
        }
        LocalDate selected = selectedCalendarDate();
        int year = selected.getYear();
        int month = selected.getMonthValue();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "sync_members", Integer.toString(year), Integer.toString(month)),
                node -> {
                    statusLabel.setText(
                            "同期: 適用 "
                                    + node.path("applied").asInt(0)
                                    + " / スキップ "
                                    + node.path("skipped").asInt(0));
                    shell.refreshAttendanceReadiness();
                    refreshLocalReadiness();
                },
                true,
                null);
    }

    @FXML
    private void onSave() {
        saveEditsAsync(null);
    }

    @FXML
    private void onExportMaster() {
        if (shell == null) {
            return;
        }
        runAsync(
                shell.buildAttendanceDataIoRequest("export_master"),
                node -> {
                    statusLabel.setText(
                            "master 出力: " + node.path("sheets_updated").toString());
                    shell.refreshAttendanceReadiness();
                    refreshLocalReadiness();
                },
                false,
                null);
    }

    @FXML
    private void onRefresh() {
        loadGridFromPython();
        refreshLocalReadiness();
    }

    @FXML
    private void onOpenMasterXlsm() {
        if (shell == null) {
            return;
        }
        if (!shell.openMasterWorkbookInDesktop("[member-attendance]")) {
            statusLabel.setText("master.xlsm が見つかりません");
        }
    }

    @FXML
    private void onOpenViewXlsx() {
        if (shell == null) {
            return;
        }
        if (!shell.openAttendanceViewXlsxInDesktop("[member-attendance]")) {
            statusLabel.setText("勤怠_表示用.xlsx が見つかりません");
        }
    }

    private void loadGridFromPython() {
        if (shell == null) {
            return;
        }
        LocalDate selected = selectedCalendarDate();
        long gen = loadGeneration.incrementAndGet();
        int year = selected.getYear();
        int month = selected.getMonthValue();
        // #region agent log
        debugCalendarLog(
                "H3",
                "loadGridStart",
                Map.of("gen", gen, "year", year, "month", month, "selectedDate", selected.toString()));
        // #endregion
        if (gridPane != null) {
            gridPane.setGridLoading(true);
            // #region agent log
            debugCalendarLog("H6", "gridLoadingOn", Map.of("gen", gen, "year", year, "month", month));
            // #endregion
        }
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "member_grid", Integer.toString(year), Integer.toString(month)),
                node -> {
                    if (gen != loadGeneration.get()) {
                        return;
                    }
                    if (gridPane != null) {
                        gridPane.loadFromMemberGridJson(node);
                    }
                    // #region agent log
                    debugCalendarLog(
                            "H3",
                            "loadGridDone",
                            Map.of(
                                    "gen",
                                    gen,
                                    "currentGen",
                                    loadGeneration.get(),
                                    "year",
                                    year,
                                    "month",
                                    month,
                                    "days",
                                    node.path("dates").size(),
                                    "members",
                                    node.path("members").size()));
                    // #endregion
                    statusLabel.setText(
                            "読込 "
                                    + year
                                    + "/"
                                    + month
                                    + " メンバー="
                                    + node.path("members").size()
                                    + " revision="
                                    + node.path("member_attendance_revision").asInt(0));
                },
                false,
                null,
                gen,
                null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshGridAfter,
            Path tempPatchFile) {
        runAsync(req, onOk, refreshGridAfter, tempPatchFile, null, null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshGridAfter,
            Path tempPatchFile,
            Long gridLoadGen,
            Consumer<Boolean> onFinished) {
        if (statusLabel != null) {
            statusLabel.setText("処理中…");
        }
        setToolbarBusy(true);
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            setToolbarBusy(false);
                                            if (tempPatchFile != null) {
                                                try {
                                                    Files.deleteIfExists(tempPatchFile);
                                                } catch (Exception ignored) {
                                                    // ignore
                                                }
                                            }
                                            if (err != null) {
                                                statusLabel.setText("エラー: " + err.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog("[member-attendance] " + err);
                                                }
                                                finishGridLoadingOverlay(gridLoadGen);
                                                if (onFinished != null) {
                                                    onFinished.accept(false);
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                statusLabel.setText("失敗");
                                                finishGridLoadingOverlay(gridLoadGen);
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
                                                    if (shell != null) {
                                                        shell.appendLog(
                                                                "[member-attendance] exit="
                                                                        + cap.exitCode()
                                                                        + " "
                                                                        + cap.stdout());
                                                    }
                                                    finishGridLoadingOverlay(gridLoadGen);
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                    return;
                                                }
                                                if (cap.exitCode() != 0) {
                                                    statusLabel.setText(
                                                            "失敗 exit=" + cap.exitCode());
                                                    if (shell != null) {
                                                        shell.appendLog(
                                                                "[member-attendance] "
                                                                        + cap.stdout());
                                                    }
                                                    finishGridLoadingOverlay(gridLoadGen);
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                    return;
                                                }
                                                onOk.accept(node);
                                                finishGridLoadingOverlay(gridLoadGen);
                                                if (onFinished != null) {
                                                    onFinished.accept(true);
                                                }
                                                if (refreshGridAfter) {
                                                    loadGridFromPython();
                                                }
                                            } catch (Exception e) {
                                                statusLabel.setText(e.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog(
                                                            "[member-attendance] " + e);
                                                }
                                                finishGridLoadingOverlay(gridLoadGen);
                                                if (onFinished != null) {
                                                    onFinished.accept(false);
                                                }
                                            }
                                        }));
    }

    private void finishGridLoadingOverlay(Long gridLoadGen) {
        if (gridLoadGen == null || gridPane == null) {
            return;
        }
        if (gridLoadGen != loadGeneration.get()) {
            return;
        }
        gridPane.setGridLoading(false);
        // #region agent log
        debugCalendarLog(
                "H6",
                "gridLoadingOff",
                Map.of("gen", gridLoadGen, "currentGen", loadGeneration.get()));
        // #endregion
    }

    private void setToolbarBusy(boolean busy) {
        // #region agent log
        debugCalendarLog("H5", "setToolbarBusy", Map.of("busy", busy));
        // #endregion
        if (syncButton != null) {
            syncButton.setDisable(busy);
        }
        if (saveButton != null) {
            saveButton.setDisable(busy);
        }
        if (exportMasterButton != null) {
            exportMasterButton.setDisable(busy);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(busy);
        }
    }
}
