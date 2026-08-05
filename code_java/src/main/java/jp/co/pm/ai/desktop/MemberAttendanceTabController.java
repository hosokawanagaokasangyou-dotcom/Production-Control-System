package jp.co.pm.ai.desktop;

import java.time.LocalDate;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;

/** メンバー勤怠（カレンダー方式）編集タブ。 */
public class MemberAttendanceTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    @FXML
    private Spinner<Integer> yearSpinner;

    @FXML
    private Spinner<Integer> monthSpinner;

    @FXML
    private Label statusLabel;

    @FXML
    private Button syncButton;

    @FXML
    private Button exportMasterButton;

    @FXML
    private Button refreshButton;

    private MainShellController shell;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        LocalDate today = LocalDate.now();
        if (yearSpinner != null) {
            yearSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(
                            2020, 2040, today.getYear()));
        }
        if (monthSpinner != null) {
            monthSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 12, today.getMonthValue()));
        }
        onRefresh();
    }

    @FXML
    private void onSyncFromCompanyCalendar() {
        if (shell == null) {
            return;
        }
        int year = yearSpinner.getValue();
        int month = monthSpinner.getValue();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "sync_members", Integer.toString(year), Integer.toString(month)),
                node ->
                        statusLabel.setText(
                                "同期: 適用 "
                                        + node.path("applied").asInt(0)
                                        + " / スキップ "
                                        + node.path("skipped").asInt(0)));
    }

    @FXML
    private void onExportMaster() {
        if (shell == null) {
            return;
        }
        runAsync(
                shell.buildAttendanceDataIoRequest("export_master"),
                node ->
                        statusLabel.setText(
                                "master 出力: " + node.path("sheets_updated").toString()));
    }

    @FXML
    private void onRefresh() {
        if (shell == null) {
            return;
        }
        runAsync(
                shell.buildAttendanceDataIoRequest("status"),
                node ->
                        statusLabel.setText(
                                "JSON: "
                                        + node.path("json_path").asText("")
                                        + " exists="
                                        + node.path("json_exists").asBoolean(false)));
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req, java.util.function.Consumer<JsonNode> onOk) {
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
                                            if (err != null) {
                                                statusLabel.setText("エラー: " + err.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog("[member-attendance] " + err);
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                statusLabel.setText("失敗");
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
                                                    return;
                                                }
                                                onOk.accept(node);
                                            } catch (Exception e) {
                                                statusLabel.setText(e.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog(
                                                            "[member-attendance] " + e);
                                                }
                                            }
                                        }));
    }

    private void setToolbarBusy(boolean busy) {
        if (syncButton != null) {
            syncButton.setDisable(busy);
        }
        if (exportMasterButton != null) {
            exportMasterButton.setDisable(busy);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(busy);
        }
    }
}
