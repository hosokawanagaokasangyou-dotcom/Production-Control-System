package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.YearMonth;
import java.util.HashMap;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.EditableCompanyCalendarPane;

/** 会社カレンダー編集タブ。 */
public class CompanyCalendarTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    @FXML
    private VBox calendarHost;

    @FXML
    private Spinner<Integer> yearSpinner;

    @FXML
    private Label statusLabel;

    @FXML
    private Button fetchHolidaysButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button exportMasterButton;

    private MainShellController shell;
    private EditableCompanyCalendarPane calendarPane;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        if (calendarHost != null && calendarPane == null) {
            calendarPane = new EditableCompanyCalendarPane();
            calendarHost.getChildren().add(calendarPane);
        }
        int y = java.time.LocalDate.now().getYear();
        if (yearSpinner != null) {
            yearSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(2020, 2040, y));
            yearSpinner
                    .valueProperty()
                    .addListener(
                            (obs, oldY, newY) -> {
                                if (calendarPane != null && newY != null) {
                                    YearMonth cur = calendarPane.displayedYearMonth();
                                    int month =
                                            cur != null
                                                    ? cur.getMonthValue()
                                                    : java.time.LocalDate.now().getMonthValue();
                                    calendarPane.setDisplayedYearMonth(YearMonth.of(newY, month));
                                }
                                refreshFromPython();
                            });
        }
        refreshFromPython();
    }

    @FXML
    private void onFetchHolidays() {
        if (shell == null) {
            return;
        }
        int year = yearSpinner.getValue();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "fetch_holidays", Integer.toString(year), "--weekends"),
                node ->
                        statusLabel.setText(
                                "祝日取得: 適用 " + node.path("applied").asInt(0) + " 日"),
                true);
    }

    @FXML
    private void onSave() {
        if (shell == null || calendarPane == null) {
            return;
        }
        try {
            int year = yearSpinner.getValue();
            Map<String, Object> patch = new HashMap<>();
            patch.put("year", year);
            patch.put("days", calendarPane.exportDaysJsonForYear(year));
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-attendance-patch-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildAttendanceDataIoRequest(
                            "merge_company_calendar", "--patch-file", tmp.toString()),
                    node -> statusLabel.setText("保存完了: " + node.path("json_path").asText("")),
                    true);
        } catch (Exception e) {
            statusLabel.setText("エラー: " + e.getMessage());
        }
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
                                "master 出力: " + node.path("sheets_updated").toString()),
                false);
    }

    @FXML
    private void onRefresh() {
        refreshFromPython();
    }

    private void refreshFromPython() {
        if (shell == null) {
            return;
        }
        int year =
                yearSpinner != null
                        ? yearSpinner.getValue()
                        : java.time.LocalDate.now().getYear();
        runAsync(
                shell.buildAttendanceDataIoRequest("company_calendar", Integer.toString(year)),
                node -> {
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
                                            days.put(e.getKey(), row);
                                        });
                        calendarPane.setDisplayedYearMonth(
                                YearMonth.of(
                                        year,
                                        Math.min(
                                                java.time.LocalDate.now().getMonthValue(), 12)));
                        calendarPane.setDaysFromJson(days);
                    }
                    statusLabel.setText("読込 revision=" + node.path("revision").asInt(0));
                },
                false);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshAfter) {
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
                                                shell.appendLog("[company-calendar] " + err);
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
                                                    shell.appendLog(
                                                            "[company-calendar] exit="
                                                                    + cap.exitCode()
                                                                    + " "
                                                                    + cap.stdout());
                                                    return;
                                                }
                                                if (cap.exitCode() != 0) {
                                                    statusLabel.setText(
                                                            "失敗 exit=" + cap.exitCode());
                                                    shell.appendLog(
                                                            "[company-calendar] "
                                                                    + cap.stdout());
                                                    return;
                                                }
                                                onOk.accept(node);
                                                if (refreshAfter) {
                                                    refreshFromPython();
                                                }
                                            } catch (Exception e) {
                                                statusLabel.setText("JSON 解析失敗: " + e.getMessage());
                                                shell.appendLog("[company-calendar] " + e);
                                            }
                                        }));
    }

    private void setToolbarBusy(boolean busy) {
        if (fetchHolidaysButton != null) {
            fetchHolidaysButton.setDisable(busy);
        }
        if (saveButton != null) {
            saveButton.setDisable(busy);
        }
        if (exportMasterButton != null) {
            exportMasterButton.setDisable(busy);
        }
    }
}
