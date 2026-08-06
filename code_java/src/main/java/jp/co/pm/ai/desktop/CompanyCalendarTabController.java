package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicLong;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.AttendanceSyncStatusPane;
import jp.co.pm.ai.desktop.ui.EditableCompanyCalendarPane;
import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;

/** 会社カレンダー編集タブ。 */
public class CompanyCalendarTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

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
    private Button fetchHolidaysButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button exportMasterButton;

    @FXML
    private Spinner<Integer> cellSizeSpinner;

    private MainShellController shell;
    private EditableCompanyCalendarPane calendarPane;
    private AttendanceSyncStatusPane syncStatusPane;
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private final PauseTransition fiscalDebounce = new PauseTransition(Duration.millis(350));
    private boolean attendanceLoadEnabled = false;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        if (statusHost != null && syncStatusPane == null) {
            syncStatusPane = new AttendanceSyncStatusPane();
            statusHost.getChildren().add(syncStatusPane);
        }
        if (calendarHost != null && calendarPane == null) {
            calendarPane = new EditableCompanyCalendarPane();
            calendarHost.getChildren().add(calendarPane);
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
            fiscalYearSpinner.valueProperty().addListener((obs, o, n) -> scheduleFiscalReload());
        }
        if (fiscalStartMonthSpinner != null) {
            fiscalStartMonthSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 12, 4));
            fiscalStartMonthSpinner
                    .valueProperty()
                    .addListener((obs, o, n) -> updateFiscalStartDaySpinnerMax());
            fiscalStartMonthSpinner.valueProperty().addListener((obs, o, n) -> scheduleFiscalReload());
        }
        if (fiscalStartDaySpinner != null) {
            fiscalStartDaySpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 31, 1));
            fiscalStartDaySpinner.valueProperty().addListener((obs, o, n) -> scheduleFiscalReload());
        }
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

    private void applyFiscalSettingsToPane() {
        if (calendarPane != null) {
            calendarPane.setFiscalYear(currentFiscalYearLabel(), currentFiscalPeriod());
        }
    }

    @FXML
    private void onSetupWizard() {
        if (shell != null) {
            AttendanceSetupWizard.show(
                    shell,
                    ok -> {
                        if (ok) {
                            refreshFromPython();
                            shell.refreshAttendanceReadiness();
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
                },
                err -> {
                    if (statusLabel != null) {
                        statusLabel.setText("状態取得失敗: " + err);
                    }
                });
    }

    @FXML
    private void onFetchHolidays() {
        if (shell == null) {
            return;
        }
        int fiscalYear = currentFiscalYearLabel();
        FiscalYearPeriod period = currentFiscalPeriod();
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "fetch_holidays_fiscal",
                        Integer.toString(fiscalYear),
                        Integer.toString(period.startMonth()),
                        Integer.toString(period.startDay()),
                        "--weekends"),
                node -> {
                    statusLabel.setText(
                            "祝日取得: 適用 "
                                    + node.path("applied").asInt(0)
                                    + " 日 / スキップ "
                                    + node.path("skipped").asInt(0));
                    shell.refreshAttendanceReadiness();
                    refreshLocalReadiness();
                },
                true,
                null);
    }

    @FXML
    private void onSave() {
        if (shell == null || calendarPane == null) {
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
                    node -> {
                        statusLabel.setText("保存完了: " + node.path("json_path").asText(""));
                        shell.refreshAttendanceReadiness();
                        refreshLocalReadiness();
                    },
                    true,
                    tmp);
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
        refreshFromPython();
        refreshLocalReadiness();
    }

    @FXML
    private void onOpenMasterXlsm() {
        if (shell == null) {
            return;
        }
        if (!shell.openMasterWorkbookInDesktop("[company-calendar]")) {
            statusLabel.setText("master.xlsm が見つかりません");
        }
    }

    @FXML
    private void onOpenViewXlsx() {
        if (shell == null) {
            return;
        }
        if (!shell.openAttendanceViewXlsxInDesktop("[company-calendar]")) {
            statusLabel.setText("勤怠_表示用.xlsx が見つかりません");
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
                            "読込 "
                                    + period.rangeLabel(fiscalYear)
                                    + " revision="
                                    + node.path("revision").asInt(0));
                },
                false,
                null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshAfter,
            Path tempPatchFile) {
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
        if (fiscalYearSpinner != null) {
            fiscalYearSpinner.setDisable(busy);
        }
        if (fiscalStartMonthSpinner != null) {
            fiscalStartMonthSpinner.setDisable(busy);
        }
        if (fiscalStartDaySpinner != null) {
            fiscalStartDaySpinner.setDisable(busy);
        }
    }
}
