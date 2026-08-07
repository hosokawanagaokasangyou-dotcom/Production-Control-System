package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.EditableMachineCalendarGridPane;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;
import jp.co.pm.ai.desktop.ui.InlineMonthCalendarPane;

/** 機械カレンダー（JSON 正本）編集タブ。 */
public class MachineCalendarTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    @FXML private VBox gridHost;
    @FXML private VBox monthCalendarHost;
    @FXML private Label statusLabel;
    @FXML private Button saveButton;
    @FXML private Button importMasterButton;
    @FXML private Button refreshButton;
    @FXML private Spinner<Integer> cellSizeSpinner;

    private MainShellController shell;
    private EditableMachineCalendarGridPane gridPane;
    private InlineMonthCalendarPane monthCalendar;
    private ButtonAttentionGlow saveButtonGlow;
    private final AtomicLong loadGeneration = new AtomicLong(0);

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        LocalDate today = LocalDate.now();
        if (gridHost != null && gridPane == null) {
            gridPane = new EditableMachineCalendarGridPane();
            gridPane.setDirtyListener(this::applyGridDirtyState);
            gridHost.getChildren().add(gridPane);
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        installCellSizeSpinner();
        applyGridCellSize(shell.attendanceGridCellSizePx());
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
        monthCalendar.setSelectedDate(today);
        monthCalendar.selectedDateProperty().addListener((obs, o, n) -> {
            if (n != null) {
                loadGridFromPython();
            }
        });
        monthCalendarHost.getChildren().add(monthCalendar);
    }

    private LocalDate selectedDate() {
        return monthCalendar != null && monthCalendar.getSelectedDate() != null
                ? monthCalendar.getSelectedDate()
                : LocalDate.now();
    }

    @FXML
    private void onSave() {
        if (shell == null || gridPane == null) {
            return;
        }
        if (!FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "機械カレンダー保存",
                "編集内容を machine-calendar-data.json（正本）に保存します。",
                "保存")) {
            return;
        }
        try {
            Map<String, Object> patch = gridPane.exportPatchJson();
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-machine-calendar-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildMachineCalendarIoRequest("merge", "--patch-file", tmp.toString()),
                    node -> {
                        gridPane.captureSavedBaseline();
                        applyGridDirtyState(false);
                        statusLabel.setText(
                                "保存完了: "
                                        + node.path("json_path").asText("")
                                        + " ("
                                        + node.path("applied").asInt(0)
                                        + " セル)");
                    },
                    tmp);
        } catch (Exception e) {
            statusLabel.setText("エラー: " + e.getMessage());
        }
    }

    @FXML
    private void onImportFromMaster() {
        if (shell == null) {
            return;
        }
        runAsync(
                shell.buildMachineCalendarIoRequest("import_from_master"),
                node -> {
                    statusLabel.setText(
                            "master 取込: 列="
                                    + node.path("columns").asInt(0)
                                    + " スロット="
                                    + node.path("occupancy_slots").asInt(0));
                    loadGridFromPython();
                },
                null);
    }

    @FXML
    private void onRefresh() {
        loadGridFromPython();
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
                    statusLabel.setText(
                            "読込 "
                                    + d
                                    + " 列="
                                    + node.path("columns").size()
                                    + " 行="
                                    + node.path("rows").size());
                },
                null);
    }

    private void applyGridDirtyState(boolean dirty) {
        if (saveButtonGlow != null) {
            if (dirty) {
                saveButtonGlow.startIfIdle();
            } else {
                saveButtonGlow.stop();
            }
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
                                                statusLabel.setText("エラー: " + err.getMessage());
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
                                                    return;
                                                }
                                                onOk.accept(node);
                                            } catch (Exception e) {
                                                statusLabel.setText(e.getMessage());
                                            }
                                        }));
    }
}
