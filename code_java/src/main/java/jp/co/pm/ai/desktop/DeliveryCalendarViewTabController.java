package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TabPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;

/**
 * Displays JSON from {@code pm_ai_delivery_calendar_view.py} (delivery calendar + plan comparison).
 * Source file uses ASCII and Unicode escapes only (safe on CP932 mounts).
 */
public final class DeliveryCalendarViewTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Map<String, String> COMPARE_HEADER_JP = compareHeaderJp();

    private static Map<String, String> compareHeaderJp() {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("machine_key", "\u6a5f\u68b0\u30ad\u30fc");
        m.put("machine_display", "\u6a5f\u68b0\u540d\u8868\u793a");
        m.put("request_id", "\u4f9d\u983cNO");
        m.put("calendar_date", "\u66c6\u65e5");
        m.put("qty_dispatch_json", "\u6570\u91cf_\u7d50\u679c\u914d\u53f0\u8868");
        m.put(
                "qty_task_input_aladdin",
                "\u6570\u91cf_\u30bf\u30b9\u30af\u5165\u529b\u30a2\u30e9\u30b8\u30f3");
        m.put("delta", "\u5dee\u5206");
        return Map.copyOf(m);
    }

    /**
     * Child stderr is merged into stdout ({@link PythonProcessRunner#runCaptureAsync}); probe scripts may
     * print logging lines before the JSON line. Scan object-shaped lines from bottom to top (same as
     * {@code ActualsStatusTabController.parseActualsPayloadRoot}).
     */
    private static JsonNode parseDeliveryCalendarPayloadRoot(String trimmed)
            throws JsonProcessingException {
        String[] lines = trimmed.split("\\R", -1);
        JsonProcessingException lastLineFailure = null;
        for (int i = lines.length - 1; i >= 0; i--) {
            String ln = lines[i].trim();
            if (ln.isEmpty() || !ln.startsWith("{")) {
                continue;
            }
            try {
                return JSON.readTree(ln);
            } catch (JsonProcessingException e) {
                lastLineFailure = e;
            }
        }
        try {
            return JSON.readTree(trimmed);
        } catch (JsonProcessingException e) {
            if (lastLineFailure != null) {
                throw lastLineFailure;
            }
            throw e;
        }
    }

    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label metaLabel;

    @FXML
    private TabPane innerTabPane;

    @FXML
    private StackPane mainSpreadsheetHost;

    @FXML
    private StackPane compareSpreadsheetHost;

    private MainShellController shell;

    private final SpreadsheetView mainSpreadsheet = new SpreadsheetView();

    private final SpreadsheetView compareSpreadsheet = new SpreadsheetView();

    private Supplier<RunRequest> requestFactory;

    @FXML
    private void initialize() {
        StackPane.setAlignment(mainSpreadsheet, Pos.TOP_LEFT);
        mainSpreadsheetHost.getChildren().setAll(mainSpreadsheet);
        VBox.setVgrow(mainSpreadsheetHost, Priority.ALWAYS);
        mainSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        mainSpreadsheet.prefWidthProperty().bind(mainSpreadsheetHost.widthProperty());
        mainSpreadsheet.prefHeightProperty().bind(mainSpreadsheetHost.heightProperty());

        StackPane.setAlignment(compareSpreadsheet, Pos.TOP_LEFT);
        compareSpreadsheetHost.getChildren().setAll(compareSpreadsheet);
        VBox.setVgrow(compareSpreadsheetHost, Priority.ALWAYS);
        compareSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        compareSpreadsheet.prefWidthProperty().bind(compareSpreadsheetHost.widthProperty());
        compareSpreadsheet.prefHeightProperty().bind(compareSpreadsheetHost.heightProperty());

        SpreadsheetThemeBridge.install(mainSpreadsheet);
        SpreadsheetThemeBridge.install(compareSpreadsheet);
        mainSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        compareSpreadsheet.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(mainSpreadsheet);
        SpreadsheetTabularSupport.installFullRowDataSelection(compareSpreadsheet);

        mainSpreadsheet.setGrid(new GridBase(0, 0));
        compareSpreadsheet.setGrid(new GridBase(0, 0));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.requestFactory = shell::buildDeliveryCalendarRequest;
    }

    @FXML
    private void onRefreshButtonAction() {
        if (requestFactory == null) {
            statusLabel.setText("\u521d\u671f\u5316\u5f85\u3061");
            return;
        }
        refreshButton.setDisable(true);
        statusLabel.setText("\u53d6\u5f97\u4e2d...");
        RunRequest req = requestFactory.get();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            refreshButton.setDisable(false);
                                            if (err != null) {
                                                statusLabel.setText("error: " + err.getMessage());
                                                if (shell != null) {
                                                    shell.appendLog("[delivery-calendar] " + err.getMessage());
                                                }
                                                return;
                                            }
                                            if (cap == null) {
                                                statusLabel.setText("no result");
                                                return;
                                            }
                                            statusLabel.setText("exit=" + cap.exitCode());
                                            applyPayload(cap.stdout());
                                        }));
    }

    private void applyPayload(String stdout) {
        metaLabel.setText("");
        String trimmed = stdout != null ? stdout.trim() : "";
        if (trimmed.isEmpty()) {
            statusLabel.setText("empty stdout");
            return;
        }
        try {
            JsonNode root = parseDeliveryCalendarPayloadRoot(trimmed);
            if (!root.path("ok").asBoolean(false)) {
                statusLabel.setText(root.path("error").asText("failed"));
            }
            JsonNode meta = root.get("meta");
            if (meta != null && meta.isObject()) {
                StringBuilder sb = new StringBuilder();
                if (meta.hasNonNull("processingPlanPath")) {
                    sb.append("PM_AI_PROCESSING_PLAN_PATH: ")
                            .append(meta.get("processingPlanPath").asText())
                            .append("\n");
                }
                if (meta.hasNonNull("dispatchJsonPath")) {
                    sb.append("\u7d50\u679c_\u914d\u53f0\u8868.json: ")
                            .append(meta.get("dispatchJsonPath").asText());
                }
                metaLabel.setText(sb.toString());
            }

            JsonNode mainCal = root.get("mainCalendar");
            if (mainCal != null && mainCal.isObject()) {
                loadMainCalendar(mainCal);
            }
            JsonNode cmp = root.get("planCompareTable");
            if (cmp != null && cmp.isObject()) {
                loadCompareTable(cmp);
            }

            SpreadsheetTabularSupport.applyColumnFiltersWithDialog(mainSpreadsheet);
            SpreadsheetTabularSupport.applyColumnFiltersWithDialog(compareSpreadsheet);
            SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(mainSpreadsheet);
            SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(compareSpreadsheet);
            int fixMain = 0;
            if (mainCal != null && mainCal.has("columns") && mainCal.get("columns").isArray()) {
                fixMain = Math.min(6, mainCal.get("columns").size());
            }
            SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(mainSpreadsheet, fixMain);
            SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(compareSpreadsheet, 3);
        } catch (Exception e) {
            statusLabel.setText("parse: " + e.getMessage());
            if (shell != null) {
                shell.appendLog("[delivery-calendar] parse " + e.getMessage());
            }
        }
    }

    private void loadMainCalendar(JsonNode mainCal) {
        List<String> headers = new ArrayList<>();
        JsonNode cols = mainCal.get("columns");
        if (cols != null && cols.isArray()) {
            for (JsonNode c : cols) {
                headers.add(c.asText(""));
            }
        }
        ObservableList<ObservableList<String>> rowObs = FXCollections.observableArrayList();
        JsonNode rows = mainCal.get("rows");
        if (rows != null && rows.isArray()) {
            for (JsonNode row : rows) {
                ObservableList<String> line = FXCollections.observableArrayList();
                JsonNode cells = row.get("cells");
                if (cells != null && cells.isArray()) {
                    for (JsonNode cell : cells) {
                        line.add(cell.asText(""));
                    }
                }
                rowObs.add(line);
            }
        }
        GridBase grid = SpreadsheetTabularSupport.buildReadOnlyPlainGrid(headers, rowObs);
        mainSpreadsheet.setGrid(grid);
        mainSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
    }

    private void loadCompareTable(JsonNode cmp) {
        List<String> headers = new ArrayList<>();
        JsonNode cols = cmp.get("columns");
        if (cols != null && cols.isArray()) {
            for (JsonNode c : cols) {
                String key = c.asText("");
                headers.add(COMPARE_HEADER_JP.getOrDefault(key, key));
            }
        }
        ObservableList<ObservableList<String>> rowObs = FXCollections.observableArrayList();
        JsonNode rows = cmp.get("rows");
        if (rows != null && rows.isArray()) {
            for (JsonNode row : rows) {
                ObservableList<String> line = FXCollections.observableArrayList();
                if (row.isArray()) {
                    for (JsonNode cell : row) {
                        line.add(cell.asText(""));
                    }
                }
                rowObs.add(line);
            }
        }
        GridBase grid = SpreadsheetTabularSupport.buildReadOnlyPlainGrid(headers, rowObs);
        compareSpreadsheet.setGrid(grid);
        compareSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
    }
}
