package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Supplier;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.GridPane;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** Master workbook read summary (Python JSON); layout {@code MasterReadSummaryTab.fxml}. */
public final class MasterReadSummaryTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String SCRIPT = "master_read_summary.py";

    public static final class SheetCheckRow {
        private String category;
        private String sheetName;
        private String status;
        private String note;

        public String getCategory() {
            return category;
        }

        public void setCategory(String category) {
            this.category = category;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public String getStatus() {
            return status;
        }

        public void setStatus(String status) {
            this.status = status;
        }

        public String getNote() {
            return note;
        }

        public void setNote(String note) {
            this.note = note;
        }
    }

    @FXML
    private Button refreshButton;

    @FXML
    private Button openExcelButton;

    @FXML
    private Button copyJsonButton;

    @FXML
    private Label statusLabel;

    @FXML
    private TextField resolvedPathField;

    @FXML
    private Label fileExistsLabel;

    @FXML
    private Label cwdLabel;

    @FXML
    private GridPane envGrid;

    @FXML
    private GridPane mainGrid;

    @FXML
    private GridPane attendanceGrid;

    @FXML
    private TableView<SheetCheckRow> sheetTable;

    @FXML
    private TableColumn<SheetCheckRow, String> colCategory;

    @FXML
    private TableColumn<SheetCheckRow, String> colSheetName;

    @FXML
    private TableColumn<SheetCheckRow, String> colPresent;

    @FXML
    private TableColumn<SheetCheckRow, String> colNote;

    @FXML
    private ListView<String> warningsList;

    private MainShellController shell;

    private Supplier<RunRequest> requestFactory;

    private String lastRawJson = "";

    private Path lastResolvedPath;

    @FXML
    private void initialize() {
        colCategory.setCellValueFactory(new PropertyValueFactory<>("category"));
        colSheetName.setCellValueFactory(new PropertyValueFactory<>("sheetName"));
        colPresent.setCellValueFactory(new PropertyValueFactory<>("status"));
        colNote.setCellValueFactory(new PropertyValueFactory<>("note"));
        sheetTable.setItems(FXCollections.observableArrayList());
        warningsList.setItems(FXCollections.observableArrayList());
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.requestFactory = shell::buildMasterReadSummaryRequest;
    }

    @FXML
    private void onRefreshAction() {
        refreshButton.setDisable(true);
        openExcelButton.setDisable(true);
        statusLabel.setText("\u53d6\u5f97\u4e2d...");
        RunRequest req = requestFactory.get();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            refreshButton.setDisable(false);
                                            openExcelButton.setDisable(false);
                                            if (err != null) {
                                                statusLabel.setText("error: " + err.getMessage());
                                                shell.appendLog("[master-summary] " + err.getMessage());
                                                clearDisplay();
                                                return;
                                            }
                                            if (cap == null) {
                                                statusLabel.setText("no result");
                                                return;
                                            }
                                            statusLabel.setText("exit=" + cap.exitCode());
                                            applyStdout(cap.stdout(), cap.exitCode());
                                        }));
    }

    private void clearDisplay() {
        resolvedPathField.clear();
        fileExistsLabel.setText("");
        cwdLabel.setText("");
        envGrid.getChildren().clear();
        mainGrid.getChildren().clear();
        attendanceGrid.getChildren().clear();
        sheetTable.getItems().clear();
        warningsList.getItems().clear();
        lastRawJson = "";
        lastResolvedPath = null;
    }

    private void applyStdout(String stdout, int exitCode) {
        String trimmed = stdout != null ? stdout.trim() : "";
        lastRawJson = trimmed;
        if (trimmed.isEmpty()) {
            clearDisplay();
            statusLabel.setText("exit=" + exitCode + " (empty stdout)");
            return;
        }
        try {
            JsonNode root = JSON.readTree(trimmed);
            applyJson(root);
            statusLabel.setText(
                    "exit="
                            + exitCode
                            + (root.path("ok").asBoolean(true) ? " OK" : " (\u8981\u78ba\u8a8d)"));
        } catch (Exception e) {
            statusLabel.setText("JSON parse error: " + e.getMessage());
            shell.appendLog("[master-summary] parse: " + e.getMessage());
            clearDisplay();
        }
    }

    private void applyJson(JsonNode root) {
        resolvedPathField.setText(root.path("resolved_path").asText(""));
        boolean exists = root.path("file_exists").asBoolean(false);
        fileExistsLabel.setText(exists ? "\u30d5\u30a1\u30a4\u30eb: \u3042\u308a" : "\u30d5\u30a1\u30a4\u30eb: \u306a\u3057");
        cwdLabel.setText("cwd: " + root.path("cwd").asText(""));

        String rp = root.path("resolved_path").asText("");
        if (!rp.isEmpty()) {
            try {
                lastResolvedPath = Path.of(rp).toAbsolutePath().normalize();
            } catch (Exception ignored) {
                lastResolvedPath = null;
            }
        } else {
            lastResolvedPath = null;
        }

        envGrid.getChildren().clear();
        int row = 0;
        addGridRow(
                envGrid,
                row++,
                "MASTER_USE_SPEED_SHEET",
                root.path("master_use_speed_sheet_env").asText(""));
        JsonNode speed = root.path("speed");
        boolean sen = speed.path("enabled").asBoolean(false);
        addGridRow(
                envGrid,
                row++,
                "\u52a0\u5de5\u901f\u5ea6\u4e0a\u66f8\u304d\uff08speed\uff09",
                sen ? "\u6709\u52b9" : "\u7121\u52b9");
        addGridRow(
                envGrid,
                row++,
                "speed \u30b7\u30fc\u30c8",
                speed.path("sheet_name").asText(""));
        addGridRow(
                envGrid,
                row++,
                "\u30c7\u30fc\u30bf\u958b\u59cb\u5217\uff081\u8d77\u7b97\uff09",
                String.valueOf(speed.path("first_data_col_1based").asInt(0)));
        addGridRow(
                envGrid,
                row++,
                "\u8aad\u307f\u8fbc\u307f\u4ef6\u6570\uff08\u5de5\u7a0b+\u6a5f\u68b0\u30ad\u30fc\uff09",
                String.valueOf(speed.path("lookup_entry_count").asInt(0)));

        mainGrid.getChildren().clear();
        row = 0;
        JsonNode ms = root.path("main_sheet");
        addGridRow(
                mainGrid,
                row++,
                "\u89e3\u6c7a\u30b7\u30fc\u30c8\u540d",
                ms.path("resolved_name").asText("\u2014"));
        JsonNode fo = ms.path("factory_operating");
        addGridRow(
                mainGrid,
                row++,
                "\u5de5\u5834\u67a2 A12",
                formatTimeCell(fo, "a12", fo.path("effective").asBoolean(false)));
        addGridRow(
                mainGrid,
                row++,
                "\u5de5\u5834\u67a2 B12",
                formatTimeCell(fo, "b12", fo.path("effective").asBoolean(false)));
        JsonNode rs = ms.path("regular_shift");
        addGridRow(
                mainGrid,
                row++,
                "\u5b9a\u5e38 A15",
                formatTimeCell(rs, "a15", rs.path("effective").asBoolean(false)));
        addGridRow(
                mainGrid,
                row++,
                "\u5b9a\u5e38 B15",
                formatTimeCell(rs, "b15", rs.path("effective").asBoolean(false)));

        ObservableList<SheetCheckRow> sheetRows = FXCollections.observableArrayList();
        for (JsonNode ch : root.withArray("sheet_checks")) {
            SheetCheckRow r = new SheetCheckRow();
            String k = ch.path("key").asText("");
            r.setCategory(sheetKeyLabelJa(k));
            r.setSheetName(ch.path("sheet_name").asText(""));
            boolean pr = ch.path("present").asBoolean(false);
            r.setStatus(pr ? "\u3042\u308a" : "\u306a\u3057");
            String note = ch.path("note").asText("");
            r.setNote("\u2014".equals(note) || note.isEmpty() ? (pr ? "" : "\u306a\u3057") : note);
            sheetRows.add(r);
        }
        sheetTable.setItems(sheetRows);

        attendanceGrid.getChildren().clear();
        JsonNode att = root.path("attendance");
        addGridRow(
                attendanceGrid,
                0,
                "skills \u30e1\u30f3\u30d0\u30fc\u6570",
                String.valueOf(att.path("skills_member_count").asInt(0)));
        addGridRow(
                attendanceGrid,
                1,
                "\u52e4\u6020\u30b7\u30fc\u30c8\u6570\uff08\u4e00\u81f4\uff09",
                String.valueOf(att.path("attendance_sheets_matched").asInt(0)));

        ObservableList<String> warns = FXCollections.observableArrayList();
        for (JsonNode w : root.withArray("warnings")) {
            warns.add(w.asText(""));
        }
        if (root.path("openpyxl_skip").asBoolean(false)) {
            warns.add(
                    0,
                    "openpyxl: incompatible workbook marker (\u8a73\u7d30\u306f Python \u5074\u8b66\u544a)");
        }
        if (warns.isEmpty()) {
            warns.add("\u8b66\u544a\u306a\u3057");
        }
        warningsList.setItems(warns);
    }

    private static String formatTimeCell(JsonNode parent, String field, boolean effective) {
        String t = parent.path(field).asText("");
        if (t == null || t.isEmpty() || "null".equals(t)) {
            return effective ? "\u2014" : "\uff08\u672a\u8a2d\u5b9a\u30fb\u65e2\u5b9a\u5024\u4f7f\u7528\u53ef\uff09";
        }
        return t;
    }

    private static void addGridRow(GridPane grid, int row, String label, String value) {
        grid.add(new Label(label), 0, row);
        Label val = new Label(value != null ? value : "");
        val.setWrapText(true);
        grid.add(val, 1, row);
    }

    private static String sheetKeyLabelJa(String key) {
        if (key == null) {
            return "";
        }
        return switch (key) {
            case "skills" -> "skills";
            case "need" -> "need";
            case "machine_calendar" -> "\u6a5f\u68b0\u30ab\u30ec\u30f3\u30c0\u30fc";
            case "team_combinations" -> "\u7d44\u307f\u5408\u308f\u305b\u8868";
            case "speed" -> "speed";
            case "machine_daily_startup" -> "\u8a2d\u5b9a_\u6a5f\u68b0_\u65e5\u6b21\u59cb\u696d\u6e96\u5099";
            default -> key;
        };
    }

    @FXML
    private void onOpenExcelAction() {
        Path target = lastResolvedPath;
        if (target == null || !Files.isRegularFile(target)) {
            Map<String, String> ui = shell.snapshotUiEnv();
            target =
                    AppPaths.resolveMasterWorkbookPathResolved(
                            ui, shell.effectiveTaskInputWorkbookPathForShell());
        }
        if (!Files.isRegularFile(target)) {
            shell.appendLog("[master-summary] Excel: file not found: " + target);
            statusLabel.setText("\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093");
            return;
        }
        try {
            DesktopFileOpener.openFile(target);
            shell.appendLog("[master-summary] opened: " + target);
        } catch (Exception e) {
            shell.appendLog("[master-summary] open failed: " + e.getMessage());
            statusLabel.setText("open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onCopyJsonAction() {
        if (lastRawJson == null || lastRawJson.isEmpty()) {
            return;
        }
        ClipboardContent cc = new ClipboardContent();
        cc.putString(lastRawJson);
        Clipboard.getSystemClipboard().setContent(cc);
        shell.appendLog("[master-summary] JSON copied to clipboard");
    }

    static String scriptName() {
        return SCRIPT;
    }
}
