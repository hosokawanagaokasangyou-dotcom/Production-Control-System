package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
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
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** Master workbook read summary (Python JSON); layout {@code MasterReadSummaryTab.fxml}. */
public final class MasterReadSummaryTabController {

    /** Cursor debug session NDJSON (agent-debug-ndjson-logging.mdc). */
    private static final String AGENT_DEBUG_SESSION_ID = "6f550e";

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String SCRIPT = "master_read_summary.py";

    /**
     * {@link JsonNode#path(String)} が MissingNode を返してもよいが、{@link
     * com.fasterxml.jackson.databind.node.ObjectNode#withArray(String)} は ObjectNode 専用で欠損時に例外になる。
     */
    private static boolean notObjectSection(JsonNode n) {
        return n == null || n.isMissingNode() || n.isNull() || !n.isObject();
    }

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

    /** speed.lookup_sample row for {@link #speedSampleTable}. */
    public static final class SpeedSampleRow {
        private final String process;
        private final String machine;
        private final String speedMPerMin;

        public SpeedSampleRow(String process, String machine, String speedMPerMin) {
            this.process = process;
            this.machine = machine;
            this.speedMPerMin = speedMPerMin;
        }

        public String getProcess() {
            return process;
        }

        public String getMachine() {
            return machine;
        }

        public String getSpeedMPerMin() {
            return speedMPerMin;
        }
    }

    /** need_base_required_sample row. */
    public static final class NeedBaseRow {
        private final String combo;
        private final String required;

        public NeedBaseRow(String combo, String required) {
            this.combo = combo;
            this.required = required;
        }

        public String getCombo() {
            return combo;
        }

        public String getRequired() {
            return required;
        }
    }

    /** need_rules_detail row. */
    public static final class NeedRuleRow {
        private final String order;
        private final String condition;
        private final String overrideCount;

        public NeedRuleRow(String order, String condition, String overrideCount) {
            this.order = order;
            this.condition = condition;
            this.overrideCount = overrideCount;
        }

        public String getOrder() {
            return order;
        }

        public String getCondition() {
            return condition;
        }

        public String getOverrideCount() {
            return overrideCount;
        }
    }

    /** calendar_top_dates row. */
    public static final class CalendarTopRow {
        private final String date;
        private final String intervals;

        public CalendarTopRow(String date, String intervals) {
            this.date = date;
            this.intervals = intervals;
        }

        public String getDate() {
            return date;
        }

        public String getIntervals() {
            return intervals;
        }
    }

    /** exclude_rules_sheet.rules_sample row. */
    public static final class ExcludeRuleSampleRow {
        private final String process;
        private final String machine;
        private final String excludeFlag;
        private final String logicParsedShort;

        public ExcludeRuleSampleRow(
                String process, String machine, String excludeFlag, String logicParsedShort) {
            this.process = process;
            this.machine = machine;
            this.excludeFlag = excludeFlag;
            this.logicParsedShort = logicParsedShort;
        }

        public String getProcess() {
            return process;
        }

        public String getMachine() {
            return machine;
        }

        public String getExcludeFlag() {
            return excludeFlag;
        }

        public String getLogicParsedShort() {
            return logicParsedShort;
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
    private GridPane skillsNeedGrid;

    @FXML
    private GridPane teamComboGrid;

    @FXML
    private GridPane machineCalendarGrid;

    @FXML
    private GridPane appConfigGrid;

    @FXML
    private GridPane excludeRulesGrid;

    @FXML
    private TableView<ExcludeRuleSampleRow> excludeRulesSampleTable;

    @FXML
    private TableColumn<ExcludeRuleSampleRow, String> colExProcess;

    @FXML
    private TableColumn<ExcludeRuleSampleRow, String> colExMachine;

    @FXML
    private TableColumn<ExcludeRuleSampleRow, String> colExFlag;

    @FXML
    private TableColumn<ExcludeRuleSampleRow, String> colExParsed;

    @FXML
    private GridPane planningConstantsGrid;

    @FXML
    private TableView<SpeedSampleRow> speedSampleTable;

    @FXML
    private TableColumn<SpeedSampleRow, String> colSpeedProcess;

    @FXML
    private TableColumn<SpeedSampleRow, String> colSpeedMachine;

    @FXML
    private TableColumn<SpeedSampleRow, String> colSpeedMpm;

    @FXML
    private TableView<NeedBaseRow> needBaseTable;

    @FXML
    private TableColumn<NeedBaseRow, String> colNeedCombo;

    @FXML
    private TableColumn<NeedBaseRow, String> colNeedReq;

    @FXML
    private TableView<NeedRuleRow> needRulesTable;

    @FXML
    private TableColumn<NeedRuleRow, String> colRuleOrder;

    @FXML
    private TableColumn<NeedRuleRow, String> colRuleCond;

    @FXML
    private TableColumn<NeedRuleRow, String> colRuleOv;

    @FXML
    private TableView<CalendarTopRow> calendarTopDatesTable;

    @FXML
    private TableColumn<CalendarTopRow, String> colCalDate;

    @FXML
    private TableColumn<CalendarTopRow, String> colCalIv;

    @FXML
    private ListView<String> allSheetsList;

    @FXML
    private ListView<String> attendanceSheetsList;

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
        allSheetsList.setItems(FXCollections.observableArrayList());
        attendanceSheetsList.setItems(FXCollections.observableArrayList());
        speedSampleTable.setItems(FXCollections.observableArrayList());
        needBaseTable.setItems(FXCollections.observableArrayList());
        needRulesTable.setItems(FXCollections.observableArrayList());
        calendarTopDatesTable.setItems(FXCollections.observableArrayList());
        colSpeedProcess.setCellValueFactory(new PropertyValueFactory<>("process"));
        colSpeedMachine.setCellValueFactory(new PropertyValueFactory<>("machine"));
        colSpeedMpm.setCellValueFactory(new PropertyValueFactory<>("speedMPerMin"));
        colNeedCombo.setCellValueFactory(new PropertyValueFactory<>("combo"));
        colNeedReq.setCellValueFactory(new PropertyValueFactory<>("required"));
        colRuleOrder.setCellValueFactory(new PropertyValueFactory<>("order"));
        colRuleCond.setCellValueFactory(new PropertyValueFactory<>("condition"));
        colRuleOv.setCellValueFactory(new PropertyValueFactory<>("overrideCount"));
        colCalDate.setCellValueFactory(new PropertyValueFactory<>("date"));
        colCalIv.setCellValueFactory(new PropertyValueFactory<>("intervals"));
        excludeRulesSampleTable.setItems(FXCollections.observableArrayList());
        colExProcess.setCellValueFactory(new PropertyValueFactory<>("process"));
        colExMachine.setCellValueFactory(new PropertyValueFactory<>("machine"));
        colExFlag.setCellValueFactory(new PropertyValueFactory<>("excludeFlag"));
        colExParsed.setCellValueFactory(new PropertyValueFactory<>("logicParsedShort"));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.requestFactory = shell::buildMasterReadSummaryRequest;
    }

    @FXML
    private void onRefreshAction() {
        refreshButton.setDisable(true);
        openExcelButton.setDisable(true);
        statusLabel.setText("取得中...");
        RunRequest req = requestFactory.get();
        // #region agent log
        if (shell != null) {
            Map<String, Object> data = new LinkedHashMap<>();
            data.put("pythonExe", req.pythonExecutable().toString());
            data.put("scriptDir", req.scriptDirectory().toString());
            data.put("scriptFile", req.scriptFileName());
            String wb = req.taskInputWorkbook();
            data.put("workbookArgChars", wb != null ? wb.length() : 0);
            data.put(
                    "workbookArgTail",
                    wb != null && wb.length() > 120 ? wb.substring(wb.length() - 120) : wb);
            AgentDebugLog.appendStructured(
                    shell.snapshotUiEnv(),
                    AGENT_DEBUG_SESSION_ID,
                    "D",
                    "MasterReadSummaryTabController.onRefreshAction",
                    "RunRequest before PythonProcessRunner.runCaptureAsync",
                    data);
        }
        // #endregion
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
                                                // #region agent log
                                                Map<String, Object> fail = new LinkedHashMap<>();
                                                fail.put("errorClass", err.getClass().getSimpleName());
                                                fail.put("errorMessage", err.getMessage());
                                                AgentDebugLog.appendStructured(
                                                        shell.snapshotUiEnv(),
                                                        AGENT_DEBUG_SESSION_ID,
                                                        "D",
                                                        "MasterReadSummaryTabController.onRefreshAction",
                                                        "runCaptureAsync failed before stdout",
                                                        fail);
                                                // #endregion
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
        skillsNeedGrid.getChildren().clear();
        teamComboGrid.getChildren().clear();
        machineCalendarGrid.getChildren().clear();
        appConfigGrid.getChildren().clear();
        excludeRulesGrid.getChildren().clear();
        if (excludeRulesSampleTable != null) {
            excludeRulesSampleTable.getItems().clear();
        }
        planningConstantsGrid.getChildren().clear();
        speedSampleTable.getItems().clear();
        needBaseTable.getItems().clear();
        needRulesTable.getItems().clear();
        calendarTopDatesTable.getItems().clear();
        allSheetsList.getItems().clear();
        attendanceSheetsList.getItems().clear();
        sheetTable.getItems().clear();
        warningsList.getItems().clear();
        lastRawJson = "";
        lastResolvedPath = null;
    }

    private void applyStdout(String stdout, int exitCode) {
        String trimmed = stdout != null ? stdout.trim() : "";
        String jsonPayload = extractJsonPayload(trimmed);
        lastRawJson = jsonPayload;
        // #region agent log
        {
            Map<String, Object> data = new LinkedHashMap<>();
            data.put("exitCode", exitCode);
            data.put("stdoutChars", trimmed.length());
            String[] split = trimmed.isEmpty() ? new String[0] : trimmed.split("\\R", -1);
            data.put("stdoutLineCount", split.length);
            data.put("hasBraceJsonLine", hasExtractableJsonLine(trimmed));
            data.put("stdoutContainsBrace", trimmed.contains("{"));
            data.put("jsonPayloadChars", jsonPayload.length());
            data.put("jsonPayloadStartsTraceback", jsonPayload.trim().startsWith("Traceback"));
            data.put("jsonPayloadPreview", debugPreview(jsonPayload, 500));
            AgentDebugLog.appendStructured(
                    shell != null ? shell.snapshotUiEnv() : Map.of(),
                    AGENT_DEBUG_SESSION_ID,
                    "A",
                    "MasterReadSummaryTabController.applyStdout",
                    "stdout summary before JSON.readTree",
                    data);
        }
        // #endregion
        if (jsonPayload.isEmpty()) {
            clearDisplay();
            statusLabel.setText("exit=" + exitCode + " (empty stdout)");
            return;
        }
        try {
            JsonNode root = JSON.readTree(jsonPayload);
            applyJson(root);
            statusLabel.setText(
                    "exit="
                            + exitCode
                            + (root.path("ok").asBoolean(true) ? " OK" : " (要確認)"));
        } catch (Exception e) {
            statusLabel.setText("JSON parse error: " + e.getMessage());
            shell.appendLog("[master-summary] parse: " + e.getMessage());
            // #region agent log
            Map<String, Object> errData = new LinkedHashMap<>();
            errData.put("exceptionClass", e.getClass().getSimpleName());
            errData.put("exceptionMessage", e.getMessage());
            errData.put(
                    "mergedStderrIntoStdoutNote",
                    "PythonProcessRunner merges stderr into stdout (redirectErrorStream)");
            AgentDebugLog.appendStructured(
                    shell != null ? shell.snapshotUiEnv() : Map.of(),
                    AGENT_DEBUG_SESSION_ID,
                    "B",
                    "MasterReadSummaryTabController.applyStdout",
                    "JSON.parse failed",
                    errData);
            // #endregion
            clearDisplay();
        }
    }

    /** Same brace-line detection as {@link #extractJsonPayload(String)} (single-line JSON object). */
    private static boolean hasExtractableJsonLine(String stdout) {
        if (stdout == null || stdout.isBlank()) {
            return false;
        }
        String[] lines = stdout.split("\\R", -1);
        for (int i = lines.length - 1; i >= 0; i--) {
            String t = lines[i].trim();
            if (t.length() >= 2 && t.startsWith("{") && t.endsWith("}")) {
                return true;
            }
        }
        return false;
    }

    private static String debugPreview(String s, int maxChars) {
        if (s == null) {
            return "";
        }
        String t = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n");
        if (t.length() <= maxChars) {
            return t;
        }
        return t.substring(0, maxChars) + "...(trunc)";
    }

    /**
     * planning_core may emit log lines to stdout before JSON; take the last line that looks like one JSON
     * object (starts with "{", ends with "}").
     */
    static String extractJsonPayload(String stdout) {
        if (stdout == null || stdout.isBlank()) {
            return "";
        }
        String[] lines = stdout.split("\\R", -1);
        for (int i = lines.length - 1; i >= 0; i--) {
            String t = lines[i].trim();
            if (t.length() >= 2 && t.startsWith("{") && t.endsWith("}")) {
                return t;
            }
        }
        return stdout.trim();
    }

    private void applyJson(JsonNode root) {
        resolvedPathField.setText(root.path("resolved_path").asText(""));
        boolean exists = root.path("file_exists").asBoolean(false);
        fileExistsLabel.setText(exists ? "ファイル: あり" : "ファイル: なし");
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
        int         row = 0;
        addGridRow(
                envGrid,
                row++,
                "MASTER_WORKBOOK_FILE",
                root.path("master_workbook_file_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "PM_AI_MASTER_WORKBOOK",
                root.path("pm_ai_master_workbook_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "MASTER_USE_SPEED_SHEET",
                root.path("master_use_speed_sheet_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "TEAM_ASSIGN_USE_MASTER_COMBO_SHEET",
                root.path("team_assign_use_master_combo_sheet_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW",
                root.path("team_assign_ignore_need_surplus_row_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "MASTER_SPEED_SHEET_NAME",
                root.path("speed").path("master_speed_sheet_name_env").asText(""));
        addGridRow(
                envGrid,
                row++,
                "PM_AI_EXCLUDE_RULES_JSON",
                root.path("pm_ai_exclude_rules_json_env").asText(""));
        JsonNode speed = root.path("speed");
        boolean sen = speed.path("enabled").asBoolean(false);
        addGridRow(
                envGrid,
                row++,
                "加工速度上書き（speed）",
                sen ? "有効" : "無効");
        addGridRow(
                envGrid,
                row++,
                "speed シート",
                speed.path("sheet_name").asText(""));
        addGridRow(
                envGrid,
                row++,
                "データ開始列（1起算）",
                String.valueOf(speed.path("first_data_col_1based").asInt(0)));
        addGridRow(
                envGrid,
                row++,
                "読み込み件数（工程+機械キー）",
                String.valueOf(speed.path("lookup_entry_count").asInt(0)));

        applySkillsNeedSection(root.path("skills_need"));

        applyTeamComboSection(root.path("team_combinations"));

        mainGrid.getChildren().clear();
        row = 0;
        JsonNode ms = root.path("main_sheet");
        addGridRow(
                mainGrid,
                row++,
                "解決シート名",
                ms.path("resolved_name").asText("—"));
        JsonNode fo = ms.path("factory_operating");
        addGridRow(
                mainGrid,
                row++,
                "工場枢 A12",
                formatTimeCell(fo, "a12", fo.path("effective").asBoolean(false)));
        addGridRow(
                mainGrid,
                row++,
                "工場枢 B12",
                formatTimeCell(fo, "b12", fo.path("effective").asBoolean(false)));
        JsonNode rs = ms.path("regular_shift");
        addGridRow(
                mainGrid,
                row++,
                "定常 A15",
                formatTimeCell(rs, "a15", rs.path("effective").asBoolean(false)));
        addGridRow(
                mainGrid,
                row++,
                "定常 B15",
                formatTimeCell(rs, "b15", rs.path("effective").asBoolean(false)));

        ObservableList<SheetCheckRow> sheetRows = FXCollections.observableArrayList();
        for (JsonNode ch : root.path("sheet_checks")) {
            SheetCheckRow r = new SheetCheckRow();
            String k = ch.path("key").asText("");
            r.setCategory(sheetKeyLabelJa(k));
            r.setSheetName(ch.path("sheet_name").asText(""));
            boolean pr = ch.path("present").asBoolean(false);
            r.setStatus(pr ? "あり" : "なし");
            String note = ch.path("note").asText("");
            r.setNote("—".equals(note) || note.isEmpty() ? (pr ? "" : "なし") : note);
            sheetRows.add(r);
        }
        sheetTable.setItems(sheetRows);

        applyAppConfigSection(root.path("app_config_sheet"));
        applyExcludeRulesSection(root.path("exclude_rules_sheet"));
        applyPlanningConstantsSection(root.path("planning_constants"));
        applyDetailTables(root);
        applyAllSheetsList(root.path("all_sheet_names"));

        applyMachineCalendarSection(root);

        attendanceGrid.getChildren().clear();
        JsonNode att = root.path("attendance");
        addGridRow(
                attendanceGrid,
                0,
                "skills メンバー数",
                String.valueOf(att.path("skills_member_count").asInt(0)));
        addGridRow(
                attendanceGrid,
                1,
                "勤怠シート数（一致）",
                String.valueOf(att.path("attendance_sheets_matched").asInt(0)));
        applyAttendanceSheetsList(att.path("matched_sheet_names"));

        ObservableList<String> warns = FXCollections.observableArrayList();
        for (JsonNode w : root.path("warnings")) {
            warns.add(w.asText(""));
        }
        if (root.path("openpyxl_skip").asBoolean(false)) {
            warns.add(
                    0,
                    "openpyxl: incompatible workbook marker (詳細は Python 側警告)");
        }
        if (warns.isEmpty()) {
            warns.add("警告なし");
        }
        warningsList.setItems(warns);
    }

    private void applySkillsNeedSection(JsonNode sn) {
        skillsNeedGrid.getChildren().clear();
        if (notObjectSection(sn)) {
            addGridRow(skillsNeedGrid, 0, "状態", "—");
            return;
        }
        int row = 0;
        String fmt = sn.path("skills_sheet_format").asText("");
        addGridRow(
                skillsNeedGrid,
                row++,
                "skills 形式",
                formatSkillsSheetFormatJa(fmt));
        if (sn.hasNonNull("skip_reason")) {
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "スキップ",
                    formatSkillsNeedSkipJa(sn.path("skip_reason").asText("")));
        }
        if (sn.hasNonNull("validation_error")) {
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "検証エラー",
                    sn.path("validation_error").asText(""));
        }
        if (sn.hasNonNull("parse_error")) {
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "解析エラー",
                    sn.path("parse_error").asText(""));
        }
        boolean loaded = sn.path("loaded").asBoolean(false);
        addGridRow(
                skillsNeedGrid,
                row++,
                "load_skills_and_needs",
                loaded ? "成功（データあり）" : "未収集または空");
        if (sn.path("empty_result").asBoolean(false)) {
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "補足",
                    sn.path("note").asText(
                            "メンバー・設備列が空"));
        }
        if (loaded) {
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "メンバー数",
                    String.valueOf(sn.path("members_count").asInt(0)));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "設備列（工程×機械等）",
                    String.valueOf(sn.path("equipment_columns").asInt(0)));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "need 工程×機械列数",
                    String.valueOf(sn.path("need_combo_columns").asInt(0)));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "特別指定ルール数",
                    String.valueOf(sn.path("need_rules_count").asInt(0)));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "配台時追加人数（組合せキー）",
                    String.valueOf(sn.path("surplus_combo_entries").asInt(0)));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "メンバー（最多40人）",
                    jsonArrayJoin(sn.path("member_names_sample")));
            addGridRow(
                    skillsNeedGrid,
                    row++,
                    "need キー（最多40）",
                    jsonArrayJoin(sn.path("need_combo_keys_sample")));
            if (sn.has("equipment_columns_sample")) {
                addGridRow(
                        skillsNeedGrid,
                        row++,
                        "設備列サンプル（先頭48）",
                        jsonArrayJoin(sn.path("equipment_columns_sample")));
            }
            JsonNode surplusSample = sn.path("surplus_sample");
            if (surplusSample.isArray() && surplusSample.size() > 0) {
                addGridRow(
                        skillsNeedGrid,
                        row++,
                        "配台時追加上限サンプル",
                        String.valueOf(surplusSample.size())
                                + " 件");
            }
        }
    }

    private void applyAppConfigSection(JsonNode ac) {
        appConfigGrid.getChildren().clear();
        if (notObjectSection(ac)) {
            addGridRow(appConfigGrid, 0, "状態", "—");
            return;
        }
        int row = 0;
        if (!ac.path("present").asBoolean(false)) {
            addGridRow(
                    appConfigGrid,
                    row++,
                    "シート",
                    ac.path("note").asText("なし"));
            return;
        }
        if (ac.path("openpyxl_cells_unreadable").asBoolean(false)) {
            addGridRow(
                    appConfigGrid,
                    row++,
                    "注意",
                    ac.path("note").asText("openpyxl"));
        }
        addGridRow(
                appConfigGrid,
                row++,
                "A列 依頼NO（トレース）件数",
                String.valueOf(ac.path("trace_task_ids_count").asInt(0)));
        addGridRow(
                appConfigGrid,
                row++,
                "A列 サンプル",
                jsonArrayJoin(ac.path("trace_task_ids_sample")));
        addGridRow(
                appConfigGrid,
                row++,
                "B列 依頼NO（デバッグ）件数",
                String.valueOf(ac.path("debug_task_ids_count").asInt(0)));
        addGridRow(
                appConfigGrid,
                row++,
                "B列 サンプル",
                jsonArrayJoin(ac.path("debug_task_ids_sample")));
        addGridRow(
                appConfigGrid,
                row++,
                "Gemini 有効モデル数",
                String.valueOf(ac.path("gemini_models_enabled_count").asInt(0)));
        addGridRow(
                appConfigGrid,
                row++,
                "Gemini モデルリスト",
                jsonArrayJoin(ac.path("gemini_models_enabled_sample")));
    }

    private void applyExcludeRulesSection(JsonNode ex) {
        excludeRulesGrid.getChildren().clear();
        if (excludeRulesSampleTable != null) {
            excludeRulesSampleTable.getItems().clear();
        }
        if (notObjectSection(ex)) {
            addGridRow(excludeRulesGrid, 0, "状態", "—");
            return;
        }
        int row = 0;
        if (!ex.path("present").asBoolean(false)) {
            addGridRow(
                    excludeRulesGrid,
                    row++,
                    "シート",
                    ex.path("read_error")
                            .asText(ex.path("error")
                                    .asText(ex.path("note").asText("なし"))));
            return;
        }
        addGridRow(
                excludeRulesGrid,
                row++,
                "期待シート名",
                ex.path("expected_sheet_name").asText("—"));
        addGridRow(
                excludeRulesGrid,
                row++,
                "解決後シート名",
                ex.path("resolved_sheet_name").asText("—"));
        addGridRow(
                excludeRulesGrid,
                row++,
                "行数/列数",
                ex.path("rows").asInt(0) + " / " + ex.path("cols").asInt(0));
        addGridRow(
                excludeRulesGrid,
                row++,
                "ルール行（工程名あり）",
                String.valueOf(ex.path("rules_count").asInt(0)));
        addGridRow(
                excludeRulesGrid,
                row++,
                "最初5列が空でない行数",
                String.valueOf(ex.path("non_empty_rows_scan").asInt(0)));
        if (ex.path("openpyxl_io_skipped_by_marker").asBoolean(false)) {
            addGridRow(
                    excludeRulesGrid,
                    row++,
                    "openpyxl",
                    ex.path("note").asText("—"));
        }
        if (ex.hasNonNull("header_error")) {
            addGridRow(
                    excludeRulesGrid,
                    row++,
                    "見出し",
                    ex.path("header_error").asText(""));
        }
        String envNote =
                ex.path("pm_ai_exclude_rules_json_set").asBoolean(false)
                        ? "（JSON 優先のため段階1はExcel行を見ない場合あり）"
                        : "";
        addGridRow(
                excludeRulesGrid,
                row++,
                "PM_AI_EXCLUDE_RULES_JSON",
                (ex.path("pm_ai_exclude_rules_json_path").asText("").isEmpty()
                                ? "（未設定）"
                                : ex.path("pm_ai_exclude_rules_json_path").asText(""))
                        + envNote);
        addGridRow(
                excludeRulesGrid,
                row++,
                "段階1との関係",
                ex.path("stage1_effective_source_note").asText(""));

        ObservableList<ExcludeRuleSampleRow> exRows = FXCollections.observableArrayList();
        for (JsonNode r : ex.path("rules_sample")) {
            exRows.add(
                    new ExcludeRuleSampleRow(
                            r.path("process").asText(""),
                            r.path("machine").asText(""),
                            r.path("exclude_flag").asText(""),
                            summarizeExcludeParsedJson(r.path("logic_json_parsed"))));
        }
        if (excludeRulesSampleTable != null) {
            excludeRulesSampleTable.setItems(exRows);
        }
    }

    private static String summarizeExcludeParsedJson(JsonNode parsed) {
        if (parsed == null || parsed.isNull() || parsed.isMissingNode()) {
            return "";
        }
        if (parsed.isObject()) {
            int ver = parsed.path("version").asInt(0);
            String mode = parsed.path("mode").asText("");
            boolean reqAll = parsed.path("require_all").asBoolean(false);
            String tail = "";
            JsonNode conditions = parsed.path("conditions");
            if (conditions.isArray()) {
                tail = ", conds~" + conditions.size();
            }
            return "v" + ver + " " + mode + (mode.equals("conditions") ? " all=" + reqAll : "")
                    + tail;
        }
        return parsed.toString();
    }

    private void applyPlanningConstantsSection(JsonNode pc) {
        planningConstantsGrid.getChildren().clear();
        if (notObjectSection(pc)) {
            addGridRow(planningConstantsGrid, 0, "状態", "—");
            return;
        }
        JsonNode rr = pc.path("config_task_ids_row_range");
        String rng =
                rr.size() >= 2
                        ? rr.get(0).asInt() + "–" + rr.get(1).asInt()
                        : "—";
        addGridRow(
                planningConstantsGrid,
                0,
                "機械カレンダースロット（分）",
                String.valueOf(pc.path("machine_calendar_slot_minutes").asInt(0)));
        addGridRow(
                planningConstantsGrid,
                1,
                "解決後 speed シート名",
                pc.path("resolved_speed_sheet_name").asText(""));
        addGridRow(
                planningConstantsGrid,
                2,
                "設定 依頼NO行範囲（1起算 Excel行）",
                rng);
    }

    private void applyDetailTables(JsonNode root) {
        JsonNode sp = root.path("speed").path("lookup_sample");
        ObservableList<SpeedSampleRow> speedRows = FXCollections.observableArrayList();
        for (JsonNode r : sp) {
            speedRows.add(
                    new SpeedSampleRow(
                            r.path("process").asText(""),
                            r.path("machine").asText(""),
                            r.path("speed_m_per_min").asText("")));
        }
        speedSampleTable.setItems(speedRows);

        JsonNode sn = root.path("skills_need");
        ObservableList<NeedBaseRow> nb = FXCollections.observableArrayList();
        for (JsonNode r : sn.path("need_base_required_sample")) {
            nb.add(
                    new NeedBaseRow(
                            r.path("combo").asText(""),
                            String.valueOf(r.path("required").asInt(0))));
        }
        needBaseTable.setItems(nb);

        ObservableList<NeedRuleRow> nr = FXCollections.observableArrayList();
        for (JsonNode r : sn.path("need_rules_detail")) {
            nr.add(
                    new NeedRuleRow(
                            String.valueOf(r.path("order").asInt(0)),
                            r.path("condition").asText(""),
                            String.valueOf(r.path("override_entry_count").asInt(0))));
        }
        needRulesTable.setItems(nr);

        ObservableList<CalendarTopRow> ct = FXCollections.observableArrayList();
        for (JsonNode r : sn.path("calendar_top_dates")) {
            ct.add(
                    new CalendarTopRow(
                            r.path("date").asText(""),
                            String.valueOf(r.path("intervals").asInt(0))));
        }
        calendarTopDatesTable.setItems(ct);
    }

    private void applyAllSheetsList(JsonNode arr) {
        ObservableList<String> items = FXCollections.observableArrayList();
        if (arr != null && arr.isArray()) {
            for (JsonNode x : arr) {
                items.add(x.asText(""));
            }
        }
        allSheetsList.setItems(items);
    }

    private void applyAttendanceSheetsList(JsonNode arr) {
        ObservableList<String> items = FXCollections.observableArrayList();
        if (arr != null && arr.isArray()) {
            for (JsonNode x : arr) {
                items.add(x.asText(""));
            }
        }
        attendanceSheetsList.setItems(items);
    }

    private void applyTeamComboSection(JsonNode tc) {
        teamComboGrid.getChildren().clear();
        if (notObjectSection(tc)) {
            addGridRow(teamComboGrid, 0, "状態", "—");
            return;
        }
        int row = 0;
        if (tc.hasNonNull("load_error")) {
            addGridRow(
                    teamComboGrid,
                    row++,
                    "読込エラー",
                    tc.path("load_error").asText(""));
            return;
        }
        addGridRow(
                teamComboGrid,
                row++,
                "環境で組み合わせ表を使用",
                tc.path("env_use_sheet_enabled").asBoolean(false)
                        ? "はい"
                        : "いいえ");
        addGridRow(
                teamComboGrid,
                row++,
                "工程+機械キー数（別）",
                String.valueOf(tc.path("distinct_equipment_keys").asInt(0)));
        addGridRow(
                teamComboGrid,
                row++,
                "プリセット行合計",
                String.valueOf(tc.path("preset_rows_total").asInt(0)));
        addGridRow(
                teamComboGrid,
                row++,
                "サンプルキー（最多20）",
                jsonArrayJoin(tc.path("sample_equipment_keys")));
    }

    private void applyMachineCalendarSection(JsonNode root) {
        machineCalendarGrid.getChildren().clear();
        int row = 0;
        JsonNode mcs = root.path("machine_calendar_sheet");
        if (mcs.path("present").asBoolean(false)) {
            addGridRow(
                    machineCalendarGrid,
                    row++,
                    "シート行数/列数",
                    mcs.path("rows").asInt(0) + " / " + mcs.path("cols").asInt(0));
            if (mcs.has("equipment_like_columns")) {
                addGridRow(
                        machineCalendarGrid,
                        row++,
                        "設備列見込み",
                        String.valueOf(mcs.path("equipment_like_columns").asInt(0)));
            }
            if (mcs.has("two_row_equipment_header")) {
                addGridRow(
                        machineCalendarGrid,
                        row++,
                        "ヘッダ",
                        mcs.path("two_row_equipment_header").asBoolean(false)
                                ? "2行（工程+機械）"
                                : "1行");
            }
        } else {
            String note = mcs.path("note").asText("");
            String err = mcs.path("read_error").asText("");
            addGridRow(
                    machineCalendarGrid,
                    row++,
                    "機械カレンダー",
                    (!note.isEmpty() ? note : err.isEmpty() ? "なし" : err));
        }
        JsonNode sn = root.path("skills_need");
        if (sn.path("loaded").asBoolean(false)) {
            if (sn.has("machine_calendar_days_with_blocks")) {
                addGridRow(
                        machineCalendarGrid,
                        row++,
                        "占有がある日（論理同一）",
                        String.valueOf(sn.path("machine_calendar_days_with_blocks").asInt(0)));
            }
            if (sn.has("machine_calendar_intervals")) {
                addGridRow(
                        machineCalendarGrid,
                        row++,
                        "占有区間合計（論理同一）",
                        String.valueOf(sn.path("machine_calendar_intervals").asInt(0)));
            }
            if (sn.hasNonNull("machine_calendar_occ_error")) {
                addGridRow(
                        machineCalendarGrid,
                        row++,
                        "占有集計エラー",
                        sn.path("machine_calendar_occ_error").asText(""));
            }
        }
        JsonNode md = root.path("machine_daily_startup");
        if (md.path("present").asBoolean(false)) {
            addGridRow(
                    machineCalendarGrid,
                    row++,
                    "設定_機械_日次始業準備（データ行）",
                    String.valueOf(md.path("data_rows").asInt(0)));
        } else if (md.hasNonNull("read_error")) {
            addGridRow(
                    machineCalendarGrid,
                    row++,
                    "設定_機械_日次始業準備",
                    md.path("read_error").asText(""));
        } else {
            addGridRow(
                    machineCalendarGrid,
                    row++,
                    "設定_機械_日次始業準備",
                    "シートなし");
        }
    }

    private static String jsonArrayJoin(JsonNode arr) {
        if (arr == null || !arr.isArray() || arr.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (JsonNode x : arr) {
            if (sb.length() > 0) {
                sb.append(", ");
            }
            sb.append(x.asText(""));
        }
        return sb.toString();
    }

    private static String formatSkillsSheetFormatJa(String fmt) {
        if (fmt == null || fmt.isEmpty()) {
            return "—";
        }
        return switch (fmt) {
            case "two_row_header" -> "2行ヘッダ（工程+機械）";
            case "single_row_header" -> "1行ヘッダ（旧互換）";
            case "unknown" -> "不明";
            default -> fmt;
        };
    }

    private static String formatSkillsNeedSkipJa(String key) {
        if (key == null || key.isEmpty()) {
            return "";
        }
        return switch (key) {
            case "openpyxl_incompatible_marker" ->
                    "openpyxl 非対話マーカーのため未実行";
            case "missing_skills_or_need_sheet" ->
                    "skills / need シートがありません";
            default -> key;
        };
    }

    private static String formatTimeCell(JsonNode parent, String field, boolean effective) {
        String t = parent.path(field).asText("");
        if (t == null || t.isEmpty() || "null".equals(t)) {
            return effective ? "—" : "（未設定・既定値使用可）";
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
            case "machine_calendar" -> "機械カレンダー";
            case "team_combinations" -> "組み合わせ表";
            case "speed" -> "speed";
            case "machine_daily_startup" -> "設定_機械_日次始業準備";
            case "app_config" -> "設定";
            case "exclude_dispatch_rules" -> "設定_配台不要工程";
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
            statusLabel.setText("ファイルが見つかりません");
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
