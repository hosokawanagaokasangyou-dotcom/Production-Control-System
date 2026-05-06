package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;

/**
 * PM_AI_EXCLUDE_RULES_JSON editor; static structure in {@code ExcludeRulesTab.fxml}.
 *
 * <p>ルールは JSON（配列または {@code {"rules":[]}}）と表の両方から編集でき、保存時は直近で編集した側をファイルへ書き出します。
 */
public final class ExcludeRulesTabController {

    private static final String AGENT_DEBUG_SESSION_ID = "504628";

    /** Python {@code planning_core._core} と同じ列名（Excel 見出し）。 */
    private static final String COL_PROCESS = "工程名";

    private static final String COL_MACHINE = "機械名";

    private static final String COL_FLAG = "配台不要";

    private static final String COL_LOGIC_JA = "配台不要ロジック";

    /** 昔のブック／JSON で誤記された見出しに対する読み取りフォールバック。 */
    private static final String COL_LOGIC_JA_LEGACY = "配台不能ロジック";

    private static final String COL_LOGIC_JSON = "ロジック式";

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private static final String HINT_TEXT =
            "PM_AI_EXCLUDE_RULES_JSON が実在かつ有効なら"
                    + " 設定_配台不要工程 の Excel 保守を省略可。"
                    + " 環境変数タブに同名を追加してパスを共有してください。"
                    + " 表と JSON の両方で編集できます。保存は直近で触れた側を書き出します。";

    private Stage ownerStage;

    private MainShellController shell;

    /** {@link #bodyArea} へのプログラム代入ではテキスト「変更」とみなさない。 */
    private boolean suppressBodyDirty;

    /** true のとき保存前に表から JSON を生成する（表で編集したあと）。 */
    private boolean preferTableOnSave;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField pathField;

    @FXML
    private TextArea bodyArea;

    @FXML
    private SplitPane editorSplit;

    @FXML
    private TableView<ExcludeRuleTableRow> rulesTable;

    private final ObservableList<ExcludeRuleTableRow> ruleRows = FXCollections.observableArrayList();

    @FXML
    private void initialize() {
        pathField.setPromptText("PM_AI_EXCLUDE_RULES_JSON — .json フルパス");
        bodyArea.setPromptText(
                "[\n  { \"工程名\": \"...\", \"機械名\": \"...\", ... }\n]\n"
                        + "または {\"rules\":[...]}");
        hintLabel.setText(HINT_TEXT);
        bodyArea.setStyle("-fx-font-family: monospace");

        rulesTable.setItems(ruleRows);
        rulesTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        rulesTable.setEditable(true);

        TableColumn<ExcludeRuleTableRow, String> colP = new TableColumn<>(COL_PROCESS);
        colP.setCellValueFactory(d -> d.getValue().processProperty());
        colP.setCellFactory(TextFieldTableCell.forTableColumn());
        colP.setOnEditCommit(e -> {
            e.getRowValue().processProperty().set(e.getNewValue());
            tableEdited();
        });
        colP.setPrefWidth(110);

        TableColumn<ExcludeRuleTableRow, String> colM = new TableColumn<>(COL_MACHINE);
        colM.setCellValueFactory(d -> d.getValue().machineProperty());
        colM.setCellFactory(TextFieldTableCell.forTableColumn());
        colM.setOnEditCommit(e -> {
            e.getRowValue().machineProperty().set(e.getNewValue());
            tableEdited();
        });
        colM.setPrefWidth(140);

        TableColumn<ExcludeRuleTableRow, String> colF = new TableColumn<>(COL_FLAG);
        colF.setCellValueFactory(d -> d.getValue().flagProperty());
        colF.setCellFactory(TextFieldTableCell.forTableColumn());
        colF.setOnEditCommit(e -> {
            e.getRowValue().flagProperty().set(e.getNewValue());
            tableEdited();
        });
        colF.setPrefWidth(72);

        TableColumn<ExcludeRuleTableRow, String> colLj = new TableColumn<>(COL_LOGIC_JA);
        colLj.setCellValueFactory(d -> d.getValue().logicJaProperty());
        colLj.setCellFactory(TextFieldTableCell.forTableColumn());
        colLj.setOnEditCommit(e -> {
            e.getRowValue().logicJaProperty().set(e.getNewValue());
            tableEdited();
        });
        colLj.setPrefWidth(160);

        TableColumn<ExcludeRuleTableRow, String> colJe = new TableColumn<>(COL_LOGIC_JSON);
        colJe.setCellValueFactory(d -> d.getValue().logicJsonProperty());
        colJe.setCellFactory(TextFieldTableCell.forTableColumn());
        colJe.setOnEditCommit(e -> {
            e.getRowValue().logicJsonProperty().set(e.getNewValue());
            tableEdited();
        });
        colJe.setPrefWidth(220);

        rulesTable.getColumns().setAll(colP, colM, colF, colLj, colJe);

        rulesTable.focusedProperty().addListener((obs, was, now) -> {
            if (Boolean.TRUE.equals(now)) {
                preferTableOnSave = true;
            }
        });
        bodyArea.focusedProperty().addListener((obs, was, now) -> {
            if (Boolean.TRUE.equals(now)) {
                preferTableOnSave = false;
            }
        });
        bodyArea
                .textProperty()
                .addListener(
                        (obs, prev, cur) -> {
                            if (!suppressBodyDirty) {
                                preferTableOnSave = false;
                            }
                        });

        javafx.application.Platform.runLater(
                () -> {
                    if (editorSplit != null) {
                        editorSplit.setDividerPositions(0.48);
                    }
                });
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
    }

    @FXML
    private void onFromEnvButtonAction() {
        String p = shell.snapshotUiEnv().get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
        pathField.setText(p != null ? p : "");
    }

    @FXML
    private void onPickButtonAction() {
        FileChooser fc = new FileChooser();
        fc.setTitle("PM_AI_EXCLUDE_RULES_JSON");
        FileChooserForEnvKey.apply(fc, AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
        String cur = pathField.getText();
        if (cur != null && !cur.isBlank()) {
            try {
                Path p = Path.of(cur.trim());
                if (Files.isRegularFile(p) && p.getParent() != null) {
                    fc.setInitialDirectory(p.getParent().toFile());
                }
            } catch (Exception ignored) {
            }
        }
        var f = fc.showOpenDialog(ownerStage);
        if (f != null) {
            pathField.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void onLoadButtonAction() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        if (p.isEmpty()) {
            shell.appendLog("[exclude-json] path empty");
            return;
        }
        try {
            String s = Files.readString(Path.of(p), StandardCharsets.UTF_8);
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("path", p);
                d.put("charsRead", s.length());
                d.put("source", "onLoadButtonAction");
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H2",
                        "ExcludeRulesTabController.onLoadButtonAction",
                        "load read ok",
                        d);
            }
            // #endregion
            setBodyAndSyncTable(s, "[exclude-json] load ok: " + p);
        } catch (IOException ex) {
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("path", p);
                d.put("error", ex.getClass().getSimpleName());
                d.put("message", ex.getMessage());
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H4",
                        "ExcludeRulesTabController.onLoadButtonAction",
                        "load read failed",
                        d);
            }
            // #endregion
            shell.appendLog("[exclude-json] load error: " + ex.getMessage());
        }
    }

    @FXML
    private void onSaveButtonAction() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        if (p.isEmpty()) {
            shell.appendLog("[exclude-json] path empty (set PM_AI_EXCLUDE_RULES_JSON or type path)");
            return;
        }
        try {
            if (preferTableOnSave) {
                applyTableToBodyInternal();
            }
            Files.writeString(Path.of(p), bodyArea.getText(), StandardCharsets.UTF_8);
            shell.appendLog("[exclude-json] save ok: " + p);
        } catch (IOException ex) {
            shell.appendLog("[exclude-json] save error: " + ex.getMessage());
        }
    }

    /**
     * Loads JSON into the editor when a session path points at an existing file (next-launch restore).
     */
    void tryStartupLoadFromPathField() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        // #region agent log
        {
            Path fpProbe = p.isEmpty() ? null : Path.of(p);
            boolean exists = fpProbe != null && Files.exists(fpProbe);
            boolean regular = fpProbe != null && Files.isRegularFile(fpProbe);
            Map<String, Object> d = new LinkedHashMap<>();
            d.put("path", p);
            d.put("pathEmpty", p.isEmpty());
            d.put("exists", exists);
            d.put("isRegularFile", regular);
            AgentDebugLog.appendStructured(
                    dbgUiEnv(),
                    AGENT_DEBUG_SESSION_ID,
                    "H1",
                    "ExcludeRulesTabController.tryStartupLoadFromPathField",
                    "startup load entry",
                    d);
        }
        // #endregion
        if (p.isEmpty()) {
            return;
        }
        try {
            Path fp = Path.of(p);
            if (!Files.isRegularFile(fp)) {
                // #region agent log
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H1",
                        "ExcludeRulesTabController.tryStartupLoadFromPathField",
                        "skip load: not a regular file",
                        Map.of("path", p));
                // #endregion
                return;
            }
            String jsonText = Files.readString(fp, StandardCharsets.UTF_8);
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("path", p);
                d.put("charsRead", jsonText.length());
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H1",
                        "ExcludeRulesTabController.tryStartupLoadFromPathField",
                        "startup read ok before sync",
                        d);
            }
            // #endregion
            setBodyAndSyncTable(jsonText, "[exclude-json] restored session: " + p);
        } catch (IOException ex) {
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("path", p);
                d.put("error", ex.getClass().getSimpleName());
                d.put("message", ex.getMessage());
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H4",
                        "ExcludeRulesTabController.tryStartupLoadFromPathField",
                        "session restore io error",
                        d);
            }
            // #endregion
            shell.appendLog("[exclude-json] session restore load error: " + ex.getMessage());
        }
    }

    String snapshotExcludeRulesPath() {
        return pathField.getText() != null ? pathField.getText().trim() : "";
    }

    void restoreDesktopSessionPath(String path) {
        if (path != null && !path.isBlank()) {
            pathField.setText(path.trim());
        }
    }

    @FXML
    private void onValidateButtonAction() {
        String t = bodyArea.getText() != null ? bodyArea.getText().trim() : "";
        if (preferTableOnSave) {
            applyTableToBodyInternal();
            t = bodyArea.getText() != null ? bodyArea.getText().trim() : "";
        }
        if (t.isEmpty()) {
            shell.appendLog("[exclude-json] body empty");
            return;
        }
        try {
            JsonNode n = JSON.readTree(t);
            if (n.isObject() && n.has("rules") && n.get("rules").isArray()) {
                shell.appendLog(
                        "[exclude-json] ok: object with rules[" + n.get("rules").size() + "]");
            } else if (n.isArray()) {
                shell.appendLog("[exclude-json] ok: array len=" + n.size());
            } else {
                shell.appendLog(
                        "[exclude-json] parse ok but expected array or {\"rules\":[]}, got: "
                                + n.getNodeType());
            }
        } catch (Exception ex) {
            shell.appendLog("[exclude-json] invalid: " + ex.getMessage());
        }
    }

    @FXML
    private void onFormatButtonAction() {
        String t = bodyArea.getText() != null ? bodyArea.getText().trim() : "";
        if (t.isEmpty()) {
            shell.appendLog("[exclude-json] body empty");
            return;
        }
        try {
            JsonNode n = JSON.readTree(t);
            String pretty = JSON.writerWithDefaultPrettyPrinter().writeValueAsString(n);
            setBodyTextProgrammatically(pretty);
            preferTableOnSave = false;
            syncJsonToTableInternal();
            shell.appendLog("[exclude-json] formatted");
        } catch (Exception ex) {
            shell.appendLog("[exclude-json] format error: " + ex.getMessage());
        }
    }

    @FXML
    private void onAddRowButtonAction() {
        ruleRows.add(new ExcludeRuleTableRow());
        tableEdited();
    }

    @FXML
    private void onRemoveRowButtonAction() {
        var sel = rulesTable.getSelectionModel().getSelectedItems();
        if (sel == null || sel.isEmpty()) {
            shell.appendLog("[exclude-json] 削除する行を選択してください");
            return;
        }
        ruleRows.removeAll(new ArrayList<>(sel));
        tableEdited();
    }

    @FXML
    private void onJsonToTableButtonAction() {
        try {
            syncJsonToTableInternal();
            preferTableOnSave = false;
            shell.appendLog("[exclude-json] JSON→表 反映しました");
        } catch (Exception ex) {
            shell.appendLog("[exclude-json] JSON→表 エラー: " + ex.getMessage());
        }
    }

    @FXML
    private void onTableToJsonButtonAction() {
        try {
            applyTableToBodyInternal();
            preferTableOnSave = false;
            shell.appendLog("[exclude-json] 表→JSON 反映しました");
        } catch (Exception ex) {
            shell.appendLog("[exclude-json] 表→JSON エラー: " + ex.getMessage());
        }
    }

    private void setBodyAndSyncTable(String jsonText, String logOk) {
        setBodyTextProgrammatically(jsonText);
        preferTableOnSave = false;
        try {
            syncJsonToTableInternal();
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("tableRows", ruleRows.size());
                d.put("bodyChars", bodyArea.getText() != null ? bodyArea.getText().length() : 0);
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H3",
                        "ExcludeRulesTabController.setBodyAndSyncTable",
                        "sync json to table ok",
                        d);
            }
            // #endregion
        } catch (Exception ex) {
            // #region agent log
            {
                Map<String, Object> d = new LinkedHashMap<>();
                d.put("error", ex.getClass().getSimpleName());
                d.put("message", ex.getMessage());
                d.put("bodyCharsAfterFail", bodyArea.getText() != null ? bodyArea.getText().length() : 0);
                AgentDebugLog.appendStructured(
                        dbgUiEnv(),
                        AGENT_DEBUG_SESSION_ID,
                        "H3",
                        "ExcludeRulesTabController.setBodyAndSyncTable",
                        "sync json to table failed",
                        d);
            }
            // #endregion
            shell.appendLog("[exclude-json] load後の表反映スキップ: " + ex.getMessage());
        }
        shell.appendLog(logOk);
    }

    private Map<String, String> dbgUiEnv() {
        return shell != null ? shell.snapshotUiEnv() : Map.of();
    }

    private void setBodyTextProgrammatically(String text) {
        suppressBodyDirty = true;
        try {
            bodyArea.setText(text != null ? text : "");
        } finally {
            suppressBodyDirty = false;
        }
    }

    private void tableEdited() {
        preferTableOnSave = true;
    }

    private void syncJsonToTableInternal() throws IOException {
        String t = bodyArea.getText() != null ? bodyArea.getText().trim() : "";
        ruleRows.clear();
        if (t.isEmpty()) {
            return;
        }
        JsonNode root = JSON.readTree(t);
        JsonNode rules;
        if (root.isArray()) {
            rules = root;
        } else if (root.isObject() && root.has("rules")) {
            rules = root.get("rules");
        } else {
            throw new IOException("JSON は配列か {\"rules\":[]} 形式である必要があります");
        }
        if (!rules.isArray()) {
            throw new IOException("rules が配列ではありません");
        }
        for (JsonNode r : rules) {
            if (!r.isObject()) {
                continue;
            }
            ruleRows.add(rowFromJsonObject((ObjectNode) r));
        }
    }

    private static ExcludeRuleTableRow rowFromJsonObject(ObjectNode o) {
        ExcludeRuleTableRow row = new ExcludeRuleTableRow();
        row.processProperty().set(cellText(o, COL_PROCESS));
        row.machineProperty().set(cellText(o, COL_MACHINE));
        row.flagProperty().set(cellText(o, COL_FLAG));
        String lj = cellText(o, COL_LOGIC_JA);
        if (lj.isEmpty()) {
            lj = cellText(o, COL_LOGIC_JA_LEGACY);
        }
        row.logicJaProperty().set(lj);
        row.logicJsonProperty().set(cellText(o, COL_LOGIC_JSON));
        return row;
    }

    private static String cellText(ObjectNode o, String key) {
        if (!o.has(key) || o.get(key).isNull()) {
            return "";
        }
        JsonNode v = o.get(key);
        if (v.isTextual()) {
            return v.asText("");
        }
        if (v.isNumber() || v.isBoolean()) {
            return v.asText();
        }
        return v.toString();
    }

    private void applyTableToBodyInternal() {
        ArrayNode arr = JSON.createArrayNode();
        for (ExcludeRuleTableRow row : ruleRows) {
            String proc = nz(row.processProperty().get());
            if (proc.isEmpty()) {
                continue;
            }
            ObjectNode rec = JSON.createObjectNode();
            putCell(rec, COL_PROCESS, proc);
            putCell(rec, COL_MACHINE, nz(row.machineProperty().get()));
            putCell(rec, COL_FLAG, nz(row.flagProperty().get()));
            putCell(rec, COL_LOGIC_JA, nz(row.logicJaProperty().get()));
            putCell(rec, COL_LOGIC_JSON, nz(row.logicJsonProperty().get()));
            arr.add(rec);
        }
        ObjectNode wrapper = JSON.createObjectNode();
        wrapper.set("rules", arr);
        try {
            String out = JSON.writerWithDefaultPrettyPrinter().writeValueAsString(wrapper);
            setBodyTextProgrammatically(out);
        } catch (IOException ex) {
            throw new IllegalStateException(ex);
        }
    }

    private static void putCell(ObjectNode rec, String key, String value) {
        if (value.isEmpty()) {
            rec.putNull(key);
        } else {
            rec.put(key, value);
        }
    }

    private static String nz(String s) {
        return s != null ? s.trim() : "";
    }

    /** 1 行分（配台不要ルール）。 */
    public static final class ExcludeRuleTableRow {
        private final StringProperty process = new SimpleStringProperty("");
        private final StringProperty machine = new SimpleStringProperty("");
        private final StringProperty flag = new SimpleStringProperty("");
        private final StringProperty logicJa = new SimpleStringProperty("");
        private final StringProperty logicJson = new SimpleStringProperty("");

        public StringProperty processProperty() {
            return process;
        }

        public StringProperty machineProperty() {
            return machine;
        }

        public StringProperty flagProperty() {
            return flag;
        }

        public StringProperty logicJaProperty() {
            return logicJa;
        }

        public StringProperty logicJsonProperty() {
            return logicJson;
        }
    }
}
