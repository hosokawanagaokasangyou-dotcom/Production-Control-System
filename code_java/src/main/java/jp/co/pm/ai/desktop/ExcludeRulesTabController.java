package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.BiConsumer;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.Region;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.crypto.GeminiCredentialsV2Crypto;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;

/**
 * PM_AI_EXCLUDE_RULES_JSON editor; static structure in {@code ExcludeRulesTab.fxml}.
 *
 * <p>ルールは JSON（配列または {@code {"rules":[]}}）と表の両方から編集でき、保存時は直近で編集した側をファイルへ書き出します。
 */
public final class ExcludeRulesTabController {

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
                    + " 表と JSON の両方で編集できます。保存は直近で触れた側を書き出します。"
                    + " 「ロジック式をAI生成」は選択行の配台不要ロジックを Gemini で JSON 化します（"
                    + "code/gemini_credentials.encrypted.json または GEMINI_CREDENTIALS_JSON）。";

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

    @FXML
    private Button aiCompileLogicJsonButton;

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
        colLj.setCellFactory(c -> wrappingTextAreaCell(c, (row, txt) -> {
            row.logicJaProperty().set(txt);
            tableEdited();
        }));
        colLj.setPrefWidth(280);

        TableColumn<ExcludeRuleTableRow, String> colJe = new TableColumn<>(COL_LOGIC_JSON);
        colJe.setCellValueFactory(d -> d.getValue().logicJsonProperty());
        colJe.setCellFactory(c -> wrappingTextAreaCell(c, (row, txt) -> {
            row.logicJsonProperty().set(txt);
            tableEdited();
        }));
        colJe.setPrefWidth(320);

        rulesTable.getColumns().setAll(colP, colM, colF, colLj, colJe);
        rulesTable.setFixedCellSize(96);

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
            shell.showWarningDialog("読込", "JSON のパスが空です。");
            return;
        }
        Optional<Path> resolvedOpt = resolveExistingExcludeRulesJsonPath(p);
        if (resolvedOpt.isEmpty()) {
            shell.appendLog(
                    "[exclude-json] 読込失敗: ファイルなし（入力パスおよび code/json 等の既定候補）。パス="
                            + p);
            shell.showWarningDialog(
                    "読込",
                    "ファイルが見つかりません（入力パスおよび既定候補）。\nパス: " + p);
            return;
        }
        Path fp = resolvedOpt.get();
        Path requestedAbs = Path.of(p).toAbsolutePath().normalize();
        if (!fp.equals(requestedAbs)) {
            pathField.setText(fp.toString());
            shell.appendLog("[exclude-json] 実在パスへ切替えて読込: " + fp);
        }
        try {
            String s = Files.readString(fp, StandardCharsets.UTF_8);
            setBodyAndSyncTable(s, "[exclude-json] load ok: " + fp);
            shell.showInformationDialog("読込完了", "配台不要ルール JSON を読み込みました。\n" + fp);
        } catch (IOException ex) {
            shell.appendLog("[exclude-json] load error: " + ex.getMessage());
            shell.showErrorDialog(
                    "読込エラー",
                    ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onSaveButtonAction() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        if (p.isEmpty()) {
            shell.appendLog("[exclude-json] path empty (set PM_AI_EXCLUDE_RULES_JSON or type path)");
            shell.showWarningDialog("保存", "保存先のパスが空です（PM_AI_EXCLUDE_RULES_JSON または手入力）。");
            return;
        }
        try {
            if (preferTableOnSave) {
                applyTableToBodyInternal();
            }
            Files.writeString(Path.of(p), bodyArea.getText(), StandardCharsets.UTF_8);
            shell.appendLog("[exclude-json] save ok: " + p);
            shell.showInformationDialog("保存完了", "配台不要ルール JSON を保存しました。\n" + p);
        } catch (IOException ex) {
            shell.appendLog("[exclude-json] save error: " + ex.getMessage());
            shell.showErrorDialog(
                    "保存エラー",
                    ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    /**
     * Loads JSON into the editor when a session path points at an existing file (next-launch restore).
     */
    void tryStartupLoadFromPathField() {
        String p = pathField.getText() != null ? pathField.getText().trim() : "";
        if (p.isEmpty()) {
            return;
        }
        Optional<Path> resolvedOpt = resolveExistingExcludeRulesJsonPath(p);
        if (resolvedOpt.isEmpty()) {
            shell.appendLog(
                    "[exclude-json] セッション復元スキップ: ファイルなし（入力パスおよび code/json 等の既定候補）。パス="
                            + p);
            return;
        }
        Path fp = resolvedOpt.get();
        Path requestedAbs = Path.of(p).toAbsolutePath().normalize();
        if (!fp.equals(requestedAbs)) {
            pathField.setText(fp.toString());
            shell.appendLog("[exclude-json] 実在パスへ切替え（復元）: " + fp);
        }
        try {
            String jsonText = Files.readString(fp, StandardCharsets.UTF_8);
            setBodyAndSyncTable(jsonText, "[exclude-json] restored session: " + fp);
        } catch (IOException ex) {
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
    private void onAiCompileLogicJsonButtonAction() {
        if (shell == null) {
            return;
        }
        var selected = rulesTable.getSelectionModel().getSelectedItems();
        if (selected == null || selected.isEmpty()) {
            shell.appendLog("[exclude-json] AI: 行を選択してください");
            shell.showWarningDialog("AI ロジック式", "変換する行を表で選択してください。");
            return;
        }
        List<ExcludeRuleTableRow> targets = new ArrayList<>();
        for (ExcludeRuleTableRow r : selected) {
            String ja = r.logicJaProperty().get();
            if (ja != null && !ja.strip().isEmpty()) {
                targets.add(r);
            }
        }
        if (targets.isEmpty()) {
            shell.showWarningDialog("AI ロジック式", "選択行に配台不要ロジックの本文がありません。");
            return;
        }
        String apiKey = loadGeminiApiKeyOrNull();
        if (apiKey == null) {
            return;
        }
        if (aiCompileLogicJsonButton != null) {
            aiCompileLogicJsonButton.setDisable(true);
        }
        shell.appendLog("[exclude-json] AI: " + targets.size() + " 行を変換中…");
        Duration timeout = Duration.ofSeconds(120);
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                for (ExcludeRuleTableRow row : targets) {
                                    final ExcludeRuleTableRow r = row;
                                    String json =
                                            ExcludeRuleLogicGeminiService.compileToCompactJson(
                                                    r.logicJaProperty().get(),
                                                    r.processProperty().get(),
                                                    r.machineProperty().get(),
                                                    apiKey,
                                                    timeout);
                                    final String fj = json;
                                    Platform.runLater(
                                            () -> {
                                                r.logicJsonProperty().set(fj);
                                                tableEdited();
                                            });
                                }
                                Platform.runLater(
                                        () ->
                                                shell.appendLog(
                                                        "[exclude-json] AI: "
                                                                + targets.size()
                                                                + " 行のロジック式を更新しました"));
                            } catch (Exception ex) {
                                String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
                                Platform.runLater(
                                        () -> {
                                            shell.appendLog("[exclude-json] AI エラー: " + msg);
                                            shell.showErrorDialog("AI ロジック式", msg);
                                        });
                            } finally {
                                Platform.runLater(
                                        () -> {
                                            if (aiCompileLogicJsonButton != null) {
                                                aiCompileLogicJsonButton.setDisable(false);
                                            }
                                        });
                            }
                        },
                        "exclude-rules-ai-logic");
        worker.setDaemon(true);
        worker.start();
    }

    private void setBodyAndSyncTable(String jsonText, String logOk) {
        setBodyTextProgrammatically(jsonText);
        preferTableOnSave = false;
        try {
            syncJsonToTableInternal();
        } catch (Exception ex) {
            shell.appendLog("[exclude-json] load後の表反映スキップ: " + ex.getMessage());
        }
        shell.appendLog(logOk);
    }

    /**
     * 入力パスが誤ったサブフォルダ（例: Production-Control-System）を挟んでいても、リポジトリの {@code
     * code/json/stage1_exclude_rules.json} 等へ読込をフォールバックする。
     */
    private Optional<Path> resolveExistingExcludeRulesJsonPath(String requested) {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        List<Path> candidates = new ArrayList<>();
        if (requested != null && !requested.isBlank()) {
            candidates.add(Path.of(requested.trim()));
        }
        String fileName = AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME;
        if (requested != null && !requested.isBlank()) {
            Path fn = Path.of(requested.trim()).getFileName();
            if (fn != null && !fn.toString().isBlank()) {
                fileName = fn.toString();
            }
        }
        Path repo = AppPaths.resolveRepoRoot(ui);
        candidates.add(repo.resolve("code").resolve("json").resolve(fileName));
        candidates.add(AppPaths.stage1ExcludeRulesJsonPath(ui));
        candidates.add(AppPaths.stage1ExcludeRulesJsonPathLegacyUnderPython(ui));
        Path parent = repo.getParent();
        if (parent != null) {
            candidates.add(parent.resolve("code").resolve("json").resolve(fileName));
        }
        Set<Path> seen = new LinkedHashSet<>();
        for (Path c : candidates) {
            Path n = c.toAbsolutePath().normalize();
            if (!seen.add(n)) {
                continue;
            }
            if (Files.isRegularFile(n)) {
                return Optional.of(n);
            }
        }
        return Optional.empty();
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

    /**
     * 長文列用: 右端で折り返す {@link TextArea} をセルに載せ、フォーカス喪失時にモデルへ反映する。
     */
    private TableCell<ExcludeRuleTableRow, String> wrappingTextAreaCell(
            TableColumn<ExcludeRuleTableRow, String> column,
            BiConsumer<ExcludeRuleTableRow, String> applyText) {
        return new TableCell<>() {
            private final TextArea area = new TextArea();
            private ExcludeRuleTableRow boundRow;

            {
                area.setWrapText(true);
                area.setPrefRowCount(4);
                area.setMinHeight(Region.USE_PREF_SIZE);
                area.setMaxWidth(Double.MAX_VALUE);
                area
                        .focusedProperty()
                        .addListener(
                                (obs, was, now) -> {
                                    if (Boolean.FALSE.equals(now)
                                            && boundRow != null
                                            && getTableRow() != null
                                            && boundRow == getTableRow().getItem()) {
                                        applyText.accept(boundRow, area.getText());
                                    }
                                });
            }

            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || getTableRow() == null || getTableRow().getItem() == null) {
                    boundRow = null;
                    setGraphic(null);
                    return;
                }
                boundRow = getTableRow().getItem();
                String next = item == null ? "" : item;
                if (!area.isFocused() && !Objects.equals(next, area.getText())) {
                    area.setText(next);
                }
                if (area.prefWidthProperty().isBound()) {
                    area.prefWidthProperty().unbind();
                }
                area.prefWidthProperty().bind(column.widthProperty().subtract(14));
                setGraphic(area);
            }
        };
    }

    /**
     * {@link jp.co.pm.ai.desktop.ApiModelBenchmarkTabController} と同じ解決（環境変数または既定パス）。
     */
    private static Path resolveGeminiCredentialsPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(AppPaths.KEY_GEMINI_CREDENTIALS_JSON);
        if (raw != null && !raw.isBlank()) {
            return Path.of(raw.strip()).toAbsolutePath().normalize();
        }
        return AppPaths.resolveRepoRoot(u)
                .resolve("code")
                .resolve("gemini_credentials.encrypted.json")
                .toAbsolutePath()
                .normalize();
    }

    /** 復号済み API キー。失敗時はダイアログのみで {@code null}。 */
    private String loadGeminiApiKeyOrNull() {
        Path credPath = resolveGeminiCredentialsPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(credPath)) {
            shell.showWarningDialog(
                    "AI ロジック式",
                    "認証ファイルが見つかりません。\n"
                            + credPath
                            + "\n環境変数タブで GEMINI_CREDENTIALS_JSON を設定するか、既定パスにファイルを置いてください。");
            return null;
        }
        try {
            String json = Files.readString(credPath, StandardCharsets.UTF_8);
            String apiKey =
                    GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                            json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
            if (apiKey == null || apiKey.isBlank()) {
                shell.showErrorDialog("AI ロジック式", "復号した API キーが空です。");
                return null;
            }
            return apiKey;
        } catch (Exception ex) {
            shell.showErrorDialog(
                    "AI ロジック式",
                    "認証 JSON の読込・復号に失敗しました。\n" + ex.getMessage());
            return null;
        }
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
