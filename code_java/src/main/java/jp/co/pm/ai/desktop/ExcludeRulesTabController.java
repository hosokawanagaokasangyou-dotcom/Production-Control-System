package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;

/**
 * PM_AI_EXCLUDE_RULES_JSON editor; static structure in {@code ExcludeRulesTab.fxml}.
 */
public final class ExcludeRulesTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String HINT_TEXT =
            "PM_AI_EXCLUDE_RULES_JSON が実在かつ有効なら"
                    + " 設定_配台不要工程 の Excel 保守を省略可。"
                    + " 環境変数タブに同名を追加してパスを共有してください。";

    private Stage ownerStage;

    private MainShellController shell;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField pathField;

    @FXML
    private TextArea bodyArea;

    @FXML
    private Button fromEnvButton;

    @FXML
    private Button pickButton;

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button validateButton;

    @FXML
    private void initialize() {
        pathField.setPromptText("PM_AI_EXCLUDE_RULES_JSON — .json フルパス");
        bodyArea.setPromptText(
                "[\n  { \"工程名\": \"...\", \"機械名\": \"...\", ... }\n]\n"
                        + "または {\"rules\":[...]}");
        hintLabel.setText(HINT_TEXT);
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
            bodyArea.setText(s);
            shell.appendLog("[exclude-json] load ok: " + p);
        } catch (IOException ex) {
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
        if (p.isEmpty()) {
            return;
        }
        try {
            Path fp = Path.of(p);
            if (!Files.isRegularFile(fp)) {
                return;
            }
            bodyArea.setText(Files.readString(fp, StandardCharsets.UTF_8));
            shell.appendLog("[exclude-json] restored session: " + p);
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
}
