package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;

/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    @FXML
    private TextField workbookField;

    @FXML
    private TextField pythonExeField;

    @FXML
    private TextField scriptDirField;

    @FXML
    private TextArea logArea;

    @FXML
    private Label statusLabel;

    @FXML
    private Button browseWbButton;

    @FXML
    private Button detectWbButton;

    @FXML
    private Button refreshDirButton;

    @FXML
    private HBox stageActionsBox;

    void bindShell(MainShellController shell) {
        browseWbButton.setOnAction(e -> shell.pickWorkbook());
        detectWbButton.setOnAction(e -> workbookField.setText(shell.resolveTaskInputWorkbookFromEnv()));
        refreshDirButton.setOnAction(e -> scriptDirField.setText(shell.resolvePythonScriptDirFromEnv()));
        shell.attachStageButtons(stageActionsBox);
    }

    TextField getWorkbookField() {
        return workbookField;
    }

    TextField getPythonExeField() {
        return pythonExeField;
    }

    TextField getScriptDirField() {
        return scriptDirField;
    }

    TextArea getLogArea() {
        return logArea;
    }

    Label getStatusLabel() {
        return statusLabel;
    }

    void appendLog(String line) {
        logArea.appendText(line + "\n");
    }
}
