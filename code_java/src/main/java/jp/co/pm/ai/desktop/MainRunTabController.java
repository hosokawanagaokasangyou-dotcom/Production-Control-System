package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    private MainShellController shell;

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
    private Button stage1RunButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private Button peekSheetsButton;

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    @FXML
    private void onBrowseWorkbookButtonAction() {
        shell.pickWorkbook();
    }

    @FXML
    private void onDetectWorkbookButtonAction() {
        workbookField.setText(shell.resolveTaskInputWorkbookFromEnv());
    }

    @FXML
    private void onRefreshScriptDirButtonAction() {
        scriptDirField.setText(shell.resolvePythonScriptDirFromEnv());
    }

    @FXML
    private void onStage1RunButtonAction() {
        shell.triggerStage1();
    }

    @FXML
    private void onStage2RunButtonAction() {
        shell.triggerStage2();
    }

    @FXML
    private void onPeekSheetsButtonAction() {
        shell.triggerPeekSheets();
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
