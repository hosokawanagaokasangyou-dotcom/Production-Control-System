package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TextArea;
import javafx.scene.control.Toggle;
import javafx.scene.control.ToggleGroup;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * リポジトリ直下の {@link AppPaths#SPECIAL_RULES_SUMMARY_MD} と {@link AppPaths#SPECIAL_RULES_ENUMERATED_MD} を閲覧する。
 */
public final class SpecialRulesTabController {

    private MainShellController shell;

    @FXML
    private ToggleGroup sourceToggle;

    @FXML
    private RadioButton summaryRadio;

    @FXML
    private RadioButton enumeratedRadio;

    @FXML
    private Label pathLabel;

    @FXML
    private TextArea bodyArea;

    @FXML
    private void initialize() {
        summaryRadio.setUserData(Boolean.TRUE);
        enumeratedRadio.setUserData(Boolean.FALSE);
        sourceToggle
                .selectedToggleProperty()
                .addListener(
                        (obs, prev, cur) -> {
                            if (shell != null && cur != null) {
                                loadCurrent(false);
                            }
                        });
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        loadCurrent(false);
    }

    @FXML
    private void onReloadAction() {
        loadCurrent(true);
    }

    private void loadCurrent(boolean userCompletionDialog) {
        if (shell == null) {
            return;
        }
        Toggle sel = sourceToggle != null ? sourceToggle.getSelectedToggle() : null;
        boolean summary =
                sel == null || sel == summaryRadio || Boolean.TRUE.equals(sel.getUserData());
        Path path =
                summary
                        ? AppPaths.resolveSpecialRulesSummaryMd(shell.snapshotUiEnv())
                        : AppPaths.resolveSpecialRulesEnumeratedMd(shell.snapshotUiEnv());
        pathLabel.setText(path.toString());
        try {
            if (Files.isRegularFile(path)) {
                bodyArea.setText(Files.readString(path, StandardCharsets.UTF_8));
                shell.appendLog("[special-rules] load ok: " + path);
                if (userCompletionDialog) {
                    shell.showInformationDialog("再読込完了", "特別ルールを読み込みました。\n" + path);
                }
            } else {
                bodyArea.setText(
                        "ファイルが見つかりません。\n"
                                + path
                                + "\n\nリポジトリ根（PM_AI_REPO_ROOT）を環境変数タブで確認してください。");
                shell.appendLog("[special-rules] missing: " + path);
                if (userCompletionDialog) {
                    shell.showWarningDialog("再読込", "ファイルが見つかりません。\n" + path);
                }
            }
        } catch (IOException ex) {
            bodyArea.setText("読込エラー: " + ex.getMessage());
            shell.appendLog("[special-rules] read error: " + ex.getMessage());
            if (userCompletionDialog) {
                shell.showErrorDialog(
                        "読込エラー",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }
}
