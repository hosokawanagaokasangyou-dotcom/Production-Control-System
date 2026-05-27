package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.rules.SpecialRulesBuilderTabController;
import jp.co.pm.ai.desktop.dispatch.rules.SpecialRulesTestLabTabController;
import jp.co.pm.ai.desktop.dispatch.rules.migration.DispatchRuleMigrationService;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;
import jp.co.pm.ai.desktop.dispatch.rules.trace.SpecialRulesTraceTabController;

/** Special rules tab: markdown + builder + test lab + trace + JSON. */
public final class SpecialRulesTabController {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private MainShellController shell;

    @FXML private TabPane innerTabPane;
    @FXML private Label markdownPathLabel;
    @FXML private TextArea summaryBodyArea;
    @FXML private TextArea enumeratedBodyArea;
    @FXML private TextArea jsonBodyArea;

    @FXML private SpecialRulesBuilderTabController builderTabController;
    @FXML private SpecialRulesTestLabTabController testLabTabController;
    @FXML private SpecialRulesTraceTabController traceTabController;

    @FXML
    private void initialize() {
        if (innerTabPane != null) {
            innerTabPane
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (b != null && "ルール試走".equals(b.getText()) && testLabTabController != null) {
                                    testLabTabController.refreshPickers();
                                }
                            });
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        loadMarkdown(true);
        loadMarkdown(false);
        reloadJsonEditor();
        if (builderTabController != null) {
            builderTabController.bindShell(shell);
        }
        if (testLabTabController != null && builderTabController != null) {
            testLabTabController.bindShell(shell, builderTabController);
            testLabTabController.refreshPickers();
        }
        if (traceTabController != null) {
            traceTabController.bindShell(shell);
        }
    }

    void reloadTraceFromDisk() {
        if (traceTabController != null) {
            traceTabController.reloadFromDisk();
        }
    }

    @FXML
    private void onReloadMarkdownAction() {
        Tab sel = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedItem() : null;
        boolean summary = sel == null || "要約".equals(sel.getText());
        loadMarkdown(summary);
    }

    @FXML
    private void onReloadJsonAction() {
        reloadJsonEditor();
    }

    @FXML
    private void onSaveJsonAction() {
        if (shell == null || jsonBodyArea == null) {
            return;
        }
        try {
            Path work = DispatchRulePaths.resolveWorkJson(shell.snapshotUiEnv());
            Files.createDirectories(work.getParent());
            Files.writeString(work, jsonBodyArea.getText(), StandardCharsets.UTF_8);
            shell.dispatchRulesAppendLog("[dispatch-rules] JSON tab saved: " + work);
            if (builderTabController != null) {
                builderTabController.bindShell(shell);
            }
        } catch (IOException ex) {
            shell.showErrorDialog("保存エラー", ex.getMessage());
        }
    }

    private void loadMarkdown(boolean summary) {
        if (shell == null) {
            return;
        }
        Path path =
                summary
                        ? AppPaths.resolveSpecialRulesSummaryMd(shell.snapshotUiEnv())
                        : AppPaths.resolveSpecialRulesEnumeratedMd(shell.snapshotUiEnv());
        if (markdownPathLabel != null) {
            markdownPathLabel.setText(path.toString());
        }
        TextArea area = summary ? summaryBodyArea : enumeratedBodyArea;
        if (area == null) {
            return;
        }
        try {
            if (Files.isRegularFile(path)) {
                area.setText(Files.readString(path, StandardCharsets.UTF_8));
            } else {
                area.setText("ファイルが見つかりません: " + path);
            }
        } catch (IOException ex) {
            area.setText("読込エラー: " + ex.getMessage());
        }
    }

    private void reloadJsonEditor() {
        if (shell == null || jsonBodyArea == null) {
            return;
        }
        DispatchRulePaths.ensureWorkJsonFromRepoIfMissing(shell.dispatchRulesUiEnv());
        Path work = DispatchRulePaths.resolveWorkJson(shell.dispatchRulesUiEnv());
        try {
            if (Files.isRegularFile(work)) {
                var raw = JSON.readTree(Files.readString(work, StandardCharsets.UTF_8));
                var migrated = DispatchRuleMigrationService.migrate((com.fasterxml.jackson.databind.node.ObjectNode) raw);
                jsonBodyArea.setText(JSON.writerWithDefaultPrettyPrinter().writeValueAsString(migrated));
            } else {
                jsonBodyArea.setText(JSON.writerWithDefaultPrettyPrinter().writeValueAsString(new DispatchRuleDocument()));
            }
        } catch (IOException ex) {
            jsonBodyArea.setText("読込エラー: " + ex.getMessage());
        }
    }
}
