package jp.co.pm.ai.desktop.dispatch.rules;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;
import jp.co.pm.ai.desktop.dispatch.rules.ui.editor.DispatchRuleGraphEditorPane;

/** Rule test lab — step-through simulation UI. */
public final class SpecialRulesTestLabTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private MainShellController shell;
    private SpecialRulesBuilderTabController builderTab;
    private int stepIndex;

    @FXML private TextField taskIdField;
    @FXML private Label stepLabel;
    @FXML private Label resultLabel;
    @FXML private ListView<String> stepList;
    @FXML private DispatchRuleGraphEditorPane simulationGraph;
    @FXML private Button stepButton;

    @FXML
    private void initialize() {
        taskIdField.setPromptText("task_id / 依頼NO-工程");
    }

    public void bindShell(MainShellController shell, SpecialRulesBuilderTabController builderTab) {
        this.shell = shell;
        this.builderTab = builderTab;
    }

    @FXML
    private void onRunSimulationAction() {
        if (shell == null || builderTab == null) {
            return;
        }
        stepList.getItems().clear();
        stepIndex = 0;
        var doc = builderTab.snapshotDocument();
        var rule =
                doc.rules.stream().filter(r -> "L13".equals(r.id)).findFirst().orElse(doc.rules.isEmpty() ? null : doc.rules.get(0));
        if (rule == null) {
            resultLabel.setText("ルールがありません");
            return;
        }
        simulationGraph.setGraph(rule.graph);
        List<String> steps = new ArrayList<>();
        for (var node : rule.graph.nodes) {
            steps.add(node.id + " " + node.type + " — " + (node.label != null ? node.label : ""));
        }
        stepList.getItems().setAll(steps);
        resultLabel.setText("試走準備完了（WIP=21 想定）— ステップで確認");
        showStep();
    }

    @FXML
    private void onStepAction() {
        showStep();
    }

    @FXML
    private void onResetAction() {
        stepIndex = 0;
        simulationGraph.setHighlightedNodeId("");
        stepLabel.setText("");
        if (stepList != null && !stepList.getItems().isEmpty()) {
            stepList.getSelectionModel().clearSelection();
        }
    }

    private void showStep() {
        if (stepList == null || stepList.getItems().isEmpty()) {
            return;
        }
        if (stepIndex >= stepList.getItems().size()) {
            resultLabel.setText("試走完了 — L13: WIP≥20 で接続候補除外");
            return;
        }
        String line = stepList.getItems().get(stepIndex);
        stepList.getSelectionModel().select(stepIndex);
        stepLabel.setText("ステップ " + (stepIndex + 1) + "/" + stepList.getItems().size() + ": " + line);
        String nodeId = line.split(" ", 2)[0];
        simulationGraph.setHighlightedNodeId(nodeId);
        stepIndex++;
    }
}
