package jp.co.pm.ai.desktop.dispatch.rules;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.Slider;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.PlanInputTabController;
import jp.co.pm.ai.desktop.dispatch.rules.planinput.DispatchRulePlanInputTaskSource;
import jp.co.pm.ai.desktop.dispatch.rules.simulation.DispatchRuleSimulationResult;
import jp.co.pm.ai.desktop.dispatch.rules.simulation.DispatchRuleSimulationService;
import jp.co.pm.ai.desktop.dispatch.rules.simulation.DispatchRuleSimulationStep;
import jp.co.pm.ai.desktop.dispatch.rules.ui.DispatchRuleRollStageVisualPane;
import jp.co.pm.ai.desktop.dispatch.rules.ui.editor.DispatchRuleGraphEditorPane;

/** Rule test lab — Python simulation with step playback. */
public final class SpecialRulesTestLabTabController {

    private static final double WIP_THRESHOLD = 20.0;

    private MainShellController shell;
    private SpecialRulesBuilderTabController builderTab;
    private int stepIndex;
    private List<DispatchRuleSimulationStep> simulationSteps = List.of();
    private javafx.animation.Timeline playTimeline;
    private final DispatchRulePlanInputTaskSource taskSource = new DispatchRulePlanInputTaskSource();

    @FXML private ComboBox<String> taskPicker;
    @FXML private ComboBox<String> rulePicker;
    @FXML private CheckBox allRollsCheckBox;
    @FXML private Spinner<Integer> initialWipSpinner;
    @FXML private Label stepLabel;
    @FXML private Label resultLabel;
    @FXML private Label speedLabel;
    @FXML private ListView<String> stepList;
    @FXML private DispatchRuleGraphEditorPane simulationGraph;
    @FXML private DispatchRuleRollStageVisualPane rollStageVisual;
    @FXML private Button stepButton;
    @FXML private Slider speedSlider;

    @FXML
    private void initialize() {
        if (speedSlider != null) {
            speedSlider.valueProperty()
                    .addListener((o, a, b) -> {
                        if (speedLabel != null && b != null) {
                            speedLabel.setText(String.format("%.1fx", b.doubleValue()));
                        }
                    });
        }
        if (initialWipSpinner != null) {
            initialWipSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 40, 5));
        }
        if (rollStageVisual != null) {
            rollStageVisual.clear();
        }
    }

    public void bindShell(MainShellController shell, SpecialRulesBuilderTabController builderTab) {
        this.shell = shell;
        this.builderTab = builderTab;
        refreshPickers();
    }

    public void refreshPickers() {
        if (shell == null || builderTab == null) {
            return;
        }
        if (rulePicker != null) {
            List<String> ids =
                    builderTab.snapshotDocument().rules.stream().map(r -> r.id + " " + r.name).toList();
            rulePicker.getItems().setAll(ids);
            if (!ids.isEmpty()) {
                rulePicker.getSelectionModel().select(0);
            }
        }
        if (taskPicker != null) {
            PlanInputTabController planInput = shell.planInputTabControllerForDispatchRollUnit();
            taskSource.reload(shell.dispatchRulesUiEnv(), planInput);
            taskPicker.getItems().setAll(taskSource.labels());
            if (!taskPicker.getItems().isEmpty()) {
                taskPicker.getSelectionModel().select(0);
            }
            if (resultLabel != null) {
                resultLabel.setText(
                        taskPicker.getItems().isEmpty()
                                ? taskSource.sourceDescription()
                                : "タスク: "
                                        + taskPicker.getItems().size()
                                        + " 件（"
                                        + taskSource.sourceDescription()
                                        + "）");
            }
        }
    }

    @FXML
    private void onRefreshTasksAction() {
        refreshPickers();
    }

    @FXML
    private void onRunSimulationAction() {
        if (shell == null || builderTab == null) {
            return;
        }
        stopPlay();
        stepList.getItems().clear();
        stepIndex = 0;
        simulationSteps = List.of();
        clearPlaybackVisuals();
        String ruleLabel = rulePicker != null ? rulePicker.getValue() : null;
        String ruleId = ruleLabel != null && ruleLabel.contains(" ") ? ruleLabel.split(" ", 2)[0] : ruleLabel;
        Map<String, String> taskRow = resolveSelectedTaskRow();
        if (taskRow.isEmpty()) {
            resultLabel.setText(
                    "タスクを選択してください。"
                            + (taskSource.sourceDescription().isBlank()
                                    ? ""
                                    : " （" + taskSource.sourceDescription() + "）"));
            return;
        }
        var doc = builderTab.snapshotDocument();
        var rule =
                doc.rules.stream()
                        .filter(r -> ruleId == null || ruleId.equals(r.id))
                        .findFirst()
                        .orElse(doc.rules.isEmpty() ? null : doc.rules.get(0));
        if (rule == null) {
            resultLabel.setText("ルールがありません");
            return;
        }
        simulationGraph.setGraph(rule.graph);
        int rollCount = DispatchRulePlanInputTaskSource.parseRollCount(taskRow);
        boolean allRolls = allRollsCheckBox == null || allRollsCheckBox.isSelected();
        int initialWip = initialWipSpinner != null ? initialWipSpinner.getValue() : 0;
        Map<String, String> secTaskRow = Map.of();
        boolean connectionSecPipeline = false;
        if (allRolls && DispatchRulePlanInputTaskSource.isConnectionProcess(taskRow)) {
            String requestNo = taskRow.getOrDefault("依頼NO", "");
            var secOpt = taskSource.findSecRowForRequest(requestNo);
            if (secOpt.isPresent()) {
                secTaskRow = secOpt.get();
                connectionSecPipeline = true;
            }
        }
        final boolean pipelineMode = connectionSecPipeline;
        resultLabel.setText(
                allRolls
                        ? (pipelineMode
                                ? "接続→SEC 全ロール試走実行中…（"
                                        + rollCount
                                        + " ロール・WIP初期 "
                                        + initialWip
                                        + "）"
                                : "全ロール試走実行中…（"
                                        + rollCount
                                        + " ロール・WIP初期 "
                                        + initialWip
                                        + "）")
                        : "試走実行中…");
        final String selectedRuleId = rule.id;
        final Map<String, String> taskCopy = new LinkedHashMap<>(taskRow);
        final Map<String, String> secTaskCopy = new LinkedHashMap<>(secTaskRow);
        final Map<String, Object> overrides =
                allRolls
                        ? Map.of("metrics", Map.of("initial_wip", initialWip))
                        : Map.of(
                                "metrics",
                                Map.of("wip_connection_sec", initialWip + 20, "request_roll_diff", 10));
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                DispatchRuleSimulationResult result =
                                        DispatchRuleSimulationService.simulate(
                                                shell.resolveStagePythonExecutablePath(),
                                                shell.dispatchRulesUiEnv(),
                                                doc,
                                                taskCopy,
                                                secTaskCopy,
                                                selectedRuleId,
                                                overrides,
                                                allRolls,
                                                shell::dispatchRulesAppendLog);
                                Platform.runLater(
                                        () ->
                                                applySimulationResult(
                                                        result, allRolls, rollCount, pipelineMode));
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                resultLabel.setText(
                                                        "試走失敗: "
                                                                + (ex.getMessage() != null
                                                                        ? ex.getMessage()
                                                                        : ex.toString())));
                            }
                        },
                        "dispatch-rule-simulation");
        worker.setDaemon(true);
        worker.start();
    }

    private void applySimulationResult(
            DispatchRuleSimulationResult result, boolean allRolls, int rollCount, boolean pipelineMode) {
        simulationSteps = result.steps();
        List<String> lines = new ArrayList<>();
        for (DispatchRuleSimulationStep s : simulationSteps) {
            String prefix =
                    s.rollIndex() > 0 ? "[R" + s.rollIndex() + "] " : "";
            String phase =
                    s.flowPhase() != null && !s.flowPhase().isBlank()
                            ? s.flowPhase().toUpperCase() + " "
                            : "";
            lines.add(prefix + phase + s.sequence() + " " + stepKey(s) + " " + s.summaryJa());
        }
        stepList.getItems().setAll(lines);
        stepIndex = 0;
        clearPlaybackVisuals();
        if (stepLabel != null) {
            stepLabel.setText(
                    simulationSteps.isEmpty()
                            ? ""
                            : allRolls
                                    ? (pipelineMode
                                            ? "接続→SEC 全ロール試走準備完了（"
                                                    + rollCount
                                                    + " ロール）— 再生または1ステップ"
                                            : "全ロール試走準備完了（"
                                                    + rollCount
                                                    + " ロール）— 再生または1ステップ")
                                    : "試走準備完了 — 再生または1ステップ");
        }
        if (simulationSteps.isEmpty()) {
            resultLabel.setText("グラフ試走ステップがありません（ルール未選択・グラフ空・無効）");
        } else if (allRolls) {
            resultLabel.setText(
                    result.summaryJa()
                            + "（"
                            + result.steps().size()
                            + " ステップ / "
                            + (result.rollTotal() > 0 ? result.rollTotal() : rollCount)
                            + " ロール）");
        } else {
            resultLabel.setText(result.summaryJa() + "（ステップ " + simulationSteps.size() + " 件）");
        }
    }

    @FXML
    private void onStepAction() {
        if (simulationSteps.isEmpty()) {
            if (resultLabel != null) {
                resultLabel.setText("ステップがありません。先に「試走開始」を実行してください。");
            }
            return;
        }
        if (stepIndex >= simulationSteps.size()) {
            stepIndex = 0;
            clearPlaybackVisuals();
        }
        showStep();
    }

    @FXML
    private void onPlayAction() {
        if (simulationSteps.isEmpty()) {
            if (resultLabel != null) {
                resultLabel.setText("再生するステップがありません。先に「試走開始」を実行してください。");
            }
            return;
        }
        if (playTimeline != null
                && playTimeline.getStatus() == javafx.animation.Animation.Status.RUNNING) {
            stopPlay();
            if (resultLabel != null) {
                resultLabel.setText("再生を停止しました");
            }
            return;
        }
        stopPlay();
        stepIndex = 0;
        if (stepList != null) {
            stepList.getSelectionModel().clearSelection();
        }
        clearPlaybackVisuals();
        if (stepLabel != null) {
            stepLabel.setText("");
        }
        showStep();
        double speed = speedSlider != null ? speedSlider.getValue() : 1.0;
        long intervalMs = Math.max(150L, Math.round(600.0 / speed));
        playTimeline =
                new javafx.animation.Timeline(
                        new javafx.animation.KeyFrame(
                                javafx.util.Duration.millis(intervalMs),
                                e -> {
                                    if (stepIndex >= simulationSteps.size()) {
                                        stopPlay();
                                        if (resultLabel != null) {
                                            resultLabel.setText("試走完了");
                                        }
                                        return;
                                    }
                                    showStep();
                                }));
        playTimeline.setCycleCount(javafx.animation.Animation.INDEFINITE);
        playTimeline.play();
    }

    @FXML
    private void onResetAction() {
        stopPlay();
        stepIndex = 0;
        clearPlaybackVisuals();
        if (stepLabel != null) {
            stepLabel.setText("");
        }
        if (stepList != null && !stepList.getItems().isEmpty()) {
            stepList.getSelectionModel().clearSelection();
        }
    }

    private void stopPlay() {
        if (playTimeline != null) {
            playTimeline.stop();
            playTimeline = null;
        }
    }

    private void clearPlaybackVisuals() {
        simulationGraph.setHighlightedNodeId("");
        simulationGraph.clearWipOverlay();
        if (rollStageVisual != null) {
            rollStageVisual.clear();
        }
    }

    private Map<String, String> resolveSelectedTaskRow() {
        if (shell == null || taskPicker == null) {
            return Map.of();
        }
        String label = taskPicker.getValue();
        if (label == null || label.isBlank()) {
            return Map.of();
        }
        PlanInputTabController planInput = shell.planInputTabControllerForDispatchRollUnit();
        Optional<Map<String, String>> fromMemory =
                planInput != null ? planInput.findPlanRowMapByLabel(label) : Optional.empty();
        if (fromMemory.isPresent()) {
            return fromMemory.get();
        }
        return taskSource.findRowByLabel(label).orElse(Map.of());
    }

    private void showStep() {
        if (stepList == null || simulationSteps.isEmpty()) {
            return;
        }
        if (stepIndex >= simulationSteps.size()) {
            resultLabel.setText("試走完了");
            return;
        }
        DispatchRuleSimulationStep step = simulationSteps.get(stepIndex);
        stepList.getSelectionModel().select(stepIndex);
        stepList.scrollTo(stepIndex);
        String rollPrefix =
                step.rollIndex() > 0 ? "ロール " + step.rollIndex() + "/" + step.rollTotal() + " — " : "";
        stepLabel.setText(
                rollPrefix
                        + "ステップ "
                        + (stepIndex + 1)
                        + "/"
                        + simulationSteps.size()
                        + ": "
                        + step.summaryJa());
        if (rollStageVisual != null
                && (step.rollIndex() > 0
                        || step.preInputRawRolls() > 0
                        || step.connectionRolls() > 0
                        || step.secBeforeWipRolls() > 0
                        || step.secCompleteRolls() > 0)) {
            rollStageVisual.update(
                    step.preInputRawRolls(),
                    step.connectionRolls(),
                    step.secBeforeWipRolls(),
                    step.secCompleteRolls(),
                    step.wipCount(),
                    WIP_THRESHOLD,
                    step.rollIndex(),
                    step.rollTotal());
        }
        double wip = step.wipCount();
        if (wip > 0) {
            simulationGraph.setWipOverlay(wip, WIP_THRESHOLD);
        }
        if (step.rollAccumulateStep()) {
            simulationGraph.setHighlightedNodeId("");
        } else {
            simulationGraph.setHighlightedNodeId(step.nodeId());
        }
        stepIndex++;
    }

    private static String stepKey(DispatchRuleSimulationStep s) {
        if (s.rollAccumulateStep()) {
            if ("connection".equals(s.flowPhase())) {
                return "接続+";
            }
            if ("sec".equals(s.flowPhase())) {
                return "SEC+";
            }
            return "WIP+";
        }
        return s.nodeId();
    }
}
