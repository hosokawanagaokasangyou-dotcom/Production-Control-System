package jp.co.pm.ai.desktop.dispatch.rules.trace;

import java.util.List;
import java.util.Map;

import javafx.fxml.FXML;
import javafx.scene.control.ListView;

import jp.co.pm.ai.desktop.MainShellController;

/** Application trace child tab. */
public final class SpecialRulesTraceTabController {

    @FXML private ListView<String> eventList;

    private MainShellController shell;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        reload();
    }

    @FXML
    private void onReloadAction() {
        reload();
    }

    private void reload() {
        if (shell == null || eventList == null) {
            return;
        }
        try {
            List<DispatchRuleTraceLoader.ApplicationEvent> events =
                    DispatchRuleTraceLoader.loadFromWorkDir(shell.dispatchRulesUiEnv());
            eventList.getItems().setAll(
                    events.stream()
                            .map(
                                    e ->
                                            e.ruleId()
                                                    + " "
                                                    + e.taskId()
                                                    + " "
                                                    + e.effect()
                                                    + " "
                                                    + e.summaryJa())
                            .toList());
        } catch (Exception ex) {
            eventList.getItems().setAll("（sidecar 未生成 — 段階2 実行後に表示）");
        }
    }
}
