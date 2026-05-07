package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.dispatch.MachineCalendarBlockIndex;

/**
 * Preview / export {@link AppPaths#MACHINE_CALENDAR_BLOCKS_JSON_BASENAME} for interactive dispatch blocking.
 */
public final class MachineCalendarTabController {

    private MainShellController shell;

    @FXML
    private Button exportButton;

    @FXML
    private Button reloadButton;

    @FXML
    private Label statusLabel;

    @FXML
    private ProgressBar progressBar;

    @FXML
    private Label pathLabel;

    @FXML
    private TextArea jsonArea;

    @FXML
    private void initialize() {
        jsonArea.setStyle("-fx-font-family: monospace;");
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        Platform.runLater(this::refreshPathLabelAndPreview);
    }

    private Path resolveJsonPath() {
        if (shell == null) {
            return null;
        }
        return AppPaths.resolveMachineCalendarBlocksJsonPath(shell.snapshotUiEnv());
    }

    private Path resolvePythonExe() {
        String py =
                shell.snapshotUiEnv().getOrDefault(AppPaths.KEY_PM_AI_PYTHON, "").trim();
        if (!py.isEmpty()) {
            return Path.of(py);
        }
        return Path.of("python3");
    }

    private void refreshPathLabelAndPreview() {
        Path p = resolveJsonPath();
        if (pathLabel != null) {
            pathLabel.setText(p != null ? p.toString() : "");
        }
        reloadFromDiskUi();
    }

    private void reloadFromDiskUi() {
        Path p = resolveJsonPath();
        if (p == null || jsonArea == null) {
            return;
        }
        try {
            if (!Files.isRegularFile(p)) {
                jsonArea.setText("");
                if (statusLabel != null) {
                    statusLabel.setText("ファイルなし");
                }
                return;
            }
            String s = Files.readString(p, StandardCharsets.UTF_8);
            jsonArea.setText(s);
            if (statusLabel != null) {
                statusLabel.setText(s.length() + " chars");
            }
        } catch (Exception e) {
            jsonArea.setText("");
            if (statusLabel != null) {
                statusLabel.setText("読込エラー");
            }
            if (shell != null) {
                shell.appendLog("[machine-calendar-json] read failed: " + e.getMessage());
            }
        }
    }

    @FXML
    private void onReloadFromDiskAction() {
        refreshPathLabelAndPreview();
    }

    @FXML
    private void onExportFromMasterAction() {
        if (shell == null) {
            return;
        }
        showProgress(true);
        final MainShellController shellRef = shell;
        Task<MachineCalendarBlockIndex.LoadOutcome> task =
                new Task<>() {
                    @Override
                    protected MachineCalendarBlockIndex.LoadOutcome call() throws Exception {
                        Path jsonOut =
                                AppPaths.resolveMachineCalendarBlocksJsonPath(shellRef.snapshotUiEnv());
                        Path primary =
                                AppPaths.resolveMasterWorkbookPathResolved(
                                        shellRef.snapshotUiEnv(),
                                        shellRef.effectiveTaskInputWorkbookPathForShell());
                        Path summary = AppPaths.summaryAiDispatchXlsmPath(shellRef.snapshotUiEnv());
                        Path pyExe = resolvePythonExe();
                        Path pyDir = AppPaths.resolvePythonScriptDir(shellRef.snapshotUiEnv());
                        // #region agent log
                        try {
                            Map<String, Object> d = new LinkedHashMap<>();
                            d.put("pyExe", pyExe.toString());
                            d.put("pyDir", pyDir.toString());
                            d.put("primaryMaster", primary.toString());
                            d.put("summaryXlsm", summary != null ? summary.toString() : "");
                            d.put("jsonOut", jsonOut.toString());
                            AgentDebugLog.appendStructured(
                                    shellRef.snapshotUiEnv(),
                                    "ecf65d",
                                    "H1",
                                    "MachineCalendarTabController.onExportFromMasterAction",
                                    "before exportWithSummaryFallbackToJsonFile",
                                    d);
                        } catch (Throwable ignored) {
                        }
                        // #endregion
                        return MachineCalendarBlockIndex.exportWithSummaryFallbackToJsonFile(
                                primary, summary, pyExe, pyDir, jsonOut);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    MachineCalendarBlockIndex.LoadOutcome lo = task.getValue();
                    showProgress(false);
                    refreshPathLabelAndPreview();
                    if (lo.pythonJsonError() != null) {
                        shellRef.appendLog(
                                "[machine-calendar-json] python error field: " + lo.pythonJsonError());
                    }
                    if (lo.pythonDiagnosticsJson() != null) {
                        shellRef.appendLog(
                                "[machine-calendar-json] diagnostics: " + lo.pythonDiagnosticsJson());
                    }
                    shellRef.appendLog(
                            "[machine-calendar-json] wrote blocks; empty="
                                    + lo.index().isEmpty());
                    shellRef.notifyMachineCalendarJsonUpdated();
                });
        task.setOnFailed(
                e -> {
                    showProgress(false);
                    Throwable ex = task.getException();
                    if (shellRef != null) {
                        shellRef.appendLog(
                                "[machine-calendar-json] export failed: "
                                        + (ex != null ? ex.getMessage() : ""));
                    }
                });
        new Thread(task, "machine-calendar-json-export").start();
    }

    private void showProgress(boolean on) {
        if (progressBar != null) {
            progressBar.setManaged(on);
            progressBar.setVisible(on);
            if (on) {
                progressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
            }
        }
        if (exportButton != null) {
            exportButton.setDisable(on);
        }
        if (reloadButton != null) {
            reloadButton.setDisable(on);
        }
    }
}
