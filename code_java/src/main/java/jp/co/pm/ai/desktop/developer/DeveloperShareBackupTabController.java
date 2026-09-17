package jp.co.pm.ai.desktop.developer;

import java.awt.Desktop;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryShareBackupStore;
import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * 湖南・国分の工場共有 DATA をローカルと共有の両方へバックアップする。
 */
public final class DeveloperShareBackupTabController {

    private MainShellController shell;
    private boolean running;

    @FXML private Label localPathLabel;
    @FXML private Label konanPathLabel;
    @FXML private Label kokubuPathLabel;
    @FXML private Label statusLabel;
    @FXML private TextArea logArea;
    @FXML private Button backupButton;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshPaths();
    }

    public void onMainShellTabSelected() {
        refreshPaths();
    }

    @FXML
    private void onBackupAction() {
        if (running || shell == null) {
            return;
        }
        running = true;
        if (backupButton != null) {
            backupButton.setDisable(true);
        }
        setStatus("バックアップ中…");
        Map<String, String> ui = shell.snapshotUiEnv();
        Task<FactoryShareBackupStore.Result> task =
                new Task<>() {
                    @Override
                    protected FactoryShareBackupStore.Result call() throws Exception {
                        return FactoryShareBackupStore.backupNow(ui, null);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    running = false;
                    if (backupButton != null) {
                        backupButton.setDisable(false);
                    }
                    FactoryShareBackupStore.Result result = task.getValue();
                    showResult(result);
                    if (shell != null) {
                        shell.recordOperatorAction(
                                "developer",
                                "factory_share_backup",
                                result != null && result.ok() ? "ok" : "error",
                                result == null ? "失敗" : result.generationId());
                    }
                });
        task.setOnFailed(
                e -> {
                    running = false;
                    if (backupButton != null) {
                        backupButton.setDisable(false);
                    }
                    Throwable err = task.getException();
                    String msg = err == null ? "失敗" : String.valueOf(err.getMessage());
                    setStatus("失敗: " + msg);
                    appendLog("失敗: " + msg);
                    if (shell != null) {
                        shell.recordOperatorAction(
                                "developer", "factory_share_backup", "error", msg);
                    }
                });
        Thread t = new Thread(task, "factory-share-backup");
        t.setDaemon(true);
        t.start();
    }

    @FXML
    private void onOpenLocalAction() {
        openDir(FactoryShareBackupStore.resolveLocalBackupRoot());
    }

    @FXML
    private void onOpenKonanSharedAction() {
        if (shell == null) {
            return;
        }
        openDir(FactoryShareBackupStore.resolveSharedBackupRoot(shell.snapshotUiEnv(), FactorySite.KONAN));
    }

    @FXML
    private void onOpenKokubuSharedAction() {
        if (shell == null) {
            return;
        }
        openDir(
                FactoryShareBackupStore.resolveSharedBackupRoot(
                        shell.snapshotUiEnv(), FactorySite.KOKUBU));
    }

    private void refreshPaths() {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path local = FactoryShareBackupStore.resolveLocalBackupRoot();
        Path konanShared = FactoryShareBackupStore.resolveSharedBackupRoot(ui, FactorySite.KONAN);
        Path kokubuShared = FactoryShareBackupStore.resolveSharedBackupRoot(ui, FactorySite.KOKUBU);
        Path konanSrc = FactoryShareBackupStore.resolveSource(ui, FactorySite.KONAN);
        Path kokubuSrc = FactoryShareBackupStore.resolveSource(ui, FactorySite.KOKUBU);
        if (localPathLabel != null) {
            localPathLabel.setText("ローカル: " + local);
        }
        if (konanPathLabel != null) {
            konanPathLabel.setText("湖南 ソース→共有: " + konanSrc + " → " + konanShared);
        }
        if (kokubuPathLabel != null) {
            kokubuPathLabel.setText("国分 ソース→共有: " + kokubuSrc + " → " + kokubuShared);
        }
    }

    private void showResult(FactoryShareBackupStore.Result result) {
        if (result == null) {
            setStatus("結果が空です");
            return;
        }
        StringBuilder sb = new StringBuilder();
        sb.append("世代 ").append(result.generationId()).append('\n');
        sb.append("ローカル ").append(result.localGenerationRoot()).append('\n');
        for (FactoryShareBackupStore.SiteResult site : result.sites()) {
            sb.append(site.site().displayLabelJa())
                    .append(" ローカル ")
                    .append(site.copiedToLocal())
                    .append("件");
            if (site.localError() != null && !site.localError().isBlank()) {
                sb.append(" 失敗: ").append(site.localError());
            }
            sb.append(" / 共有 ").append(site.copiedToShared()).append("件");
            if (site.sharedError() != null && !site.sharedError().isBlank()) {
                sb.append(" 失敗: ").append(site.sharedError());
            }
            sb.append('\n');
            if (site.localDest() != null) {
                sb.append("  local: ").append(site.localDest()).append('\n');
            }
            if (site.sharedDest() != null) {
                sb.append("  share: ").append(site.sharedDest()).append('\n');
            }
        }
        String text = sb.toString();
        appendLog(text);
        setStatus(result.ok() ? "完了 世代 " + result.generationId() : "一部失敗 世代 " + result.generationId());
    }

    private void setStatus(String text) {
        if (statusLabel != null) {
            statusLabel.setText(text);
        }
    }

    private void appendLog(String text) {
        if (logArea == null || text == null) {
            return;
        }
        Platform.runLater(
                () -> {
                    if (!logArea.getText().isEmpty()) {
                        logArea.appendText("\n");
                    }
                    logArea.appendText(text);
                });
    }

    private void openDir(Path dir) {
        if (dir == null) {
            return;
        }
        try {
            Files.createDirectories(dir);
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(dir.toFile());
            }
        } catch (Exception ex) {
            setStatus("フォルダを開けません: " + (ex.getMessage() != null ? ex.getMessage() : ex));
        }
    }
}
