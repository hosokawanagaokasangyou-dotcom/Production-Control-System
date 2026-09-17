package jp.co.pm.ai.desktop.developer;

import java.awt.Desktop;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TextArea;
import javafx.scene.layout.HBox;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.FactoryShareBackupStore;
import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * 湖南・国分の工場共有 DATA をローカルと共有の両方へバックアップする。
 */
public final class DeveloperShareBackupTabController {

    private MainShellController shell;
    private boolean running;
    private int probeSeq;

    @FXML private Label localPathLabel;
    @FXML private Label konanSourceLabel;
    @FXML private Label konanSharedLabel;
    @FXML private Label kokubuSourceLabel;
    @FXML private Label kokubuSharedLabel;
    @FXML private Label statusLabel;
    @FXML private TextArea logArea;
    @FXML private Button backupButton;
    @FXML private Button openLocalButton;
    @FXML private Button openKonanButton;
    @FXML private Button openKokubuButton;
    @FXML private HBox loadingChip;
    @FXML private ProgressIndicator loadingIndicator;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshPaths();
    }

    public void onMainShellTabSelected() {
        refreshPaths();
        probeSourceExistence();
    }

    @FXML
    private void onBackupAction() {
        if (running || shell == null) {
            return;
        }
        if (!confirmBackup()) {
            return;
        }
        setRunning(true);
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
                    FactoryShareBackupStore.Result result = task.getValue();
                    showResult(result);
                    recordBackupAction(result);
                    setRunning(false);
                });
        task.setOnFailed(
                e -> {
                    Throwable err = task.getException();
                    String msg = err == null ? "失敗" : String.valueOf(err.getMessage());
                    setStatus("失敗: " + msg);
                    appendLog("失敗: " + msg);
                    recordBackupAction(null);
                    setRunning(false);
                });
        task.setOnCancelled(e -> setRunning(false));
        Thread t = new Thread(task, "factory-share-backup");
        t.setDaemon(true);
        try {
            t.start();
        } catch (RuntimeException ex) {
            setRunning(false);
            setStatus("開始できません: " + (ex.getMessage() != null ? ex.getMessage() : ex));
        }
    }

    @FXML
    private void onOpenLocalAction() {
        openDir(FactoryShareBackupStore.resolveLocalBackupRoot(), true, "ローカルバックアップ");
    }

    @FXML
    private void onOpenKonanSharedAction() {
        if (shell == null) {
            return;
        }
        openDir(
                FactoryShareBackupStore.resolveSharedBackupRoot(
                        shell.snapshotUiEnv(), FactorySite.KONAN),
                false,
                "湖南共有バックアップ");
    }

    @FXML
    private void onOpenKokubuSharedAction() {
        if (shell == null) {
            return;
        }
        openDir(
                FactoryShareBackupStore.resolveSharedBackupRoot(
                        shell.snapshotUiEnv(), FactorySite.KOKUBU),
                false,
                "国分共有バックアップ");
    }

    private boolean confirmBackup() {
        Map<String, String> ui = shell.snapshotUiEnv();
        Path konanSrc = FactoryShareBackupStore.resolveSource(ui, FactorySite.KONAN);
        Path kokubuSrc = FactoryShareBackupStore.resolveSource(ui, FactorySite.KOKUBU);
        String message =
                "湖南と国分の共有 DATA を、このPCと各工場の「バックアップ」へコピーします。\n"
                        + "世代は"
                        + FactoryShareBackupStore.MAX_GENERATIONS
                        + "件まで残し、古いものは削除します。所要時間は共有フォルダの大きさによります。\n\n"
                        + "湖南ソース: "
                        + konanSrc
                        + "\n国分ソース: "
                        + kokubuSrc
                        + "\n\n実行しますか？";
        Alert alert =
                new Alert(
                        Alert.AlertType.CONFIRMATION, message, ButtonType.OK, ButtonType.CANCEL);
        alert.setTitle("工場共有バックアップ");
        alert.setHeaderText("湖南と国分の共有フォルダをバックアップします");
        shell.prepareDialogForMainTheme(alert);
        Optional<ButtonType> ans = alert.showAndWait();
        return ans.isPresent() && ans.get() == ButtonType.OK;
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
        if (konanSourceLabel != null) {
            konanSourceLabel.setText("湖南ソース: " + konanSrc);
        }
        if (konanSharedLabel != null) {
            konanSharedLabel.setText("湖南共有先: " + konanShared);
        }
        if (kokubuSourceLabel != null) {
            kokubuSourceLabel.setText("国分ソース: " + kokubuSrc);
        }
        if (kokubuSharedLabel != null) {
            kokubuSharedLabel.setText("国分共有先: " + kokubuShared);
        }
    }

    private void probeSourceExistence() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        int seq = ++probeSeq;
        Thread t =
                new Thread(
                        () -> {
                            boolean konan =
                                    Files.isDirectory(
                                            FactoryShareBackupStore.resolveSource(
                                                    ui, FactorySite.KONAN));
                            boolean kokubu =
                                    Files.isDirectory(
                                            FactoryShareBackupStore.resolveSource(
                                                    ui, FactorySite.KOKUBU));
                            Platform.runLater(
                                    () -> {
                                        if (seq != probeSeq) {
                                            return;
                                        }
                                        appendExistence(konanSourceLabel, konan);
                                        appendExistence(kokubuSourceLabel, kokubu);
                                    });
                        },
                        "factory-share-backup-probe");
        t.setDaemon(true);
        t.start();
    }

    private static void appendExistence(Label label, boolean exists) {
        if (label == null) {
            return;
        }
        String text = label.getText();
        if (text == null) {
            text = "";
        }
        text = text.replace(" （あり）", "").replace(" （なし）", "");
        label.setText(text + (exists ? " （あり）" : " （なし）"));
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
            if (site.localError() != null) {
                sb.append(" 失敗: ").append(site.localError());
            }
            sb.append(" / 共有 ").append(site.copiedToShared()).append("件");
            if (site.sharedError() != null) {
                sb.append(" 失敗: ").append(site.sharedError());
            }
            sb.append('\n');
            if (site.localDest() != null) {
                sb.append("  ローカル: ").append(site.localDest()).append('\n');
            }
            if (site.sharedDest() != null) {
                sb.append("  共有: ").append(site.sharedDest()).append('\n');
            }
        }
        appendLog(sb.toString());
        if (result.ok()) {
            setStatus("完了 世代 " + result.generationId());
        } else {
            String fail = result.failureSummaryJa();
            setStatus(
                    "一部失敗 世代 "
                            + result.generationId()
                            + (fail.isBlank() ? "" : "（" + fail + "）"));
        }
        probeSourceExistence();
    }

    private void recordBackupAction(FactoryShareBackupStore.Result result) {
        if (shell == null) {
            return;
        }
        if (result == null) {
            shell.recordOperatorAction("developer", "factory_share_backup", "error", "失敗");
            return;
        }
        shell.recordOperatorAction(
                "developer",
                "factory_share_backup",
                result.ok() ? "ok" : "error",
                result.ok() ? result.generationId() : result.generationId() + " " + result.failureSummaryJa());
    }

    private void setRunning(boolean on) {
        running = on;
        if (backupButton != null) {
            backupButton.setDisable(on);
        }
        if (openLocalButton != null) {
            openLocalButton.setDisable(on);
        }
        if (openKonanButton != null) {
            openKonanButton.setDisable(on);
        }
        if (openKokubuButton != null) {
            openKokubuButton.setDisable(on);
        }
        if (loadingChip != null) {
            loadingChip.setVisible(on);
            loadingChip.setManaged(on);
        }
        if (loadingIndicator != null) {
            loadingIndicator.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
        }
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
        Runnable add =
                () -> {
                    if (!logArea.getText().isEmpty()) {
                        logArea.appendText("\n");
                    }
                    logArea.appendText(text);
                };
        if (Platform.isFxApplicationThread()) {
            add.run();
        } else {
            Platform.runLater(add);
        }
    }

    private void openDir(Path dir, boolean createIfMissing, String kind) {
        if (dir == null) {
            return;
        }
        setStatus(kind + " を開いています…");
        Thread t =
                new Thread(
                        () -> {
                            try {
                                if (createIfMissing) {
                                    Files.createDirectories(dir);
                                } else if (!Files.isDirectory(dir)) {
                                    Platform.runLater(
                                            () -> {
                                                setStatus(kind + " がありません");
                                                warn(
                                                        kind + " を開けません",
                                                        "フォルダがありません。先にバックアップを実行してください。\n"
                                                                + dir);
                                            });
                                    return;
                                }
                                if (!Desktop.isDesktopSupported()
                                        || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                                    Platform.runLater(
                                            () -> {
                                                setStatus("この環境ではフォルダを開けません");
                                                warn("開けません", "この環境では外部アプリ起動に対応していません。");
                                            });
                                    return;
                                }
                                Desktop.getDesktop().open(dir.toFile());
                                Platform.runLater(() -> setStatus(kind + " を開きました"));
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            String msg =
                                                    ex.getMessage() != null
                                                            ? ex.getMessage()
                                                            : ex.toString();
                                            setStatus(kind + " を開けません: " + msg);
                                            warn(kind + " を開けません", msg);
                                        });
                            }
                        },
                        "factory-share-backup-open");
        t.setDaemon(true);
        t.start();
    }

    private void warn(String title, String body) {
        Alert alert = new Alert(Alert.AlertType.WARNING, body, ButtonType.OK);
        alert.setTitle(title);
        alert.setHeaderText(null);
        if (shell != null) {
            shell.prepareDialogForMainTheme(alert);
        }
        alert.showAndWait();
    }
}
