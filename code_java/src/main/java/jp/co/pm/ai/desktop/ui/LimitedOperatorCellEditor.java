package jp.co.pm.ai.desktop.ui;

import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.io.SkillsSheetMemberReader;

/** 入力1表の「担当OP_限定」セル用非同期編集処理。 */
public final class LimitedOperatorCellEditor {

    private static final LimitedOperatorLoadCoordinator LOAD_COORDINATOR =
            new LimitedOperatorLoadCoordinator();
    private static final ExecutorService LOAD_EXECUTOR =
            Executors.newSingleThreadExecutor(
                    runnable -> {
                        Thread thread = new Thread(runnable, "limited-operator-skills-loader");
                        thread.setDaemon(true);
                        return thread;
                    });

    private LimitedOperatorCellEditor() {}

    public static void editAsync(
            MainShellController shell,
            Window owner,
            List<String> headers,
            List<String> row,
            String currentValue,
            double anchorScreenX,
            double anchorScreenY,
            Consumer<String> onConfirmed) {
        LimitedOperatorEditContext context =
                LimitedOperatorEditContext.fromRow(headers, row);
        try {
            context.validateComplete();
        } catch (IllegalArgumentException ex) {
            shell.showErrorDialog("担当OP_限定", ex.getMessage());
            return;
        }

        Path master = shell.resolveMasterWorkbookIfPresent();
        if (master == null) {
            shell.showErrorDialog(
                    "担当OP_限定",
                    "master.xlsm を解決できません。環境変数 PM_AI_MASTER_WORKBOOK を確認してください。");
            return;
        }

        Dialog<Void> busyDialog = createBusyDialog(owner);
        busyDialog.show();
        boolean started =
                LOAD_COORDINATOR.submit(
                        () ->
                                SkillsSheetMemberReader.readQualifiedMemberDisplayNames(
                                        master,
                                        context.processName(),
                                        context.machineName()),
                        LOAD_EXECUTOR,
                        Platform::runLater,
                        candidates -> {
                            if (!busyDialog.isShowing()) {
                                return;
                            }
                            busyDialog.close();
                            try {
                                LimitedOperatorChecklistDialog.edit(
                                                owner,
                                                candidates,
                                                currentValue,
                                                anchorScreenX,
                                                anchorScreenY)
                                        .ifPresent(onConfirmed);
                            } catch (Exception ex) {
                                showLoadError(shell, ex);
                            }
                        },
                        failure -> {
                            if (!busyDialog.isShowing()) {
                                return;
                            }
                            busyDialog.close();
                            showLoadError(shell, failure);
                        });
        if (!started) {
            busyDialog.close();
            shell.showInformationDialog(
                    "担当OP_限定",
                    "資格候補を読み込み中です。完了してから再度操作してください。");
        }
    }

    static Dialog<Void> createBusyDialog(Window owner) {
        Dialog<Void> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        dialog.setTitle("担当OP_限定");
        dialog.setHeaderText("資格候補を読み込み中");
        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        dialog.getDialogPane()
                .setContent(new VBox(10, progress, new Label("master.xlsm の skills を確認しています。")));
        dialog.getDialogPane().setPrefWidth(360);
        dialog.getDialogPane()
                .getButtonTypes()
                .add(new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE));
        return dialog;
    }

    private static void showLoadError(MainShellController shell, Throwable failure) {
        shell.showErrorDialog(
                "担当OP_限定",
                failure.getMessage() != null ? failure.getMessage() : failure.toString());
    }
}
