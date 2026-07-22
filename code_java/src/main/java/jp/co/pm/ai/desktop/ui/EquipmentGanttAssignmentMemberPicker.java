package jp.co.pm.ai.desktop.ui;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.io.SkillsSheetMemberReader;

/** 設備ガント担当割当の追加用・資格者1名選択ダイアログ。 */
public final class EquipmentGanttAssignmentMemberPicker {

    private static final LimitedOperatorLoadCoordinator LOAD_COORDINATOR =
            new LimitedOperatorLoadCoordinator();
    private static final ExecutorService LOAD_EXECUTOR =
            Executors.newSingleThreadExecutor(
                    runnable -> {
                        Thread thread =
                                new Thread(runnable, "equipment-gantt-assignment-skills-loader");
                        thread.setDaemon(true);
                        return thread;
                    });

    private EquipmentGanttAssignmentMemberPicker() {}

    public static void pickSingleMemberAsync(
            MainShellController shell,
            Window owner,
            String processName,
            String machineName,
            double anchorScreenX,
            double anchorScreenY,
            Consumer<String> onSelectedFullName) {
        Path master = shell != null ? shell.resolveMasterWorkbookIfPresent() : null;
        if (master == null) {
            if (shell != null) {
                shell.showErrorDialog(
                        "担当を追加",
                        "master.xlsm を解決できません。環境変数 PM_AI_MASTER_WORKBOOK を確認してください。");
            }
            return;
        }

        javafx.scene.control.Dialog<Void> busyDialog = LimitedOperatorCellEditor.createBusyDialog(owner);
        busyDialog.setTitle("担当を追加");
        busyDialog.setHeaderText("資格候補を読み込み中");
        busyDialog.show();
        String proc = processName != null ? processName.strip() : "";
        String mach = machineName != null ? machineName.strip() : "";
        boolean started =
                LOAD_COORDINATOR.submit(
                        () -> loadCandidates(master, proc, mach),
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
                                                "",
                                                anchorScreenX,
                                                anchorScreenY)
                                        .ifPresent(
                                                encoded -> {
                                                    List<String> names =
                                                            LimitedOperatorJsonCodec.decode(encoded);
                                                    if (names.isEmpty()) {
                                                        return;
                                                    }
                                                    if (names.size() != 1) {
                                                        shell.showWarningDialog(
                                                                "担当を追加",
                                                                "1名のみ選択してください。");
                                                        return;
                                                    }
                                                    onSelectedFullName.accept(names.getFirst());
                                                });
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
                    "担当を追加", "資格候補を読み込み中です。完了してから再度操作してください。");
        }
    }

    private static List<String> loadCandidates(Path master, String processName, String machineName)
            throws IOException {
        List<String> qualified =
                SkillsSheetMemberReader.readQualifiedMemberDisplayNames(
                        master, processName, machineName);
        if (!qualified.isEmpty()) {
            return qualified;
        }
        return SkillsSheetMemberReader.readMemberDisplayNames(master);
    }

    private static void showLoadError(MainShellController shell, Throwable failure) {
        shell.showErrorDialog(
                "担当を追加",
                failure.getMessage() != null ? failure.getMessage() : failure.toString());
    }
}
