package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Consumer;

import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.ui.ConflictSaveDialog;

/**
 * コントローラから呼ぶ保存前競合ゲート。
 *
 * @return true なら保存を続行、false なら中止（再読込済みの場合も含む）
 */
public final class SaveConflictUiGate {

    private SaveConflictUiGate() {}

    public static boolean allowSave(
            Window owner,
            String screenTitle,
            FingerprintBaseline baseline,
            ConflictDiffSummarizer summarizer,
            Runnable onReload,
            Consumer<String> onStatusError) {
        if (baseline == null) {
            return true;
        }
        ConflictCheckResult check = SaveConflictChecker.check(baseline);
        if (check.kind() == ConflictCheckResult.Kind.IO_ERROR) {
            String msg = "保存中止: 競合確認に失敗しました — " + check.errorMessage();
            if (onStatusError != null) {
                onStatusError.accept(msg);
            } else {
                Alert a = new Alert(Alert.AlertType.ERROR, msg, ButtonType.OK);
                a.showAndWait();
            }
            return false;
        }
        if (check.kind() == ConflictCheckResult.Kind.OK) {
            return true;
        }
        Map<Path, byte[]> disk;
        try {
            disk = SaveConflictGate.readDiskBytes(baseline);
        } catch (IOException e) {
            String msg = "保存中止: " + e.getMessage();
            if (onStatusError != null) {
                onStatusError.accept(msg);
            }
            return false;
        }
        String summary =
                SaveConflictGate.summarizeOrFallback(summarizer, baseline, disk, check);
        ConflictSaveChoice choice = ConflictSaveDialog.show(owner, screenTitle, summary);
        if (choice == ConflictSaveChoice.CANCEL) {
            return false;
        }
        if (choice == ConflictSaveChoice.RELOAD) {
            if (onReload != null) {
                onReload.run();
            }
            return false;
        }
        return true;
    }
}
