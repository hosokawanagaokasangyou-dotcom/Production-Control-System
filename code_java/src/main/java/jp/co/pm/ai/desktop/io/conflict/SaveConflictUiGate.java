package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Consumer;

import javafx.application.Platform;
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
            reportError(
                    onStatusError,
                    "保存中止: 競合確認用の読込指紋がありません。再読込してから保存してください。");
            return false;
        }
        if (!Platform.isFxApplicationThread()) {
            reportError(
                    onStatusError,
                    "保存中止: 競合確認は JavaFX アプリケーションスレッドで実行してください。");
            return false;
        }
        ConflictCheckResult check = SaveConflictChecker.check(baseline);
        if (check.kind() == ConflictCheckResult.Kind.IO_ERROR) {
            reportError(
                    onStatusError,
                    "保存中止: 競合確認に失敗しました — " + check.errorMessage());
            return false;
        }
        if (check.kind() == ConflictCheckResult.Kind.OK) {
            return true;
        }
        Map<Path, byte[]> disk;
        try {
            disk = SaveConflictGate.readDiskBytes(baseline);
        } catch (IOException e) {
            reportError(onStatusError, "保存中止: " + e.getMessage());
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

    private static void reportError(Consumer<String> onStatusError, String msg) {
        if (onStatusError != null) {
            onStatusError.accept(msg);
            return;
        }
        Alert a = new Alert(Alert.AlertType.ERROR, msg, ButtonType.OK);
        a.showAndWait();
    }
}
