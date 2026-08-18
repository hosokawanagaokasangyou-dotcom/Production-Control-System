package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.security.SecureRandom;

import javafx.application.Platform;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AdminTabCredentialsStore;

/** ユーザー管理者タブ解錠ダイアログ（PMD / RDP 共通）。 */
public final class AdminTabUnlockSupport {

    /** 画面表示・入力するロックNOの桁数。 */
    public static final int LOCK_NO_DIGITS = 4;

    private static final SecureRandom LOCK_NO_RANDOM = new SecureRandom();

    @FunctionalInterface
    public interface DialogPreparer {
        void prepare(Dialog<?> dialog);
    }

    private AdminTabUnlockSupport() {}

    /** ダイアログ表示用のロックNO（ゼロ埋めの数字列）。 */
    public static String generateLockNo() {
        int n = LOCK_NO_RANDOM.nextInt(10_000);
        return String.format("%04d", n);
    }

    /** 表示中のロックNOと入力値を照合する（前後空白は無視）。 */
    public static boolean verifyLockNo(String displayed, String entered) {
        if (displayed == null || displayed.isEmpty()) {
            return false;
        }
        String input = entered != null ? entered.strip() : "";
        return displayed.equals(input);
    }

    /**
     * 保存済み解錠が有効ならダイアログなしで true。未保存・失効時はダイアログを表示する。
     *
     * <p>OK で認証失敗した場合は保存済み解錠を削除する。
     */
    public static boolean ensureUnlocked(Stage stage, DialogPreparer preparer) {
        if (stage == null) {
            return false;
        }
        if (AdminTabCredentialsStore.hasValidSavedUnlock()) {
            return true;
        }
        Dialog<ButtonType> dialog = new Dialog<>();
        if (preparer != null) {
            preparer.prepare(dialog);
        }
        dialog.setTitle("ユーザー管理者");
        dialog.setHeaderText(null);
        String lockNo = generateLockNo();
        Label hint = new Label("ユーザー管理者タブを開くには、表示されているロックNOを入力してください。");
        hint.setWrapText(true);
        Label lockNoValue = new Label(lockNo);
        lockNoValue.setStyle("-fx-font-size: 28px; -fx-font-weight: bold;");
        TextField inputField = new TextField();
        inputField.setPromptText("ロックNO");
        VBox box =
                new VBox(
                        8,
                        hint,
                        new Label("ロックNO:"),
                        lockNoValue,
                        new Label("ロックNO入力:"),
                        inputField);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        focusInputWhenDialogShown(dialog, inputField);
        var ans = dialog.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return false;
        }
        if (verifyLockNo(lockNo, inputField.getText())) {
            try {
                AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
            } catch (IOException ignored) {
                // 解錠自体は成功。次回起動で再プロンプトされる。
            }
            return true;
        }
        AdminTabCredentialsStore.clearSavedUnlock();
        return false;
    }

    private static void focusInputWhenDialogShown(Dialog<?> dialog, javafx.scene.Node input) {
        if (dialog == null || input == null) {
            return;
        }
        dialog.setOnShown(e -> Platform.runLater(input::requestFocus));
    }
}
