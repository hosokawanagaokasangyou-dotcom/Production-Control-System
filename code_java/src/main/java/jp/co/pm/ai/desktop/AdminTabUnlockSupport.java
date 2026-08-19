package jp.co.pm.ai.desktop;

import java.io.IOException;

import javafx.application.Platform;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

/** ユーザー管理者タブ解錠ダイアログ（PMD / RDP 共通）。 */
public final class AdminTabUnlockSupport {

    @FunctionalInterface
    public interface DialogPreparer {
        void prepare(Dialog<?> dialog);
    }

    private AdminTabUnlockSupport() {}

    /** フォールバックパスワード（前後空白は無視）。 */
    public static boolean verifyFallbackPassword(String entered) {
        String input = entered != null ? entered.strip() : "";
        return FactoryOperatorUserStore.ADMIN_TAB_PASSWORD.equals(input);
    }

    /**
     * セッション操作者が管理者ならダイアログなしで true。それ以外はパスワード入力。
     *
     * <p>管理者が未設定でも、管理者がログインしていなくても、フォールバックパスワードで入れる。
     */
    public static boolean ensureUnlocked(Stage stage, DialogPreparer preparer) {
        return ensureUnlocked(stage, preparer, null);
    }

    public static boolean ensureUnlocked(Stage stage, DialogPreparer preparer, FactorySite site) {
        if (stage == null) {
            return false;
        }
        FactorySite scope =
                site != null
                        ? site
                        : FactoryOperatorUserStore.operatorScopeForCurrentApp(java.util.Map.of(), null);
        try {
            if (FactoryOperatorUserStore.sessionOperatorIsAdmin(scope)) {
                return true;
            }
        } catch (IOException ignored) {
            // ストア読込失敗時はパスワードへフォールバック
        }
        Dialog<ButtonType> dialog = new Dialog<>();
        if (preparer != null) {
            preparer.prepare(dialog);
        }
        dialog.setTitle("ユーザー管理者");
        dialog.setHeaderText(null);
        Label hint =
                new Label(
                        "ログイン中のユーザーが管理者のときはパスワード不要です。"
                                + "管理者が未設定、または管理者が不在のときはパスワードを入力してください。");
        hint.setWrapText(true);
        PasswordField inputField = new PasswordField();
        inputField.setPromptText("パスワード");
        VBox box = new VBox(8, hint, new Label("パスワード:"), inputField);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        focusInputWhenDialogShown(dialog, inputField);
        var ans = dialog.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return false;
        }
        return verifyFallbackPassword(inputField.getText());
    }

    private static void focusInputWhenDialogShown(Dialog<?> dialog, javafx.scene.Node input) {
        if (dialog == null || input == null) {
            return;
        }
        dialog.setOnShown(e -> Platform.runLater(input::requestFocus));
    }
}
