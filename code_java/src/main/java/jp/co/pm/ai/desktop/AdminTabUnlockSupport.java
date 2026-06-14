package jp.co.pm.ai.desktop;

import java.io.IOException;

import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AdminTabCredentialsStore;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;

/** ユーザー管理者タブ解錠ダイアログ（PMD / RDP 共通）。 */
public final class AdminTabUnlockSupport {

    @FunctionalInterface
    public interface DialogPreparer {
        void prepare(Dialog<?> dialog);
    }

    private AdminTabUnlockSupport() {}

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
        Label hint =
                new Label(
                        "ユーザー管理者タブを開くには、ユーザー名 "
                                + FactoryOperatorUserStore.ADMIN_TAB_USERNAME
                                + " と管理者パスワードを入力してください。");
        hint.setWrapText(true);
        TextField userField = new TextField();
        userField.setPromptText(FactoryOperatorUserStore.ADMIN_TAB_USERNAME);
        PasswordField pf = new PasswordField();
        pf.setPromptText("管理者パスワード");
        VBox box =
                new VBox(
                        8,
                        hint,
                        new Label("ユーザー名:"),
                        userField,
                        new Label("パスワード:"),
                        pf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        var ans = dialog.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return false;
        }
        if (FactoryOperatorUserStore.verifyAdminTabAccess(userField.getText(), pf.getText())) {
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
}
