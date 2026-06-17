package jp.co.pm.ai.desktop;

import java.util.Map;

import javafx.scene.control.Dialog;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.FactorySite;

/** リモートデスクトップ・ユーザー管理タブが依存するメインシェル操作。 */
public interface DesktopShellHost {

    Map<String, String> snapshotUiEnv();

    void updateEnvTabValue(String envKey, String value);

    void appendLog(String line);

    Stage primaryStageForDialogs();

    void requireOperatorSelectionForFactory(FactorySite site, boolean startup);

    void refreshOperatorUserPresentation();

    /** 操作者変更後にリモートデスクトップタブの RPA設定 ini パス表示・読込を同期する。 */
    default void refreshRemoteDesktopOperatorContext() {}

    void prepareDialogForMainTheme(Dialog<?> dialog);

    void showWarningDialog(String title, String message);

    void showInformationDialog(String title, String message);
}
