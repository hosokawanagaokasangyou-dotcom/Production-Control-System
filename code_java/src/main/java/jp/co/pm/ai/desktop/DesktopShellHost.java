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

    /**
     * セッション中の操作者を変更する（起動時・工場切替とは別フロー。既に操作者がいる場合も選択ダイアログを出す）。
     */
    void changeSessionOperator(FactorySite site);

    void refreshOperatorUserPresentation();

    /** 操作者変更後にリモートデスクトップタブの RPA設定 ini パス表示・読込を同期する。 */
    default void refreshRemoteDesktopOperatorContext() {}

    void prepareDialogForMainTheme(Dialog<?> dialog);

    void showWarningDialog(String title, String message);

    void showInformationDialog(String title, String message);

    /**
     * グローバルステータスバーに長時間タスクの進捗を表示する。
     *
     * @param fraction 0.0–1.0。{@link Double#isNaN()} で不定（スピナー）。
     * @param detail 状況文言（メッセージに追記）
     */
    default void setGlobalLongTaskProgress(double fraction, String detail) {}

    /** {@link #setGlobalLongTaskProgress} の表示を消す。 */
    default void clearGlobalLongTaskProgress() {}
}
