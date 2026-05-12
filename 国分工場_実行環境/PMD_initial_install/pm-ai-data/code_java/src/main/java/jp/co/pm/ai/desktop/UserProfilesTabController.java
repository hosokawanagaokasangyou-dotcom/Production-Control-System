package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;

import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppVersionInfo;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.UserProfileStore;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * ユーザープロファイル（UI 設定の保存・読み出し）。保存先は {@code ~/.pm-ai-desktop/user-profiles}
 * （アップデートで上書きされない）。
 */
public final class UserProfilesTabController {

    @FXML
    private ListView<UserProfileStore.ListedProfile> profileListView;

    @FXML
    private TextField profileNameField;

    @FXML
    private Button saveButton;

    @FXML
    private Button loadButton;

    @FXML
    private Button deleteButton;

    @FXML
    private Label hintLabel;

    private MainShellController shell;

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    @FXML
    private void initialize() {
        if (hintLabel != null) {
            hintLabel.setWrapText(true);
            hintLabel.setText(
                    "現在の UI 状態（セッション・列順・環境タブの行・テーマ・ガント等）を保存します。保存上限は "
                            + UserProfileStore.MAX_PROFILES
                            + " 件です。名前を空にすると表示名は保存日時（秒まで）になります。"
                            + " プロファイルはユーザーホームのみに保存され、アプリのアップデートでは置き換わりません。"
                            + " JSON ファイルへのエクスポート／インポートで別 PC やバックアップとの共有もできます。");
        }
        if (profileListView != null) {
            profileListView.setCellFactory(
                    lv ->
                            new ListCell<>() {
                                @Override
                                protected void updateItem(UserProfileStore.ListedProfile item, boolean empty) {
                                    super.updateItem(item, empty);
                                    if (empty || item == null) {
                                        setText(null);
                                        return;
                                    }
                                    String extra =
                                            item.savedAtIso().isEmpty() ? "" : " ・ " + item.savedAtIso();
                                    setText(item.displayLabel() + extra);
                                }
                            });
        }
        refreshList();
    }

    @FXML
    private void onRefreshAction() {
        refreshList();
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        try {
            shell.persistDesktopSessionNow();
            ObjectNode sessionJson =
                    DesktopSessionStateStore.toJsonObject(shell.snapshotDesktopSessionForExport());
            var tables = TableColumnOrderPersistence.readCurrentStoreRoot();
            String name = profileNameField != null ? profileNameField.getText() : "";
            UserProfileStore.saveProfile(name, sessionJson, tables);
            refreshList();
            info("保存しました。", "");
        } catch (IllegalStateException ex) {
            warn("上限", ex.getMessage());
        } catch (IOException ex) {
            warn("保存エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onLoadAction() {
        if (shell == null) {
            return;
        }
        UserProfileStore.ListedProfile sel =
                profileListView != null ? profileListView.getSelectionModel().getSelectedItem() : null;
        if (sel == null) {
            warn("選択なし", "読み出すプロファイルを一覧から選んでください。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        confirm.setTitle("確認");
        confirm.setHeaderText(null);
        confirm.setContentText("現在の UI をこのプロファイルで置き換えます。よろしいですか？");
        if (confirm.showAndWait().isEmpty() || confirm.getResult() != ButtonType.OK) {
            return;
        }
        try {
            UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfile(sel.path());
            shell.applyUserProfileSnapshot(loaded.session(), loaded.tableColumnOrder());
            info(
                    "読み出しました。",
                    loaded.displayName().isEmpty()
                            ? loaded.savedAt()
                            : loaded.displayName());
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onDeleteAction() {
        UserProfileStore.ListedProfile sel =
                profileListView != null ? profileListView.getSelectionModel().getSelectedItem() : null;
        if (sel == null) {
            warn("選択なし", "削除するプロファイルを一覧から選んでください。");
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        confirm.setTitle("確認");
        confirm.setHeaderText(null);
        confirm.setContentText("選択したプロファイルを削除しますか？");
        if (confirm.showAndWait().isEmpty() || confirm.getResult() != ButtonType.OK) {
            return;
        }
        try {
            UserProfileStore.deleteProfile(sel.path());
            refreshList();
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onExportSelectedToFileAction() {
        UserProfileStore.ListedProfile sel =
                profileListView != null ? profileListView.getSelectionModel().getSelectedItem() : null;
        if (sel == null) {
            warn("選択なし", "エクスポートするプロファイルを一覧から選んでください。");
            return;
        }
        FileChooser fc = profileFileChooser("プロファイルをエクスポート");
        fc.setInitialFileName(sanitizeExportFileName(sel.displayLabel()) + ".json");
        java.io.File dest = fc.showSaveDialog(shell != null ? shell.primaryStageForDialogs() : null);
        if (dest == null) {
            return;
        }
        try {
            UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfile(sel.path());
            ObjectNode sessionJson = DesktopSessionStateStore.toJsonObject(loaded.session());
            UserProfileStore.exportProfileToFile(
                    dest.toPath(),
                    loaded.displayName(),
                    loaded.savedAt(),
                    sessionJson,
                    loaded.tableColumnOrder(),
                    resolveSourceAppVersion());
            info("エクスポートしました。", dest.getAbsolutePath());
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onExportCurrentToFileAction() {
        if (shell == null) {
            return;
        }
        FileChooser fc = profileFileChooser("現在の設定をエクスポート");
        fc.setInitialFileName("pm-ai-user-profile.json");
        java.io.File dest = fc.showSaveDialog(shell.primaryStageForDialogs());
        if (dest == null) {
            return;
        }
        try {
            shell.persistDesktopSessionNow();
            ObjectNode sessionJson =
                    DesktopSessionStateStore.toJsonObject(shell.snapshotDesktopSessionForExport());
            var tables = TableColumnOrderPersistence.readCurrentStoreRoot();
            String name = profileNameField != null ? profileNameField.getText() : "";
            String savedAt = LocalDateTime.now().toString();
            UserProfileStore.exportProfileToFile(
                    dest.toPath(), name != null ? name.strip() : "", savedAt, sessionJson, tables, resolveSourceAppVersion());
            info("エクスポートしました。", dest.getAbsolutePath());
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onImportFromFileApplyAction() {
        if (shell == null) {
            return;
        }
        FileChooser fc = profileFileChooser("プロファイル JSON を選択");
        java.io.File src = fc.showOpenDialog(shell.primaryStageForDialogs());
        if (src == null) {
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        if (shell.primaryStageForDialogs() != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        confirm.setTitle("確認");
        confirm.setHeaderText(null);
        confirm.setContentText("現在の UI をこのファイルの内容で置き換えます。よろしいですか？");
        if (confirm.showAndWait().isEmpty() || confirm.getResult() != ButtonType.OK) {
            return;
        }
        try {
            UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfile(src.toPath());
            shell.applyUserProfileSnapshot(loaded.session(), loaded.tableColumnOrder());
            info(
                    "ファイルから適用しました。",
                    loaded.displayName().isEmpty()
                            ? loaded.savedAt()
                            : loaded.displayName());
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    @FXML
    private void onImportFromFileAddAction() {
        FileChooser fc = profileFileChooser("一覧に追加するプロファイル JSON を選択");
        java.io.File src =
                fc.showOpenDialog(shell != null ? shell.primaryStageForDialogs() : null);
        if (src == null) {
            return;
        }
        try {
            UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfile(src.toPath());
            ObjectNode sessionJson = DesktopSessionStateStore.toJsonObject(loaded.session());
            UserProfileStore.saveProfile(loaded.displayName(), sessionJson, loaded.tableColumnOrder());
            refreshList();
            info(
                    "一覧に追加しました。",
                    loaded.displayName().isEmpty()
                            ? loaded.savedAt()
                            : loaded.displayName());
        } catch (IllegalStateException ex) {
            warn("上限", ex.getMessage());
        } catch (IOException ex) {
            warn("エラー", ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    private FileChooser profileFileChooser(String title) {
        FileChooser fc = new FileChooser();
        fc.setTitle(title);
        fc.getExtensionFilters().add(new ExtensionFilter("JSON (*.json)", "*.json"));
        Path home = Paths.get(System.getProperty("user.home", "."));
        if (java.nio.file.Files.isDirectory(home)) {
            fc.setInitialDirectory(home.toFile());
        }
        return fc;
    }

    private String resolveSourceAppVersion() {
        if (shell == null) {
            return AppVersionInfo.VERSION_UNKNOWN;
        }
        return AppVersionInfo.resolveDisplayedVersion(
                Path.of(System.getProperty("user.dir", ".")), shell.snapshotUiEnv());
    }

    private static String sanitizeExportFileName(String label) {
        if (label == null || label.isBlank()) {
            return "pm-ai-user-profile";
        }
        String t = label.strip().replaceAll("[\\\\/:*?\"<>|]", "_");
        if (t.length() > 120) {
            return t.substring(0, 120);
        }
        return t;
    }

    private void refreshList() {
        if (profileListView == null) {
            return;
        }
        try {
            ObservableList<UserProfileStore.ListedProfile> items =
                    FXCollections.observableArrayList(UserProfileStore.listProfiles());
            profileListView.setItems(items);
        } catch (IOException ex) {
            profileListView.setItems(FXCollections.observableArrayList());
        }
    }

    private void warn(String title, String msg) {
        Alert a = new Alert(AlertType.WARNING);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            a.initOwner(shell.primaryStageForDialogs());
        }
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(msg);
        a.showAndWait();
    }

    private void info(String title, String msg) {
        Alert a = new Alert(AlertType.INFORMATION);
        if (shell != null && shell.primaryStageForDialogs() != null) {
            a.initOwner(shell.primaryStageForDialogs());
        }
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(msg);
        a.showAndWait();
    }
}
