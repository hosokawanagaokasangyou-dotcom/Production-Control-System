package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * 依頼書入力タブ内「リモートデスクトップ」子タブの UI。
 */
public final class RequestFormRemoteDesktopPane {

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Consumer<String> profileChangeHandler,
            Consumer<String> statusConsumer) {}

    private static final double CARD_WIDTH = 720;

    private RequestFormRemoteDesktopPane() {}

    public static VBox buildTabContent(Window owner, Context ctx) {
        Supplier<Map<String, String>> uiEnv = ctx.uiEnv() != null ? ctx.uiEnv() : () -> Map.of();
        Consumer<String> profileChangeHandler = ctx.profileChangeHandler();
        Consumer<String> status = ctx.statusConsumer() != null ? ctx.statusConsumer() : s -> {};

        Label title = new Label("リモートデスクトップ");
        title.getStyleClass().add("paper-main-title");

        Label subtitle =
                new Label(
                        "環境変数 "
                                + AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE
                                + " に指定した .rdp プロファイルで Windows リモートデスクトップ接続を起動します。"
                                + " プロファイルは環境変数タブと同期されます。");
        subtitle.getStyleClass().add("paper-main-subtitle");
        subtitle.setWrapText(true);
        subtitle.setMaxWidth(CARD_WIDTH);

        TextField profileField = new TextField();
        profileField.setEditable(false);
        profileField.setPromptText("（.rdp プロファイル未設定）");
        HBox.setHgrow(profileField, Priority.ALWAYS);

        Label statusLabel = new Label();
        statusLabel.setWrapText(true);
        statusLabel.setMaxWidth(CARD_WIDTH);
        statusLabel.getStyleClass().add("top-status");

        Runnable refreshProfileField =
                () -> {
                    String path =
                            AppPaths.resolveRequestFormRdpProfile(uiEnv.get())
                                    .map(Path::toString)
                                    .orElse("");
                    profileField.setText(path);
                };
        refreshProfileField.run();

        Button btnChoose = new Button("プロファイルを選択...");
        btnChoose.getStyleClass().add("btn-reload");
        btnChoose.setTooltip(
                new Tooltip("接続に使う .rdp ファイルを選択し、環境変数タブへ反映します。"));
        btnChoose.setOnAction(
                e -> {
                    FileChooser chooser = new FileChooser();
                    chooser.setTitle("RDP プロファイル (.rdp) を選択");
                    chooser.getExtensionFilters()
                            .add(
                                    new FileChooser.ExtensionFilter(
                                            "リモートデスクトップ接続 (*.rdp)", "*.rdp"));
                    String current = profileField.getText();
                    if (!current.isBlank()) {
                        Path p = Path.of(current);
                        Path parent = p.getParent();
                        if (parent != null && Files.isDirectory(parent)) {
                            chooser.setInitialDirectory(parent.toFile());
                        }
                    }
                    java.io.File chosen =
                            chooser.showOpenDialog(owner != null ? owner : btnChoose.getScene().getWindow());
                    if (chosen == null) {
                        return;
                    }
                    String abs = chosen.getAbsolutePath();
                    profileField.setText(abs);
                    if (profileChangeHandler != null) {
                        profileChangeHandler.accept(abs);
                    }
                    status.accept("RDP プロファイルを設定しました: " + abs);
                    statusLabel.setText("プロファイル: " + abs);
                });

        Button btnClear = new Button("クリア");
        btnClear.getStyleClass().add("btn-reload");
        btnClear.setTooltip(new Tooltip("プロファイル指定を解除します。"));
        btnClear.setOnAction(
                e -> {
                    profileField.clear();
                    if (profileChangeHandler != null) {
                        profileChangeHandler.accept("");
                    }
                    status.accept("RDP プロファイルをクリアしました。");
                    statusLabel.setText("プロファイル未設定です。");
                });

        Button btnLaunch = new Button("リモートデスクトップを起動");
        btnLaunch.getStyleClass().add("btn-reload");
        btnLaunch.setTooltip(
                new Tooltip("指定した .rdp プロファイルで mstsc.exe を起動します（Windows のみ）。"));
        btnLaunch.setOnAction(
                e -> {
                    Optional<Path> profile = AppPaths.resolveRequestFormRdpProfile(uiEnv.get());
                    if (profile.isEmpty()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "プロファイル未設定",
                                AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE
                                        + " に .rdp ファイルを指定してください。");
                        return;
                    }
                    if (!RemoteDesktopLauncher.isSupportedPlatform()) {
                        showAlert(
                                Alert.AlertType.ERROR,
                                "未対応",
                                "リモートデスクトップの起動は Windows 上のデスクトップアプリでのみ利用できます。");
                        return;
                    }
                    try {
                        RemoteDesktopLauncher.launch(profile.get());
                        String msg = "リモートデスクトップを起動しました: " + profile.get();
                        status.accept(msg);
                        statusLabel.setText(msg);
                    } catch (IOException ex) {
                        showAlert(Alert.AlertType.ERROR, "起動失敗", ex.getMessage());
                        status.accept("リモートデスクトップ起動失敗: " + ex.getMessage());
                        statusLabel.setText(ex.getMessage());
                    }
                });

        HBox profileRow = new HBox(8, profileField, btnChoose, btnClear);
        profileRow.setAlignment(Pos.CENTER_LEFT);
        profileRow.setMaxWidth(CARD_WIDTH);

        HBox actionRow = new HBox(8, btnLaunch);
        actionRow.setAlignment(Pos.CENTER_LEFT);

        VBox card = new VBox(10, title, subtitle, profileRow, actionRow, statusLabel);
        card.getStyleClass().add("settings-card");
        card.setPadding(new Insets(12));
        card.setMaxWidth(CARD_WIDTH);
        card.setPrefWidth(CARD_WIDTH);
        card.setFillWidth(true);

        VBox root = new VBox(12, card);
        root.getStyleClass().add("form-tab-container");
        root.setPadding(new Insets(12));
        root.setFillWidth(true);
        root.setMaxWidth(Double.MAX_VALUE);
        VBox.setVgrow(root, Priority.ALWAYS);
        return root;
    }

    private static void showAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
