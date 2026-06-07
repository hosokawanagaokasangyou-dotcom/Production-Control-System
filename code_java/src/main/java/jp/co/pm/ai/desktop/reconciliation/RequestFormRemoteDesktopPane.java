package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
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
import jp.co.pm.ai.desktop.io.RdpFileSigner;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * 依頼書入力タブ内「リモートデスクトップ」子タブの UI。
 */
public final class RequestFormRemoteDesktopPane {

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Consumer<String> profileChangeHandler,
            Consumer<String> companionProgramChangeHandler,
            Consumer<String> companionProgramArgsChangeHandler,
            Consumer<String> statusConsumer) {}

    private static final double CARD_WIDTH = 720;

    private RequestFormRemoteDesktopPane() {}

    public static VBox buildTabContent(Window owner, Context ctx) {
        Supplier<Map<String, String>> uiEnv = ctx.uiEnv() != null ? ctx.uiEnv() : () -> Map.of();
        Consumer<String> profileChangeHandler = ctx.profileChangeHandler();
        Consumer<String> companionProgramChangeHandler = ctx.companionProgramChangeHandler();
        Consumer<String> companionProgramArgsChangeHandler = ctx.companionProgramArgsChangeHandler();
        Consumer<String> status = ctx.statusConsumer() != null ? ctx.statusConsumer() : s -> {};

        Label title = new Label("リモートデスクトップ");
        title.getStyleClass().add("paper-main-title");

        Label subtitle =
                new Label(
                        "環境変数 "
                                + AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE
                                + " の .rdp で接続します。"
                                + " 「リモート起動プログラム」は接続先サーバー上の exe パスで、"
                                + " 起動時に .rdp へ RemoteApp 設定として書き込みます（環境変数タブと同期）。");
        subtitle.getStyleClass().add("paper-main-subtitle");
        subtitle.setWrapText(true);
        subtitle.setMaxWidth(CARD_WIDTH);

        TextField profileField = new TextField();
        profileField.setEditable(false);
        profileField.setPromptText("（.rdp プロファイル未設定）");
        HBox.setHgrow(profileField, Priority.ALWAYS);

        Label profileCaption = fieldCaption("RDP プロファイル");

        Label companionCaption = fieldCaption("リモート起動プログラム（接続先のパス）");
        TextField companionProgramField = new TextField();
        companionProgramField.setPromptText("例: C:\\Windows\\System32\\notepad.exe");
        HBox.setHgrow(companionProgramField, Priority.ALWAYS);

        Label argsCaption = fieldCaption("引数");
        TextField companionArgsField = new TextField();
        companionArgsField.setPromptText("空白区切り。空白を含む引数は \"...\" で囲む");
        HBox.setHgrow(companionArgsField, Priority.ALWAYS);

        Label statusLabel = new Label();
        statusLabel.setWrapText(true);
        statusLabel.setMaxWidth(CARD_WIDTH);
        statusLabel.getStyleClass().add("top-status");

        Runnable refreshFromEnv =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Optional<Path> configured = AppPaths.resolveRequestFormRdpProfile(ui);
                    if (configured.isEmpty()) {
                        profileField.setText("");
                    } else {
                        Path preferred =
                                RdpFileSigner.resolvePreferredSignedProfilePath(configured.get(), ui);
                        profileField.setText(preferred.toString());
                        if (!preferred.equals(configured.get()) && profileChangeHandler != null) {
                            profileChangeHandler.accept(preferred.toString());
                        }
                    }
                    companionProgramField.setText(
                            ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM, ""));
                    companionArgsField.setText(
                            ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS, ""));
                };
        refreshFromEnv.run();

        Runnable persistCompanionFields =
                () -> {
                    if (companionProgramChangeHandler != null) {
                        companionProgramChangeHandler.accept(companionProgramField.getText().trim());
                    }
                    if (companionProgramArgsChangeHandler != null) {
                        companionProgramArgsChangeHandler.accept(companionArgsField.getText().trim());
                    }
                };
        companionProgramField.focusedProperty()
                .addListener((obs, wasFocused, focused) -> {
                    if (wasFocused && !focused) {
                        persistCompanionFields.run();
                    }
                });
        companionArgsField.focusedProperty()
                .addListener((obs, wasFocused, focused) -> {
                    if (wasFocused && !focused) {
                        persistCompanionFields.run();
                    }
                });

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

        Button btnChooseProgram = new Button("プログラムを選択...");
        btnChooseProgram.getStyleClass().add("btn-reload");
        btnChooseProgram.setTooltip(
                new Tooltip(
                        "接続先サーバー上の exe パスを入力します。"
                                + " ローカル PC と同じパスなら参照用にファイル選択できます。"));
        btnChooseProgram.setOnAction(
                e -> {
                    FileChooser chooser = new FileChooser();
                    chooser.setTitle("同時起動プログラムを選択");
                    chooser.getExtensionFilters()
                            .addAll(
                                    new FileChooser.ExtensionFilter("実行ファイル (*.exe)", "*.exe"),
                                    new FileChooser.ExtensionFilter("すべてのファイル", "*.*"));
                    String current = companionProgramField.getText();
                    if (!current.isBlank()) {
                        Path p = Path.of(current);
                        Path parent = p.getParent();
                        if (parent != null && Files.isDirectory(parent)) {
                            chooser.setInitialDirectory(parent.toFile());
                        }
                    }
                    java.io.File chosen =
                            chooser.showOpenDialog(
                                    owner != null ? owner : btnChooseProgram.getScene().getWindow());
                    if (chosen == null) {
                        return;
                    }
                    String abs = chosen.getAbsolutePath();
                    companionProgramField.setText(abs);
                    persistCompanionFields.run();
                    status.accept("リモート起動プログラムを設定しました: " + abs);
                    statusLabel.setText("リモート起動: " + abs);
                });

        Button btnClearProgram = new Button("クリア");
        btnClearProgram.getStyleClass().add("btn-reload");
        btnClearProgram.setTooltip(new Tooltip("リモート起動プログラム指定を解除します（.rdp からも削除）。"));
        btnClearProgram.setOnAction(
                e -> {
                    companionProgramField.clear();
                    persistCompanionFields.run();
                    status.accept("リモート起動プログラムをクリアしました。");
                    statusLabel.setText("リモート起動プログラム未設定です。");
                });

        Button btnSignWizard = new Button("RDP 署名ウィザード...");
        btnSignWizard.getStyleClass().add("btn-reload");
        btnSignWizard.setTooltip(
                new Tooltip(
                        "Windows の rdpsign.exe で .rdp を署名し、"
                                + " GPO 信頼設定用の SHA-1 サムプリントを案内します（Windows のみ）。"));
        btnSignWizard.setOnAction(
                e ->
                        RdpProfileSignWizard.show(
                                owner != null ? owner : btnSignWizard.getScene().getWindow(),
                                RdpProfileSignWizard.initialProfileFromUi(uiEnv.get()),
                                uiEnv.get(),
                                msg -> {
                                    status.accept(msg);
                                    statusLabel.setText(msg);
                                },
                                path -> {
                                    profileField.setText(path);
                                    if (profileChangeHandler != null) {
                                        profileChangeHandler.accept(path);
                                    }
                                }));

        Button btnLaunch = new Button("リモートデスクトップを起動");
        btnLaunch.getStyleClass().add("btn-reload");
        btnLaunch.setTooltip(
                new Tooltip(
                        "指定 .rdp で mstsc.exe を起動します。"
                                + " リモート起動プログラムが設定されていれば .rdp へ書き込んでから接続します。"));
        btnLaunch.setOnAction(
                e -> {
                    persistCompanionFields.run();
                    Map<String, String> ui = new HashMap<>(uiEnv.get());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM,
                            companionProgramField.getText().trim());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS,
                            companionArgsField.getText().trim());
                    Optional<Path> profile = AppPaths.resolveRequestFormRdpProfile(ui);
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
                        Path configured = profile.get();
                        Path preferred =
                                RdpFileSigner.resolvePreferredSignedProfilePath(configured, ui);
                        RemoteDesktopLauncher.LaunchOutcome outcome =
                                RemoteDesktopLauncher.launch(preferred, ui);
                        if (!preferred.equals(configured) && profileChangeHandler != null) {
                            profileChangeHandler.accept(preferred.toString());
                            profileField.setText(preferred.toString());
                        }
                        String msg = "リモートデスクトップを起動しました: " + outcome.rdpProfile();
                        if (outcome.remoteStartupSummary().isPresent()) {
                            msg += "\n.rdp に組込: " + outcome.remoteStartupSummary().get();
                        }
                        if (outcome.signatureRemoved()) {
                            msg += "\n（.rdp 編集により署名が無効化されました。必要なら署名ウィザードを実行してください）";
                        }
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

        HBox companionRow =
                new HBox(8, companionProgramField, btnChooseProgram, btnClearProgram);
        companionRow.setAlignment(Pos.CENTER_LEFT);
        companionRow.setMaxWidth(CARD_WIDTH);

        HBox argsRow = new HBox(8, companionArgsField);
        argsRow.setAlignment(Pos.CENTER_LEFT);
        argsRow.setMaxWidth(CARD_WIDTH);

        HBox actionRow = new HBox(8, btnLaunch, btnSignWizard);
        actionRow.setAlignment(Pos.CENTER_LEFT);

        VBox card =
                new VBox(
                        10,
                        title,
                        subtitle,
                        profileCaption,
                        profileRow,
                        companionCaption,
                        companionRow,
                        argsCaption,
                        argsRow,
                        actionRow,
                        statusLabel);
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

    private static Label fieldCaption(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("paper-main-subtitle");
        return label;
    }

    private static void showAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
