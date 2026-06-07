package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.RdpFileSigner;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherDeployer;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherIni;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * 依頼書入力タブ内「リモートデスクトップ」子タブの UI。
 */
public final class RequestFormRemoteDesktopPane {

    private record SlotRowFields(TextField programField, TextField argsField) {}

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
                        "接続先 PC のタスクスケジューラが "
                                + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                                + " を起動し、共有フォルダの "
                                + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                + " に従ってプログラムを起動します。"
                                + " ローカルからは .rdp で mstsc を起動します。");
        subtitle.getStyleClass().add("paper-main-subtitle");
        subtitle.setWrapText(true);
        subtitle.setMaxWidth(CARD_WIDTH);

        Label iniPathLabel = new Label();
        iniPathLabel.getStyleClass().add("paper-main-subtitle");
        iniPathLabel.setWrapText(true);
        iniPathLabel.setMaxWidth(CARD_WIDTH);

        Label rapStatusLabel = new Label();
        rapStatusLabel.setWrapText(true);
        rapStatusLabel.setMaxWidth(CARD_WIDTH);
        rapStatusLabel.getStyleClass().add("top-status");

        Label launcherPathLabel = new Label();
        launcherPathLabel.getStyleClass().add("paper-main-subtitle");
        launcherPathLabel.setWrapText(true);
        HBox.setHgrow(launcherPathLabel, Priority.ALWAYS);

        Button btnCopyLauncherPath = new Button("パスをコピー");
        btnCopyLauncherPath.getStyleClass().add("btn-reload");
        btnCopyLauncherPath.setTooltip(
                new Tooltip("ランチャー exe の UNC フルパスをクリップボードへコピーします。"));
        btnCopyLauncherPath.setOnAction(
                e -> {
                    String path = AppPaths.resolveRdpLauncherExe(uiEnv.get()).toString();
                    copyToClipboard(path);
                    rapStatusLabel.setText("ランチャー exe パスをクリップボードへコピーしました: " + path);
                    status.accept("ランチャー exe パスをクリップボードへコピーしました。");
                });

        HBox launcherPathRow = new HBox(8, launcherPathLabel, btnCopyLauncherPath);
        launcherPathRow.setAlignment(Pos.CENTER_LEFT);
        launcherPathRow.setMaxWidth(CARD_WIDTH);

        Label deployStatusLabel = new Label();
        deployStatusLabel.setWrapText(true);
        deployStatusLabel.setMaxWidth(CARD_WIDTH);
        deployStatusLabel.getStyleClass().add("top-status");

        Spinner<Integer> slotSpinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, RdpRemoteLauncherIni.MAX_SLOTS, 1));
        slotSpinner.setEditable(true);
        slotSpinner.setPrefWidth(80);

        List<SlotRowFields> slotFields = new ArrayList<>();
        VBox slotBox = new VBox(8);
        slotBox.setFillWidth(true);

        Runnable refreshPaths =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    iniPathLabel.setText(
                            "RAP設定.ini: " + AppPaths.resolveRdpLauncherIni(ui).toString());
                    launcherPathLabel.setText(
                            "ランチャー exe: " + AppPaths.resolveRdpLauncherExe(ui).toString());
                };
        refreshPaths.run();

        java.util.function.IntConsumer appendSlotRow =
                slot -> {
                    Label caption = fieldCaption("スロット " + slot);
                    TextField programField = new TextField();
                    programField.setPromptText("接続先の exe パス（保存時に \"...\" で囲みます）");
                    HBox.setHgrow(programField, Priority.ALWAYS);

                    TextField argsField = new TextField();
                    argsField.setPromptText("引数（空白区切り。空白を含む引数は \"...\"）");
                    HBox.setHgrow(argsField, Priority.ALWAYS);

                    Button browse = new Button("参照...");
                    browse.getStyleClass().add("btn-reload");
                    browse.setTooltip(
                            new Tooltip("ローカル PC 上の exe を参照（接続先 UNC パスは手入力）。"));
                    browse.setOnAction(
                            e -> {
                                FileChooser chooser = new FileChooser();
                                chooser.setTitle("スロット " + slot + " のプログラム");
                                chooser.getExtensionFilters()
                                        .add(
                                                new FileChooser.ExtensionFilter(
                                                        "実行ファイル (*.exe)", "*.exe"));
                                java.io.File chosen =
                                        chooser.showOpenDialog(
                                                owner != null ? owner : browse.getScene().getWindow());
                                if (chosen != null) {
                                    programField.setText(chosen.getAbsolutePath());
                                }
                            });

                    HBox programRow = new HBox(8, programField, browse);
                    programRow.setAlignment(Pos.CENTER_LEFT);
                    programRow.setMaxWidth(CARD_WIDTH);

                    Label argsCaption = fieldCaption("  引数");
                    HBox argsRow = new HBox(8, argsField);
                    argsRow.setAlignment(Pos.CENTER_LEFT);
                    argsRow.setMaxWidth(CARD_WIDTH);

                    slotBox.getChildren().addAll(caption, programRow, argsCaption, argsRow);
                    slotFields.add(new SlotRowFields(programField, argsField));
                };

        for (int i = 1; i <= 3; i++) {
            appendSlotRow.accept(i);
        }

        Runnable loadIniFromShare =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    refreshPaths.run();
                    try {
                        RdpRemoteLauncherIni ini = RdpRemoteLauncherIni.load(iniPath);
                        slotSpinner.getValueFactory().setValue(ini.selectedSlot());
                        int visible = ini.visibleSlotCount();
                        while (slotFields.size() < visible) {
                            appendSlotRow.accept(slotFields.size() + 1);
                        }
                        for (int i = 1; i <= slotFields.size(); i++) {
                            RdpRemoteLauncherIni.Command command = ini.getSlotCommand(i);
                            SlotRowFields row = slotFields.get(i - 1);
                            row.programField().setText(command.executable());
                            row.argsField().setText(command.arguments());
                        }
                        rapStatusLabel.setText("読込しました: " + iniPath);
                    } catch (IOException ex) {
                        rapStatusLabel.setText("読込失敗: " + ex.getMessage());
                        showAlert(Alert.AlertType.ERROR, "読込失敗", ex.getMessage());
                    }
                };

        Runnable ensureLauncherDeployed =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    RdpRemoteLauncherDeployer.DeployOutcome outcome =
                            RdpRemoteLauncherDeployer.ensureDeployed(ui, status);
                    outcome.message().ifPresent(deployStatusLabel::setText);
                };

        Runnable saveIniToShare =
                () -> {
                    ensureLauncherDeployed.run();
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
                    ini.setSelectedSlot(slotSpinner.getValue());
                    for (int i = 1; i <= slotFields.size(); i++) {
                        SlotRowFields row = slotFields.get(i - 1);
                        ini.setSlotCommand(
                                i, row.programField().getText(), row.argsField().getText());
                    }
                    String validation = ini.validateMessageForSave();
                    if (validation != null) {
                        showAlert(Alert.AlertType.WARNING, "保存できません", validation);
                        rapStatusLabel.setText(validation);
                        return;
                    }
                    try {
                        ini.save(iniPath);
                        rapStatusLabel.setText("保存しました: " + iniPath);
                        status.accept("RAP設定.ini を保存しました: " + iniPath);
                    } catch (IOException ex) {
                        rapStatusLabel.setText("保存失敗: " + ex.getMessage());
                        showAlert(Alert.AlertType.ERROR, "保存失敗", ex.getMessage());
                    }
                };

        Button btnReloadIni = new Button("再読込");
        btnReloadIni.getStyleClass().add("btn-reload");
        btnReloadIni.setOnAction(
                e -> {
                    Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
                    confirm.setTitle("再読込");
                    confirm.setHeaderText(null);
                    confirm.setContentText("共有フォルダから RAP設定.ini を再読込します。未保存の変更は失われます。");
                    confirm.showAndWait().ifPresent(
                            choice -> {
                                if (choice == ButtonType.OK) {
                                    loadIniFromShare.run();
                                }
                            });
                });

        Button btnSaveIni = new Button("保存");
        btnSaveIni.getStyleClass().add("btn-reload");
        btnSaveIni.setOnAction(e -> saveIniToShare.run());

        Button btnAddSlot = new Button("+ 行を追加");
        btnAddSlot.getStyleClass().add("btn-reload");
        btnAddSlot.setOnAction(
                e -> {
                    if (slotFields.size() >= RdpRemoteLauncherIni.MAX_SLOTS) {
                        return;
                    }
                    appendSlotRow.accept(slotFields.size() + 1);
                    btnAddSlot.setDisable(slotFields.size() >= RdpRemoteLauncherIni.MAX_SLOTS);
                });

        Button btnOpenFolder = new Button("フォルダを開く");
        btnOpenFolder.getStyleClass().add("btn-reload");
        btnOpenFolder.setOnAction(
                e -> {
                    Map<String, String> ui = uiEnv.get();
                    Path dir = AppPaths.resolveRdpLauncherDeployDir(ui);
                    if (!Files.isDirectory(dir)) {
                        DirectoryChooser chooser = new DirectoryChooser();
                        chooser.setTitle("RAP設定.ini のフォルダ");
                        java.io.File chosen =
                                chooser.showDialog(
                                        owner != null ? owner : btnOpenFolder.getScene().getWindow());
                        if (chosen != null) {
                            openFolder(chosen.toPath());
                        }
                        return;
                    }
                    openFolder(dir);
                });

        HBox slotSpinnerRow = new HBox(8, new Label("起動プログラム番号:"), slotSpinner);
        slotSpinnerRow.setAlignment(Pos.CENTER_LEFT);

        HBox rapActionRow = new HBox(8, btnReloadIni, btnSaveIni, btnOpenFolder, btnAddSlot);
        rapActionRow.setAlignment(Pos.CENTER_LEFT);

        VBox rapContent =
                new VBox(
                        8,
                        iniPathLabel,
                        launcherPathRow,
                        deployStatusLabel,
                        slotSpinnerRow,
                        slotBox,
                        rapActionRow,
                        rapStatusLabel);
        rapContent.setFillWidth(true);

        TitledPane rapPane = new TitledPane("RAP設定（接続先起動プログラム）", rapContent);
        rapPane.setExpanded(true);
        rapPane.setMaxWidth(CARD_WIDTH);

        // --- RDP profile section ---
        TextField profileField = new TextField();
        profileField.setEditable(false);
        profileField.setPromptText("（.rdp プロファイル未設定）");
        HBox.setHgrow(profileField, Priority.ALWAYS);

        TextField companionProgramField = new TextField();
        companionProgramField.setPromptText("例: C:\\Windows\\System32\\notepad.exe");
        HBox.setHgrow(companionProgramField, Priority.ALWAYS);

        TextField companionArgsField = new TextField();
        companionArgsField.setPromptText("空白区切り");
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
                    refreshPaths.run();
                };

        Runnable persistCompanionFields =
                () -> {
                    if (companionProgramChangeHandler != null) {
                        companionProgramChangeHandler.accept(companionProgramField.getText().trim());
                    }
                    if (companionProgramArgsChangeHandler != null) {
                        companionProgramArgsChangeHandler.accept(companionArgsField.getText().trim());
                    }
                };

        companionProgramField
                .focusedProperty()
                .addListener((obs, wasFocused, focused) -> {
                    if (wasFocused && !focused) {
                        persistCompanionFields.run();
                    }
                });
        companionArgsField
                .focusedProperty()
                .addListener((obs, wasFocused, focused) -> {
                    if (wasFocused && !focused) {
                        persistCompanionFields.run();
                    }
                });

        Button btnLaunch = new Button("リモートデスクトップを起動");
        btnLaunch.getStyleClass().add("btn-reload");
        btnLaunch.setTooltip(
                new Tooltip(
                        "指定 .rdp で mstsc.exe を起動します。"
                                + " 接続先の起動プログラムは RAP設定.ini + タスクスケジューラが参照します。"));
        btnLaunch.setOnAction(
                e -> {
                    persistCompanionFields.run();
                    ensureLauncherDeployed.run();
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

        Button btnChoose = new Button("プロファイルを選択...");
        btnChoose.getStyleClass().add("btn-reload");
        btnChoose.setOnAction(
                e -> {
                    FileChooser chooser = new FileChooser();
                    chooser.setTitle("RDP プロファイル (.rdp) を選択");
                    chooser.getExtensionFilters()
                            .add(
                                    new FileChooser.ExtensionFilter(
                                            "リモートデスクトップ接続 (*.rdp)", "*.rdp"));
                    java.io.File chosen =
                            chooser.showOpenDialog(owner != null ? owner : btnChoose.getScene().getWindow());
                    if (chosen == null) {
                        return;
                    }
                    profileField.setText(chosen.getAbsolutePath());
                    if (profileChangeHandler != null) {
                        profileChangeHandler.accept(chosen.getAbsolutePath());
                    }
                });

        Button btnSignWizard = new Button("RDP 署名ウィザード...");
        btnSignWizard.getStyleClass().add("btn-reload");
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

        HBox profileRow = new HBox(8, profileField, btnChoose);
        profileRow.setAlignment(Pos.CENTER_LEFT);
        profileRow.setMaxWidth(CARD_WIDTH);

        HBox actionRow = new HBox(8, btnLaunch, btnSignWizard);
        actionRow.setAlignment(Pos.CENTER_LEFT);

        VBox rdpContent =
                new VBox(
                        8,
                        fieldCaption("RDP プロファイル"),
                        profileRow,
                        actionRow,
                        statusLabel);
        TitledPane rdpPane = new TitledPane("リモートデスクトップ接続", rdpContent);
        rdpPane.setExpanded(true);
        rdpPane.setMaxWidth(CARD_WIDTH);

        VBox policyContent =
                new VBox(
                        8,
                        fieldCaption("リモート起動プログラム（.rdp 組込・PM_AI_RDP_EMBED_STARTUP_IN_PROFILE=1 時）"),
                        companionProgramField,
                        fieldCaption("引数"),
                        companionArgsField);
        TitledPane policyPane = new TitledPane("ポリシー有効化後用（.rdp 組込）", policyContent);
        policyPane.setExpanded(false);
        policyPane.setMaxWidth(CARD_WIDTH);

        VBox card = new VBox(10, title, subtitle, rapPane, rdpPane, policyPane);
        card.getStyleClass().add("settings-card");
        card.setPadding(new Insets(12));
        card.setMaxWidth(CARD_WIDTH);
        card.setPrefWidth(CARD_WIDTH);
        card.setFillWidth(true);

        Runnable onTabShow =
                () -> {
                    refreshFromEnv.run();
                    ensureLauncherDeployed.run();
                    loadIniFromShare.run();
                };
        onTabShow.run();

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

    private static void openFolder(Path dir) {
        if (!RemoteDesktopLauncher.isSupportedPlatform()) {
            return;
        }
        try {
            new ProcessBuilder("explorer.exe", dir.toString()).start();
        } catch (IOException ignored) {
            // ignore
        }
    }

    private static void copyToClipboard(String text) {
        if (text == null || text.isBlank()) {
            return;
        }
        ClipboardContent content = new ClipboardContent();
        content.putString(text);
        Clipboard.getSystemClipboard().setContent(content);
    }

    private static void showAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
