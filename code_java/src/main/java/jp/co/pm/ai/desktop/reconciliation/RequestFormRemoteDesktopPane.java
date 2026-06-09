package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.MalformedInputException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.IntConsumer;
import java.util.function.Supplier;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ListCell;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.SplitPane;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.util.StringConverter;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AladdinRpaLaunchArgs;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.io.RdpFileSigner;
import jp.co.pm.ai.desktop.io.RdpLaunchProfile;
import jp.co.pm.ai.desktop.io.RdpLaunchProfileCatalog;
import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings;
import jp.co.pm.ai.desktop.io.RdpMstscSessionMonitor;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherDeployer;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherIni;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * メインシェル「リモートデスクトップ」タブの UI（RAP 設定・RDP 起動）。
 */
public final class RequestFormRemoteDesktopPane {

    private record ProfileRowFields(
            int number,
            TextField nameField,
            TextField descriptionField,
            TextField categoryField,
            TextField programField,
            TextField argsField,
            CheckBox chkRpaEternal) {}

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Consumer<String> profileChangeHandler,
            Consumer<String> companionProgramChangeHandler,
            Consumer<String> companionProgramArgsChangeHandler,
            Consumer<Map<String, String>> displayEnvChangeHandler,
            IntConsumer launchProfileNumberChangeHandler,
            Consumer<String> statusConsumer) {}

    private static final double CARD_WIDTH = 720;

    private static final String LAUNCHER_LOG_DEFAULT_FONT_FAMILY_LABEL = "システム既定";

    private static final List<Double> LAUNCHER_LOG_PRESET_FONT_SIZES =
            List.of(9d, 10d, 11d, 12d, 13d, 14d, 15d, 16d, 18d, 20d, 22d, 24d);

    /** 構築結果。{@link #scheduleInitialRefresh()} はタブ初回表示時のみ呼ぶ（UNC I/O を初回マウントから分離）。 */
    public record TabContent(SplitPane root, Runnable scheduleInitialRefresh) {}

    private RequestFormRemoteDesktopPane() {}

    public static TabContent buildTabContent(Window owner, Context ctx) {
        Supplier<Map<String, String>> uiEnv = ctx.uiEnv() != null ? ctx.uiEnv() : () -> Map.of();
        Consumer<String> profileChangeHandler = ctx.profileChangeHandler();
        Consumer<String> companionProgramChangeHandler = ctx.companionProgramChangeHandler();
        Consumer<String> companionProgramArgsChangeHandler = ctx.companionProgramArgsChangeHandler();
        Consumer<Map<String, String>> displayEnvChangeHandler = ctx.displayEnvChangeHandler();
        IntConsumer launchProfileNumberChangeHandler = ctx.launchProfileNumberChangeHandler();
        Consumer<String> status = ctx.statusConsumer() != null ? ctx.statusConsumer() : s -> {};
        AtomicReference<Thread> rdpSessionWatchThread = new AtomicReference<>();
        AtomicBoolean launcherDeployInProgress = new AtomicBoolean(false);
        AtomicBoolean launcherDeployReady = new AtomicBoolean(false);
        AtomicReference<Runnable> pendingLaunchAfterDeploy = new AtomicReference<>();
        AtomicBoolean rdpSessionActive = new AtomicBoolean(false);

        Label title = new Label("リモートデスクトップ");
        title.getStyleClass().add("pm-rdp-page-title");

        Label subtitle =
                new Label(
                        "接続先 PC のタスクスケジューラが "
                                + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                                + " を起動し、共有フォルダの "
                                + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                + " に従ってプログラムを起動します。"
                                + " ローカルからは .rdp で mstsc を起動します。");
        subtitle.getStyleClass().add("pm-rdp-page-subtitle");
        subtitle.setWrapText(true);
        subtitle.setMaxWidth(CARD_WIDTH);

        VBox headerBlock = new VBox(6, title, subtitle);
        headerBlock.getStyleClass().add("pm-rdp-header");
        headerBlock.setMaxWidth(CARD_WIDTH);

        Label iniPathLabel = new Label();
        iniPathLabel.getStyleClass().add("pm-rdp-meta-label");
        iniPathLabel.setWrapText(true);
        iniPathLabel.setMaxWidth(CARD_WIDTH);

        Label rapStatusLabel = new Label();
        rapStatusLabel.setWrapText(true);
        rapStatusLabel.setMaxWidth(CARD_WIDTH);
        rapStatusLabel.getStyleClass().add("pm-rdp-status-label");

        Label launcherPathLabel = new Label();
        launcherPathLabel.getStyleClass().add("pm-rdp-meta-label");
        launcherPathLabel.setWrapText(true);
        launcherPathLabel.setMaxWidth(CARD_WIDTH);

        Button btnCopyLauncherPath = new Button("パスをコピー");
        styleSecondaryButton(btnCopyLauncherPath);
        btnCopyLauncherPath.setTooltip(
                new Tooltip("ランチャー exe の UNC フルパスをクリップボードへコピーします。"));
        btnCopyLauncherPath.setOnAction(
                e -> {
                    String path = AppPaths.resolveRdpLauncherExe(uiEnv.get()).toString();
                    copyToClipboard(path);
                    rapStatusLabel.setText("ランチャー exe パスをクリップボードへコピーしました: " + path);
                    status.accept("ランチャー exe パスをクリップボードへコピーしました。");
                });

        Button btnForceDeployLauncher = new Button("ランチャーを強制転送");
        styleSecondaryButton(btnForceDeployLauncher);
        btnForceDeployLauncher.setTooltip(
                new Tooltip(
                        "版比較に関係なく、同梱 "
                                + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                                + " をサマリ Excel 同階層へ上書きコピーします。"));

        Label launcherLogPathLabel = new Label();
        launcherLogPathLabel.getStyleClass().add("pm-rdp-meta-label");
        launcherLogPathLabel.setWrapText(true);
        launcherLogPathLabel.setMaxWidth(CARD_WIDTH);

        HBox launcherActionRow =
                new HBox(8, btnCopyLauncherPath, btnForceDeployLauncher);
        launcherActionRow.setAlignment(Pos.CENTER_LEFT);
        launcherActionRow.setMaxWidth(CARD_WIDTH);

        VBox launcherPathBlock =
                new VBox(4, launcherPathLabel, launcherLogPathLabel, launcherActionRow);
        launcherPathBlock.setFillWidth(true);
        launcherPathBlock.setMaxWidth(CARD_WIDTH);

        Label deployStatusLabel = new Label();
        deployStatusLabel.setWrapText(true);
        deployStatusLabel.setMaxWidth(CARD_WIDTH);
        deployStatusLabel.getStyleClass().add("pm-rdp-status-label");

        List<ProfileRowFields> profileFields = new ArrayList<>();
        VBox profileBox = new VBox(10);
        profileBox.setFillWidth(true);

        Map<Integer, RdpLaunchProfile> profileMetadataByNumber = new HashMap<>();

        ComboBox<Integer> launchProfileCombo = new ComboBox<>();
        launchProfileCombo.setPrefWidth(CARD_WIDTH);
        launchProfileCombo.setMaxWidth(CARD_WIDTH);
        launchProfileCombo.setButtonCell(createLaunchProfileListCell(profileFields, profileMetadataByNumber));
        launchProfileCombo.setCellFactory(
                listView -> createLaunchProfileListCell(profileFields, profileMetadataByNumber));

        Label launchProfileDetail = new Label();
        launchProfileDetail.getStyleClass().add("pm-rdp-profile-detail");
        launchProfileDetail.setWrapText(true);
        launchProfileDetail.setMaxWidth(CARD_WIDTH);

        Runnable[] refreshLaunchProfileCombo = new Runnable[1];
        refreshLaunchProfileCombo[0] =
                () -> {
                    Integer previous = launchProfileCombo.getValue();
                    launchProfileCombo.getItems().clear();
                    for (ProfileRowFields row : profileFields) {
                        launchProfileCombo.getItems().add(row.number());
                    }
                    int fallback = AppPaths.resolveRdpLaunchProfileNumber(uiEnv.get());
                    Integer target =
                            previous != null && launchProfileCombo.getItems().contains(previous)
                                    ? previous
                                    : (launchProfileCombo.getItems().contains(fallback)
                                            ? fallback
                                            : (launchProfileCombo.getItems().isEmpty()
                                                    ? null
                                                    : launchProfileCombo.getItems().getFirst()));
                    launchProfileCombo.setValue(target);
                    updateLaunchProfileDetail(
                            launchProfileDetail, target, profileFields, profileMetadataByNumber);
                };

        launchProfileCombo
                .valueProperty()
                .addListener(
                        (obs, was, now) -> {
                            updateLaunchProfileDetail(
                                    launchProfileDetail, now, profileFields, profileMetadataByNumber);
                            if (now != null && launchProfileNumberChangeHandler != null) {
                                launchProfileNumberChangeHandler.accept(now);
                            }
                        });

        CheckBox chkDisconnectOnChildExit = new CheckBox("子プロセス終了時に RDP を切断");
        chkDisconnectOnChildExit.setSelected(true);
        chkDisconnectOnChildExit.setTooltip(
                new Tooltip(
                        "接続先ランチャーが起動したプログラムの終了を待ち、終了後に RDP セッションを切断します。"
                                + " RAP設定.ini の「"
                                + RdpRemoteLauncherIni.DISCONNECT_ON_CHILD_EXIT_KEY
                                + "」に保存されます。"));

        Runnable refreshPaths =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path launcherExe = AppPaths.resolveRdpLauncherExe(ui);
                    iniPathLabel.setText(
                            "RAP設定.ini: " + AppPaths.resolveRdpLauncherIni(ui).toString());
                    launcherPathLabel.setText("ランチャー exe: " + launcherExe.toString());
                    Path sharedLog = AppPaths.resolveRdpLauncherSharedLogPath(ui);
                    launcherLogPathLabel.setText(
                            "接続先ランチャーログ（共有フォルダ）: "
                                    + sharedLog
                                    + " ／ 接続先 TEMP: %TEMP%\\PM-AI-RDP-Launcher\\launcher-"
                                    + java.time.LocalDate.now()
                                            .format(
                                                    java.time.format.DateTimeFormatter
                                                            .BASIC_ISO_DATE)
                                    + ".log");
                };
        refreshPaths.run();

        java.util.function.IntConsumer appendProfileRow =
                profileNumber -> {
                    Label header = new Label("プロファイル " + profileNumber);
                    header.getStyleClass().add("pm-rdp-profile-card-title");

                    TextField nameField = new TextField();
                    nameField.setPromptText("例: アラジン 工程マスタ取得");
                    HBox.setHgrow(nameField, Priority.ALWAYS);

                    TextField descriptionField = new TextField();
                    descriptionField.setPromptText("この接続で行う作業の説明（RPA シナリオは接続先で別設定）");
                    HBox.setHgrow(descriptionField, Priority.ALWAYS);

                    TextField categoryField = new TextField();
                    categoryField.setPromptText("例: マスタ更新 / データ取得");
                    categoryField.setPrefWidth(180);

                    TextField programField = new TextField();
                    programField.setPromptText("接続先の RPA exe パス（RAP設定.ini スロット " + profileNumber + "）");
                    HBox.setHgrow(programField, Priority.ALWAYS);

                    TextField argsField = new TextField();
                    argsField.setPromptText(
                            "任意: --scenario path.ardrpa（空なら Aladdin 本体のみ起動）");
                    HBox.setHgrow(argsField, Priority.ALWAYS);

                    Button browseArgs = new Button("参照");
                    styleSecondaryButton(browseArgs);
                    browseArgs.setTooltip(
                            new Tooltip("シナリオ .ardrpa を参照（接続先 UNC パスは手入力）。"));
                    browseArgs.setOnAction(
                            e -> {
                                FileChooser chooser = new FileChooser();
                                chooser.setTitle("プロファイル " + profileNumber + " の RPA シナリオ");
                                chooser.getExtensionFilters()
                                        .add(
                                                new FileChooser.ExtensionFilter(
                                                        "Aladdin シナリオ (*.ardrpa)", "*.ardrpa"));
                                java.io.File chosen =
                                        chooser.showOpenDialog(
                                                owner != null ? owner : browseArgs.getScene().getWindow());
                                if (chosen != null) {
                                    argsField.setText(
                                            RdpRemoteLauncherIni.formatScenarioArgument(
                                                    chosen.getAbsolutePath()));
                                }
                            });

                    HBox argsRow = new HBox(8, argsField, browseArgs);
                    argsRow.setAlignment(Pos.CENTER_LEFT);
                    argsRow.setMaxWidth(CARD_WIDTH);
                    HBox.setHgrow(argsField, Priority.ALWAYS);

                    CheckBox chkRpaEternal = new CheckBox("RPA に --eternal を付与");
                    chkRpaEternal.setSelected(false);
                    chkRpaEternal.setTooltip(
                            new Tooltip(
                                    "オンにすると RPA 起動引数へ --eternal を追加します。"
                                            + " シナリオが無い場合や、シナリオ終了後もプロセスを終了させたくないときに使います。"
                                            + " 接続先ランチャーは --id / --password を先に付与し、--eternal は末尾に付けます。"));

                    Button browse = new Button("参照");
                    styleSecondaryButton(browse);
                    browse.setTooltip(
                            new Tooltip("ローカル PC 上の exe を参照（接続先 UNC パスは手入力）。"));
                    browse.setOnAction(
                            e -> {
                                FileChooser chooser = new FileChooser();
                                chooser.setTitle("プロファイル " + profileNumber + " の RPA プログラム");
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
                    HBox.setHgrow(programField, Priority.ALWAYS);

                    HBox categoryRow = new HBox(8, fieldCaption("区分"), categoryField);
                    categoryRow.setAlignment(Pos.CENTER_LEFT);

                    VBox cardInner =
                            new VBox(
                                    8,
                                    header,
                                    fieldCaption("名称"),
                                    nameField,
                                    fieldCaption("説明"),
                                    descriptionField,
                                    categoryRow,
                                    fieldCaption("RPA プログラム"),
                                    programRow,
                                    fieldCaption("RPA 引数"),
                                    argsRow,
                                    chkRpaEternal);
                    cardInner.getStyleClass().add("pm-rdp-profile-card");
                    cardInner.setFillWidth(true);
                    cardInner.setMaxWidth(CARD_WIDTH);

                    Runnable refreshCombo = refreshLaunchProfileCombo[0];
                    nameField
                            .focusedProperty()
                            .addListener(
                                    (obs, wasFocused, focused) -> {
                                        if (wasFocused && !focused && refreshCombo != null) {
                                            refreshCombo.run();
                                        }
                                    });

                    profileBox.getChildren().add(cardInner);
                    profileFields.add(
                            new ProfileRowFields(
                                    profileNumber,
                                    nameField,
                                    descriptionField,
                                    categoryField,
                                    programField,
                                    argsField,
                                    chkRpaEternal));
                };

        for (int i = 1; i <= 3; i++) {
            appendProfileRow.accept(i);
        }
        refreshLaunchProfileCombo[0].run();

        Runnable[] refreshIniFilePreview = new Runnable[1];
        Runnable[] refreshLauncherLog = new Runnable[1];

        Runnable loadIniFromShare =
                () ->
                        scheduleIniReloadFromShare(
                                uiEnv,
                                launchProfileCombo,
                                profileFields,
                                profileMetadataByNumber,
                                appendProfileRow,
                                rapStatusLabel,
                                refreshPaths,
                                chkDisconnectOnChildExit,
                                refreshIniFilePreview[0],
                                refreshLauncherLog[0],
                                refreshLaunchProfileCombo[0]);

        Runnable saveIniToShareBody =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    Path profilesPath = AppPaths.resolveRdpLaunchProfilesFile(ui);
                    RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
                    Integer selectedProfile = launchProfileCombo.getValue();
                    int profileNumber =
                            selectedProfile != null
                                    ? selectedProfile
                                    : AppPaths.resolveRdpLaunchProfileNumber(ui);
                    ini.setSelectedSlot(profileNumber);
                    ini.setDisconnectOnChildExit(chkDisconnectOnChildExit.isSelected());
                    for (ProfileRowFields row : profileFields) {
                        ini.setSlotCommand(
                                row.number(),
                                row.programField().getText(),
                                RdpRemoteLauncherIni.mergeEternalFlag(
                                        RdpRemoteLauncherIni.normalizeScenarioArguments(
                                                row.argsField().getText()),
                                        row.chkRpaEternal().isSelected()));
                    }
                    String validation = ini.validateMessageForSave();
                    if (validation != null) {
                        showAlert(Alert.AlertType.WARNING, "保存できません", validation);
                        rapStatusLabel.setText(validation);
                        return;
                    }
                    try {
                        ini.save(iniPath);
                        List<RdpLaunchProfile> profiles = collectProfilesFromRows(profileFields);
                        RdpLaunchProfileCatalog.save(profilesPath, profiles);
                        for (RdpLaunchProfile profile : profiles) {
                            profileMetadataByNumber.put(profile.number(), profile);
                        }
                        refreshLaunchProfileCombo[0].run();
                        rapStatusLabel.setText(
                                "保存しました: "
                                        + iniPath
                                        + " / "
                                        + profilesPath);
                        status.accept(
                                "RAP設定.ini と起動プロファイル JSON を保存しました: "
                                        + iniPath);
                        refreshRightPanePreviews(refreshIniFilePreview[0], refreshLauncherLog[0]);
                    } catch (IOException ex) {
                        rapStatusLabel.setText("保存失敗: " + ex.getMessage());
                        showAlert(Alert.AlertType.ERROR, "保存失敗", ex.getMessage());
                    }
                };

        Runnable[] updateLaunchButtonState = new Runnable[1];
        Runnable[] refreshAladdinCredentialsUi = new Runnable[1];
        Runnable[] saveIniToShare = new Runnable[1];

        Button btnReloadIni = new Button("再読込");
        styleSecondaryButton(btnReloadIni);
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
        styleSecondaryButton(btnSaveIni);
        btnSaveIni.setOnAction(e -> saveIniToShare[0].run());

        Button btnAddProfile = new Button("+ プロファイルを追加");
        styleSecondaryButton(btnAddProfile);
        btnAddProfile.setOnAction(
                e -> {
                    if (profileFields.size() >= RdpRemoteLauncherIni.MAX_SLOTS) {
                        return;
                    }
                    appendProfileRow.accept(profileFields.size() + 1);
                    btnAddProfile.setDisable(
                            profileFields.size() >= RdpRemoteLauncherIni.MAX_SLOTS);
                    refreshLaunchProfileCombo[0].run();
                });

        Button btnOpenFolder = new Button("フォルダを開く");
        styleSecondaryButton(btnOpenFolder);
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

        Button btnSuppressLaunchSlot = new Button("起動番号を 0 に");
        styleSecondaryButton(btnSuppressLaunchSlot);
        btnSuppressLaunchSlot.setTooltip(
                new Tooltip(
                        "RAP設定.ini の「"
                                + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                + "」を 0 に書き込み、接続先タスクスケジューラによる RPA 自動起動を抑止します。"
                                + " プロファイル定義（スロット行）は保持されます。"));
        btnSuppressLaunchSlot.setOnAction(
                e -> {
                    Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
                    confirm.setTitle("起動番号 0（抑止）");
                    confirm.setHeaderText(null);
                    confirm.setContentText(
                            "共有フォルダの RAP設定.ini に "
                                    + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                    + "=0 を書き込みます。"
                                    + " タスクスケジューラ経由の RPA 起動が抑止されます。続行しますか？");
                    confirm.showAndWait()
                            .ifPresent(
                                    choice -> {
                                        if (choice == ButtonType.OK) {
                                            scheduleTaskSchedulerSuppress(
                                                    uiEnv,
                                                    rapStatusLabel,
                                                    status,
                                                    loadIniFromShare);
                                        }
                                    });
                });

        HBox rapActionRow =
                new HBox(
                        8,
                        btnReloadIni,
                        btnSaveIni,
                        btnSuppressLaunchSlot,
                        btnOpenFolder,
                        btnAddProfile);
        rapActionRow.setAlignment(Pos.CENTER_LEFT);

        VBox rapContent =
                new VBox(
                        8,
                        iniPathLabel,
                        launcherPathBlock,
                        deployStatusLabel,
                        profileBox,
                        chkDisconnectOnChildExit,
                        rapActionRow,
                        rapStatusLabel);
        rapContent.setFillWidth(true);

        TitledPane rapPane = new TitledPane("RAP設定（起動プロファイルと RPA コマンド）", rapContent);
        rapPane.setExpanded(true);
        rapPane.setMaxWidth(CARD_WIDTH);
        rapPane.getStyleClass().add("pm-rdp-section-pane");

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
        statusLabel.getStyleClass().add("pm-rdp-status-label");

        CheckBox chkFullScreen = new CheckBox("全画面で起動");
        chkFullScreen.setTooltip(
                new Tooltip("オフのときは下記の幅・高さでウィンドウ表示（配台計画システムの背面で作業しやすい）。"));

        Spinner<Integer> desktopWidthSpinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(800, 3840, 1280, 10));
        desktopWidthSpinner.setEditable(true);
        desktopWidthSpinner.setPrefWidth(100);

        Spinner<Integer> desktopHeightSpinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(600, 2160, 800, 10));
        desktopHeightSpinner.setEditable(true);
        desktopHeightSpinner.setPrefWidth(100);

        Runnable applyDisplayFieldsFromUi =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    chkFullScreen.setSelected(RdpLaunchDisplaySettings.resolveFullScreen(ui));
                    desktopWidthSpinner
                            .getValueFactory()
                            .setValue(RdpLaunchDisplaySettings.resolveWidth(ui));
                    desktopHeightSpinner
                            .getValueFactory()
                            .setValue(RdpLaunchDisplaySettings.resolveHeight(ui));
                    updateDisplayControlsEnabled(
                            chkFullScreen, desktopWidthSpinner, desktopHeightSpinner);
                };

        Runnable persistDisplayFields =
                () -> {
                    Map<String, String> values = new HashMap<>();
                    values.put(
                            AppPaths.KEY_PM_AI_RDP_FULLSCREEN,
                            chkFullScreen.isSelected() ? "1" : "0");
                    values.put(
                            AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH,
                            String.valueOf(desktopWidthSpinner.getValue()));
                    values.put(
                            AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT,
                            String.valueOf(desktopHeightSpinner.getValue()));
                    if (displayEnvChangeHandler != null) {
                        displayEnvChangeHandler.accept(values);
                    }
                };

        chkFullScreen
                .selectedProperty()
                .addListener(
                        (obs, was, selected) -> {
                            updateDisplayControlsEnabled(
                                    chkFullScreen, desktopWidthSpinner, desktopHeightSpinner);
                            persistDisplayFields.run();
                        });
        desktopWidthSpinner
                .valueProperty()
                .addListener((obs, was, now) -> persistDisplayFields.run());
        desktopHeightSpinner
                .valueProperty()
                .addListener((obs, was, now) -> persistDisplayFields.run());

        HBox displaySizeRow =
                new HBox(
                        8,
                        new Label("幅:"),
                        desktopWidthSpinner,
                        new Label("高さ:"),
                        desktopHeightSpinner);
        displaySizeRow.setAlignment(Pos.CENTER_LEFT);

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
        btnLaunch.getStyleClass().add("pm-rdp-connect-button");
        btnLaunch.setMaxWidth(Double.MAX_VALUE);
        btnLaunch.setTooltip(
                new Tooltip(
                        "指定 .rdp で mstsc.exe を起動します。"
                                + " 接続先の起動プログラムは RAP設定.ini + タスクスケジューラが参照します。"));
        btnLaunch.setOnAction(
                e -> {
                    if (launcherDeployInProgress.get()) {
                        showAlert(
                                Alert.AlertType.INFORMATION,
                                "転送中",
                                "ランチャー転送が完了するまで、リモートデスクトップは起動できません。");
                        return;
                    }
                    if (rdpSessionActive.get()) {
                        showAlert(
                                Alert.AlertType.INFORMATION,
                                "接続中",
                                "リモートデスクトップは既に起動しています。接続が終了するまで再実行できません。");
                        return;
                    }
                    persistCompanionFields.run();
                    persistDisplayFields.run();
                    Map<String, String> uiGate = uiEnv.get();
                    FactorySite siteGate = GlobalInitSettingTarget.loadEffective(uiGate);
                    String sessionOperator = FactoryOperatorUserStore.sessionOperatorName();
                    if (sessionOperator.isBlank()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "操作者未選択",
                                "起動時に操作者を選んでからリモートデスクトップに接続してください。");
                        return;
                    }
                    if (FactoryOperatorUserStore.isGuestOperator(sessionOperator)) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "ゲスト",
                                "ゲスト操作者ではリモートデスクトップに接続できません。");
                        return;
                    }
                    try {
                        FactoryOperatorUserStore.configureFromUi(uiGate, siteGate);
                        if (!FactoryOperatorUserStore.hasAladdinCredentials(
                                siteGate, sessionOperator)) {
                            showAlert(
                                    Alert.AlertType.WARNING,
                                    "アラジン資格情報未設定",
                                    "リモートデスクトップタブでアラジン ログイン ID とパスワードを保存してから接続してください。");
                            return;
                        }
                    } catch (IOException credEx) {
                        showAlert(
                                Alert.AlertType.ERROR,
                                "資格情報の確認に失敗",
                                credEx.getMessage());
                        return;
                    }
                    Integer selectedLaunchProfile = launchProfileCombo.getValue();
                    if (selectedLaunchProfile == null) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "プロファイル未選択",
                                "起動プロファイルを選択してください。");
                        return;
                    }
                    ProfileRowFields launchRow = findProfileRow(profileFields, selectedLaunchProfile);
                    if (launchRow == null) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "プロファイル未設定",
                                "プロファイル "
                                        + selectedLaunchProfile
                                        + " が見つかりません。");
                        return;
                    }
                    if (launchRow.programField().getText().trim().isEmpty()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "RPA 未設定",
                                "プロファイル "
                                        + selectedLaunchProfile
                                        + " の RPA プログラムが空です。"
                                        + " RAP設定で exe パスを登録して保存してください。");
                        return;
                    }
                    RdpLaunchProfile launchProfileMeta =
                            profileMetadataByNumber.getOrDefault(
                                    selectedLaunchProfile,
                                    buildProfileFromRow(launchRow));
                    Map<String, String> ui = new HashMap<>(uiEnv.get());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM,
                            companionProgramField.getText().trim());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS,
                            companionArgsField.getText().trim());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_FULLSCREEN,
                            chkFullScreen.isSelected() ? "1" : "0");
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH,
                            String.valueOf(desktopWidthSpinner.getValue()));
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT,
                            String.valueOf(desktopHeightSpinner.getValue()));
                    applyProfileLaunchOverrides(ui, launchProfileMeta);
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

                    int launchSlot = selectedLaunchProfile;
                    if (launchProfileNumberChangeHandler != null) {
                        launchProfileNumberChangeHandler.accept(launchSlot);
                    }
                    String launchProfileLabel = profileComboLabel(launchRow, launchProfileMeta);
                    Path launcherIniPath = AppPaths.resolveRdpLauncherIni(ui);
                    try {
                        if (launchProfileMeta.disconnectOnChildExit() != null) {
                            RdpRemoteLauncherIni preIni = RdpRemoteLauncherIni.load(launcherIniPath);
                            preIni.setDisconnectOnChildExit(launchProfileMeta.disconnectOnChildExit());
                            preIni.save(launcherIniPath);
                        }
                        FactoryOperatorUserStore.syncLauncherCredentialsJsonToDeployDir(ui);
                        RdpRemoteLauncherIni.writeTaskSchedulerSlotBeforeConnect(
                                launcherIniPath, launchSlot, ui);
                        RdpRemoteLauncherIni.writeOperatorContext(
                                launcherIniPath, sessionOperator);
                        refreshRightPanePreviews(refreshIniFilePreview[0], refreshLauncherLog[0]);
                    } catch (IOException iniEx) {
                        showAlert(
                                Alert.AlertType.ERROR,
                                "RAP設定.ini の更新に失敗",
                                "接続前に起動プログラム番号を書き込めませんでした: "
                                        + iniEx.getMessage());
                        return;
                    }

                    Runnable performLaunch =
                            () -> {
                                try {
                                    Path configured = profile.get();
                                    Path preferred =
                                            RdpFileSigner.resolvePreferredSignedProfilePath(
                                                    configured, ui);
                                    if (!rdpSessionActive.compareAndSet(false, true)) {
                                        showAlert(
                                                Alert.AlertType.INFORMATION,
                                                "接続中",
                                                "リモートデスクトップは既に起動しています。");
                                        return;
                                    }
                                    updateLaunchButtonState[0].run();
                                    RemoteDesktopLauncher.LaunchOutcome outcome =
                                            RemoteDesktopLauncher.launch(preferred, ui);
                                    if (!preferred.equals(configured)
                                            && profileChangeHandler != null) {
                                        profileChangeHandler.accept(preferred.toString());
                                        profileField.setText(preferred.toString());
                                    }
                                    String msg =
                                            "リモートデスクトップを起動しました: "
                                                    + outcome.rdpProfile();
                                    msg +=
                                            "\n表示: "
                                                    + RdpLaunchDisplaySettings.formatSummary(ui);
                                    msg +=
                                            "\n起動プロファイル: "
                                                    + launchProfileLabel
                                                    + "（"
                                                    + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                                    + "="
                                                    + launchSlot
                                                    + "）";
                                    if (outcome.remoteStartupSummary().isPresent()) {
                                        msg +=
                                                "\n.rdp に組込: "
                                                        + outcome.remoteStartupSummary().get();
                                    }
                                    if (outcome.signatureRemoved()) {
                                        msg +=
                                                "\n（.rdp 編集により署名が無効化されました。必要なら署名ウィザードを実行してください）";
                                    }
                                    msg += "\n接続終了を監視中…";
                                    status.accept(msg);
                                    statusLabel.setText(msg);
                                    RdpMstscSessionMonitor.watchAfterLaunch(
                                            outcome.rdpProfile(),
                                            outcome.mstscProcessId(),
                                            outcome.mstscPidMarkerFile(),
                                            rdpSessionWatchThread,
                                            event ->
                                                    Platform.runLater(
                                                            () -> {
                                                    rdpSessionActive.set(false);
                                                    updateLaunchButtonState[0].run();
                                                    try {
                                                        Path iniPath =
                                                                AppPaths.resolveRdpLauncherIni(
                                                                        ui);
                                                        RdpRemoteLauncherIni
                                                                .writeTaskSchedulerSuppress(
                                                                        iniPath, ui);
                                                        status.accept(
                                                                "RAP設定.ini の起動プログラム番号を "
                                                                        + RdpRemoteLauncherIni
                                                                                .SLOT_DISABLED
                                                                        + " に設定しました（タスクスケジューラ抑止・保険）。");
                                                        refreshRightPanePreviews(
                                                                refreshIniFilePreview[0],
                                                                refreshLauncherLog[0]);
                                                    } catch (IOException suppressEx) {
                                                        status.accept(
                                                                "起動プログラム番号の抑止（0）設定に失敗: "
                                                                        + suppressEx
                                                                                .getMessage());
                                                    }
                                                                String endMsg =
                                                                        formatRdpSessionEndMessage(
                                                                                event);
                                                                status.accept(endMsg);
                                                                statusLabel.setText(endMsg);
                                                            }));
                                } catch (IOException ex) {
                                    rdpSessionActive.set(false);
                                    updateLaunchButtonState[0].run();
                                    showAlert(
                                            Alert.AlertType.ERROR, "起動失敗", ex.getMessage());
                                    status.accept("リモートデスクトップ起動失敗: " + ex.getMessage());
                                    statusLabel.setText(ex.getMessage());
                                }
                            };

                    if (!launcherDeployReady.get()) {
                        pendingLaunchAfterDeploy.set(performLaunch);
                        scheduleLauncherDeploy(
                                uiEnv,
                                false,
                                deployStatusLabel,
                                btnForceDeployLauncher,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                updateLaunchButtonState[0],
                                status,
                                () -> {
                                    Runnable pending = pendingLaunchAfterDeploy.getAndSet(null);
                                    if (launcherDeployReady.get() && pending != null) {
                                        pending.run();
                                    } else if (pending != null) {
                                        showAlert(
                                                Alert.AlertType.ERROR,
                                                "転送失敗",
                                                deployStatusLabel.getText().isBlank()
                                                        ? "ランチャーの転送に失敗したため、リモートデスクトップを起動できません。"
                                                                + " 接続先セッション終了後に「ランチャーを強制転送」を試してください。"
                                                        : deployStatusLabel.getText());
                                    }
                                });
                        return;
                    }
                    performLaunch.run();
                });

        Button btnChoose = new Button("プロファイルを選択");
        styleSecondaryButton(btnChoose);
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

        Button btnSignWizard = new Button("RDP 署名ウィザード");
        styleSecondaryButton(btnSignWizard);
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
        HBox.setHgrow(profileField, Priority.ALWAYS);
        HBox.setHgrow(btnChoose, Priority.NEVER);

        HBox signWizardRow = new HBox(btnSignWizard);
        signWizardRow.setAlignment(Pos.CENTER);

        VBox launchBlock = new VBox(10, btnLaunch, signWizardRow);
        launchBlock.getStyleClass().add("pm-rdp-launch-block");
        launchBlock.setAlignment(Pos.CENTER);
        launchBlock.setFillWidth(true);
        launchBlock.setMaxWidth(CARD_WIDTH);

        saveIniToShare[0] =
                () ->
                        scheduleLauncherDeploy(
                                uiEnv,
                                false,
                                deployStatusLabel,
                                btnForceDeployLauncher,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                updateLaunchButtonState[0],
                                status,
                                saveIniToShareBody);

        btnForceDeployLauncher.setOnAction(
                e ->
                        scheduleLauncherDeploy(
                                uiEnv,
                                true,
                                deployStatusLabel,
                                btnForceDeployLauncher,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                updateLaunchButtonState[0],
                                status,
                                null));

        Label connectSectionTitle = new Label("リモートデスクトップ接続");
        connectSectionTitle.getStyleClass().add("pm-rdp-section-title");

        Label operatorNameLabel = new Label();
        operatorNameLabel.getStyleClass().add("pm-rdp-meta-label");
        operatorNameLabel.setWrapText(true);
        operatorNameLabel.setMaxWidth(CARD_WIDTH);

        TextField aladdinLoginIdField = new TextField();
        aladdinLoginIdField.setPromptText("例: 000585");
        HBox.setHgrow(aladdinLoginIdField, Priority.ALWAYS);

        PasswordField aladdinPasswordField = new PasswordField();
        aladdinPasswordField.setPromptText("変更時のみ入力");
        HBox.setHgrow(aladdinPasswordField, Priority.ALWAYS);

        Label credentialsStatusLabel = new Label();
        credentialsStatusLabel.getStyleClass().add("pm-rdp-status-label");
        credentialsStatusLabel.setWrapText(true);
        credentialsStatusLabel.setMaxWidth(CARD_WIDTH);

        Label rpaInjectedArgsHint =
                new Label(
                        "RPA 引数が空のときは Aladdin_RPA_Studio.exe のみ起動します。"
                                + " シナリオ実行時は "
                                + AladdinRpaLaunchArgs.SCENARIO_FLAG
                                + " と .ardrpa パスを入力してください。"
                                + " "
                                + AladdinRpaLaunchArgs.ID_FLAG
                                + " / "
                                + AladdinRpaLaunchArgs.PASSWORD_FLAG
                                + " は資格情報から自動付与され、引数欄には書きません。");
        rpaInjectedArgsHint.getStyleClass().add("pm-rdp-page-subtitle");
        rpaInjectedArgsHint.setWrapText(true);
        rpaInjectedArgsHint.setMaxWidth(CARD_WIDTH);

        Button btnSaveAladdinCredentials = new Button("資格情報を保存");
        styleSecondaryButton(btnSaveAladdinCredentials);

        btnSaveAladdinCredentials.setOnAction(
                e -> {
                    Map<String, String> ui = uiEnv.get();
                    FactorySite site = GlobalInitSettingTarget.loadEffective(ui);
                    String operator = FactoryOperatorUserStore.sessionOperatorName();
                    if (operator.isBlank()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "操作者未選択",
                                "起動時に操作者を選んでから資格情報を保存してください。");
                        return;
                    }
                    if (FactoryOperatorUserStore.isGuestOperator(operator)) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "ゲスト",
                                "ゲスト操作者にはアラジン資格情報を設定できません。");
                        return;
                    }
                    String loginId = aladdinLoginIdField.getText().trim();
                    String password = aladdinPasswordField.getText();
                    if (loginId.isEmpty() || password.isBlank()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "入力不足",
                                "アラジン ログイン ID とパスワードを入力してください。");
                        return;
                    }
                    try {
                        FactoryOperatorUserStore.configureFromUi(ui, site);
                        FactoryOperatorUserStore.setAladdinCredentials(
                                site, operator, loginId, password);
                        FactoryOperatorUserStore.syncLauncherCredentialsJsonToDeployDir(ui);
                        aladdinPasswordField.clear();
                        if (refreshAladdinCredentialsUi[0] != null) {
                            refreshAladdinCredentialsUi[0].run();
                        }
                        status.accept("アラジン資格情報を保存しました: " + operator);
                    } catch (Exception ex) {
                        showAlert(Alert.AlertType.ERROR, "保存失敗", ex.getMessage());
                    }
                });

        HBox credentialsSaveRow = new HBox(8, btnSaveAladdinCredentials);
        credentialsSaveRow.setAlignment(Pos.CENTER_LEFT);

        VBox connectHero =
                new VBox(
                        12,
                        connectSectionTitle,
                        operatorNameLabel,
                        fieldCaption("アラジン ログイン ID"),
                        aladdinLoginIdField,
                        fieldCaption("アラジン パスワード"),
                        aladdinPasswordField,
                        credentialsSaveRow,
                        credentialsStatusLabel,
                        rpaInjectedArgsHint,
                        fieldCaption("起動プロファイル"),
                        launchProfileCombo,
                        launchProfileDetail,
                        fieldCaption("RDP プロファイル (.rdp)"),
                        profileRow,
                        chkFullScreen,
                        displaySizeRow,
                        launchBlock,
                        statusLabel);
        connectHero.getStyleClass().add("pm-rdp-connect-hero");
        connectHero.setFillWidth(true);
        connectHero.setMaxWidth(CARD_WIDTH);

        refreshAladdinCredentialsUi[0] =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    FactorySite site = GlobalInitSettingTarget.loadEffective(ui);
                    try {
                        FactoryOperatorUserStore.configureFromUi(ui, site);
                        String operator = FactoryOperatorUserStore.sessionOperatorName();
                        if (operator.isBlank()) {
                            operatorNameLabel.setText("操作者: （起動時に操作者を選んでください）");
                            aladdinLoginIdField.clear();
                            aladdinPasswordField.clear();
                            credentialsStatusLabel.setText("");
                        } else if (FactoryOperatorUserStore.isGuestOperator(operator)) {
                            operatorNameLabel.setText("操作者: " + operator + "（ゲストは接続不可）");
                            aladdinLoginIdField.clear();
                            aladdinPasswordField.clear();
                            credentialsStatusLabel.setText("");
                        } else {
                            operatorNameLabel.setText("操作者: " + operator);
                            String loginId =
                                    FactoryOperatorUserStore.aladdinLoginIdFor(site, operator);
                            aladdinLoginIdField.setText(loginId);
                            aladdinPasswordField.clear();
                            String loginIdPreview =
                                    FactoryOperatorUserStore.aladdinLoginIdFor(site, operator);
                            if (FactoryOperatorUserStore.hasAladdinCredentials(site, operator)) {
                                credentialsStatusLabel.setText(
                                        "アラジン資格情報: 設定済。"
                                                + " RPA 起動引数に付与: "
                                                + AladdinRpaLaunchArgs.ID_FLAG
                                                + " "
                                                + loginIdPreview
                                                + " "
                                                + AladdinRpaLaunchArgs.PASSWORD_FLAG
                                                + " ****");
                            } else {
                                credentialsStatusLabel.setText(
                                        "アラジン資格情報: 未設定（保存後に接続できます）");
                            }
                        }
                    } catch (IOException ex) {
                        credentialsStatusLabel.setText("資格情報の読込に失敗: " + ex.getMessage());
                    }
                    if (updateLaunchButtonState[0] != null) {
                        updateLaunchButtonState[0].run();
                    }
                };

        updateLaunchButtonState[0] =
                () ->
                        applyRdpLaunchButtonState(
                                uiEnv,
                                btnLaunch,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                rdpSessionActive);
        refreshAladdinCredentialsUi[0].run();

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
        policyPane.getStyleClass().add("pm-rdp-section-pane");

        VBox card = new VBox(16, headerBlock, connectHero, rapPane, policyPane);
        card.getStyleClass().add("pm-rdp-settings-stack");
        card.setPadding(new Insets(16));
        card.setMaxWidth(CARD_WIDTH);
        card.setPrefWidth(CARD_WIDTH);
        card.setFillWidth(true);

        Label iniPreviewTitle = new Label("RAP設定.ini（共有ファイル）");
        iniPreviewTitle.getStyleClass().add("pm-rdp-ini-panel-title");

        Label iniPreviewMetaLabel = new Label("読込待ち…");
        iniPreviewMetaLabel.setWrapText(true);
        iniPreviewMetaLabel.getStyleClass().add("pm-rdp-page-subtitle");

        TextArea iniPreviewArea = new TextArea();
        iniPreviewArea.setEditable(false);
        iniPreviewArea.setWrapText(false);
        iniPreviewArea.setPromptText("共有 UNC 上の RAP設定.ini の生テキストがここに表示されます。");
        iniPreviewArea.getStyleClass().add("pm-remote-desktop-ini-preview");
        VBox.setVgrow(iniPreviewArea, Priority.ALWAYS);

        Button btnRefreshIniPreview = new Button("プレビュー更新");
        styleSecondaryButton(btnRefreshIniPreview);
        refreshIniFilePreview[0] =
                () ->
                        scheduleIniFilePreviewRefresh(
                                uiEnv, iniPreviewArea, iniPreviewMetaLabel);
        btnRefreshIniPreview.setOnAction(e -> refreshIniFilePreview[0].run());

        HBox iniPreviewActions = new HBox(8, btnRefreshIniPreview);
        iniPreviewActions.setAlignment(Pos.CENTER_LEFT);

        VBox iniPreviewPane =
                new VBox(10, iniPreviewTitle, iniPreviewMetaLabel, iniPreviewActions, iniPreviewArea);
        iniPreviewPane.getStyleClass().add("pm-rdp-ini-panel");
        iniPreviewPane.setPadding(new Insets(16));
        iniPreviewPane.setFillWidth(true);
        VBox.setVgrow(iniPreviewPane, Priority.ALWAYS);

        AtomicReference<String> launcherLogFullText = new AtomicReference<>("");
        AtomicReference<String> launcherLogBaseMeta = new AtomicReference<>("読込待ち…");

        Label launcherLogTitle = new Label("接続先ランチャーログ");
        launcherLogTitle.getStyleClass().add("pm-rdp-ini-panel-title");

        Label launcherLogMetaLabel = new Label("読込待ち…");
        launcherLogMetaLabel.setWrapText(true);
        launcherLogMetaLabel.getStyleClass().add("pm-rdp-page-subtitle");

        TextField launcherLogSearchField = new TextField();
        launcherLogSearchField.setPromptText("ログ行を部分一致で絞り込み（空欄ですべて表示）");
        HBox.setHgrow(launcherLogSearchField, Priority.ALWAYS);

        ComboBox<String> launcherLogFontFamilyCombo = new ComboBox<>();
        launcherLogFontFamilyCombo.setPrefWidth(120);
        List<String> launcherLogFontFamilies = new ArrayList<>();
        launcherLogFontFamilies.add(LAUNCHER_LOG_DEFAULT_FONT_FAMILY_LABEL);
        List<String> installedFontFamilies = new ArrayList<>(Font.getFamilies());
        Collections.sort(installedFontFamilies);
        launcherLogFontFamilies.addAll(installedFontFamilies);
        launcherLogFontFamilyCombo.getItems().setAll(launcherLogFontFamilies);
        launcherLogFontFamilyCombo.getSelectionModel().selectFirst();

        ComboBox<Double> launcherLogFontSizeCombo = new ComboBox<>();
        launcherLogFontSizeCombo.setPrefWidth(72);
        launcherLogFontSizeCombo.getItems().setAll(LAUNCHER_LOG_PRESET_FONT_SIZES);
        launcherLogFontSizeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(Double object) {
                        if (object == null) {
                            return "";
                        }
                        if (object == Math.rint(object)) {
                            return String.valueOf(object.intValue());
                        }
                        return object.toString();
                    }

                    @Override
                    public Double fromString(String string) {
                        if (string == null || string.isBlank()) {
                            return null;
                        }
                        return Double.valueOf(string.trim());
                    }
                });
        launcherLogFontSizeCombo.setValue(12d);

        TextArea launcherLogArea = new TextArea();
        launcherLogArea.setEditable(false);
        launcherLogArea.setWrapText(false);
        launcherLogArea.setPromptText("共有フォルダ上の launcher-yyyyMMdd.log がここに表示されます。");
        launcherLogArea.getStyleClass().add("pm-remote-desktop-launcher-log");
        VBox.setVgrow(launcherLogArea, Priority.ALWAYS);

        Runnable applyLauncherLogFont =
                () ->
                        applyLauncherLogAreaFont(
                                launcherLogArea,
                                launcherLogFontFamilyCombo.getValue(),
                                launcherLogFontSizeCombo.getValue());
        launcherLogFontFamilyCombo
                .valueProperty()
                .addListener((obs, was, now) -> applyLauncherLogFont.run());
        launcherLogFontSizeCombo
                .valueProperty()
                .addListener((obs, was, now) -> applyLauncherLogFont.run());
        applyLauncherLogFont.run();

        Runnable applyLauncherLogFilter =
                () ->
                        applyLauncherLogSearchFilter(
                                launcherLogArea,
                                launcherLogMetaLabel,
                                launcherLogBaseMeta.get(),
                                launcherLogFullText.get(),
                                launcherLogSearchField.getText());
        launcherLogSearchField
                .textProperty()
                .addListener((obs, was, now) -> applyLauncherLogFilter.run());

        Button btnRefreshLauncherLog = new Button("ログ更新");
        styleSecondaryButton(btnRefreshLauncherLog);
        refreshLauncherLog[0] =
                () ->
                        scheduleLauncherLogRefresh(
                                uiEnv,
                                launcherLogFullText,
                                launcherLogBaseMeta,
                                applyLauncherLogFilter);
        btnRefreshLauncherLog.setOnAction(e -> refreshLauncherLog[0].run());

        HBox launcherLogToolbar =
                new HBox(
                        8,
                        new Label("検索"),
                        launcherLogSearchField,
                        new Label("フォント"),
                        launcherLogFontFamilyCombo,
                        new Label("サイズ"),
                        launcherLogFontSizeCombo,
                        btnRefreshLauncherLog);
        launcherLogToolbar.setAlignment(Pos.CENTER_LEFT);
        launcherLogSearchField.setMaxWidth(Double.MAX_VALUE);

        VBox launcherLogPane =
                new VBox(
                        10,
                        launcherLogTitle,
                        launcherLogMetaLabel,
                        launcherLogToolbar,
                        launcherLogArea);
        launcherLogPane.getStyleClass().add("pm-rdp-launcher-log-panel");
        launcherLogPane.setPadding(new Insets(16));
        launcherLogPane.setFillWidth(true);
        VBox.setVgrow(launcherLogPane, Priority.ALWAYS);

        SplitPane rightPaneSplit = new SplitPane(iniPreviewPane, launcherLogPane);
        rightPaneSplit.setOrientation(Orientation.VERTICAL);
        rightPaneSplit.setDividerPositions(0.58);
        SplitPane.setResizableWithParent(iniPreviewPane, Boolean.TRUE);
        SplitPane.setResizableWithParent(launcherLogPane, Boolean.TRUE);
        rightPaneSplit.getStyleClass().add("pm-remote-desktop-right-split");
        VBox.setVgrow(rightPaneSplit, Priority.ALWAYS);

        ScrollPane leftScroll = new ScrollPane(card);
        leftScroll.setFitToWidth(true);
        leftScroll.getStyleClass().add("pm-rdp-form-scroll");

        SplitPane splitPane = new SplitPane(leftScroll, rightPaneSplit);
        splitPane.setDividerPositions(0.52);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(rightPaneSplit, Boolean.TRUE);
        splitPane.getStyleClass().add("pm-remote-desktop-split");

        Runnable scheduleInitialRefresh =
                () ->
                        scheduleSharedSettingsRefresh(
                                uiEnv,
                                profileField,
                                profileChangeHandler,
                                companionProgramField,
                                companionArgsField,
                                refreshPaths,
                                deployStatusLabel,
                                status,
                                launchProfileCombo,
                                profileFields,
                                profileMetadataByNumber,
                                appendProfileRow,
                                rapStatusLabel,
                                applyDisplayFieldsFromUi,
                                chkDisconnectOnChildExit,
                                btnForceDeployLauncher,
                                btnLaunch,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                rdpSessionActive,
                                updateLaunchButtonState[0],
                                refreshAladdinCredentialsUi[0],
                                refreshIniFilePreview[0],
                                refreshLauncherLog[0],
                                refreshLaunchProfileCombo[0]);

        return new TabContent(splitPane, scheduleInitialRefresh);
    }

    private static void scheduleSharedSettingsRefresh(
            Supplier<Map<String, String>> uiEnv,
            TextField profileField,
            Consumer<String> profileChangeHandler,
            TextField companionProgramField,
            TextField companionArgsField,
            Runnable refreshPaths,
            Label deployStatusLabel,
            Consumer<String> status,
            ComboBox<Integer> launchProfileCombo,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber,
            IntConsumer appendProfileRow,
            Label rapStatusLabel,
            Runnable applyDisplayFieldsFromUi,
            CheckBox chkDisconnectOnChildExit,
            Button btnForceDeployLauncher,
            Button btnLaunch,
            AtomicBoolean launcherDeployInProgress,
            AtomicBoolean launcherDeployReady,
            AtomicBoolean rdpSessionActive,
            Runnable updateLaunchButtonState,
            Runnable refreshAladdinCredentialsUi,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshLaunchProfileCombo) {
        Platform.runLater(
                () -> {
                    if (launcherDeployInProgress.compareAndSet(false, true)) {
                        deployStatusLabel.setText("ランチャーを転送中…");
                        btnForceDeployLauncher.setDisable(true);
                        updateLaunchButtonState.run();
                    }
                });
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    RdpRemoteLauncherDeployer.DeployOutcome deploy =
                            RdpRemoteLauncherDeployer.ensureDeployed(ui, null);
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    Path profilesPath = AppPaths.resolveRdpLaunchProfilesFile(ui);
                    List<RdpLaunchProfile> catalogProfiles =
                            RdpLaunchProfileCatalog.loadOrDefaults(profilesPath);
                    RdpRemoteLauncherIni ini = null;
                    IOException iniError = null;
                    try {
                        ini = RdpRemoteLauncherIni.load(iniPath);
                    } catch (IOException ex) {
                        iniError = ex;
                    }
                    Optional<Path> configured = AppPaths.resolveRequestFormRdpProfile(ui);
                    Path preferredProfile = null;
                    if (configured.isPresent()) {
                        preferredProfile =
                                RdpFileSigner.resolvePreferredSignedProfilePath(configured.get(), ui);
                    }
                    final RdpRemoteLauncherIni loadedIni = ini;
                    final IOException loadError = iniError;
                    final Path resolvedIniPath = iniPath;
                    final Path resolvedProfilesPath = profilesPath;
                    final List<RdpLaunchProfile> resolvedCatalogProfiles = catalogProfiles;
                    final Path resolvedPreferred = preferredProfile;
                    final Optional<Path> resolvedConfigured = configured;
                    final Optional<String> deployMessage = deploy.message();
                    final boolean deploySucceeded = deploy.succeeded();
                    Platform.runLater(
                            () -> {
                                launcherDeployInProgress.set(false);
                                launcherDeployReady.set(deploySucceeded);
                                btnForceDeployLauncher.setDisable(false);
                                updateLaunchButtonState.run();
                                refreshPaths.run();
                                deployMessage.ifPresent(
                                        msg -> {
                                            deployStatusLabel.setText(msg);
                                            status.accept(msg);
                                        });
                                if (resolvedConfigured.isEmpty()) {
                                    profileField.setText("");
                                } else {
                                    Path preferred =
                                            resolvedPreferred != null
                                                    ? resolvedPreferred
                                                    : resolvedConfigured.get();
                                    profileField.setText(preferred.toString());
                                    if (!preferred.equals(resolvedConfigured.get())
                                            && profileChangeHandler != null) {
                                        profileChangeHandler.accept(preferred.toString());
                                    }
                                }
                                companionProgramField.setText(
                                        ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM, ""));
                                companionArgsField.setText(
                                        ui.getOrDefault(
                                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS, ""));
                                applyDisplayFieldsFromUi.run();
                                if (loadError != null) {
                                    rapStatusLabel.setText("読込失敗: " + loadError.getMessage());
                                    showAlert(Alert.AlertType.ERROR, "読込失敗", loadError.getMessage());
                                    if (refreshIniFilePreview != null || refreshLauncherLog != null) {
                                        refreshRightPanePreviews(
                                                refreshIniFilePreview, refreshLauncherLog);
                                    }
                                    return;
                                }
                                if (loadedIni == null) {
                                    return;
                                }
                                applyLoadedSettings(
                                        loadedIni,
                                        resolvedIniPath,
                                        resolvedProfilesPath,
                                        resolvedCatalogProfiles,
                                        launchProfileCombo,
                                        profileFields,
                                        profileMetadataByNumber,
                                        appendProfileRow,
                                        rapStatusLabel,
                                        chkDisconnectOnChildExit,
                                        uiEnv.get(),
                                        refreshLaunchProfileCombo);
                                if (refreshAladdinCredentialsUi != null) {
                                    refreshAladdinCredentialsUi.run();
                                }
                                refreshRightPanePreviews(refreshIniFilePreview, refreshLauncherLog);
                            });
                });
    }

    private static void scheduleIniReloadFromShare(
            Supplier<Map<String, String>> uiEnv,
            ComboBox<Integer> launchProfileCombo,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber,
            IntConsumer appendProfileRow,
            Label rapStatusLabel,
            Runnable refreshPaths,
            CheckBox chkDisconnectOnChildExit,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshLaunchProfileCombo) {
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    Path profilesPath = AppPaths.resolveRdpLaunchProfilesFile(ui);
                    RdpRemoteLauncherIni ini = null;
                    IOException iniError = null;
                    List<RdpLaunchProfile> catalogProfiles =
                            RdpLaunchProfileCatalog.loadOrDefaults(profilesPath);
                    try {
                        ini = RdpRemoteLauncherIni.load(iniPath);
                    } catch (IOException ex) {
                        iniError = ex;
                    }
                    final RdpRemoteLauncherIni loadedIni = ini;
                    final IOException loadError = iniError;
                    final Path resolvedIniPath = iniPath;
                    final Path resolvedProfilesPath = profilesPath;
                    final List<RdpLaunchProfile> resolvedCatalogProfiles = catalogProfiles;
                    Platform.runLater(
                            () -> {
                                refreshPaths.run();
                                if (loadError != null) {
                                    rapStatusLabel.setText("読込失敗: " + loadError.getMessage());
                                    showAlert(Alert.AlertType.ERROR, "読込失敗", loadError.getMessage());
                                    if (refreshIniFilePreview != null || refreshLauncherLog != null) {
                                        refreshRightPanePreviews(
                                                refreshIniFilePreview, refreshLauncherLog);
                                    }
                                    return;
                                }
                                applyLoadedSettings(
                                        loadedIni,
                                        resolvedIniPath,
                                        resolvedProfilesPath,
                                        resolvedCatalogProfiles,
                                        launchProfileCombo,
                                        profileFields,
                                        profileMetadataByNumber,
                                        appendProfileRow,
                                        rapStatusLabel,
                                        chkDisconnectOnChildExit,
                                        ui,
                                        refreshLaunchProfileCombo);
                                refreshRightPanePreviews(refreshIniFilePreview, refreshLauncherLog);
                            });
                });
    }

    private static void scheduleIniFilePreviewRefresh(
            Supplier<Map<String, String>> uiEnv,
            TextArea previewArea,
            Label metaLabel) {
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    String body;
                    String meta;
                    try {
                        if (!Files.isRegularFile(iniPath)) {
                            body = "";
                            meta = "ファイル未作成: " + iniPath;
                        } else {
                            body = Files.readString(iniPath, StandardCharsets.UTF_8);
                            var modified =
                                    Files.getLastModifiedTime(iniPath)
                                            .toInstant()
                                            .atZone(java.time.ZoneId.systemDefault())
                                            .toLocalDateTime()
                                            .format(
                                                    java.time.format.DateTimeFormatter.ofPattern(
                                                            "yyyy-MM-dd HH:mm:ss"));
                            meta =
                                    iniPath
                                            + " · "
                                            + Files.size(iniPath)
                                            + " bytes · 更新 "
                                            + modified;
                        }
                    } catch (IOException ex) {
                        body = "";
                        meta = "読込失敗: " + ex.getMessage() + " (" + iniPath + ")";
                    }
                    final String previewBody = body;
                    final String previewMeta = meta;
                    Platform.runLater(
                            () -> {
                                previewArea.setText(previewBody);
                                metaLabel.setText(previewMeta);
                            });
                });
    }

    private static void refreshRightPanePreviews(Runnable refreshIniFilePreview, Runnable refreshLauncherLog) {
        if (refreshIniFilePreview != null) {
            refreshIniFilePreview.run();
        }
        if (refreshLauncherLog != null) {
            refreshLauncherLog.run();
        }
    }

    private static void scheduleLauncherLogRefresh(
            Supplier<Map<String, String>> uiEnv,
            AtomicReference<String> fullTextHolder,
            AtomicReference<String> baseMetaHolder,
            Runnable applyFilter) {
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path logPath = AppPaths.resolveRdpLauncherSharedLogPath(ui);
                    String body;
                    String meta;
                    try {
                        if (!Files.isRegularFile(logPath)) {
                            body = "";
                            meta = "ファイル未作成: " + logPath;
                        } else {
                            body = readLauncherLogText(logPath);
                            var modified =
                                    Files.getLastModifiedTime(logPath)
                                            .toInstant()
                                            .atZone(java.time.ZoneId.systemDefault())
                                            .toLocalDateTime()
                                            .format(
                                                    java.time.format.DateTimeFormatter.ofPattern(
                                                            "yyyy-MM-dd HH:mm:ss"));
                            long lineCount = body.isEmpty() ? 0 : body.lines().count();
                            meta =
                                    logPath
                                            + " · "
                                            + Files.size(logPath)
                                            + " bytes · "
                                            + lineCount
                                            + " 行 · 更新 "
                                            + modified;
                        }
                    } catch (IOException ex) {
                        body = "";
                        meta = "読込失敗: " + ex.getMessage() + " (" + logPath + ")";
                    }
                    final String logBody = body;
                    final String logMeta = meta;
                    Platform.runLater(
                            () -> {
                                fullTextHolder.set(logBody);
                                baseMetaHolder.set(logMeta);
                                applyFilter.run();
                            });
                });
    }

    private static String readLauncherLogText(Path logPath) throws IOException {
        try {
            return Files.readString(logPath, StandardCharsets.UTF_8);
        } catch (MalformedInputException ex) {
            return Files.readString(logPath, Charset.forName("MS932"));
        }
    }

    private static void applyLauncherLogSearchFilter(
            TextArea area,
            Label metaLabel,
            String baseMeta,
            String fullText,
            String searchRaw) {
        String full = fullText != null ? fullText : "";
        String search = searchRaw != null ? searchRaw.trim() : "";
        String meta = baseMeta != null ? baseMeta : "";
        if (search.isEmpty()) {
            area.setText(full);
            metaLabel.setText(meta);
            return;
        }
        if (full.isEmpty()) {
            area.setText("");
            metaLabel.setText(meta + " · 検索「" + search + "」 0 / 0 行");
            return;
        }
        String searchLower = search.toLowerCase(Locale.ROOT);
        StringBuilder filtered = new StringBuilder();
        int matchCount = 0;
        int totalLines = 0;
        for (String line : full.split("\n", -1)) {
            if (line.isEmpty() && full.isEmpty()) {
                break;
            }
            totalLines++;
            if (line.toLowerCase(Locale.ROOT).contains(searchLower)) {
                if (matchCount > 0) {
                    filtered.append('\n');
                }
                filtered.append(line);
                matchCount++;
            }
        }
        area.setText(filtered.toString());
        metaLabel.setText(
                meta + " · 検索「" + search + "」 " + matchCount + " / " + totalLines + " 行");
    }

    private static void applyLauncherLogAreaFont(
            TextArea area, String familyChoice, Double sizePoints) {
        double size =
                sizePoints != null && sizePoints > 0 && Double.isFinite(sizePoints)
                        ? sizePoints
                        : Font.getDefault().getSize();
        if (familyChoice == null
                || familyChoice.equals(LAUNCHER_LOG_DEFAULT_FONT_FAMILY_LABEL)) {
            area.setStyle(
                    String.format(
                            Locale.ROOT,
                            "-fx-font-family: \"Consolas\", \"Cascadia Mono\", \"Courier New\", monospace; -fx-font-size: %.0fpx;",
                            size));
            return;
        }
        String escaped = familyChoice.replace("'", "\\'");
        area.setStyle(
                String.format(
                        Locale.ROOT,
                        "-fx-font-family: '%s'; -fx-font-size: %.0fpx;",
                        escaped,
                        size));
    }

    private static void applyLoadedSettings(
            RdpRemoteLauncherIni ini,
            Path iniPath,
            Path profilesPath,
            List<RdpLaunchProfile> catalogProfiles,
            ComboBox<Integer> launchProfileCombo,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber,
            IntConsumer appendProfileRow,
            Label rapStatusLabel,
            CheckBox chkDisconnectOnChildExit,
            Map<String, String> ui,
            Runnable refreshLaunchProfileCombo) {
        profileMetadataByNumber.clear();
        for (RdpLaunchProfile profile : catalogProfiles) {
            profileMetadataByNumber.put(profile.number(), profile);
        }

        int loadedSlot = ini.selectedSlot();
        int comboTarget;
        if (AppPaths.hasSavedRdpLaunchProfileNumber(ui)) {
            comboTarget = AppPaths.resolveRdpLaunchProfileNumber(ui);
        } else if (loadedSlot >= 1 && loadedSlot <= RdpRemoteLauncherIni.MAX_SLOTS) {
            comboTarget = loadedSlot;
        } else {
            comboTarget = 1;
        }
        chkDisconnectOnChildExit.setSelected(ini.disconnectOnChildExit());

        int visibleFromIni = ini.visibleSlotCount();
        int visibleFromCatalog =
                catalogProfiles.stream().mapToInt(RdpLaunchProfile::number).max().orElse(0);
        int visible =
                Math.min(
                        RdpRemoteLauncherIni.MAX_SLOTS,
                        Math.max(3, Math.max(visibleFromIni, visibleFromCatalog)));
        while (profileFields.size() < visible) {
            appendProfileRow.accept(profileFields.size() + 1);
        }

        for (ProfileRowFields row : profileFields) {
            RdpRemoteLauncherIni.Command command = ini.getSlotCommand(row.number());
            row.programField().setText(command.executable());
            String slotArguments = command.arguments();
            boolean eternalFromIni = RdpRemoteLauncherIni.hasEternalFlag(slotArguments);
            row.argsField()
                    .setText(
                            RdpRemoteLauncherIni.argumentsForUiDisplayWithoutManagedFlags(
                                    slotArguments));
            RdpLaunchProfile meta =
                    profileMetadataByNumber.getOrDefault(row.number(), RdpLaunchProfile.empty(row.number()));
            row.nameField().setText(meta.name());
            row.descriptionField().setText(meta.description());
            row.categoryField().setText(meta.category());
            boolean eternal =
                    meta.rpaEternal() != null ? meta.rpaEternal() : eternalFromIni;
            row.chkRpaEternal().setSelected(eternal);
        }

        if (refreshLaunchProfileCombo != null) {
            refreshLaunchProfileCombo.run();
            if (launchProfileCombo.getItems().contains(comboTarget)) {
                launchProfileCombo.setValue(comboTarget);
            }
        } else {
            launchProfileCombo.setValue(comboTarget);
        }

        if (loadedSlot == RdpRemoteLauncherIni.SLOT_DISABLED) {
            rapStatusLabel.setText(
                    "読込: タスクスケジューラ抑止中（"
                            + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                            + "=0）。"
                            + " 次回接続時にプロファイル "
                            + comboTarget
                            + " を ini へ書き込みます: "
                            + iniPath
                            + " / "
                            + profilesPath);
        } else {
            rapStatusLabel.setText("読込しました: " + iniPath + " / " + profilesPath);
        }
    }

    private static void scheduleTaskSchedulerSuppress(
            Supplier<Map<String, String>> uiEnv,
            Label rapStatusLabel,
            Consumer<String> status,
            Runnable reloadAfterSuppress) {
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = AppPaths.resolveRdpLauncherIni(ui);
                    IOException error = null;
                    try {
                        RdpRemoteLauncherIni.writeTaskSchedulerSuppress(iniPath, ui);
                    } catch (IOException ex) {
                        error = ex;
                    }
                    final IOException suppressError = error;
                    Platform.runLater(
                            () -> {
                                if (suppressError != null) {
                                    String msg =
                                            "起動プログラム番号の 0 設定に失敗: "
                                                    + suppressError.getMessage();
                                    rapStatusLabel.setText(msg);
                                    status.accept(msg);
                                    showAlert(Alert.AlertType.ERROR, "抑止失敗", msg);
                                    return;
                                }
                                String msg =
                                        RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                                + "="
                                                + RdpRemoteLauncherIni.SLOT_DISABLED
                                                + " を設定しました（タスクスケジューラ抑止）: "
                                                + iniPath;
                                rapStatusLabel.setText(msg);
                                status.accept(msg);
                                if (reloadAfterSuppress != null) {
                                    reloadAfterSuppress.run();
                                }
                            });
                });
    }

    private static void runBackgroundThenFx(Runnable backgroundWork) {
        Thread worker = new Thread(backgroundWork, "request-form-rdp-io");
        worker.setDaemon(true);
        worker.start();
    }

    private static void scheduleLauncherDeploy(
            Supplier<Map<String, String>> uiEnv,
            boolean force,
            Label deployStatusLabel,
            Button btnForceDeployLauncher,
            AtomicBoolean deployInProgress,
            AtomicBoolean deployReady,
            Runnable updateLaunchButtonState,
            Consumer<String> status,
            Runnable whenDone) {
        if (!deployInProgress.compareAndSet(false, true)) {
            Platform.runLater(
                    () -> deployStatusLabel.setText("ランチャー転送が既に実行中です。"));
            return;
        }
        Platform.runLater(
                () -> {
                    deployStatusLabel.setText("ランチャーを転送中…");
                    btnForceDeployLauncher.setDisable(true);
                    updateLaunchButtonState.run();
                });
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    RdpRemoteLauncherDeployer.DeployOutcome outcome =
                            force
                                    ? RdpRemoteLauncherDeployer.forceDeploy(ui)
                                    : RdpRemoteLauncherDeployer.ensureDeployed(ui);
                    Platform.runLater(
                            () -> {
                                deployInProgress.set(false);
                                deployReady.set(outcome.succeeded());
                                outcome.message().ifPresent(deployStatusLabel::setText);
                                outcome.message().ifPresent(status::accept);
                                btnForceDeployLauncher.setDisable(false);
                                updateLaunchButtonState.run();
                                if (whenDone != null) {
                                    whenDone.run();
                                }
                            });
                });
    }

    private static void applyRdpLaunchButtonState(
            Supplier<Map<String, String>> uiEnv,
            Button btnLaunch,
            AtomicBoolean deployInProgress,
            AtomicBoolean deployReady,
            AtomicBoolean sessionActive) {
        btnLaunch
                .getStyleClass()
                .removeAll(
                        "pm-rdp-connect-pending",
                        "pm-rdp-connect-deploying",
                        "pm-rdp-connect-active");
        if (deployInProgress.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-deploying");
            btnLaunch.setDisable(true);
            btnLaunch.setText("ランチャー転送中…");
            return;
        }
        if (sessionActive.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-active");
            btnLaunch.setDisable(true);
            btnLaunch.setText("リモートデスクトップ接続中");
            return;
        }
        if (!deployReady.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-pending");
            btnLaunch.setDisable(false);
            btnLaunch.setText("ランチャー転送してから起動");
            return;
        }
        if (!rdpLaunchCredentialsReady(uiEnv)) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-pending");
            btnLaunch.setDisable(true);
            btnLaunch.setText("アラジン資格情報を保存してください");
            return;
        }
        btnLaunch.setDisable(false);
        btnLaunch.setText("リモートデスクトップを起動");
    }

    private static boolean rdpLaunchCredentialsReady(Supplier<Map<String, String>> uiEnv) {
        Map<String, String> ui = uiEnv != null ? uiEnv.get() : Map.of();
        FactorySite site = GlobalInitSettingTarget.loadEffective(ui);
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank() || FactoryOperatorUserStore.isGuestOperator(operator)) {
            return false;
        }
        try {
            FactoryOperatorUserStore.configureFromUi(ui, site);
            return FactoryOperatorUserStore.hasAladdinCredentials(site, operator);
        } catch (IOException ex) {
            return false;
        }
    }

    private static String formatRdpSessionEndMessage(RdpMstscSessionMonitor.SessionEndEvent event) {
        return switch (event.reason()) {
            case MSTSC_EXIT ->
                    "リモートデスクトップ接続が終了しました（接続先 RPA 完了後の RDP 切断を含む）。";
            case PROCESS_NOT_FOUND ->
                    "mstsc プロセスを特定できなかったため、接続終了を監視できませんでした: "
                            + event.rdpProfile();
            case MONITOR_INTERRUPTED -> "リモートデスクトップ接続の監視が中断されました。";
        };
    }

    private static void updateDisplayControlsEnabled(
            CheckBox chkFullScreen,
            Spinner<Integer> desktopWidthSpinner,
            Spinner<Integer> desktopHeightSpinner) {
        boolean windowed = !chkFullScreen.isSelected();
        desktopWidthSpinner.setDisable(!windowed);
        desktopHeightSpinner.setDisable(!windowed);
    }

    private static void pinButtonPrefWidth(Button button) {
        button.setMinWidth(Region.USE_PREF_SIZE);
        button.setMaxWidth(Region.USE_PREF_SIZE);
    }

    private static void styleSecondaryButton(Button button) {
        button.getStyleClass().add("pm-rdp-secondary-button");
        pinButtonPrefWidth(button);
    }

    private static Label fieldCaption(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("pm-rdp-field-caption");
        return label;
    }

    private static ListCell<Integer> createLaunchProfileListCell(
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber) {
        return new ListCell<>() {
            @Override
            protected void updateItem(Integer item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                    return;
                }
                ProfileRowFields row = findProfileRow(profileFields, item);
                RdpLaunchProfile meta =
                        profileMetadataByNumber.getOrDefault(
                                item, row != null ? buildProfileFromRow(row) : RdpLaunchProfile.empty(item));
                setText(row != null ? profileComboLabel(row, meta) : meta.displayLabel());
            }
        };
    }

    private static void updateLaunchProfileDetail(
            Label detailLabel,
            Integer profileNumber,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber) {
        if (profileNumber == null) {
            detailLabel.setText("");
            return;
        }
        ProfileRowFields row = findProfileRow(profileFields, profileNumber);
        RdpLaunchProfile meta =
                profileMetadataByNumber.getOrDefault(
                        profileNumber,
                        row != null ? buildProfileFromRow(row) : RdpLaunchProfile.empty(profileNumber));
        String detail = meta.detailText();
        if (detail.isBlank() && row != null && !row.programField().getText().isBlank()) {
            detail = "RPA: " + row.programField().getText().trim();
        }
        detailLabel.setText(detail.isBlank() ? "（説明未設定）" : detail);
    }

    private static ProfileRowFields findProfileRow(
            List<ProfileRowFields> profileFields, int number) {
        for (ProfileRowFields row : profileFields) {
            if (row.number() == number) {
                return row;
            }
        }
        return null;
    }

    private static String profileComboLabel(ProfileRowFields row, RdpLaunchProfile meta) {
        String name = row.nameField().getText().trim();
        if (name.isEmpty()) {
            name = meta.name().isBlank() ? "（名称未設定）" : meta.name();
        }
        return row.number() + ": " + name;
    }

    private static RdpLaunchProfile buildProfileFromRow(ProfileRowFields row) {
        return new RdpLaunchProfile(
                row.number(),
                row.nameField().getText(),
                row.descriptionField().getText(),
                row.categoryField().getText(),
                null,
                null,
                null,
                null,
                row.chkRpaEternal().isSelected());
    }

    private static List<RdpLaunchProfile> collectProfilesFromRows(List<ProfileRowFields> rows) {
        List<RdpLaunchProfile> profiles = new ArrayList<>();
        for (ProfileRowFields row : rows) {
            profiles.add(buildProfileFromRow(row));
        }
        return profiles;
    }

    private static void applyProfileLaunchOverrides(
            Map<String, String> ui, RdpLaunchProfile profile) {
        if (profile.fullScreen() != null) {
            ui.put(AppPaths.KEY_PM_AI_RDP_FULLSCREEN, profile.fullScreen() ? "1" : "0");
        }
        if (profile.desktopWidth() != null) {
            ui.put(
                    AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH,
                    String.valueOf(profile.desktopWidth()));
        }
        if (profile.desktopHeight() != null) {
            ui.put(
                    AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT,
                    String.valueOf(profile.desktopHeight()));
        }
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
        alert.setHeaderText(title);
        Label body = new Label(message);
        body.setWrapText(true);
        body.setMaxWidth(520);
        body.setMinWidth(360);
        alert.getDialogPane().setContent(body);
        alert.showAndWait();
    }
}
