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
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.ListCell;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.util.Duration;
import javafx.util.StringConverter;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.RemoteDesktopStandaloneBootstrap;
import jp.co.pm.ai.desktop.config.AladdinRpaLaunchArgs;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.RdpFileSigner;
import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings;
import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings.LaunchDisplay;
import jp.co.pm.ai.desktop.io.RdpLaunchProfile;
import jp.co.pm.ai.desktop.io.RdpLaunchProfileCatalog;
import jp.co.pm.ai.desktop.io.RdpLaunchProfileQuickLaunch;
import jp.co.pm.ai.desktop.io.RdpLaunchProfileSequentialRun;
import jp.co.pm.ai.desktop.io.RdpMstscCloser;
import jp.co.pm.ai.desktop.io.RdpMstscSessionMonitor;
import jp.co.pm.ai.desktop.io.RdpPreviewSettings;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherDeployer;
import jp.co.pm.ai.desktop.io.RdpLaunchSharedPathValidator;
import jp.co.pm.ai.desktop.io.RdpRemoteLauncherIni;
import jp.co.pm.ai.desktop.io.RdpSessionEndAction;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * メインシェル「リモートデスクトップ」タブの UI（RAP 設定・RDP 起動）。
 */
public final class RequestFormRemoteDesktopPane {

    private record ProfileRowFields(
            int number,
            VBox card,
            TextField nameField,
            TextField descriptionField,
            TextField categoryField,
            TextField programField,
            TextField argsField,
            CheckBox chkRpaEternal,
            CheckBox chkFullScreen,
            Spinner<Integer> desktopWidthSpinner,
            Spinner<Integer> desktopHeightSpinner) {}

    private enum RdpConnectPhase {
        IDLE,
        REMOTE_SIGN_OUT,
        RPA_SESSION
    }

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Consumer<String> profileChangeHandler,
            Consumer<String> companionProgramChangeHandler,
            Consumer<String> companionProgramArgsChangeHandler,
            Consumer<String> previewInTabChangeHandler,
            IntConsumer launchProfileNumberChangeHandler,
            Consumer<String> launcherDeployDirChangeHandler,
            Consumer<String> statusConsumer) {}

    private static final double CARD_WIDTH = 720;

    private static final String LAUNCHER_LOG_DEFAULT_FONT_FAMILY_LABEL = "システム既定";

    private static final List<Double> LAUNCHER_LOG_PRESET_FONT_SIZES =
            List.of(9d, 10d, 11d, 12d, 13d, 14d, 15d, 16d, 18d, 20d, 22d, 24d);

    /** 右ペイン上下分割の初期比率（上 = ini プレビュー）。 */
    private static final double RIGHT_PANE_TOP_RATIO_INITIAL = 1.0 / 3.0;

    private static final double RIGHT_PANE_MIN_HEIGHT = 72;

    /** 「取得データ最新ファイル」自動更新の間隔（リモートデスクトップタブを開いている間のみ動作）。 */
    private static final Duration FETCHED_FILES_AUTO_REFRESH_INTERVAL = Duration.seconds(5);

    /**
     * 構築結果。{@link #scheduleInitialRefresh()} はタブ初回表示時のみ呼ぶ（UNC I/O を初回マウントから分離）。
     * {@link #setTabActive(boolean)} はメインタブの選択・非選択に合わせて呼び、
     * リモートデスクトップタブを開いている間だけ「取得データ最新ファイル」を5秒おきに自動更新する。
     */
    public record TabContent(
            SplitPane root,
            Runnable scheduleInitialRefresh,
            Runnable onSessionOperatorChanged,
            Consumer<Boolean> setTabActive) {}

    private RequestFormRemoteDesktopPane() {}

    private static Path resolveRdpIniPath(Map<String, String> ui) {
        return AppPaths.resolveRdpLauncherIni(
                ui, FactoryOperatorUserStore.resolveRdpLauncherOperatorName(ui));
    }

    private static Path resolveExistingRdpIniPath(Map<String, String> ui) {
        return AppPaths.resolveExistingRdpLauncherIni(
                ui, FactoryOperatorUserStore.resolveRdpLauncherOperatorName(ui));
    }

    public static TabContent buildTabContent(Window owner, Context ctx) {
        Supplier<Map<String, String>> uiEnv = ctx.uiEnv() != null ? ctx.uiEnv() : () -> Map.of();
        Consumer<String> profileChangeHandler = ctx.profileChangeHandler();
        Consumer<String> companionProgramChangeHandler = ctx.companionProgramChangeHandler();
        Consumer<String> companionProgramArgsChangeHandler = ctx.companionProgramArgsChangeHandler();
        Consumer<String> previewInTabChangeHandler = ctx.previewInTabChangeHandler();
        IntConsumer launchProfileNumberChangeHandler = ctx.launchProfileNumberChangeHandler();
        Consumer<String> launcherDeployDirChangeHandler = ctx.launcherDeployDirChangeHandler();
        Consumer<String> status = ctx.statusConsumer() != null ? ctx.statusConsumer() : s -> {};
        AtomicReference<Thread> rdpSessionWatchThread = new AtomicReference<>();
        AtomicBoolean launcherDeployInProgress = new AtomicBoolean(false);
        AtomicBoolean launcherDeployReady = new AtomicBoolean(false);
        AtomicReference<Runnable> pendingLaunchAfterDeploy = new AtomicReference<>();
        AtomicBoolean rdpSessionActive = new AtomicBoolean(false);
        AtomicReference<RdpConnectPhase> rdpConnectPhase =
                new AtomicReference<>(RdpConnectPhase.IDLE);
        AtomicBoolean sequentialRunActive = new AtomicBoolean(false);
        AtomicReference<List<Integer>> sequentialQueue = new AtomicReference<>(List.of());
        AtomicInteger sequentialIndex = new AtomicInteger(0);
        List<Integer> sequentialSelectionOrder = new ArrayList<>();
        Runnable[] refreshDisplayPreview = new Runnable[1];
        Runnable[] refreshFetchedFilesTable = new Runnable[1];
        Runnable[] scrollToFetchedFilesSection = new Runnable[1];
        final RdpRightPanePreviewController[] rightPanePreviewHolder = new RdpRightPanePreviewController[1];

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

        Button btnJumpToFetchedFiles = new Button("取得ファイル一覧へ ↓");
        styleSecondaryButton(btnJumpToFetchedFiles);
        btnJumpToFetchedFiles.setTooltip(
                new Tooltip("RPA 取得先フォルダの最新ファイル一覧（ページ最下部）へスクロールします。"));

        VBox headerBlock = new VBox(6, title, subtitle, btnJumpToFetchedFiles);
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
                                + " を配備先フォルダへ上書きコピーします。"));

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

        List<Button> quickLaunchProfileButtons = new ArrayList<>();
        GridPane quickLaunchProfileGrid = new GridPane();
        quickLaunchProfileGrid.setHgap(8);
        quickLaunchProfileGrid.setVgap(8);
        quickLaunchProfileGrid.setMaxWidth(CARD_WIDTH);
        for (int slot = 0; slot < RdpLaunchProfileQuickLaunch.BUTTON_SLOT_COUNT; slot++) {
            Button quickBtn = new Button();
            quickBtn.getStyleClass().add("pm-rdp-quick-launch-button");
            quickBtn.setMaxWidth(Double.MAX_VALUE);
            quickBtn.setWrapText(true);
            quickBtn.setVisible(false);
            quickBtn.setManaged(false);
            quickLaunchProfileButtons.add(quickBtn);
            quickLaunchProfileGrid.add(quickBtn, slot % 2, slot / 2);
            GridPane.setHgrow(quickBtn, Priority.ALWAYS);
        }

        CheckBox chkSequentialMode = new CheckBox("連続実行モード");
        chkSequentialMode.setTooltip(
                new Tooltip(
                        "オンにするとクイック起動ボタンで実行順を選び、"
                                + "接続が終了するたびに次のタスクを自動起動します。"
                                + " プロファイル 99（接続先サインアウト）は先頭のみ選択できます。"));
        Button btnClearSequentialSelection = new Button("選択をクリア");
        styleSecondaryButton(btnClearSequentialSelection);
        btnClearSequentialSelection.setDisable(true);
        btnClearSequentialSelection.setTooltip(
                new Tooltip("連続実行の選択順をリセットします。"));
        Label sequentialStatusLabel = new Label();
        sequentialStatusLabel.getStyleClass().add("pm-rdp-page-subtitle");
        sequentialStatusLabel.setWrapText(true);
        sequentialStatusLabel.setMaxWidth(CARD_WIDTH);
        HBox sequentialModeRow =
                new HBox(8, chkSequentialMode, btnClearSequentialSelection);
        sequentialModeRow.setAlignment(Pos.CENTER_LEFT);
        sequentialModeRow.setMaxWidth(CARD_WIDTH);

        Runnable[] updateLaunchButtonState = new Runnable[1];

        Runnable[] refreshQuickLaunchSequentialLabels = new Runnable[1];
        refreshQuickLaunchSequentialLabels[0] =
                () -> {
                    for (Button quickBtn : quickLaunchProfileButtons) {
                        if (!quickBtn.isVisible()) {
                            continue;
                        }
                        Object baseObj = quickBtn.getProperties().get("pmQuickBaseLabel");
                        String baseLabel = baseObj instanceof String s ? s : quickBtn.getText();
                        Integer profileNumber = (Integer) quickBtn.getUserData();
                        int orderIdx =
                                profileNumber != null
                                        ? RdpLaunchProfileSequentialRun.selectionOrderIndex(
                                                sequentialSelectionOrder, profileNumber)
                                        : -1;
                        quickBtn.setText(
                                RdpLaunchProfileSequentialRun.quickButtonLabel(
                                        baseLabel, orderIdx));
                        quickBtn.getStyleClass().remove("pm-rdp-quick-launch-selected");
                        if (orderIdx > 0) {
                            quickBtn.getStyleClass().add("pm-rdp-quick-launch-selected");
                        }
                    }
                    btnClearSequentialSelection.setDisable(sequentialSelectionOrder.isEmpty());
                    if (!sequentialRunActive.get()) {
                        if (chkSequentialMode.isSelected() && !sequentialSelectionOrder.isEmpty()) {
                            sequentialStatusLabel.setText(
                                    "実行順: "
                                            + sequentialSelectionOrder.stream()
                                                    .map(
                                                            n ->
                                                                    RdpLaunchProfileSequentialRun
                                                                            .selectionOrderMarker(
                                                                                    sequentialSelectionOrder
                                                                                            .indexOf(
                                                                                                    n)
                                                                                            + 1)
                                                                            + " プロファイル "
                                                                            + n)
                                                    .reduce((a, b) -> a + " → " + b)
                                                    .orElse(""));
                        } else if (!chkSequentialMode.isSelected()) {
                            sequentialStatusLabel.setText("");
                        }
                    }
                    if (updateLaunchButtonState[0] != null) {
                        updateLaunchButtonState[0].run();
                    }
                };

        Runnable[] updateProfileManageButtons = new Runnable[1];

        Runnable[] refreshLaunchProfileCombo = new Runnable[1];
        refreshLaunchProfileCombo[0] =
                () -> {
                    Integer previous = launchProfileCombo.getValue();
                    launchProfileCombo.getItems().clear();
                    List<Integer> catalogOrder = new ArrayList<>();
                    for (ProfileRowFields row : profileFields) {
                        if (isProfileDeleted(row, profileMetadataByNumber)) {
                            continue;
                        }
                        launchProfileCombo.getItems().add(row.number());
                        catalogOrder.add(row.number());
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

                    List<Integer> quickNumbers =
                            RdpLaunchProfileQuickLaunch.quickLaunchProfileNumbers(catalogOrder);
                    for (int slot = 0; slot < quickLaunchProfileButtons.size(); slot++) {
                        Button quickBtn = quickLaunchProfileButtons.get(slot);
                        if (slot < quickNumbers.size()) {
                            int profileNumber = quickNumbers.get(slot);
                            ProfileRowFields row = findProfileRow(profileFields, profileNumber);
                            RdpLaunchProfile meta =
                                    row != null
                                            ? buildProfileFromRow(row)
                                            : profileMetadataByNumber.getOrDefault(
                                                    profileNumber,
                                                    RdpLaunchProfile.empty(profileNumber));
                            String label =
                                    row != null
                                            ? profileComboLabel(row, meta)
                                            : meta.displayLabel();
                            quickBtn.setText(RdpLaunchProfileQuickLaunch.buttonLabel(label));
                            quickBtn.getProperties().put("pmQuickBaseLabel", quickBtn.getText());
                            quickBtn.setUserData(profileNumber);
                            quickBtn.setTooltip(new Tooltip(label));
                            quickBtn.setVisible(true);
                            quickBtn.setManaged(true);
                        } else {
                            quickBtn.setUserData(null);
                            quickBtn.setTooltip(null);
                            quickBtn.setVisible(false);
                            quickBtn.setManaged(false);
                        }
                    }
                    if (updateLaunchButtonState[0] != null) {
                        updateLaunchButtonState[0].run();
                    }
                    if (updateProfileManageButtons[0] != null) {
                        updateProfileManageButtons[0].run();
                    }
                    if (refreshQuickLaunchSequentialLabels[0] != null) {
                        refreshQuickLaunchSequentialLabels[0].run();
                    }
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
                            if (refreshDisplayPreview[0] != null) {
                                refreshDisplayPreview[0].run();
                            }
                        });

        ComboBox<RdpSessionEndAction> cmbSessionEndAction =
                new ComboBox<>(
                        FXCollections.observableArrayList(RdpSessionEndAction.values()));
        cmbSessionEndAction.setValue(RdpSessionEndAction.SIGN_OUT);
        cmbSessionEndAction.setMaxWidth(CARD_WIDTH);
        cmbSessionEndAction.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(RdpSessionEndAction action) {
                        return action == null ? "" : action.displayLabel();
                    }

                    @Override
                    public RdpSessionEndAction fromString(String string) {
                        return RdpSessionEndAction.fromIniValue(string, RdpSessionEndAction.SIGN_OUT);
                    }
                });
        cmbSessionEndAction.setTooltip(
                new Tooltip(
                        "接続先ランチャーが起動したプログラムの終了を待ち、終了後の RDP セッション操作を選びます。"
                                + " 「なし」は操作しません。「切断」はセッションを残して接続のみ終了します。"
                                + " 「サインアウト」はセッションを終了します。"
                                + " RPA設定.ini の「"
                                + RdpRemoteLauncherIni.SESSION_END_ACTION_KEY
                                + "」に保存されます。"));
        Label sessionEndActionLabel = new Label("子プロセス終了時のセッション操作");
        HBox sessionEndActionRow =
                new HBox(8, sessionEndActionLabel, cmbSessionEndAction);
        sessionEndActionRow.setAlignment(Pos.CENTER_LEFT);

        Label deployDirSectionTitle = new Label("ランチャー配備先フォルダ");
        deployDirSectionTitle.getStyleClass().add("pm-rdp-section-title");

        Label deployDirWarningLabel = new Label();
        deployDirWarningLabel.getStyleClass().add("pm-rdp-status-label");
        deployDirWarningLabel.setWrapText(true);
        deployDirWarningLabel.setMaxWidth(CARD_WIDTH);

        TextField deployDirField = new TextField();
        deployDirField.setPromptText("共有フォルダ（UNC 可）");
        deployDirField.setMaxWidth(CARD_WIDTH);

        Runnable refreshPaths =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path launcherExe = AppPaths.resolveRdpLauncherExe(ui);
                    Path deployDir = AppPaths.resolveRdpLauncherDeployDir(ui);
                    deployDirField.setText(
                            ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR, deployDir.toString()));
                    if (RemoteDesktopStandaloneBootstrap.isActivated()
                            && ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR, "").isBlank()) {
                        deployDirWarningLabel.setText("配備先フォルダを指定してください（参照…から共有フォルダを選ぶ）。");
                    } else {
                        deployDirWarningLabel.setText("");
                    }
                    iniPathLabel.setText(
                            "RPA設定.ini: " + resolveRdpIniPath(ui).toString());
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

        Button btnBrowseDeployDir = new Button("参照…");
        styleSecondaryButton(btnBrowseDeployDir);
        btnBrowseDeployDir.setOnAction(
                e -> {
                    DirectoryChooser chooser = new DirectoryChooser();
                    chooser.setTitle("ランチャー配備先フォルダ");
                    Map<String, String> ui = uiEnv.get();
                    Path current = AppPaths.resolveRdpLauncherDeployDir(ui);
                    if (Files.isDirectory(current)) {
                        chooser.setInitialDirectory(current.toFile());
                    }
                    java.io.File chosen =
                            chooser.showDialog(
                                    owner != null ? owner : btnBrowseDeployDir.getScene().getWindow());
                    if (chosen != null) {
                        String path = chosen.toPath().toAbsolutePath().normalize().toString();
                        deployDirField.setText(path);
                        if (launcherDeployDirChangeHandler != null) {
                            launcherDeployDirChangeHandler.accept(path);
                        }
                        refreshPaths.run();
                        status.accept("ランチャー配備先を更新: " + path);
                    }
                });

        Button btnOpenDeployDir = new Button("配備先を開く");
        styleSecondaryButton(btnOpenDeployDir);
        btnOpenDeployDir.setOnAction(
                e -> {
                    Map<String, String> ui = uiEnv.get();
                    Path dir = AppPaths.resolveRdpLauncherDeployDir(ui);
                    if (!Files.isDirectory(dir)) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setTitle("配備先フォルダ");
                        alert.setHeaderText(null);
                        alert.setContentText("フォルダが存在しません: " + dir);
                        alert.showAndWait();
                        return;
                    }
                    openFolder(dir);
                });

        HBox deployDirActionRow = new HBox(8, btnBrowseDeployDir, btnOpenDeployDir);
        deployDirActionRow.setAlignment(Pos.CENTER_LEFT);
        deployDirActionRow.setMaxWidth(CARD_WIDTH);

        VBox deployDirBlock =
                new VBox(
                        4,
                        deployDirSectionTitle,
                        deployDirField,
                        deployDirActionRow,
                        deployDirWarningLabel);
        deployDirBlock.setFillWidth(true);
        deployDirBlock.setMaxWidth(CARD_WIDTH);
        refreshPaths.run();

        java.util.function.IntConsumer appendProfileRow =
                profileNumber -> {
                    final boolean signOutOnly =
                            RdpRemoteLauncherIni.isSignOutOnlyProfile(profileNumber);
                    Label header =
                            new Label(
                                    signOutOnly
                                            ? "プロファイル 99（接続先サインアウト）"
                                            : "プロファイル " + profileNumber);
                    header.getStyleClass().add("pm-rdp-profile-card-title");

                    Region headerSpacer = new Region();
                    HBox.setHgrow(headerSpacer, Priority.ALWAYS);

                    Button btnCopyToEmpty = new Button("空きへコピー");
                    styleSecondaryButton(btnCopyToEmpty);
                    btnCopyToEmpty.setTooltip(
                            new Tooltip(
                                    "このプロファイルの内容を、空の別プロファイルへコピーします。"
                                            + " コピー先が空でない場合は上書きしません。"));
                    btnCopyToEmpty.setOnAction(
                            e ->
                                    copyProfileToEmptySlot(
                                            owner != null
                                                    ? owner
                                                    : btnCopyToEmpty.getScene().getWindow(),
                                            profileNumber,
                                            profileFields,
                                            refreshLaunchProfileCombo[0]));

                    Button btnClearProfile = new Button("クリア");
                    styleSecondaryButton(btnClearProfile);
                    btnClearProfile.setTooltip(
                            new Tooltip("このプロファイルの名称・RPA 設定などをすべて空にします。"));
                    btnClearProfile.setOnAction(
                            e ->
                                    confirmAndClearProfile(
                                            owner != null
                                                    ? owner
                                                    : btnClearProfile.getScene().getWindow(),
                                            profileNumber,
                                            profileFields,
                                            refreshLaunchProfileCombo[0]));

                    if (signOutOnly) {
                        btnCopyToEmpty.setVisible(false);
                        btnCopyToEmpty.setManaged(false);
                        btnClearProfile.setVisible(false);
                        btnClearProfile.setManaged(false);
                    }

                    HBox headerRow =
                            new HBox(8, header, headerSpacer, btnCopyToEmpty, btnClearProfile);
                    headerRow.setAlignment(Pos.CENTER_LEFT);
                    headerRow.setMaxWidth(CARD_WIDTH);

                    TextField nameField = new TextField();
                    nameField.setPromptText(
                            signOutOnly ? "例: 接続先サインアウトのみ" : "例: アラジン 工程マスタ取得");
                    HBox.setHgrow(nameField, Priority.ALWAYS);

                    TextField descriptionField = new TextField();
                    descriptionField.setPromptText("この接続で行う作業の説明（RPA シナリオは接続先で別設定）");
                    HBox.setHgrow(descriptionField, Priority.ALWAYS);

                    TextField categoryField = new TextField();
                    categoryField.setPromptText("例: マスタ更新 / データ取得");
                    categoryField.setPrefWidth(180);

                    TextField programField = new TextField();
                    programField.setPromptText(
                            signOutOnly
                                    ? "接続先から参照できる PmAiRdpRemoteLauncher.exe（例: M:\\...）"
                                    : "接続先の RPA exe パス（RPA設定.ini スロット "
                                            + profileNumber
                                            + "）");
                    HBox.setHgrow(programField, Priority.ALWAYS);

                    TextField argsField = new TextField();
                    argsField.setPromptText(
                            signOutOnly
                                    ? "ランチャー引数（空ならツールバーの操作者名のみ）"
                                    : "任意: --scenario path.ardrpa（空なら Aladdin 本体のみ起動）");
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

                    Button checkArgs = new Button("存在確認");
                    styleSecondaryButton(checkArgs);
                    checkArgs.setTooltip(
                            new Tooltip(
                                    "RPA 引数内の .ardrpa シナリオパスを確認します。"
                                            + " 002  加工G の空白潰れは修復してから UNC を参照します。"));
                    checkArgs.setOnAction(
                            e ->
                                    showExistenceCheckResult(
                                            "RPA シナリオ",
                                            RdpLaunchSharedPathValidator.validateScenarioArguments(
                                                    argsField.getText()),
                                            rapStatusLabel,
                                            status));

                    HBox argsRow = new HBox(8, argsField, browseArgs, checkArgs);
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

                    CheckBox chkProfileFullScreen = new CheckBox("全画面で接続");
                    chkProfileFullScreen.setTooltip(
                            new Tooltip("オフのときは下記の幅・高さでウィンドウ表示します。"));
                    Spinner<Integer> profileWidthSpinner = createDesktopWidthSpinner(RdpLaunchDisplaySettings.DEFAULT_WIDTH);
                    Spinner<Integer> profileHeightSpinner = createDesktopHeightSpinner(RdpLaunchDisplaySettings.DEFAULT_HEIGHT);
                    chkProfileFullScreen
                            .selectedProperty()
                            .addListener(
                                    (obs, was, selected) ->
                                            updateDisplayControlsEnabled(
                                                    chkProfileFullScreen,
                                                    profileWidthSpinner,
                                                    profileHeightSpinner));

                    HBox profileDisplayRow =
                            new HBox(
                                    8,
                                    new Label("幅:"),
                                    profileWidthSpinner,
                                    new Label("高さ:"),
                                    profileHeightSpinner);
                    profileDisplayRow.setAlignment(Pos.CENTER_LEFT);

                    Button checkProgram = new Button("存在確認");
                    styleSecondaryButton(checkProgram);
                    checkProgram.setTooltip(
                            new Tooltip(
                                    "入力した RPA プログラムのパスを確認します。"
                                            + " 002  加工G の空白潰れは修復してから UNC を参照します。"));
                    checkProgram.setOnAction(
                            e ->
                                    showExistenceCheckResult(
                                            "RPA プログラム",
                                            RdpLaunchSharedPathValidator.validateProgramPath(
                                                    programField.getText()),
                                            rapStatusLabel,
                                            status));

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

                    HBox programRow = new HBox(8, programField, browse, checkProgram);
                    programRow.setAlignment(Pos.CENTER_LEFT);
                    programRow.setMaxWidth(CARD_WIDTH);
                    HBox.setHgrow(programField, Priority.ALWAYS);

                    HBox categoryRow = new HBox(8, fieldCaption("区分"), categoryField);
                    categoryRow.setAlignment(Pos.CENTER_LEFT);

                    VBox cardInner = new VBox(8);
                    cardInner.getChildren()
                            .addAll(
                                    headerRow,
                                    fieldCaption("名称"),
                                    nameField,
                                    fieldCaption("説明"),
                                    descriptionField);
                    if (!signOutOnly) {
                        cardInner.getChildren().addAll(categoryRow, fieldCaption("表示"));
                        cardInner.getChildren().addAll(chkProfileFullScreen, profileDisplayRow);
                        cardInner.getChildren()
                                .addAll(
                                        fieldCaption("RPA プログラム"),
                                        programRow,
                                        fieldCaption("RPA 引数"),
                                        argsRow,
                                        chkRpaEternal);
                    } else {
                        Label signOutInfo =
                                new Label(
                                        "通常 mstsc で接続し、接続先タスクスケジューラが "
                                                + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                                                + " 操作者名 を起動したとき、ini の "
                                                + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                                + "=99 とスロット 99="
                                                + RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS
                                                + " でサインアウトします。alternate shell は使いません。");
                        signOutInfo.setWrapText(true);
                        signOutInfo.setMaxWidth(CARD_WIDTH);
                        cardInner.getChildren().add(signOutInfo);
                    }
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
                                    cardInner,
                                    nameField,
                                    descriptionField,
                                    categoryField,
                                    programField,
                                    argsField,
                                    chkRpaEternal,
                                    chkProfileFullScreen,
                                    profileWidthSpinner,
                                    profileHeightSpinner));
                    wireProfileDisplayPreviewRefresh(
                            profileFields.get(profileFields.size() - 1), refreshDisplayPreview);
                };

        appendProfileRow.accept(RdpRemoteLauncherIni.SLOT_SIGN_OUT);
        for (int i = 1; i <= RdpRemoteLauncherIni.DEFAULT_INITIAL_RPA_PROFILE_ROWS; i++) {
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
                                cmbSessionEndAction,
                                refreshIniFilePreview[0],
                                refreshLauncherLog[0],
                                refreshLaunchProfileCombo[0]);

        Runnable saveIniToShareBody =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    String operator = FactoryOperatorUserStore.resolveRdpLauncherOperatorName(ui);
                    if (operator.isBlank()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "操作者未選択",
                                "RPA設定.ini の保存先は操作者別ファイル（{操作者名}_RPA設定.ini）です。"
                                        + " ツールバーから操作者を選んでから保存してください。");
                        rapStatusLabel.setText("操作者未選択のため保存できません。");
                        return;
                    }
                    if (FactoryOperatorUserStore.isGuestOperator(operator)) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "ゲストは保存不可",
                                "ゲスト操作者向けの RPA設定.ini は保存できません。操作者を変更してください。");
                        rapStatusLabel.setText("ゲスト操作者のため保存できません。");
                        return;
                    }
                    Path iniPath = resolveRdpIniPath(ui);
                    Path profilesPath = AppPaths.resolveRdpLaunchProfilesFile(ui);
                    RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
                    Integer selectedProfile = launchProfileCombo.getValue();
                    int profileNumber =
                            selectedProfile != null
                                    ? selectedProfile
                                    : AppPaths.resolveRdpLaunchProfileNumber(ui);
                    ini.selectLaunchProfile(profileNumber);
                    RdpSessionEndAction endAction = cmbSessionEndAction.getValue();
                    ini.setSessionEndAction(
                            endAction != null ? endAction : RdpSessionEndAction.SIGN_OUT);
                    for (ProfileRowFields row : profileFields) {
                        if (RdpRemoteLauncherIni.isSignOutOnlyProfile(row.number())) {
                            continue;
                        }
                        ini.setSlotCommand(
                                row.number(),
                                row.programField().getText(),
                                RdpRemoteLauncherIni.mergeEternalFlag(
                                        RdpRemoteLauncherIni.normalizeScenarioArguments(
                                                row.argsField().getText()),
                                        row.chkRpaEternal().isSelected()));
                    }
                    ini.setSignOutSlotCommand();
                    String validation = ini.validateMessageForSave();
                    if (validation != null) {
                        showAlert(Alert.AlertType.WARNING, "保存できません", validation);
                        rapStatusLabel.setText(validation);
                        return;
                    }
                    try {
                        ini.save(iniPath);
                        RdpRemoteLauncherIni.writeOperatorContext(iniPath, operator);
                        List<RdpLaunchProfile> profiles =
                                collectProfilesFromRows(profileFields, profileMetadataByNumber);
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
                                "RPA設定.ini と起動プロファイル JSON を保存しました: "
                                        + iniPath);
                        refreshRightPanePreviews(refreshIniFilePreview[0], refreshLauncherLog[0]);
                    } catch (IOException ex) {
                        rapStatusLabel.setText("保存失敗: " + ex.getMessage());
                        showAlert(Alert.AlertType.ERROR, "保存失敗", ex.getMessage());
                    }
                };

        Runnable[] refreshAladdinCredentialsUi = new Runnable[1];
        Runnable[] saveIniToShare = new Runnable[1];

        Button btnReloadIni = new Button("再読込");
        styleSecondaryButton(btnReloadIni);
        btnReloadIni.setOnAction(
                e -> {
                    Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
                    confirm.setTitle("再読込");
                    confirm.setHeaderText(null);
                    confirm.setContentText("共有フォルダから RPA設定.ini を再読込します。未保存の変更は失われます。");
                    confirm.showAndWait().ifPresent(
                            choice -> {
                                if (choice == ButtonType.OK) {
                                    loadIniFromShare.run();
                                }
                            });
                });

        Button btnImportFromPeer = new Button("他ユーザーからインポート");
        styleSecondaryButton(btnImportFromPeer);
        btnImportFromPeer.setTooltip(
                new Tooltip(
                        "配備先フォルダ上の他操作者 "
                                + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                + " をプレビューし、現在の操作者設定へ取り込みます。"));
        btnImportFromPeer.setOnAction(
                e ->
                        showImportOperatorRpaIniDialog(
                                owner,
                                uiEnv,
                                launchProfileCombo,
                                profileFields,
                                appendProfileRow,
                                cmbSessionEndAction,
                                rapStatusLabel,
                                status,
                                refreshLaunchProfileCombo[0],
                                refreshDisplayPreview[0],
                                null,
                                false));

        Button btnImportFromSharedIni = new Button("共通 ini からコピー");
        styleSecondaryButton(btnImportFromSharedIni);
        btnImportFromSharedIni.setTooltip(
                new Tooltip(
                        "ユーザー名なしの共通 "
                                + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                + "（レガシー共有ファイル）からプロファイルを現在の操作者設定へ取り込みます。"
                                + " 取り込み後は「保存」で "
                                + "{操作者名}_"
                                + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                + " へ書き出してください。"));
        btnImportFromSharedIni.setOnAction(
                e -> {
                    Map<String, String> ui = uiEnv.get();
                    String operator = FactoryOperatorUserStore.resolveRdpLauncherOperatorName(ui);
                    if (operator.isBlank()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "操作者未選択",
                                "コピー先の操作者が未選択です。ツールバーから操作者を選んでから再度お試しください。");
                        return;
                    }
                    Path sharedIni = AppPaths.resolveExistingSharedRdpLauncherIni(ui);
                    if (!Files.isRegularFile(sharedIni)) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "共通 ini 未作成",
                                "共通の "
                                        + AppPaths.RDP_LAUNCHER_INI_BASENAME
                                        + " が見つかりません。\n"
                                        + sharedIni);
                        return;
                    }
                    Path targetIni = resolveRdpIniPath(ui).normalize();
                    if (sharedIni.normalize().equals(targetIni)) {
                        showAlert(
                                Alert.AlertType.INFORMATION,
                                "コピー不要",
                                "現在の操作者の ini が共通ファイルと同じです。"
                                        + " 他ユーザーからインポートするか、操作者を選び直してください。");
                        return;
                    }
                    showImportOperatorRpaIniDialog(
                            owner,
                            uiEnv,
                            launchProfileCombo,
                            profileFields,
                            appendProfileRow,
                            cmbSessionEndAction,
                            rapStatusLabel,
                            status,
                            refreshLaunchProfileCombo[0],
                            refreshDisplayPreview[0],
                            sharedIni,
                            true);
                });

        Button btnSaveIni = new Button("保存");
        styleSecondaryButton(btnSaveIni);
        btnSaveIni.setOnAction(e -> saveIniToShare[0].run());

        Button btnAddProfile = new Button("+ プロファイルを追加");
        styleSecondaryButton(btnAddProfile);
        btnAddProfile.setOnAction(
                e -> {
                    int maxRpa =
                            maxRpaProfileNumberFromRows(profileFields);
                    if (maxRpa >= RdpRemoteLauncherIni.MAX_SLOTS) {
                        return;
                    }
                    appendProfileRow.accept(maxRpa + 1);
                    btnAddProfile.setDisable(
                            maxRpaProfileNumberFromRows(profileFields)
                                    >= RdpRemoteLauncherIni.MAX_SLOTS);
                    refreshLaunchProfileCombo[0].run();
                });

        Button btnSuppressLaunchSlot = new Button("起動番号を 0 に");
        styleSecondaryButton(btnSuppressLaunchSlot);
        btnSuppressLaunchSlot.setTooltip(
                new Tooltip(
                        "RPA設定.ini の「"
                                + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                + "」を 0 に書き込みます。"
                                + " 次回タスクスケジューラ起動時は RPA を抑止します（サインアウトはしません）。"
                                + " プロファイル定義（スロット行）は保持されます。"));
        btnSuppressLaunchSlot.setOnAction(
                e -> {
                    Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
                    confirm.setTitle("起動番号 0（抑止）");
                    confirm.setHeaderText(null);
                    confirm.setContentText(
                            "共有フォルダの RPA設定.ini に "
                                    + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                    + "=0 を書き込みます。"
                                    + " 次回タスクスケジューラ起動時は RPA を抑止します（サインアウトはしません）。続行しますか？");
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
                        btnImportFromPeer,
                        btnImportFromSharedIni,
                        btnSaveIni,
                        btnSuppressLaunchSlot,
                        btnAddProfile);
        rapActionRow.setAlignment(Pos.CENTER_LEFT);

        VBox rapContent =
                new VBox(
                        8,
                        iniPathLabel,
                        deployDirBlock,
                        launcherPathBlock,
                        deployStatusLabel,
                        profileBox,
                        sessionEndActionRow,
                        rapActionRow,
                        rapStatusLabel);
        rapContent.setFillWidth(true);

        TitledPane rapPane = new TitledPane("RPA設定（起動プロファイルと RPA コマンド）", rapContent);
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

        CheckBox chkPreviewInTab = new CheckBox("接続中は右ペイン上部にプレビュー表示（読み取り専用）");
        chkPreviewInTab.setSelected(RdpPreviewSettings.isPreviewInTabEnabled(uiEnv.get()));
        chkPreviewInTab.setTooltip(
                new Tooltip(
                        "低 fps で mstsc ウィンドウをキャプチャして右ペインに表示します。"
                                + " 操作は別ウィンドウで行います。取得不可のときはプレビューを自動中止します。"));
        chkPreviewInTab
                .selectedProperty()
                .addListener(
                        (obs, was, selected) -> {
                            if (previewInTabChangeHandler != null) {
                                previewInTabChangeHandler.accept(selected ? "1" : "0");
                            }
                        });

        Label displayPreviewLabel = new Label();
        displayPreviewLabel.getStyleClass().add("pm-rdp-page-subtitle");
        displayPreviewLabel.setWrapText(true);
        displayPreviewLabel.setMaxWidth(CARD_WIDTH);

        Runnable[] refreshDisplayPreviewHolder = refreshDisplayPreview;
        refreshDisplayPreviewHolder[0] =
                () -> {
                    Integer selected = launchProfileCombo.getValue();
                    if (selected != null
                            && RdpRemoteLauncherIni.isSignOutOnlyProfile(selected)) {
                        displayPreviewLabel.setText(
                                RdpRemoteLauncherIni.signOutOnlyProfileDetailText()
                                        + "（表示は接続タブの既定設定を使用）");
                        return;
                    }
                    ProfileRowFields row =
                            selected != null ? findProfileRow(profileFields, selected) : null;
                    RdpLaunchProfile profile =
                            row != null
                                    ? buildProfileFromRow(row)
                                    : (selected != null
                                            ? profileMetadataByNumber.getOrDefault(
                                                    selected, RdpLaunchProfile.empty(selected))
                                            : RdpLaunchProfile.empty(1));
                    Map<String, String> ui = uiEnv.get();
                    LaunchDisplay display =
                            RdpLaunchDisplaySettings.resolveLaunchDisplay(profile, ui);
                    displayPreviewLabel.setText(
                            "表示: "
                                    + display.summaryText()
                                    + "（起動プロファイル "
                                    + profile.number()
                                    + " の RPA設定で変更）");
                };
        refreshDisplayPreviewHolder[0].run();

        Runnable refreshDisplayUi =
                () -> {
                    chkPreviewInTab.setSelected(
                            RdpPreviewSettings.isPreviewInTabEnabled(uiEnv.get()));
                    if (refreshDisplayPreviewHolder[0] != null) {
                        refreshDisplayPreviewHolder[0].run();
                    }
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

        Button[] btnLaunchRef = new Button[1];

        chkSequentialMode
                .selectedProperty()
                .addListener(
                        (obs, was, selected) -> {
                            if (!selected) {
                                sequentialSelectionOrder.clear();
                                sequentialStatusLabel.setText("");
                            }
                            if (refreshQuickLaunchSequentialLabels[0] != null) {
                                refreshQuickLaunchSequentialLabels[0].run();
                            }
                        });
        btnClearSequentialSelection.setOnAction(
                e -> {
                    sequentialSelectionOrder.clear();
                    sequentialStatusLabel.setText("");
                    if (refreshQuickLaunchSequentialLabels[0] != null) {
                        refreshQuickLaunchSequentialLabels[0].run();
                    }
                });

        Button btnLaunch = new Button("リモートデスクトップを起動");
        btnLaunchRef[0] = btnLaunch;
        for (Button quickBtn : quickLaunchProfileButtons) {
            quickBtn.setOnAction(
                    e -> {
                        Integer profileNumber = (Integer) quickBtn.getUserData();
                        if (profileNumber == null || btnLaunchRef[0] == null) {
                            return;
                        }
                        if (chkSequentialMode.isSelected()) {
                            if (sequentialRunActive.get() || rdpSessionActive.get()) {
                                return;
                            }
                            if (!RdpLaunchProfileSequentialRun.canAddProfileToSelection(
                                            sequentialSelectionOrder, profileNumber)
                                    && !sequentialSelectionOrder.contains(profileNumber)) {
                                showAlert(
                                        Alert.AlertType.WARNING,
                                        "連続実行の順序",
                                        "起動プロファイル "
                                                + RdpRemoteLauncherIni.SLOT_SIGN_OUT
                                                + "（接続先サインアウトのみ）は、"
                                                + "他のタスクを選ぶ前に先頭として選択してください。");
                                return;
                            }
                            List<Integer> nextSelection =
                                    RdpLaunchProfileSequentialRun.toggleSelection(
                                            List.copyOf(sequentialSelectionOrder),
                                            profileNumber);
                            sequentialSelectionOrder.clear();
                            sequentialSelectionOrder.addAll(nextSelection);
                            if (refreshQuickLaunchSequentialLabels[0] != null) {
                                refreshQuickLaunchSequentialLabels[0].run();
                            }
                            return;
                        }
                        launchProfileCombo.setValue(profileNumber);
                        btnLaunchRef[0].fire();
                    });
        }
        btnLaunch.getStyleClass().add("pm-rdp-connect-button");
        btnLaunch.setMaxWidth(Double.MAX_VALUE);
        btnLaunch.setTooltip(
                new Tooltip(
                        "指定 .rdp で mstsc.exe を起動します。"
                                + " 既存の接続ウィンドウがある場合は閉じます。"
                                + " RPA タスク終了時は ini のセッション操作（サインアウト等）が接続先で実行されます。"
                                + " 接続先サインアウトのみはプロファイル 99 で実行してください。"));
        btnLaunch.setOnAction(
                e -> {
                    if (launcherDeployInProgress.get()) {
                        showAlert(
                                Alert.AlertType.INFORMATION,
                                "転送中",
                                "ランチャー転送が完了するまで、リモートデスクトップは起動できません。");
                        return;
                    }
                    persistCompanionFields.run();
                    Map<String, String> uiGate = uiEnv.get();
                    FactorySite siteGate =
                            FactoryOperatorUserStore.operatorScopeForCurrentApp(uiGate, null);
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
                    Integer selectedLaunchProfile;
                    List<Integer> sequentialStartQueue = null;
                    if (sequentialRunActive.get()) {
                        List<Integer> queue = sequentialQueue.get();
                        int idx = sequentialIndex.get();
                        if (queue == null || queue.isEmpty() || idx < 0 || idx >= queue.size()) {
                            sequentialRunActive.set(false);
                            sequentialQueue.set(List.of());
                            showAlert(
                                    Alert.AlertType.ERROR,
                                    "連続実行エラー",
                                    "連続実行のキューが不正です。最初からやり直してください。");
                            return;
                        }
                        selectedLaunchProfile = queue.get(idx);
                    } else if (chkSequentialMode.isSelected()) {
                        List<Integer> queue =
                                RdpLaunchProfileSequentialRun.normalizeSelection(
                                        sequentialSelectionOrder);
                        if (queue.isEmpty()) {
                            showAlert(
                                    Alert.AlertType.WARNING,
                                    "タスク未選択",
                                    "連続実行モードでは、クイック起動ボタンを実行順にクリックして"
                                            + "タスクを選んでから起動してください。");
                            return;
                        }
                        var signOutOrderError =
                                RdpLaunchProfileSequentialRun.validateSignOutOnlyAtHead(queue);
                        if (signOutOrderError.isPresent()) {
                            showAlert(
                                    Alert.AlertType.WARNING,
                                    "連続実行の順序",
                                    signOutOrderError.get());
                            return;
                        }
                        sequentialStartQueue = queue;
                        sequentialIndex.set(0);
                        selectedLaunchProfile = queue.getFirst();
                    } else {
                        selectedLaunchProfile = launchProfileCombo.getValue();
                    }
                    launchProfileCombo.setValue(selectedLaunchProfile);
                    if (selectedLaunchProfile == null) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "プロファイル未選択",
                                "起動プロファイルを選択してください。");
                        return;
                    }
                    final boolean signOutOnlyLaunch =
                            RdpRemoteLauncherIni.isSignOutOnlyProfile(selectedLaunchProfile);
                    if (signOutOnlyLaunch
                            && chkSequentialMode.isSelected()
                            && !sequentialRunActive.get()) {
                        var signOutOrderError =
                                RdpLaunchProfileSequentialRun.validateSignOutOnlyAtHead(
                                        sequentialSelectionOrder);
                        if (signOutOrderError.isPresent()) {
                            showAlert(
                                    Alert.AlertType.WARNING,
                                    "連続実行の順序",
                                    signOutOrderError.get());
                            return;
                        }
                    }
                    if (signOutOnlyLaunch
                            && sequentialRunActive.get()) {
                        List<Integer> activeQueue = sequentialQueue.get();
                        int activeIdx = sequentialIndex.get();
                        if (activeQueue == null
                                || activeIdx != 0
                                || activeQueue.isEmpty()
                                || !RdpRemoteLauncherIni.isSignOutOnlyProfile(
                                        activeQueue.getFirst())) {
                            showAlert(
                                    Alert.AlertType.ERROR,
                                    "連続実行エラー",
                                    "起動プロファイル "
                                            + RdpRemoteLauncherIni.SLOT_SIGN_OUT
                                            + " は連続実行の先頭ステップのみ実行できます。");
                            return;
                        }
                    }
                    if (!signOutOnlyLaunch) {
                        try {
                            FactoryOperatorUserStore.configureForCurrentApp(uiGate, siteGate);
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
                    if (!signOutOnlyLaunch && launchRow.programField().getText().trim().isEmpty()) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "RPA 未設定",
                                "プロファイル "
                                        + selectedLaunchProfile
                                        + " の RPA プログラムが空です。"
                                        + " RPA設定で exe パスを登録して保存してください。");
                        return;
                    }
                    Map<String, String> ui = new HashMap<>(uiEnv.get());
                    if (!signOutOnlyLaunch) {
                        RdpLaunchSharedPathValidator.Result sharedPathCheck =
                                RdpLaunchSharedPathValidator.validateBeforeConnect(
                                        launchRow.programField().getText(),
                                        launchRow.argsField().getText(),
                                        AppPaths.resolveRdpLauncherExe(ui));
                        if (!sharedPathCheck.ok()) {
                            showAlert(
                                    Alert.AlertType.WARNING,
                                    "共有ファイルが見つかりません",
                                    sharedPathCheck.formatBlockingMessage());
                            rapStatusLabel.setText("接続を中止しました（共有フォルダ上のファイル未確認）");
                            status.accept(sharedPathCheck.formatBlockingMessage());
                            return;
                        }
                    }
                    RdpLaunchProfile launchProfileMeta = buildProfileFromRow(launchRow);
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM,
                            companionProgramField.getText().trim());
                    ui.put(
                            AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS,
                            companionArgsField.getText().trim());
                    ui.put(
                            RdpPreviewSettings.KEY_PM_AI_RDP_PREVIEW_IN_TAB,
                            chkPreviewInTab.isSelected() ? "1" : "0");
                    boolean previewSession = chkPreviewInTab.isSelected();
                    LaunchDisplay launchDisplay =
                            RdpLaunchDisplaySettings.resolveLaunchDisplay(
                                    launchProfileMeta, ui);
                    ui = new HashMap<>(RdpLaunchDisplaySettings.applyLaunchDisplayToUi(ui, launchDisplay));
                    ui.put(AppPaths.KEY_PM_AI_OPERATOR_USER, sessionOperator);
                    final LaunchDisplay resolvedDisplay = launchDisplay;
                    final boolean previewEnabled = previewSession;
                    final Map<String, String> launchUi = ui;
                    Path launchProfilePath =
                            RdpFileSigner.resolvePreferredSignedProfilePathFromUi(launchUi);
                    if (!Files.isRegularFile(launchProfilePath)
                            || !launchProfilePath
                                    .getFileName()
                                    .toString()
                                    .toLowerCase(Locale.ROOT)
                                    .endsWith(".rdp")) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "プロファイル未設定",
                                AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE
                                        + " に .rdp ファイルを指定するか、"
                                        + "RDP 署名ウィザードで署名済みプロファイルを作成してください。");
                        return;
                    }
                    if (!RemoteDesktopLauncher.isSupportedPlatform()) {
                        showAlert(
                                Alert.AlertType.ERROR,
                                "未対応",
                                "リモートデスクトップの起動は Windows 上のデスクトップアプリでのみ利用できます。");
                        return;
                    }
                    Path configuredProfile = launchProfilePath;
                    if (warnDefaultRdpBlockedAndMaybeOpenSignWizard(
                            owner != null ? owner : btnLaunch.getScene().getWindow(),
                            launchProfilePath,
                            uiEnv.get(),
                            status,
                            statusLabel,
                            profileChangeHandler,
                            profileField)) {
                        return;
                    }

                    int launchSlot = selectedLaunchProfile;
                    if (launchProfileNumberChangeHandler != null) {
                        launchProfileNumberChangeHandler.accept(launchSlot);
                    }
                    String launchProfileLabel = profileComboLabel(launchRow, launchProfileMeta);
                    Path launcherIniPath = resolveRdpIniPath(launchUi);
                    RdpSessionEndAction profileEndAction =
                            signOutOnlyLaunch
                                    ? RdpSessionEndAction.SIGN_OUT
                                    : launchProfileMeta.resolvedSessionEndAction();
                    final String launchArgsFinal =
                            signOutOnlyLaunch
                                    ? RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS
                                    : RdpRemoteLauncherIni.mergeEternalFlag(
                                            RdpRemoteLauncherIni.normalizeScenarioArguments(
                                                    launchRow.argsField().getText().trim()),
                                            launchRow.chkRpaEternal().isSelected());
                    try {
                        FactoryOperatorUserStore.syncLauncherCredentialsJsonToDeployDir(ui);
                        RdpRemoteLauncherIni.writeOperatorContext(
                                launcherIniPath, sessionOperator);
                        if (!signOutOnlyLaunch) {
                            RdpRemoteLauncherIni.writeLaunchContextBeforeConnect(
                                    launcherIniPath,
                                    launchSlot,
                                    launchRow.programField().getText().trim(),
                                    launchArgsFinal,
                                    profileEndAction);
                        }
                        refreshRightPanePreviews(refreshIniFilePreview[0], refreshLauncherLog[0]);
                    } catch (IOException iniEx) {
                        showAlert(
                                Alert.AlertType.ERROR,
                                "RPA設定.ini の更新に失敗",
                                "接続前に起動プログラム番号を書き込めませんでした: "
                                        + iniEx.getMessage());
                        return;
                    }

                    final RdpSessionEndAction profileEndActionFinal = profileEndAction;
                    final boolean signOutOnlyLaunchFinal = signOutOnlyLaunch;

                    final List<Integer> sequentialStartQueueFinal = sequentialStartQueue;

                    Runnable performLaunch =
                            () -> {
                                try {
                                    if (sequentialStartQueueFinal != null
                                            && !sequentialRunActive.get()) {
                                        sequentialQueue.set(sequentialStartQueueFinal);
                                        sequentialRunActive.set(true);
                                    }
                                    if (sequentialRunActive.get()) {
                                        List<Integer> activeQueue = sequentialQueue.get();
                                        int activeIdx = sequentialIndex.get();
                                        if (activeQueue != null && !activeQueue.isEmpty()) {
                                            sequentialStatusLabel.setText(
                                                    RdpLaunchProfileSequentialRun.progressStatusText(
                                                            activeIdx + 1,
                                                            activeQueue.size(),
                                                            launchProfileLabel));
                                        }
                                    }
                                    Path configured = configuredProfile;
                                    Path preferred =
                                            RdpFileSigner.resolvePreferredSignedProfilePath(
                                                    configured, launchUi);

                                    long closedPid =
                                            RdpMstscCloser.prepareRelaunch(
                                                    preferred, rdpSessionWatchThread);
                                    if (closedPid > 0) {
                                        String closeMsg =
                                                signOutOnlyLaunchFinal
                                                        ? "既存のリモートデスクトップ接続ウィンドウを閉じました（PID="
                                                                + closedPid
                                                                + "）。接続先サインアウトを実行します。"
                                                        : "既存のリモートデスクトップ接続ウィンドウを閉じました（PID="
                                                                + closedPid
                                                                + "）。RPA セッションを起動します。";
                                        status.accept(closeMsg);
                                        statusLabel.setText(
                                                signOutOnlyLaunchFinal
                                                        ? "既存接続を閉じました。接続先サインアウトを実行中…"
                                                        : "既存接続を閉じました。RPA セッションを起動中…");
                                    }

                                    if (!rdpSessionActive.compareAndSet(false, true)) {
                                        showAlert(
                                                Alert.AlertType.INFORMATION,
                                                "接続中",
                                                "リモートデスクトップ接続シーケンスが既に実行中です。");
                                        return;
                                    }
                                    updateLaunchButtonState[0].run();

                                    if (signOutOnlyLaunchFinal) {
                                        rdpConnectPhase.set(RdpConnectPhase.REMOTE_SIGN_OUT);

                                        RemoteDesktopLauncher.LaunchOutcome signOutOutcome =
                                                RemoteDesktopLauncher.launchSignOutViaTaskScheduler(
                                                        preferred, launchUi);

                                        String signOutMsg =
                                                "接続先サインアウトを起動しました: "
                                                        + signOutOutcome.rdpProfile()
                                                        + "\n（タスクスケジューラ + "
                                                        + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                                        + "=99, 99="
                                                        + RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS
                                                        + "）";
                                        if (sequentialRunActive.get()) {
                                            List<Integer> activeQueue = sequentialQueue.get();
                                            int activeIdx = sequentialIndex.get();
                                            if (activeQueue != null && !activeQueue.isEmpty()) {
                                                signOutMsg +=
                                                        "\n"
                                                                + RdpLaunchProfileSequentialRun
                                                                        .progressStatusText(
                                                                                activeIdx + 1,
                                                                                activeQueue.size(),
                                                                                launchProfileLabel);
                                            }
                                        }
                                        status.accept(signOutMsg);
                                        statusLabel.setText(signOutMsg);

                                        RdpMstscSessionMonitor.watchAfterLaunch(
                                                signOutOutcome.rdpProfile(),
                                                signOutOutcome.mstscProcessId(),
                                                signOutOutcome.mstscPidMarkerFile(),
                                                rdpSessionWatchThread,
                                                event ->
                                                        Platform.runLater(
                                                                () ->
                                                                        onSignOutOnlyLaunchEnded(
                                                                                event,
                                                                                signOutOutcome,
                                                                                statusLabel,
                                                                                status,
                                                                                rdpSessionActive,
                                                                                rdpConnectPhase,
                                                                                updateLaunchButtonState[0],
                                                                                rightPanePreviewHolder,
                                                                                refreshIniFilePreview[0],
                                                                                refreshLauncherLog[0],
                                                                                refreshFetchedFilesTable[0],
                                                                                scrollToFetchedFilesSection[0],
                                                                                sequentialRunActive,
                                                                                sequentialQueue,
                                                                                sequentialIndex,
                                                                                sequentialSelectionOrder,
                                                                                sequentialStatusLabel,
                                                                                refreshQuickLaunchSequentialLabels[0],
                                                                                btnLaunchRef)));
                                    } else {
                                        rdpConnectPhase.set(RdpConnectPhase.RPA_SESSION);
                                        launchRpaSessionAndWatch(
                                                preferred,
                                                configured,
                                                launchUi,
                                                launchSlot,
                                                launchProfileLabel,
                                                resolvedDisplay,
                                                previewEnabled,
                                                profileChangeHandler,
                                                profileField,
                                                statusLabel,
                                                status,
                                                rdpSessionActive,
                                                rdpConnectPhase,
                                                rdpSessionWatchThread,
                                                updateLaunchButtonState[0],
                                                rightPanePreviewHolder,
                                                refreshIniFilePreview[0],
                                                refreshLauncherLog[0],
                                                refreshFetchedFilesTable[0],
                                                scrollToFetchedFilesSection[0],
                                                sequentialRunActive,
                                                sequentialQueue,
                                                sequentialIndex,
                                                sequentialSelectionOrder,
                                                sequentialStatusLabel,
                                                refreshQuickLaunchSequentialLabels[0],
                                                btnLaunchRef);
                                    }
                                } catch (IOException ex) {
                                    rdpConnectPhase.set(RdpConnectPhase.IDLE);
                                    rdpSessionActive.set(false);
                                    sequentialRunActive.set(false);
                                    sequentialQueue.set(List.of());
                                    updateLaunchButtonState[0].run();
                                    if (rightPanePreviewHolder[0] != null) {
                                        rightPanePreviewHolder[0].removePreviewPane();
                                    }
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

        VBox launchBlock =
                new VBox(10, btnLaunch, signWizardRow);
        launchBlock.getStyleClass().add("pm-rdp-launch-block");
        launchBlock.setAlignment(Pos.CENTER);
        launchBlock.setFillWidth(true);
        launchBlock.setMaxWidth(CARD_WIDTH);

        saveIniToShare[0] = saveIniToShareBody;

        Runnable onSessionOperatorChanged =
                () -> {
                    refreshPaths.run();
                    loadIniFromShare.run();
                    if (refreshAladdinCredentialsUi[0] != null) {
                        refreshAladdinCredentialsUi[0].run();
                    }
                };

        Button btnDeleteLaunchProfile = new Button("プロファイルを削除");
        styleSecondaryButton(btnDeleteLaunchProfile);
        btnDeleteLaunchProfile.setTooltip(
                new Tooltip(
                        "選択中の起動プロファイルを論理削除します。"
                                + " 削除済みは「削除したプロファイルを復活」から戻せます。"));

        Button btnRestoreDeletedProfiles = new Button("削除したプロファイルを復活");
        styleSecondaryButton(btnRestoreDeletedProfiles);
        btnRestoreDeletedProfiles.setTooltip(
                new Tooltip("論理削除した起動プロファイルを一覧から選んで復元します。"));

        HBox launchProfileManageRow =
                new HBox(8, btnDeleteLaunchProfile, btnRestoreDeletedProfiles);
        launchProfileManageRow.setAlignment(Pos.CENTER_LEFT);
        launchProfileManageRow.setMaxWidth(CARD_WIDTH);

        updateProfileManageButtons[0] =
                () -> {
                    Integer selected = launchProfileCombo.getValue();
                    List<RdpLaunchProfile> all =
                            collectProfilesFromRows(profileFields, profileMetadataByNumber);
                    btnDeleteLaunchProfile.setDisable(
                            selected == null || !RdpLaunchProfileCatalog.canSoftDelete(all));
                    btnRestoreDeletedProfiles.setDisable(
                            RdpLaunchProfileCatalog.deletedProfiles(all).isEmpty());
                };

        btnDeleteLaunchProfile.setOnAction(
                e ->
                        confirmAndSoftDeleteProfile(
                                owner != null
                                        ? owner
                                        : btnDeleteLaunchProfile.getScene().getWindow(),
                                launchProfileCombo.getValue(),
                                profileFields,
                                profileMetadataByNumber,
                                refreshLaunchProfileCombo[0],
                                saveIniToShare[0]));

        btnRestoreDeletedProfiles.setOnAction(
                e ->
                        showRestoreDeletedProfilesDialog(
                                owner != null
                                        ? owner
                                        : btnRestoreDeletedProfiles.getScene().getWindow(),
                                profileFields,
                                profileMetadataByNumber,
                                refreshLaunchProfileCombo[0],
                                saveIniToShare[0]));

        updateProfileManageButtons[0].run();

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
                    FactorySite site = FactoryOperatorUserStore.operatorScopeForCurrentApp(ui, null);
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
                        FactoryOperatorUserStore.configureForCurrentApp(ui, site);
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
                        fieldCaption("起動プロファイル"),
                        launchProfileCombo,
                        launchProfileManageRow,
                        sequentialModeRow,
                        quickLaunchProfileGrid,
                        sequentialStatusLabel,
                        launchProfileDetail,
                        fieldCaption("RDP プロファイル (.rdp)"),
                        profileRow,
                        chkPreviewInTab,
                        displayPreviewLabel,
                        launchBlock,
                        statusLabel,
                        fieldCaption("アラジン ログイン ID"),
                        aladdinLoginIdField,
                        fieldCaption("アラジン パスワード"),
                        aladdinPasswordField,
                        credentialsSaveRow,
                        credentialsStatusLabel,
                        rpaInjectedArgsHint);
        connectHero.getStyleClass().add("pm-rdp-connect-hero");
        connectHero.setFillWidth(true);
        connectHero.setMaxWidth(CARD_WIDTH);

        refreshAladdinCredentialsUi[0] =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    FactorySite site = FactoryOperatorUserStore.operatorScopeForCurrentApp(ui, null);
                    try {
                        FactoryOperatorUserStore.configureForCurrentApp(ui, site);
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
                                quickLaunchProfileButtons,
                                launcherDeployInProgress,
                                launcherDeployReady,
                                rdpSessionActive,
                                rdpConnectPhase,
                                chkSequentialMode,
                                sequentialSelectionOrder,
                                sequentialRunActive,
                                sequentialIndex,
                                sequentialQueue.get());
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

        Label fetchedFilesTitle = new Label("取得データ最新ファイル");
        fetchedFilesTitle.getStyleClass().add("pm-rdp-section-title");

        Label fetchedFilesHint =
                new Label(
                        "RPA 取得先フォルダ内の更新時刻が最新のファイルを表示します。"
                                + " 取得日時はファイルの更新日時（ローカル時刻）です。"
                                + " 接続終了後に自動更新されます。");
        fetchedFilesHint.getStyleClass().add("pm-rdp-page-subtitle");
        fetchedFilesHint.setWrapText(true);
        fetchedFilesHint.setMaxWidth(CARD_WIDTH);

        Label fetchedFilesMetaLabel = new Label("未確認");
        fetchedFilesMetaLabel.getStyleClass().add("pm-rdp-meta-label");
        fetchedFilesMetaLabel.setWrapText(true);
        fetchedFilesMetaLabel.setMaxWidth(CARD_WIDTH);

        TableView<RemoteDesktopLatestSourceFiles.Row> fetchedFilesTable = new TableView<>();
        fetchedFilesTable.getStyleClass().add("pm-rdp-fetched-files-table");
        fetchedFilesTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        fetchedFilesTable.setFixedCellSize(-1);
        fetchedFilesTable.setMinHeight(120);
        fetchedFilesTable.setPrefHeight(360);
        fetchedFilesTable.setMaxHeight(640);
        fetchedFilesTable.setPlaceholder(new Label("一覧を読込中…"));

        TableColumn<RemoteDesktopLatestSourceFiles.Row, String> fetchedCategoryCol =
                new TableColumn<>("種別");
        fetchedCategoryCol.setCellValueFactory(
                data ->
                        new SimpleStringProperty(
                                data.getValue().category().label()));
        fetchedCategoryCol.setPrefWidth(108);
        fetchedCategoryCol.setMinWidth(88);
        fetchedCategoryCol.setMaxWidth(140);

        TableColumn<RemoteDesktopLatestSourceFiles.Row, String> fetchedAcquiredAtCol =
                new TableColumn<>("取得日時");
        fetchedAcquiredAtCol.setCellValueFactory(
                data -> new SimpleStringProperty(data.getValue().displayAcquiredAt()));
        fetchedAcquiredAtCol.setPrefWidth(148);
        fetchedAcquiredAtCol.setMinWidth(132);
        fetchedAcquiredAtCol.setMaxWidth(180);

        TableColumn<RemoteDesktopLatestSourceFiles.Row, String> fetchedPathCol =
                new TableColumn<>("フルパス");
        fetchedPathCol.setCellValueFactory(
                data -> new SimpleStringProperty(data.getValue().displayPath()));
        fetchedPathCol.setPrefWidth(480);
        installWrappingReadOnlyTableCell(fetchedPathCol);

        fetchedFilesTable
                .getColumns()
                .addAll(fetchedCategoryCol, fetchedAcquiredAtCol, fetchedPathCol);
        ObservableList<RemoteDesktopLatestSourceFiles.Row> fetchedFilesItems =
                FXCollections.observableArrayList();
        fetchedFilesTable.setItems(fetchedFilesItems);

        Button btnRefreshFetchedFiles = new Button("最新を再確認");
        styleSecondaryButton(btnRefreshFetchedFiles);
        btnRefreshFetchedFiles.setTooltip(
                new Tooltip("共有フォルダを再スキャンして一覧を更新します。"));

        Button btnJumpToTop = new Button("↑ 先頭へ戻る");
        styleSecondaryButton(btnJumpToTop);
        btnJumpToTop.setTooltip(new Tooltip("ページ先頭へスクロールします。"));

        HBox fetchedFilesToolbar = new HBox(8, btnRefreshFetchedFiles, btnJumpToTop);
        fetchedFilesToolbar.setAlignment(Pos.CENTER_LEFT);
        fetchedFilesToolbar.setMaxWidth(CARD_WIDTH);

        VBox fetchedFilesSection =
                new VBox(
                        8,
                        fetchedFilesTitle,
                        fetchedFilesHint,
                        fetchedFilesMetaLabel,
                        fetchedFilesTable,
                        fetchedFilesToolbar);
        fetchedFilesSection.getStyleClass().add("pm-rdp-fetched-files-section");
        fetchedFilesSection.setFillWidth(true);
        fetchedFilesSection.setMaxWidth(CARD_WIDTH);

        refreshFetchedFilesTable[0] =
                () ->
                        Thread.ofVirtual()
                                .name("rdp-fetched-files-refresh")
                                .start(
                                        () -> {
                                            List<RemoteDesktopLatestSourceFiles.Row> rows =
                                                    RemoteDesktopLatestSourceFiles.resolveAll(
                                                            uiEnv.get());
                                            int found =
                                                    (int)
                                                            rows.stream()
                                                                    .filter(
                                                                            r ->
                                                                                    !r.fullPath()
                                                                                            .isBlank())
                                                                    .count();
                                            String meta =
                                                    "更新: "
                                                            + java.time.LocalDateTime.now()
                                                                    .format(
                                                                            java.time.format
                                                                                    .DateTimeFormatter
                                                                                    .ofPattern(
                                                                                            "yyyy-MM-dd HH:mm:ss"))
                                                            + " ／ 検出 "
                                                            + found
                                                            + " / "
                                                            + rows.size()
                                                            + " 件";
                                            Platform.runLater(
                                                    () -> {
                                                        fetchedFilesItems.setAll(rows);
                                                        fetchedFilesMetaLabel.setText(meta);
                                                        fetchedFilesTable.refresh();
                                                    });
                                        });

        btnRefreshFetchedFiles.setOnAction(e -> refreshFetchedFilesTable[0].run());

        Timeline[] fetchedFilesAutoRefreshTimeline = new Timeline[1];
        Consumer<Boolean> setTabActive =
                active -> {
                    if (fetchedFilesAutoRefreshTimeline[0] != null) {
                        fetchedFilesAutoRefreshTimeline[0].stop();
                        fetchedFilesAutoRefreshTimeline[0] = null;
                    }
                    if (active != null && active) {
                        Timeline timeline =
                                new Timeline(
                                        new KeyFrame(
                                                FETCHED_FILES_AUTO_REFRESH_INTERVAL,
                                                e -> refreshFetchedFilesTable[0].run()));
                        timeline.setCycleCount(Timeline.INDEFINITE);
                        timeline.play();
                        fetchedFilesAutoRefreshTimeline[0] = timeline;
                    }
                };

        VBox card =
                new VBox(
                        16,
                        headerBlock,
                        connectHero,
                        rapPane,
                        policyPane,
                        fetchedFilesSection);
        card.getStyleClass().add("pm-rdp-settings-stack");
        card.setPadding(new Insets(16));
        card.setMaxWidth(CARD_WIDTH);
        card.setPrefWidth(CARD_WIDTH);
        card.setFillWidth(true);

        Label iniPreviewTitle = new Label("RPA設定.ini（共有ファイル）");
        iniPreviewTitle.getStyleClass().add("pm-rdp-ini-panel-title");

        Label iniPreviewMetaLabel = new Label("読込待ち…");
        iniPreviewMetaLabel.setWrapText(true);
        iniPreviewMetaLabel.getStyleClass().add("pm-rdp-page-subtitle");

        TextArea iniPreviewArea = new TextArea();
        iniPreviewArea.setEditable(false);
        iniPreviewArea.setWrapText(false);
        iniPreviewArea.setPromptText("共有 UNC 上の RPA設定.ini の生テキストがここに表示されます。");
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
        iniPreviewPane.setMinHeight(RIGHT_PANE_MIN_HEIGHT);

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
        launcherLogArea.setPromptText(
                "共有フォルダ上の launcher-logs\\launcher-yyyyMMdd.log がここに表示されます。");
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

        Button btnDeleteLauncherLog = new Button("ログ削除");
        styleSecondaryButton(btnDeleteLauncherLog);
        btnDeleteLauncherLog.setTooltip(
                new Tooltip("共有フォルダ上の当日 launcher ログファイルを削除します。"));
        btnDeleteLauncherLog.setOnAction(
                e ->
                        confirmAndDeleteLauncherLog(
                                owner != null
                                        ? owner
                                        : btnDeleteLauncherLog.getScene().getWindow(),
                                uiEnv,
                                refreshLauncherLog[0],
                                status));

        HBox launcherLogToolbar =
                new HBox(
                        8,
                        new Label("検索"),
                        launcherLogSearchField,
                        new Label("フォント"),
                        launcherLogFontFamilyCombo,
                        new Label("サイズ"),
                        launcherLogFontSizeCombo,
                        btnRefreshLauncherLog,
                        btnDeleteLauncherLog);
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
        launcherLogPane.setMinHeight(RIGHT_PANE_MIN_HEIGHT);

        SplitPane rightPaneSplit = new SplitPane(iniPreviewPane, launcherLogPane);
        rightPaneSplit.setOrientation(Orientation.VERTICAL);
        rightPaneSplit.setDividerPositions(RIGHT_PANE_TOP_RATIO_INITIAL);
        SplitPane.setResizableWithParent(iniPreviewPane, Boolean.TRUE);
        SplitPane.setResizableWithParent(launcherLogPane, Boolean.TRUE);
        rightPaneSplit.getStyleClass().add("pm-remote-desktop-right-split");
        rightPanePreviewHolder[0] = new RdpRightPanePreviewController(rightPaneSplit);

        ScrollPane leftScroll = new ScrollPane(card);
        leftScroll.setFitToWidth(true);
        leftScroll.getStyleClass().add("pm-rdp-form-scroll");

        btnJumpToFetchedFiles.setOnAction(
                e -> scrollScrollPaneToNode(leftScroll, fetchedFilesSection));
        btnJumpToTop.setOnAction(e -> leftScroll.setVvalue(0));
        scrollToFetchedFilesSection[0] =
                () -> scrollScrollPaneToNode(leftScroll, fetchedFilesSection);

        SplitPane splitPane = new SplitPane(leftScroll, rightPaneSplit);
        splitPane.setDividerPositions(0.52);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(rightPaneSplit, Boolean.TRUE);
        splitPane.getStyleClass().add("pm-remote-desktop-split");

        Runnable scheduleInitialRefresh =
                () -> {
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
                                refreshDisplayUi,
                                cmbSessionEndAction,
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
                        refreshFetchedFilesTable[0].run();
                };

        return new TabContent(splitPane, scheduleInitialRefresh, onSessionOperatorChanged, setTabActive);
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
            Runnable refreshDisplayUi,
            ComboBox<RdpSessionEndAction> cmbSessionEndAction,
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
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    if (!isExplicitLauncherDeployDirConfigured(ui)) {
                        Platform.runLater(
                                () -> {
                                    launcherDeployInProgress.set(false);
                                    launcherDeployReady.set(false);
                                    btnForceDeployLauncher.setDisable(false);
                                    updateLaunchButtonState.run();
                                    refreshPaths.run();
                                    deployStatusLabel.setText("配備先フォルダを指定してください。");
                                    status.accept("ランチャー配備先が未設定のため転送をスキップしました。");
                                });
                        return;
                    }
                    boolean willCopyExe = RdpRemoteLauncherDeployer.needsExeDeploy(ui);
                    if (willCopyExe) {
                        Platform.runLater(
                                () -> {
                                    if (launcherDeployInProgress.compareAndSet(false, true)) {
                                        deployStatusLabel.setText("ランチャーを転送中…");
                                        btnForceDeployLauncher.setDisable(true);
                                        updateLaunchButtonState.run();
                                    }
                                });
                    }
                    RdpRemoteLauncherDeployer.DeployOutcome deploy =
                            RdpRemoteLauncherDeployer.ensureDeployed(ui, null);
                    Path iniPath = resolveExistingRdpIniPath(ui);
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
                    Path preferredProfile =
                            RdpFileSigner.resolvePreferredSignedProfilePathFromUi(ui);
                    if (!Files.isRegularFile(preferredProfile)) {
                        preferredProfile = null;
                    }
                    final RdpRemoteLauncherIni loadedIni = ini;
                    final IOException loadError = iniError;
                    final Path resolvedIniPath = iniPath;
                    final Path resolvedProfilesPath = profilesPath;
                    final List<RdpLaunchProfile> resolvedCatalogProfiles = catalogProfiles;
                    final Path resolvedPreferred = preferredProfile;
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
                                if (resolvedPreferred == null) {
                                    profileField.setText("");
                                } else {
                                    profileField.setText(resolvedPreferred.toString());
                                    if (profileChangeHandler != null) {
                                        profileChangeHandler.accept(resolvedPreferred.toString());
                                    }
                                }
                                companionProgramField.setText(
                                        ui.getOrDefault(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM, ""));
                                companionArgsField.setText(
                                        ui.getOrDefault(
                                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS, ""));
                                refreshDisplayUi.run();
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
                                        cmbSessionEndAction,
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
            ComboBox<RdpSessionEndAction> cmbSessionEndAction,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshLaunchProfileCombo) {
        runBackgroundThenFx(
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path iniPath = resolveExistingRdpIniPath(ui);
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
                                        cmbSessionEndAction,
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
                    Path iniPath = resolveExistingRdpIniPath(ui);
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

    private static void confirmAndDeleteLauncherLog(
            Window owner,
            Supplier<Map<String, String>> uiEnv,
            Runnable refreshLauncherLog,
            Consumer<String> status) {
        Path logPath = AppPaths.resolveRdpLauncherSharedLogPath(uiEnv.get());
        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("ログ削除");
        confirm.setHeaderText("接続先ランチャーログを削除しますか？");
        Label body = new Label(logPath.toString() + "\n\nこの操作は取り消せません。");
        body.setWrapText(true);
        body.setMaxWidth(520);
        confirm.getDialogPane().setContent(body);
        if (owner != null) {
            confirm.initOwner(owner);
        }
        confirm.showAndWait().ifPresent(
                choice -> {
                    if (choice == ButtonType.OK) {
                        scheduleLauncherLogDelete(uiEnv, logPath, refreshLauncherLog, status);
                    }
                });
    }

    private static void scheduleLauncherLogDelete(
            Supplier<Map<String, String>> uiEnv,
            Path logPath,
            Runnable refreshLauncherLog,
            Consumer<String> status) {
        runBackgroundThenFx(
                () -> {
                    String message;
                    boolean succeeded;
                    try {
                        if (!Files.exists(logPath)) {
                            succeeded = true;
                            message = "ログファイルは既に存在しません: " + logPath;
                        } else {
                            Files.delete(logPath);
                            succeeded = true;
                            message = "ランチャーログを削除しました: " + logPath;
                        }
                    } catch (IOException ex) {
                        succeeded = false;
                        message = "ログ削除に失敗: " + ex.getMessage() + " (" + logPath + ")";
                    }
                    final String resultMessage = message;
                    final boolean ok = succeeded;
                    Platform.runLater(
                            () -> {
                                if (ok) {
                                    refreshLauncherLog.run();
                                    status.accept(resultMessage);
                                } else {
                                    showAlert(Alert.AlertType.ERROR, "ログ削除", resultMessage);
                                    status.accept(resultMessage);
                                }
                            });
                });
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
            ComboBox<RdpSessionEndAction> cmbSessionEndAction,
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
        } else if (RdpRemoteLauncherIni.isSignOutIniSlot(loadedSlot)) {
            comboTarget = RdpRemoteLauncherIni.SLOT_SIGN_OUT;
        } else if (loadedSlot >= 1 && loadedSlot <= RdpRemoteLauncherIni.MAX_SLOTS) {
            comboTarget = loadedSlot;
        } else {
            comboTarget = 1;
        }
        cmbSessionEndAction.setValue(ini.sessionEndAction());

        int visibleFromIni = ini.visibleSlotCount();
        int visibleFromCatalog =
                catalogProfiles.stream().mapToInt(RdpLaunchProfile::number).max().orElse(0);
        int visible =
                Math.min(
                        RdpRemoteLauncherIni.MAX_SLOTS,
                        Math.max(
                                RdpRemoteLauncherIni.DEFAULT_INITIAL_RPA_PROFILE_ROWS,
                                Math.max(visibleFromIni, visibleFromCatalog)));
        if (findProfileRow(profileFields, RdpRemoteLauncherIni.SLOT_SIGN_OUT) == null) {
            appendProfileRow.accept(RdpRemoteLauncherIni.SLOT_SIGN_OUT);
        }
        int maxNumber = maxRpaProfileNumberFromRows(profileFields);
        while (maxNumber < visible) {
            appendProfileRow.accept(++maxNumber);
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
            if (meta.fullScreen() != null) {
                row.chkFullScreen().setSelected(meta.fullScreen());
            } else {
                row.chkFullScreen().setSelected(false);
            }
            int width =
                    meta.desktopWidth() != null
                            ? meta.desktopWidth()
                            : RdpLaunchDisplaySettings.DEFAULT_WIDTH;
            int height =
                    meta.desktopHeight() != null
                            ? meta.desktopHeight()
                            : RdpLaunchDisplaySettings.DEFAULT_HEIGHT;
            row.desktopWidthSpinner().getValueFactory().setValue(RdpLaunchDisplaySettings.clampWidth(width));
            row.desktopHeightSpinner().getValueFactory().setValue(RdpLaunchDisplaySettings.clampHeight(height));
            updateDisplayControlsEnabled(
                    row.chkFullScreen(), row.desktopWidthSpinner(), row.desktopHeightSpinner());
            applyProfileCardVisibility(row, profileMetadataByNumber);
        }

        if (refreshLaunchProfileCombo != null) {
            refreshLaunchProfileCombo.run();
            if (launchProfileCombo.getItems().contains(comboTarget)) {
                launchProfileCombo.setValue(comboTarget);
            }
        } else {
            launchProfileCombo.setValue(comboTarget);
        }

        if (RdpRemoteLauncherIni.isSignOutIniSlot(loadedSlot)) {
            rapStatusLabel.setText(
                    "読込: "
                            + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                            + "="
                            + RdpRemoteLauncherIni.SLOT_SIGN_OUT
                            + "（接続先サインアウトのみ）。"
                            + " 起動プロファイル 99 を選択中: "
                            + iniPath
                            + " / "
                            + profilesPath);
        } else if (RdpRemoteLauncherIni.isSuppressIniSlot(loadedSlot)) {
            rapStatusLabel.setText(
                    "読込: "
                            + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                            + "="
                            + RdpRemoteLauncherIni.INI_SUPPRESS_SLOT
                            + "（タスクスケジューラ抑止）。"
                            + " "
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
                    Path iniPath = resolveRdpIniPath(ui);
                    IOException error = null;
                    try {
                        RdpRemoteLauncherIni.writeTaskSchedulerSuppress(iniPath, ui);
                        RdpRemoteLauncherIni.clearSignOutOnConnectRequest(iniPath);
                    } catch (IOException ex) {
                        error = ex;
                    }
                    final IOException suppressError = error;
                    Platform.runLater(
                            () -> {
                                if (suppressError != null) {
                                    String msg =
                                            "起動プログラム番号の "
                                                    + RdpRemoteLauncherIni.INI_SIGN_OUT_SLOT
                                                    + " 設定に失敗: "
                                                    + suppressError.getMessage();
                                    rapStatusLabel.setText(msg);
                                    status.accept(msg);
                                    showAlert(Alert.AlertType.ERROR, "抑止失敗", msg);
                                    return;
                                }
                                String msg =
                                        RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                                                + "="
                                                + RdpRemoteLauncherIni.INI_SIGN_OUT_SLOT
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
        Map<String, String> ui = uiEnv.get();
        boolean willCopyExe = force || RdpRemoteLauncherDeployer.needsExeDeploy(ui);
        if (willCopyExe) {
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
        }
        runBackgroundThenFx(
                () -> {
                    RdpRemoteLauncherDeployer.DeployOutcome outcome =
                            force
                                    ? RdpRemoteLauncherDeployer.forceDeploy(ui)
                                    : RdpRemoteLauncherDeployer.ensureDeployed(ui);
                    Platform.runLater(
                            () -> {
                                if (willCopyExe) {
                                    deployInProgress.set(false);
                                    btnForceDeployLauncher.setDisable(false);
                                }
                                deployReady.set(outcome.succeeded());
                                outcome.message().ifPresent(deployStatusLabel::setText);
                                outcome.message().ifPresent(status::accept);
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
            List<Button> quickLaunchProfileButtons,
            AtomicBoolean deployInProgress,
            AtomicBoolean deployReady,
            AtomicBoolean sessionActive,
            AtomicReference<RdpConnectPhase> connectPhase,
            CheckBox sequentialMode,
            List<Integer> sequentialSelection,
            AtomicBoolean sequentialRunActive,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialQueue) {
        btnLaunch
                .getStyleClass()
                .removeAll(
                        "pm-rdp-connect-pending",
                        "pm-rdp-connect-deploying",
                        "pm-rdp-connect-active");
        boolean sequentialModeOn = sequentialMode != null && sequentialMode.isSelected();
        boolean sequentialRunning = sequentialRunActive != null && sequentialRunActive.get();
        if (sequentialMode != null) {
            sequentialMode.setDisable(deployInProgress.get() || sessionActive.get() || sequentialRunning);
        }
        if (deployInProgress.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-deploying");
            btnLaunch.setDisable(true);
            btnLaunch.setText("ランチャー転送中…");
            applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, true);
            return;
        }
        if (sessionActive.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-active");
            btnLaunch.setDisable(true);
            if (connectPhase != null
                    && connectPhase.get() == RdpConnectPhase.REMOTE_SIGN_OUT) {
                btnLaunch.setText("接続先サインアウト実行中…");
            } else if (sequentialRunning
                    && sequentialQueue != null
                    && !sequentialQueue.isEmpty()
                    && sequentialIndex != null) {
                btnLaunch.setText(
                        RdpLaunchProfileSequentialRun.launchButtonTextActive(
                                sequentialIndex.get() + 1, sequentialQueue.size()));
            } else if (sequentialModeOn) {
                btnLaunch.setText(
                        RdpLaunchProfileSequentialRun.launchButtonTextActive(1, Math.max(1, sequentialSelection != null ? sequentialSelection.size() : 1)));
            } else {
                btnLaunch.setText("リモートデスクトップ接続中");
            }
            applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, true);
            return;
        }
        if (!deployReady.get()) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-pending");
            btnLaunch.setDisable(false);
            btnLaunch.setText(
                    sequentialModeOn
                            ? RdpLaunchProfileSequentialRun.launchButtonTextIdle(
                                    sequentialSelection != null ? sequentialSelection.size() : 0)
                            : "ランチャー転送してから起動");
            applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, false);
            return;
        }
        if (!rdpLaunchCredentialsReady(uiEnv)
                && !(sequentialModeOn
                        && !RdpLaunchProfileSequentialRun.selectionRequiresAladdinCredentials(
                                sequentialSelection))) {
            btnLaunch.getStyleClass().add("pm-rdp-connect-pending");
            btnLaunch.setDisable(true);
            btnLaunch.setText(
                    sequentialModeOn
                            ? "アラジン資格情報を保存してください"
                            : "アラジン資格情報を保存してください");
            applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, true);
            return;
        }
        if (sequentialModeOn) {
            int selectedCount = sequentialSelection != null ? sequentialSelection.size() : 0;
            btnLaunch.setDisable(selectedCount <= 0);
            btnLaunch.setText(RdpLaunchProfileSequentialRun.launchButtonTextIdle(selectedCount));
            applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, false);
            return;
        }
        btnLaunch.setDisable(false);
        btnLaunch.setText("リモートデスクトップを起動");
        applyQuickLaunchButtonDisabled(quickLaunchProfileButtons, false);
    }

    private static void applyQuickLaunchButtonDisabled(
            List<Button> quickLaunchProfileButtons, boolean disabled) {
        if (quickLaunchProfileButtons == null) {
            return;
        }
        for (Button quickBtn : quickLaunchProfileButtons) {
            if (quickBtn.isVisible()) {
                quickBtn.setDisable(disabled);
            }
        }
    }

    private static boolean rdpLaunchCredentialsReady(Supplier<Map<String, String>> uiEnv) {
        Map<String, String> ui = uiEnv != null ? uiEnv.get() : Map.of();
        FactorySite site = FactoryOperatorUserStore.operatorScopeForCurrentApp(ui, null);
        String operator = FactoryOperatorUserStore.sessionOperatorName();
        if (operator.isBlank() || FactoryOperatorUserStore.isGuestOperator(operator)) {
            return false;
        }
        try {
            FactoryOperatorUserStore.configureForCurrentApp(ui, site);
            return FactoryOperatorUserStore.hasAladdinCredentials(site, operator);
        } catch (IOException ex) {
            return false;
        }
    }

    private static String formatRdpSessionEndMessage(RdpMstscSessionMonitor.SessionEndEvent event) {
        return switch (event.reason()) {
            case MSTSC_EXIT ->
                    "リモートデスクトップ接続が終了しました（接続先 RPA 完了後のセッション終了操作を含む）。";
            case PROCESS_NOT_FOUND ->
                    "mstsc プロセスを特定できなかったため、接続終了を監視できませんでした: "
                            + event.rdpProfile();
            case MONITOR_INTERRUPTED -> "リモートデスクトップ接続の監視が中断されました。";
        };
    }

    private static void deleteSignOutStagingQuietly(RemoteDesktopLauncher.LaunchOutcome signOutOutcome) {
        if (signOutOutcome == null || signOutOutcome.signOutStagingFile().isEmpty()) {
            return;
        }
        try {
            Files.deleteIfExists(signOutOutcome.signOutStagingFile().get());
        } catch (IOException ignored) {
            // ignore
        }
    }

    private static void onSignOutOnlyLaunchEnded(
            RdpMstscSessionMonitor.SessionEndEvent event,
            RemoteDesktopLauncher.LaunchOutcome signOutOutcome,
            Label statusLabel,
            Consumer<String> status,
            AtomicBoolean rdpSessionActive,
            AtomicReference<RdpConnectPhase> rdpConnectPhase,
            Runnable updateLaunchButtonState,
            RdpRightPanePreviewController[] rightPanePreviewHolder,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshFetchedFilesTable,
            Runnable scrollToFetchedFilesSection,
            AtomicBoolean sequentialRunActive,
            AtomicReference<List<Integer>> sequentialQueue,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialSelectionOrder,
            Label sequentialStatusLabel,
            Runnable refreshQuickLaunchSequentialLabels,
            Button[] btnLaunchRef) {
        deleteSignOutStagingQuietly(signOutOutcome);
        if (rdpConnectPhase.get() != RdpConnectPhase.REMOTE_SIGN_OUT) {
            return;
        }
        onSignOutOnlyPhaseEnded(
                event,
                statusLabel,
                status,
                rdpSessionActive,
                rdpConnectPhase,
                updateLaunchButtonState,
                rightPanePreviewHolder,
                refreshIniFilePreview,
                refreshLauncherLog,
                refreshFetchedFilesTable,
                scrollToFetchedFilesSection,
                sequentialRunActive,
                sequentialQueue,
                sequentialIndex,
                sequentialSelectionOrder,
                sequentialStatusLabel,
                refreshQuickLaunchSequentialLabels,
                btnLaunchRef);
    }

    private static void launchRpaSessionAndWatch(
            Path preferredProfile,
            Path configuredProfile,
            Map<String, String> launchUi,
            int launchSlot,
            String launchProfileLabel,
            LaunchDisplay resolvedDisplay,
            boolean previewEnabled,
            Consumer<String> profileChangeHandler,
            TextField profileField,
            Label statusLabel,
            Consumer<String> status,
            AtomicBoolean rdpSessionActive,
            AtomicReference<RdpConnectPhase> rdpConnectPhase,
            AtomicReference<Thread> rdpSessionWatchThread,
            Runnable updateLaunchButtonState,
            RdpRightPanePreviewController[] rightPanePreviewHolder,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshFetchedFilesTable,
            Runnable scrollToFetchedFilesSection,
            AtomicBoolean sequentialRunActive,
            AtomicReference<List<Integer>> sequentialQueue,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialSelectionOrder,
            Label sequentialStatusLabel,
            Runnable refreshQuickLaunchSequentialLabels,
            Button[] btnLaunchRef) {
        try {
            if (previewEnabled
                    && rightPanePreviewHolder[0] != null
                    && rightPanePreviewHolder[0].isSupported()) {
                rightPanePreviewHolder[0].showPreviewPane();
            }
            RemoteDesktopLauncher.LaunchOutcome outcome =
                    RemoteDesktopLauncher.launch(preferredProfile, launchUi);
            if (!preferredProfile.equals(configuredProfile) && profileChangeHandler != null) {
                profileChangeHandler.accept(preferredProfile.toString());
                profileField.setText(preferredProfile.toString());
            }
            String msg = "RPA 用セッションを起動しました: " + outcome.rdpProfile();
            msg += "\n表示: " + RdpLaunchDisplaySettings.formatSummary(resolvedDisplay);
            msg +=
                    "\n起動プロファイル: "
                            + launchProfileLabel
                            + "（"
                            + RdpRemoteLauncherIni.SELECTED_SLOT_KEY
                            + "="
                            + launchSlot
                            + "）";
            if (outcome.remoteStartupSummary().isPresent()) {
                msg += "\n.rdp に組込: " + outcome.remoteStartupSummary().get();
            }
            if (outcome.signatureRemoved()) {
                msg +=
                        "\n（.rdp 編集により署名が無効化されました。必要なら署名ウィザードを実行してください）";
            }
            if (previewEnabled && rightPanePreviewHolder[0] != null) {
                msg += "\n右ペインへプレビュー表示中…（操作は別ウィンドウ）";
            } else {
                msg += "\n接続終了を監視中…（タスク完了時に ini のセッション操作を実行）";
            }
            if (sequentialRunActive.get()) {
                List<Integer> activeQueue = sequentialQueue.get();
                int activeIdx = sequentialIndex.get();
                if (activeQueue != null && !activeQueue.isEmpty()) {
                    msg +=
                            "\n"
                                    + RdpLaunchProfileSequentialRun.progressStatusText(
                                            activeIdx + 1, activeQueue.size(), launchProfileLabel);
                }
            }
            status.accept(msg);
            statusLabel.setText(msg);
            if (previewEnabled && rightPanePreviewHolder[0] != null) {
                rightPanePreviewHolder[0].previewWhenReady(
                        outcome.mstscProcessId(),
                        outcome.mstscPidMarkerFile(),
                        outcome.rdpProfile(),
                        previewMsg -> {
                            if (previewMsg != null && !previewMsg.isBlank()) {
                                status.accept(previewMsg);
                                statusLabel.setText(previewMsg);
                            }
                        });
            }
            RdpMstscSessionMonitor.watchAfterLaunch(
                    outcome.rdpProfile(),
                    outcome.mstscProcessId(),
                    outcome.mstscPidMarkerFile(),
                    rdpSessionWatchThread,
                    rpaEvent ->
                            Platform.runLater(
                                    () ->
                                            onRpaSessionEnded(
                                                    rpaEvent,
                                                    launchUi,
                                                    statusLabel,
                                                    status,
                                                    rdpSessionActive,
                                                    rdpConnectPhase,
                                                    updateLaunchButtonState,
                                                    rightPanePreviewHolder,
                                                    refreshIniFilePreview,
                                                    refreshLauncherLog,
                                                    refreshFetchedFilesTable,
                                                    scrollToFetchedFilesSection,
                                                    sequentialRunActive,
                                                    sequentialQueue,
                                                    sequentialIndex,
                                                    sequentialSelectionOrder,
                                                    sequentialStatusLabel,
                                                    refreshQuickLaunchSequentialLabels,
                                                    btnLaunchRef)));
        } catch (IOException ex) {
            rdpConnectPhase.set(RdpConnectPhase.IDLE);
            rdpSessionActive.set(false);
            sequentialRunActive.set(false);
            sequentialQueue.set(List.of());
            updateLaunchButtonState.run();
            if (rightPanePreviewHolder[0] != null) {
                rightPanePreviewHolder[0].removePreviewPane();
            }
            showAlert(Alert.AlertType.ERROR, "起動失敗", ex.getMessage());
            status.accept("RPA セッション起動失敗: " + ex.getMessage());
            statusLabel.setText(ex.getMessage());
        }
    }

    private static void onSignOutOnlyPhaseEnded(
            RdpMstscSessionMonitor.SessionEndEvent event,
            Label statusLabel,
            Consumer<String> status,
            AtomicBoolean rdpSessionActive,
            AtomicReference<RdpConnectPhase> rdpConnectPhase,
            Runnable updateLaunchButtonState,
            RdpRightPanePreviewController[] rightPanePreviewHolder,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshFetchedFilesTable,
            Runnable scrollToFetchedFilesSection,
            AtomicBoolean sequentialRunActive,
            AtomicReference<List<Integer>> sequentialQueue,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialSelectionOrder,
            Label sequentialStatusLabel,
            Runnable refreshQuickLaunchSequentialLabels,
            Button[] btnLaunchRef) {
        rdpConnectPhase.set(RdpConnectPhase.IDLE);
        rdpSessionActive.set(false);
        updateLaunchButtonState.run();
        if (rightPanePreviewHolder[0] != null) {
            rightPanePreviewHolder[0].removePreviewPane();
        }
        String msg;
        if (event.reason() == RdpMstscSessionMonitor.EndReason.MSTSC_EXIT) {
            msg = "接続先サインアウトが完了しました。";
        } else {
            msg =
                    "接続先サインアウトが完了しませんでした: "
                            + formatRdpSessionEndMessage(event);
        }
        status.accept(msg);
        statusLabel.setText(msg);
        refreshRightPanePreviews(refreshIniFilePreview, refreshLauncherLog);
        if (refreshFetchedFilesTable != null) {
            refreshFetchedFilesTable.run();
        }
        if (scrollToFetchedFilesSection != null) {
            scrollToFetchedFilesSection.run();
        }
        advanceSequentialRunAfterStep(
                msg,
                status,
                statusLabel,
                sequentialRunActive,
                sequentialQueue,
                sequentialIndex,
                sequentialSelectionOrder,
                sequentialStatusLabel,
                refreshQuickLaunchSequentialLabels,
                btnLaunchRef);
    }

    private static void advanceSequentialRunAfterStep(
            String endMsg,
            Consumer<String> status,
            Label statusLabel,
            AtomicBoolean sequentialRunActive,
            AtomicReference<List<Integer>> sequentialQueue,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialSelectionOrder,
            Label sequentialStatusLabel,
            Runnable refreshQuickLaunchSequentialLabels,
            Button[] btnLaunchRef) {
        if (sequentialRunActive == null || !sequentialRunActive.get()) {
            return;
        }
        List<Integer> activeQueue = sequentialQueue != null ? sequentialQueue.get() : null;
        int nextIdx = sequentialIndex != null ? sequentialIndex.get() + 1 : 0;
        if (activeQueue != null && nextIdx < activeQueue.size()) {
            sequentialIndex.set(nextIdx);
            status.accept(
                    endMsg
                            + "\n連続実行: 次のタスク（"
                            + (nextIdx + 1)
                            + "/"
                            + activeQueue.size()
                            + "）を起動します…");
            if (btnLaunchRef != null && btnLaunchRef[0] != null) {
                btnLaunchRef[0].fire();
            }
        } else {
            sequentialRunActive.set(false);
            if (sequentialQueue != null) {
                sequentialQueue.set(List.of());
            }
            if (sequentialSelectionOrder != null) {
                sequentialSelectionOrder.clear();
            }
            if (sequentialStatusLabel != null) {
                sequentialStatusLabel.setText("連続実行が完了しました。");
            }
            if (refreshQuickLaunchSequentialLabels != null) {
                refreshQuickLaunchSequentialLabels.run();
            }
            status.accept(endMsg + "\n連続実行が完了しました。");
            if (statusLabel != null) {
                statusLabel.setText(endMsg + "\n連続実行が完了しました。");
            }
        }
    }

    private static void onRpaSessionEnded(
            RdpMstscSessionMonitor.SessionEndEvent event,
            Map<String, String> launchUi,
            Label statusLabel,
            Consumer<String> status,
            AtomicBoolean rdpSessionActive,
            AtomicReference<RdpConnectPhase> rdpConnectPhase,
            Runnable updateLaunchButtonState,
            RdpRightPanePreviewController[] rightPanePreviewHolder,
            Runnable refreshIniFilePreview,
            Runnable refreshLauncherLog,
            Runnable refreshFetchedFilesTable,
            Runnable scrollToFetchedFilesSection,
            AtomicBoolean sequentialRunActive,
            AtomicReference<List<Integer>> sequentialQueue,
            AtomicInteger sequentialIndex,
            List<Integer> sequentialSelectionOrder,
            Label sequentialStatusLabel,
            Runnable refreshQuickLaunchSequentialLabels,
            Button[] btnLaunchRef) {
        rdpConnectPhase.set(RdpConnectPhase.IDLE);
        rdpSessionActive.set(false);
        updateLaunchButtonState.run();
        if (rightPanePreviewHolder[0] != null) {
            rightPanePreviewHolder[0].removePreviewPane();
        }
        try {
            Path iniPath = resolveRdpIniPath(launchUi);
            RdpRemoteLauncherIni.writeTaskSchedulerSuppress(iniPath, launchUi);
            RdpRemoteLauncherIni.clearSignOutOnConnectRequest(iniPath);
            status.accept(
                    "RPA設定.ini の起動プログラム番号を "
                                                    + RdpRemoteLauncherIni.INI_SIGN_OUT_SLOT
                            + " に設定しました（タスクスケジューラ抑止・保険）。");
            refreshRightPanePreviews(refreshIniFilePreview, refreshLauncherLog);
        } catch (IOException suppressEx) {
            status.accept(
                    "起動プログラム番号の抑止（0）設定に失敗: " + suppressEx.getMessage());
        }
        String endMsg = formatRdpSessionEndMessage(event);
        status.accept(endMsg);
        statusLabel.setText(endMsg);
        if (refreshFetchedFilesTable != null) {
            refreshFetchedFilesTable.run();
        }
        if (scrollToFetchedFilesSection != null) {
            scrollToFetchedFilesSection.run();
        }
        if (sequentialRunActive.get()) {
            advanceSequentialRunAfterStep(
                    endMsg,
                    status,
                    statusLabel,
                    sequentialRunActive,
                    sequentialQueue,
                    sequentialIndex,
                    sequentialSelectionOrder,
                    sequentialStatusLabel,
                    refreshQuickLaunchSequentialLabels,
                    btnLaunchRef);
        }
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
                if (RdpRemoteLauncherIni.isSignOutOnlyProfile(item)) {
                    ProfileRowFields row = findProfileRow(profileFields, item);
                    RdpLaunchProfile meta =
                            row != null
                                    ? buildProfileFromRow(row)
                                    : profileMetadataByNumber.getOrDefault(
                                            item, RdpLaunchProfile.signOutOnlyDefault());
                    setText(row != null ? profileComboLabel(row, meta) : meta.displayLabel());
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
        if (RdpRemoteLauncherIni.isSignOutOnlyProfile(profileNumber)) {
            ProfileRowFields row = findProfileRow(profileFields, profileNumber);
            RdpLaunchProfile meta =
                    row != null
                            ? buildProfileFromRow(row)
                            : profileMetadataByNumber.getOrDefault(
                                    profileNumber, RdpLaunchProfile.signOutOnlyDefault());
            String detail = meta.detailText();
            if (row != null && !row.programField().getText().isBlank()) {
                if (!detail.isBlank()) {
                    detail += '\n';
                }
                detail += "ランチャー引数: " + RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS;
            }
            detailLabel.setText(
                    detail.isBlank()
                            ? RdpRemoteLauncherIni.signOutOnlyProfileDetailText()
                            : detail);
            return;
        }
        ProfileRowFields row = findProfileRow(profileFields, profileNumber);
        RdpLaunchProfile meta =
                row != null
                        ? buildProfileFromRow(row)
                        : profileMetadataByNumber.getOrDefault(
                                profileNumber, RdpLaunchProfile.empty(profileNumber));
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

    private static boolean isProfileRowEmpty(ProfileRowFields row) {
        if (row == null) {
            return true;
        }
        return row.nameField().getText().trim().isEmpty()
                && row.descriptionField().getText().trim().isEmpty()
                && row.categoryField().getText().trim().isEmpty()
                && row.programField().getText().trim().isEmpty()
                && row.argsField().getText().trim().isEmpty()
                && !row.chkRpaEternal().isSelected();
    }

    private static void copyProfileRowFields(ProfileRowFields source, ProfileRowFields target) {
        target.nameField().setText(source.nameField().getText());
        target.descriptionField().setText(source.descriptionField().getText());
        target.categoryField().setText(source.categoryField().getText());
        target.programField().setText(source.programField().getText());
        target.argsField().setText(source.argsField().getText());
        target.chkRpaEternal().setSelected(source.chkRpaEternal().isSelected());
        target.chkFullScreen().setSelected(source.chkFullScreen().isSelected());
        target.desktopWidthSpinner()
                .getValueFactory()
                .setValue(source.desktopWidthSpinner().getValue());
        target.desktopHeightSpinner()
                .getValueFactory()
                .setValue(source.desktopHeightSpinner().getValue());
        updateDisplayControlsEnabled(
                target.chkFullScreen(),
                target.desktopWidthSpinner(),
                target.desktopHeightSpinner());
    }

    private static void clearProfileRowFields(ProfileRowFields row) {
        row.nameField().clear();
        row.descriptionField().clear();
        row.categoryField().clear();
        row.programField().clear();
        row.argsField().clear();
        row.chkRpaEternal().setSelected(false);
        row.chkFullScreen().setSelected(false);
        row.desktopWidthSpinner()
                .getValueFactory()
                .setValue(RdpLaunchDisplaySettings.DEFAULT_WIDTH);
        row.desktopHeightSpinner()
                .getValueFactory()
                .setValue(RdpLaunchDisplaySettings.DEFAULT_HEIGHT);
        updateDisplayControlsEnabled(
                row.chkFullScreen(), row.desktopWidthSpinner(), row.desktopHeightSpinner());
    }

    private static void copyProfileToEmptySlot(
            Window owner,
            int sourceNumber,
            List<ProfileRowFields> profileFields,
            Runnable refreshCombo) {
        ProfileRowFields source = findProfileRow(profileFields, sourceNumber);
        if (source == null) {
            return;
        }
        if (isProfileRowEmpty(source)) {
            showAlert(
                    Alert.AlertType.INFORMATION,
                    "コピーできません",
                    "プロファイル " + sourceNumber + " が空のため、コピーする内容がありません。");
            return;
        }
        List<Integer> emptyTargets = new ArrayList<>();
        for (ProfileRowFields row : profileFields) {
            if (row.number() != sourceNumber && isProfileRowEmpty(row)) {
                emptyTargets.add(row.number());
            }
        }
        if (emptyTargets.isEmpty()) {
            showAlert(
                    Alert.AlertType.INFORMATION,
                    "コピーできません",
                    "プロファイル "
                            + sourceNumber
                            + " をコピーできる空のプロファイルがありません。"
                            + " コピー先が空でない場合は上書きしません。");
            return;
        }
        ChoiceDialog<Integer> dialog =
                new ChoiceDialog<>(emptyTargets.getFirst(), emptyTargets);
        dialog.setTitle("プロファイルをコピー");
        dialog.setHeaderText("プロファイル " + sourceNumber + " → 空きプロファイル");
        dialog.setContentText("コピー先のプロファイル番号:");
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.showAndWait()
                .ifPresent(
                        targetNumber -> {
                            ProfileRowFields target = findProfileRow(profileFields, targetNumber);
                            if (target == null) {
                                return;
                            }
                            if (!isProfileRowEmpty(target)) {
                                showAlert(
                                        Alert.AlertType.WARNING,
                                        "コピーできません",
                                        "プロファイル "
                                                + targetNumber
                                                + " は空ではないため、上書きしませんでした。");
                                return;
                            }
                            copyProfileRowFields(source, target);
                            if (refreshCombo != null) {
                                refreshCombo.run();
                            }
                        });
    }

    private static void confirmAndSoftDeleteProfile(
            Window owner,
            Integer profileNumber,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber,
            Runnable refreshCombo,
            Runnable persist) {
        if (profileNumber == null) {
            return;
        }
        if (RdpRemoteLauncherIni.isSignOutOnlyProfile(profileNumber)) {
            showAlert(
                    Alert.AlertType.WARNING,
                    "削除できません",
                    "起動プロファイル 99（接続先サインアウト）は削除できません。");
            return;
        }
        ProfileRowFields row = findProfileRow(profileFields, profileNumber);
        if (row == null) {
            return;
        }
        List<RdpLaunchProfile> all =
                collectProfilesFromRows(profileFields, profileMetadataByNumber);
        if (!RdpLaunchProfileCatalog.canSoftDelete(all)) {
            showAlert(
                    Alert.AlertType.WARNING,
                    "削除できません",
                    "有効なプロファイルが1件しかないため削除できません。");
            return;
        }
        RdpLaunchProfile built = buildProfileFromRow(row);
        String label = profileComboLabel(row, built);
        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("プロファイルを削除");
        confirm.setHeaderText("プロファイル " + profileNumber + " を削除しますか？");
        confirm.setContentText(
                label
                        + " を一覧から非表示にします。"
                        + " RPA 設定は保持され、「削除したプロファイルを復活」から戻せます。");
        if (owner != null) {
            confirm.initOwner(owner);
        }
        confirm.showAndWait()
                .ifPresent(
                        choice -> {
                            if (choice != ButtonType.OK) {
                                return;
                            }
                            profileMetadataByNumber.put(profileNumber, built.withDeleted(true));
                            applyProfileCardVisibility(row, profileMetadataByNumber);
                            if (refreshCombo != null) {
                                refreshCombo.run();
                            }
                            if (persist != null) {
                                persist.run();
                            }
                        });
    }

    private static void showRestoreDeletedProfilesDialog(
            Window owner,
            List<ProfileRowFields> profileFields,
            Map<Integer, RdpLaunchProfile> profileMetadataByNumber,
            Runnable refreshCombo,
            Runnable persist) {
        List<RdpLaunchProfile> deleted =
                RdpLaunchProfileCatalog.deletedProfiles(
                        collectProfilesFromRows(profileFields, profileMetadataByNumber));
        if (deleted.isEmpty()) {
            showAlert(Alert.AlertType.INFORMATION, "復活", "削除済みのプロファイルはありません。");
            return;
        }
        List<Integer> choices =
                deleted.stream().map(RdpLaunchProfile::number).sorted().toList();
        List<String> labels = new ArrayList<>();
        Map<String, Integer> labelToNumber = new HashMap<>();
        for (Integer number : choices) {
            RdpLaunchProfile meta =
                    profileMetadataByNumber.getOrDefault(number, RdpLaunchProfile.empty(number));
            String label = meta.displayLabel();
            labels.add(label);
            labelToNumber.put(label, number);
        }
        ChoiceDialog<String> dialog = new ChoiceDialog<>(labels.getFirst(), labels);
        dialog.setTitle("削除したプロファイルを復活");
        dialog.setHeaderText("復活するプロファイルを選んでください");
        dialog.setContentText("プロファイル:");
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.showAndWait()
                .ifPresent(
                        label -> {
                            Integer number = labelToNumber.get(label);
                            if (number == null) {
                                return;
                            }
                            ProfileRowFields row = findProfileRow(profileFields, number);
                            if (row == null) {
                                return;
                            }
                            RdpLaunchProfile meta =
                                    profileMetadataByNumber.getOrDefault(
                                            number, buildProfileFromRow(row));
                            profileMetadataByNumber.put(number, meta.withDeleted(false));
                            applyProfileCardVisibility(row, profileMetadataByNumber);
                            if (refreshCombo != null) {
                                refreshCombo.run();
                            }
                            if (persist != null) {
                                persist.run();
                            }
                        });
    }

    private static boolean isProfileDeleted(
            ProfileRowFields row, Map<Integer, RdpLaunchProfile> profileMetadataByNumber) {
        RdpLaunchProfile meta = profileMetadataByNumber.get(row.number());
        return meta != null && meta.isDeleted();
    }

    private static void applyProfileCardVisibility(
            ProfileRowFields row, Map<Integer, RdpLaunchProfile> profileMetadataByNumber) {
        boolean deleted = isProfileDeleted(row, profileMetadataByNumber);
        row.card().setVisible(!deleted);
        row.card().setManaged(!deleted);
    }

    private static void confirmAndClearProfile(
            Window owner,
            int profileNumber,
            List<ProfileRowFields> profileFields,
            Runnable refreshCombo) {
        ProfileRowFields row = findProfileRow(profileFields, profileNumber);
        if (row == null) {
            return;
        }
        if (isProfileRowEmpty(row)) {
            return;
        }
        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("プロファイルをクリア");
        confirm.setHeaderText("プロファイル " + profileNumber + " の内容をクリアしますか？");
        confirm.setContentText("名称・説明・区分・RPA プログラム・RPA 引数・--eternal 設定が空になります。");
        if (owner != null) {
            confirm.initOwner(owner);
        }
        confirm.showAndWait()
                .ifPresent(
                        choice -> {
                            if (choice != ButtonType.OK) {
                                return;
                            }
                            clearProfileRowFields(row);
                            if (refreshCombo != null) {
                                refreshCombo.run();
                            }
                        });
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
                row.chkFullScreen().isSelected(),
                row.desktopWidthSpinner().getValue(),
                row.desktopHeightSpinner().getValue(),
                row.chkRpaEternal().isSelected(),
                null);
    }

    private static List<RdpLaunchProfile> collectProfilesFromRows(
            List<ProfileRowFields> rows, Map<Integer, RdpLaunchProfile> profileMetadataByNumber) {
        List<RdpLaunchProfile> profiles = new ArrayList<>();
        for (ProfileRowFields row : rows) {
            RdpLaunchProfile built = buildProfileFromRow(row);
            RdpLaunchProfile stored = profileMetadataByNumber.get(row.number());
            if (stored != null && stored.isDeleted()) {
                built = built.withDeleted(true);
            }
            profiles.add(built);
        }
        return profiles;
    }

    private static void applyProfileLaunchOverrides(
            Map<String, String> ui, RdpLaunchProfile profile) {
        LaunchDisplay display =
                RdpLaunchDisplaySettings.resolveLaunchDisplay(profile, ui);
        ui.putAll(RdpLaunchDisplaySettings.applyLaunchDisplayToUi(ui, display));
    }

    private static Spinner<Integer> createDesktopWidthSpinner(int initial) {
        Spinner<Integer> spinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                RdpLaunchDisplaySettings.MIN_WIDTH,
                                RdpLaunchDisplaySettings.MAX_WIDTH,
                                RdpLaunchDisplaySettings.clampWidth(initial),
                                10));
        spinner.setEditable(true);
        spinner.setPrefWidth(90);
        return spinner;
    }

    private static Spinner<Integer> createDesktopHeightSpinner(int initial) {
        Spinner<Integer> spinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                RdpLaunchDisplaySettings.MIN_HEIGHT,
                                RdpLaunchDisplaySettings.MAX_HEIGHT,
                                RdpLaunchDisplaySettings.clampHeight(initial),
                                10));
        spinner.setEditable(true);
        spinner.setPrefWidth(90);
        return spinner;
    }

    private static void wireProfileDisplayPreviewRefresh(
            ProfileRowFields row, Runnable[] refreshDisplayPreview) {
        Runnable refresh =
                () -> {
                    if (refreshDisplayPreview[0] != null) {
                        refreshDisplayPreview[0].run();
                    }
                };
        row.chkFullScreen().selectedProperty().addListener((obs, was, now) -> refresh.run());
        row.desktopWidthSpinner().valueProperty().addListener((obs, was, now) -> refresh.run());
        row.desktopHeightSpinner().valueProperty().addListener((obs, was, now) -> refresh.run());
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

    private static void showImportOperatorRpaIniDialog(
            Window owner,
            Supplier<Map<String, String>> uiEnv,
            ComboBox<Integer> launchProfileCombo,
            List<ProfileRowFields> profileFields,
            IntConsumer appendProfileRow,
            ComboBox<RdpSessionEndAction> cmbSessionEndAction,
            Label rapStatusLabel,
            Consumer<String> status,
            Runnable refreshLaunchProfileCombo,
            Runnable refreshDisplayPreview,
            Path preferredSource,
            boolean sharedIniMode) {
        Map<String, String> ui = uiEnv.get();
        String currentOperator = FactoryOperatorUserStore.resolveRdpLauncherOperatorName(ui);
        if (currentOperator.isBlank()) {
            showAlert(
                    Alert.AlertType.WARNING,
                    "操作者未選択",
                    "インポート先の操作者が未選択です。ツールバーから操作者を選んでから再度お試しください。");
            return;
        }
        Path targetIni = resolveRdpIniPath(ui).normalize();

        Dialog<ButtonType> dialog = new Dialog<>();
        if (sharedIniMode) {
            dialog.setTitle("共通 RPA設定.ini からインポート");
            dialog.setHeaderText(
                    "ユーザー名なしの共通 "
                            + AppPaths.RDP_LAUNCHER_INI_BASENAME
                            + " からプロファイルを、現在の操作者（"
                            + currentOperator
                            + "）の設定へ反映します。");
        } else {
            dialog.setTitle("他ユーザーの RPA設定.ini をインポート");
            dialog.setHeaderText(
                    "インポート元を選び内容を確認してから、現在の操作者（"
                            + currentOperator
                            + "）の設定へ反映します。");
        }
        if (owner != null) {
            dialog.initOwner(owner);
        }

        ComboBox<AppPaths.PeerOperatorRpaIniFile> sourceCombo = new ComboBox<>();
        sourceCombo.setPrefWidth(480);
        sourceCombo.setCellFactory(
                listView ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(AppPaths.PeerOperatorRpaIniFile item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    return;
                                }
                                setText(item.displayLabel() + "  —  " + item.path());
                            }
                        });
        sourceCombo.setButtonCell(sourceCombo.getCellFactory().call(null));

        Label sourcePathLabel = new Label();
        sourcePathLabel.getStyleClass().add("pm-rdp-page-subtitle");
        sourcePathLabel.setWrapText(true);
        sourcePathLabel.setMaxWidth(560);

        TextArea previewArea = new TextArea();
        previewArea.setEditable(false);
        previewArea.setWrapText(false);
        previewArea.setPrefRowCount(14);
        previewArea.setPromptText("インポート元 ini の内容がここに表示されます。");

        Label previewMetaLabel = new Label("読込待ち…");
        previewMetaLabel.getStyleClass().add("pm-rdp-page-subtitle");
        previewMetaLabel.setWrapText(true);
        previewMetaLabel.setMaxWidth(560);

        CheckBox chkImportSlots = new CheckBox("スロット定義（exe / 引数）");
        chkImportSlots.setSelected(true);
        CheckBox chkOnlyEmptySlots = new CheckBox("空のスロットのみ上書き（追加インポート）");
        chkOnlyEmptySlots.setSelected(true);
        chkOnlyEmptySlots.setDisable(false);
        chkImportSlots
                .selectedProperty()
                .addListener(
                        (obs, was, selected) -> chkOnlyEmptySlots.setDisable(!selected));
        CheckBox chkImportSelectedSlot = new CheckBox("起動プログラム番号");
        CheckBox chkImportSessionEnd = new CheckBox("終了時セッション操作");

        Label targetLabel = new Label("反映先: " + targetIni);
        targetLabel.getStyleClass().add("pm-rdp-meta-label");
        targetLabel.setWrapText(true);
        targetLabel.setMaxWidth(560);

        Button btnBrowseSource = new Button("参照…");
        styleSecondaryButton(btnBrowseSource);

        Runnable[] refreshPreview = new Runnable[1];
        refreshPreview[0] =
                () -> {
                    AppPaths.PeerOperatorRpaIniFile selected = sourceCombo.getValue();
                    if (selected == null) {
                        previewArea.clear();
                        previewMetaLabel.setText("インポート元を選択してください。");
                        sourcePathLabel.setText("");
                        return;
                    }
                    Path path = selected.path();
                    sourcePathLabel.setText(path.toString());
                    previewMetaLabel.setText("読込中…");
                    runBackgroundThenFx(
                            () -> {
                                String body;
                                String meta;
                                try {
                                    if (!Files.isRegularFile(path)) {
                                        body = "";
                                        meta = "ファイル未作成: " + path;
                                    } else {
                                        body = Files.readString(path, StandardCharsets.UTF_8);
                                        var modified =
                                                Files.getLastModifiedTime(path)
                                                        .toInstant()
                                                        .atZone(
                                                                java.time.ZoneId.systemDefault())
                                                        .toLocalDateTime()
                                                        .format(
                                                                java.time.format.DateTimeFormatter
                                                                        .ofPattern(
                                                                                "yyyy-MM-dd HH:mm:ss"));
                                        meta =
                                                Files.size(path)
                                                        + " bytes · 更新 "
                                                        + modified;
                                    }
                                } catch (IOException ex) {
                                    body = "";
                                    meta = "読込失敗: " + ex.getMessage();
                                }
                                final String previewBody = body;
                                final String previewMeta = meta;
                                Platform.runLater(
                                        () -> {
                                            previewArea.setText(previewBody);
                                            previewMetaLabel.setText(previewMeta);
                                        });
                            });
                };

        sourceCombo.valueProperty().addListener((obs, was, now) -> refreshPreview[0].run());

        btnBrowseSource.setOnAction(
                e -> {
                    FileChooser chooser = new FileChooser();
                    chooser.setTitle("インポート元 RPA設定.ini");
                    chooser
                            .getExtensionFilters()
                            .add(new FileChooser.ExtensionFilter("ini", "*.ini"));
                    Path deployDir = AppPaths.resolveRdpLauncherDeployDir(ui);
                    if (Files.isDirectory(deployDir)) {
                        chooser.setInitialDirectory(deployDir.toFile());
                    }
                    java.io.File picked =
                            chooser.showOpenDialog(owner != null ? owner : dialog.getDialogPane().getScene().getWindow());
                    if (picked == null) {
                        return;
                    }
                    Path pickedPath = picked.toPath().toAbsolutePath().normalize();
                    if (pickedPath.equals(targetIni)) {
                        showAlert(
                                Alert.AlertType.WARNING,
                                "同じファイル",
                                "現在の操作者の ini と同じファイルはインポート元に選べません。");
                        return;
                    }
                    AppPaths.PeerOperatorRpaIniFile custom =
                            new AppPaths.PeerOperatorRpaIniFile(
                                    AppPaths.displayLabelForRpaIniFilename(
                                            pickedPath.getFileName().toString())
                                            + "（参照）",
                                    pickedPath);
                    if (!sourceCombo.getItems().contains(custom)) {
                        sourceCombo.getItems().add(custom);
                    }
                    sourceCombo.setValue(custom);
                });

        List<AppPaths.PeerOperatorRpaIniFile> peers =
                AppPaths.listPeerOperatorRpaIniFiles(ui, currentOperator);
        sourceCombo.getItems().setAll(peers);
        if (preferredSource != null) {
            selectImportSourceComboValue(sourceCombo, preferredSource);
        } else if (!peers.isEmpty()) {
            sourceCombo.getSelectionModel().selectFirst();
        } else {
            previewMetaLabel.setText(
                    "配備先に他操作者の ini が見つかりません。「参照…」でファイルを選んでください。");
        }

        HBox sourceRow = new HBox(8, sourceCombo, btnBrowseSource);
        sourceRow.setAlignment(Pos.CENTER_LEFT);
        HBox.setHgrow(sourceCombo, Priority.ALWAYS);

        VBox content =
                new VBox(
                        10,
                        fieldCaption("インポート元"),
                        sourceRow,
                        sourcePathLabel,
                        fieldCaption("内容プレビュー"),
                        previewMetaLabel,
                        previewArea,
                        fieldCaption("取り込む項目"),
                        chkImportSlots,
                        chkOnlyEmptySlots,
                        chkImportSelectedSlot,
                        chkImportSessionEnd,
                        targetLabel);
        content.setPadding(new Insets(8, 0, 0, 0));
        content.setFillWidth(true);
        ScrollPane scroll = new ScrollPane(content);
        scroll.setFitToWidth(true);
        scroll.setPrefViewportHeight(520);
        dialog.getDialogPane().setContent(scroll);
        dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane().lookupButton(ButtonType.OK).setDisable(sourceCombo.getValue() == null);

        sourceCombo
                .valueProperty()
                .addListener(
                        (obs, was, now) ->
                                dialog.getDialogPane()
                                        .lookupButton(ButtonType.OK)
                                        .setDisable(now == null));

        if (sourceCombo.getValue() != null) {
            refreshPreview[0].run();
        }

        dialog.showAndWait()
                .ifPresent(
                        choice -> {
                            if (choice != ButtonType.OK) {
                                return;
                            }
                            AppPaths.PeerOperatorRpaIniFile selected = sourceCombo.getValue();
                            if (selected == null) {
                                return;
                            }
                            if (!chkImportSlots.isSelected()
                                    && !chkImportSelectedSlot.isSelected()
                                    && !chkImportSessionEnd.isSelected()) {
                                showAlert(
                                        Alert.AlertType.WARNING,
                                        "項目未選択",
                                        "取り込む項目を 1 つ以上選んでください。");
                                return;
                            }
                            try {
                                RdpRemoteLauncherIni source =
                                        RdpRemoteLauncherIni.load(selected.path());
                                ensureProfileRowsForImport(source, profileFields, appendProfileRow);
                                int importedSlots =
                                        applyImportFromPeerIni(
                                                source,
                                                profileFields,
                                                launchProfileCombo,
                                                cmbSessionEndAction,
                                                chkImportSlots.isSelected(),
                                                chkOnlyEmptySlots.isSelected(),
                                                chkImportSelectedSlot.isSelected(),
                                                chkImportSessionEnd.isSelected());
                                if (refreshLaunchProfileCombo != null) {
                                    refreshLaunchProfileCombo.run();
                                }
                                if (refreshDisplayPreview != null) {
                                    refreshDisplayPreview.run();
                                }
                                String message =
                                        "インポートしました（"
                                                + selected.displayLabel()
                                                + "）。スロット "
                                                + importedSlots
                                                + " 件を反映。"
                                                + " 「保存」で "
                                                + targetIni
                                                + " へ書き出してください。";
                                rapStatusLabel.setText(message);
                                status.accept(message);
                            } catch (IOException ex) {
                                showAlert(
                                        Alert.AlertType.ERROR,
                                        "読込失敗",
                                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
                            }
                        });
    }

    private static void selectImportSourceComboValue(
            ComboBox<AppPaths.PeerOperatorRpaIniFile> sourceCombo, Path preferredSource) {
        if (sourceCombo == null || preferredSource == null) {
            return;
        }
        Path normalized = preferredSource.toAbsolutePath().normalize();
        for (AppPaths.PeerOperatorRpaIniFile item : sourceCombo.getItems()) {
            if (item.path().normalize().equals(normalized)) {
                sourceCombo.setValue(item);
                return;
            }
        }
        AppPaths.PeerOperatorRpaIniFile entry =
                new AppPaths.PeerOperatorRpaIniFile(
                        AppPaths.displayLabelForRpaIniFilename(
                                normalized.getFileName().toString()),
                        normalized);
        sourceCombo.getItems().add(0, entry);
        sourceCombo.setValue(entry);
    }

    /** インポート元 ini のスロット数に合わせてプロファイル行を増やす。 */
    private static void ensureProfileRowsForImport(
            RdpRemoteLauncherIni source,
            List<ProfileRowFields> profileFields,
            IntConsumer appendProfileRow) {
        if (source == null || appendProfileRow == null) {
            return;
        }
        if (findProfileRow(profileFields, RdpRemoteLauncherIni.SLOT_SIGN_OUT) == null) {
            appendProfileRow.accept(RdpRemoteLauncherIni.SLOT_SIGN_OUT);
        }
        int needed =
                Math.min(
                        RdpRemoteLauncherIni.MAX_SLOTS,
                        Math.max(source.visibleSlotCount(), source.highestDefinedSlot()));
        int maxNumber = maxRpaProfileNumberFromRows(profileFields);
        while (maxNumber < needed) {
            appendProfileRow.accept(++maxNumber);
        }
    }

    private static int maxRpaProfileNumberFromRows(List<ProfileRowFields> profileFields) {
        return RdpRemoteLauncherIni.maxRpaProfileNumber(
                profileFields.stream().map(ProfileRowFields::number).toList());
    }

    /** @return 反映したスロット件数 */
    private static int applyImportFromPeerIni(
            RdpRemoteLauncherIni source,
            List<ProfileRowFields> profileFields,
            ComboBox<Integer> launchProfileCombo,
            ComboBox<RdpSessionEndAction> cmbSessionEndAction,
            boolean importSlots,
            boolean onlyEmptySlots,
            boolean importSelectedSlot,
            boolean importSessionEnd) {
        int importedSlotCount = 0;
        if (importSlots) {
            for (ProfileRowFields row : profileFields) {
                RdpRemoteLauncherIni.Command command = source.getSlotCommand(row.number());
                if (command.executable().isBlank()) {
                    continue;
                }
                if (onlyEmptySlots && !row.programField().getText().isBlank()) {
                    continue;
                }
                row.programField().setText(command.executable());
                String slotArguments = command.arguments();
                boolean eternalFromIni = RdpRemoteLauncherIni.hasEternalFlag(slotArguments);
                row.argsField()
                        .setText(
                                RdpRemoteLauncherIni.argumentsForUiDisplayWithoutManagedFlags(
                                        slotArguments));
                row.chkRpaEternal().setSelected(eternalFromIni);
                importedSlotCount++;
            }
        }
        if (importSelectedSlot) {
            int slot = source.selectedSlot();
            if (slot >= 1
                    && slot <= RdpRemoteLauncherIni.MAX_SLOTS
                    && launchProfileCombo.getItems().contains(slot)) {
                launchProfileCombo.setValue(slot);
            }
        }
        if (importSessionEnd) {
            cmbSessionEndAction.setValue(source.sessionEndAction());
        }
        return importedSlotCount;
    }

    private static void showExistenceCheckResult(
            String subject,
            RdpLaunchSharedPathValidator.Result result,
            Label rapStatusLabel,
            Consumer<String> status) {
        if (result.ok()) {
            String message = result.formatExistenceOkMessage();
            showAlert(Alert.AlertType.INFORMATION, subject + " 存在確認 OK", message);
            if (rapStatusLabel != null) {
                rapStatusLabel.setText(subject + " の存在を確認しました。");
            }
            if (status != null) {
                status.accept(subject + " 存在確認 OK\n" + message);
            }
            return;
        }
        String message = result.formatExistenceNgMessage();
        showAlert(Alert.AlertType.WARNING, subject + " 存在確認 NG", message);
        if (rapStatusLabel != null) {
            rapStatusLabel.setText(subject + " が見つかりません。");
        }
        if (status != null) {
            status.accept(subject + " 存在確認 NG\n" + message);
        }
    }

    private static <T> void installWrappingReadOnlyTableCell(TableColumn<T, String> column) {
        column.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Label label = new Label();

                            {
                                label.setWrapText(true);
                                label.setTextOverrun(OverrunStyle.CLIP);
                                label.getStyleClass().add("pm-rdp-fetched-files-path-label");
                                label.maxWidthProperty().bind(col.widthProperty().subtract(14));
                                label.prefWidthProperty().bind(label.maxWidthProperty());
                                label.setMaxHeight(Double.MAX_VALUE);
                                setAlignment(Pos.TOP_LEFT);
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    label.setText(null);
                                    setGraphic(null);
                                    setText(null);
                                    return;
                                }
                                label.setText(item);
                                setGraphic(label);
                                setText(null);
                            }
                        });
    }

    private static void scrollScrollPaneToNode(ScrollPane scrollPane, javafx.scene.Node target) {
        Platform.runLater(
                () -> {
                    if (scrollPane.getContent() == null || target == null) {
                        return;
                    }
                    javafx.geometry.Bounds viewport = scrollPane.getViewportBounds();
                    javafx.geometry.Bounds content = scrollPane.getContent().getBoundsInLocal();
                    double contentHeight = content.getHeight();
                    if (contentHeight <= viewport.getHeight() + 1) {
                        scrollPane.setVvalue(scrollPane.getVmax());
                        return;
                    }
                    javafx.geometry.Bounds targetBounds =
                            scrollPane.getContent()
                                    .sceneToLocal(target.localToScene(target.getBoundsInLocal()));
                    double offset = targetBounds.getMinY();
                    double range = contentHeight - viewport.getHeight();
                    if (range <= 0) {
                        return;
                    }
                    double v = offset / range;
                    scrollPane.setVvalue(
                            Math.min(scrollPane.getVmax(), Math.max(0, v)));
                });
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

    /**
     * {@link AppPaths#WINDOWS_DEFAULT_RDP_FILENAME} が解決結果のとき警告し、署名ウィザードを案内する。
     *
     * @return 起動を中止すべきとき {@code true}
     */
    private static boolean warnDefaultRdpBlockedAndMaybeOpenSignWizard(
            Window owner,
            Path resolvedProfile,
            Map<String, String> uiEnv,
            Consumer<String> status,
            Label statusLabel,
            Consumer<String> profileChangeHandler,
            TextField profileField) {
        if (!RemoteDesktopLauncher.isDefaultRdpLaunchBlocked(resolvedProfile)) {
            return false;
        }
        ButtonType openWizard =
                new ButtonType("署名ウィザードを開く", ButtonBar.ButtonData.OK_DONE);
        Alert alert =
                new Alert(
                        Alert.AlertType.WARNING,
                        RemoteDesktopLauncher.DEFAULT_RDP_LAUNCH_BLOCKED_MESSAGE,
                        openWizard,
                        ButtonType.CANCEL);
        alert.setTitle("Default.rdp は使用できません");
        alert.setHeaderText("署名済みプロファイルを作成してください");
        Label body = new Label(RemoteDesktopLauncher.DEFAULT_RDP_LAUNCH_BLOCKED_MESSAGE);
        body.setWrapText(true);
        body.setMaxWidth(520);
        body.setMinWidth(360);
        alert.getDialogPane().setContent(body);
        Optional<ButtonType> choice = alert.showAndWait();
        if (choice.isPresent() && choice.get() == openWizard) {
            RdpProfileSignWizard.show(
                    owner,
                    Optional.of(resolvedProfile),
                    uiEnv,
                    msg -> {
                        status.accept(msg);
                        statusLabel.setText(msg);
                    },
                    path -> {
                        profileField.setText(path);
                        if (profileChangeHandler != null) {
                            profileChangeHandler.accept(path);
                        }
                    });
        }
        status.accept("Default.rdp のためリモートデスクトップを起動しませんでした。");
        statusLabel.setText("Default.rdp は使用できません。署名済みプロファイルを作成してください。");
        return true;
    }

    /** 配布用アプリでは {@link AppPaths#KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR} の明示指定が必須。 */
    private static boolean isExplicitLauncherDeployDirConfigured(Map<String, String> ui) {
        if (!RemoteDesktopStandaloneBootstrap.isActivated()) {
            return true;
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        return !u.getOrDefault(AppPaths.KEY_PM_AI_RPA_LAUNCHER_DEPLOY_DIR, "").isBlank()
                || !u.getOrDefault(AppPaths.KEY_PM_AI_RDP_LAUNCHER_DEPLOY_DIR, "").isBlank();
    }
}
