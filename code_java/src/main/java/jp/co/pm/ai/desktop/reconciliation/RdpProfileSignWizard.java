package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.PmAiFxApp;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.RdpFileSigner;
import jp.co.pm.ai.desktop.io.RdpFileSigner.SigningCertificate;

/**
 * .rdp プロファイルへのデジタル署名と、GPO 信頼設定手順を案内するウィザード。
 */
public final class RdpProfileSignWizard {

    private static final double WIZARD_WIDTH = 760;

    private RdpProfileSignWizard() {}

    public static void show(
            Window owner,
            Optional<Path> initialRdp,
            Consumer<String> statusConsumer,
            Consumer<String> profileChangeHandler) {
        show(owner, initialRdp, Map.of(), statusConsumer, profileChangeHandler);
    }

    public static void show(
            Window owner,
            Optional<Path> initialRdp,
            Map<String, String> uiEnv,
            Consumer<String> statusConsumer,
            Consumer<String> profileChangeHandler) {
        if (!RdpFileSigner.isSupportedPlatform()) {
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("未対応");
            alert.setHeaderText(null);
            alert.setContentText("RDP 署名ウィザードは Windows 上のデスクトップアプリでのみ利用できます。");
            alert.showAndWait();
            return;
        }

        Stage stage = new Stage();
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("RDP プロファイル署名ウィザード");

        AtomicReference<Path> rdpPath = new AtomicReference<>(initialRdp.orElse(null));
        Map<String, String> ui = uiEnv != null ? uiEnv : Map.of();
        AtomicReference<SigningCertificate> selectedCert = new AtomicReference<>();
        AtomicReference<RdpFileSigner.CertificateListResult> certQuery =
                new AtomicReference<>(new RdpFileSigner.CertificateListResult(List.of(), 0));

        Label stepBadge = new Label("ステップ 1 / 4");
        stepBadge.getStyleClass().add("overtime-sim-step-badge");

        Label titleLabel = new Label("RDP プロファイル署名");
        titleLabel.getStyleClass().add("overtime-sim-title");

        Region headerSpacer = new Region();
        HBox.setHgrow(headerSpacer, Priority.ALWAYS);
        HBox header = new HBox(12, titleLabel, headerSpacer, stepBadge);
        header.setAlignment(Pos.CENTER_LEFT);
        header.getStyleClass().add("overtime-sim-header");

        VBox stepHost = new VBox(10);
        stepHost.setFillWidth(true);
        stepHost.setMaxWidth(Double.MAX_VALUE);

        Button btnBack = new Button("戻る");
        btnBack.getStyleClass().add("btn-reload");
        Button btnNext = new Button("次へ");
        btnNext.getStyleClass().add("btn-reload");
        Button btnCancel = new Button("キャンセル");
        btnCancel.getStyleClass().add("btn-reload");
        Button btnFinish = new Button("閉じる");
        btnFinish.getStyleClass().add("btn-reload");
        btnFinish.setVisible(false);
        btnFinish.setManaged(false);

        final int[] step = {0};

        Runnable refreshStep =
                () -> {
                    stepHost.getChildren().clear();
                    stepBadge.setText("ステップ " + (step[0] + 1) + " / 4");
                    btnBack.setDisable(step[0] == 0);
                    btnNext.setVisible(step[0] < 3);
                    btnNext.setManaged(step[0] < 3);
                    btnFinish.setVisible(step[0] == 3);
                    btnFinish.setManaged(step[0] == 3);
                    switch (step[0]) {
                        case 0 -> buildStepFile(stage, stepHost, rdpPath, ui, statusConsumer);
                        case 1 -> buildStepCertificate(stage, stepHost, certQuery, selectedCert);
                        case 2 ->
                                buildStepSign(
                                        stepHost, rdpPath, ui, selectedCert, statusConsumer, profileChangeHandler);
                        case 3 -> buildStepTrust(stepHost, rdpPath, ui, selectedCert, statusConsumer);
                        default -> {}
                    }
                };

        btnBack.setOnAction(e -> {
            if (step[0] > 0) {
                step[0]--;
                refreshStep.run();
            }
        });

        btnNext.setOnAction(
                e -> {
                    if (!validateStep(step[0], rdpPath.get(), selectedCert.get())) {
                        return;
                    }
                    if (step[0] < 3) {
                        step[0]++;
                        refreshStep.run();
                    }
                });

        btnCancel.setOnAction(e -> stage.close());
        btnFinish.setOnAction(e -> stage.close());

        HBox footer = new HBox(8, btnBack, btnNext, btnFinish, btnCancel);
        footer.setAlignment(Pos.CENTER_RIGHT);
        footer.setPadding(new Insets(8, 0, 0, 0));

        BorderPane root = new BorderPane();
        root.setTop(header);
        root.setCenter(stepHost);
        root.setBottom(footer);
        BorderPane.setMargin(stepHost, new Insets(12, 0, 0, 0));
        root.setPadding(new Insets(12));
        root.setPrefWidth(WIZARD_WIDTH);
        root.setMaxWidth(WIZARD_WIDTH);

        Scene scene = new Scene(root);
        var css = PmAiFxApp.class.getResource("/jp/co/pm/ai/desktop/css/pm-ai-desktop.css");
        if (css != null) {
            scene.getStylesheets().add(css.toExternalForm());
        }
        stage.setScene(scene);
        stage.setMinWidth(WIZARD_WIDTH);

        refreshStep.run();
        loadCertificatesAsync(stage, certQuery, refreshStep, null);
        stage.showAndWait();
    }

    private static boolean validateStep(int currentStep, Path rdp, SigningCertificate cert) {
        if (currentStep == 0) {
            if (rdp == null || !Files.isRegularFile(rdp)) {
                showAlert(Alert.AlertType.WARNING, "ファイル未選択", ".rdp プロファイルを選択してください。");
                return false;
            }
            try {
                jp.co.pm.ai.desktop.io.RemoteDesktopLauncher.validateRdpProfile(rdp);
            } catch (IOException ex) {
                showAlert(Alert.AlertType.ERROR, "ファイル不正", ex.getMessage());
                return false;
            }
        }
        if (currentStep == 1 && cert == null) {
            showAlert(Alert.AlertType.WARNING, "証明書未選択", "署名に使う証明書を選択してください。");
            return false;
        }
        return true;
    }

    private static void buildStepFile(
            Stage stage,
            VBox host,
            AtomicReference<Path> rdpPath,
            Map<String, String> ui,
            Consumer<String> statusConsumer) {
        Label intro =
                new Label(
                        "署名対象の .rdp プロファイルを選びます。"
                                + " 署名後は Windows のセキュリティ警告（不明な発行元）が解消されやすくなります。"
                                + " 元 .rdp は変更せず、リポジトリルートに "
                                + RdpFileSigner.SIGNED_OUTPUT_SUFFIX
                                + " を新規作成します（再署名時のみ出力ファイルを上書き）。");
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);

        TextField pathField = new TextField();
        pathField.setEditable(false);
        HBox.setHgrow(pathField, Priority.ALWAYS);
        updatePathField(pathField, rdpPath.get());

        Label signedLabel = new Label();
        signedLabel.setWrapText(true);
        signedLabel.setMaxWidth(Double.MAX_VALUE);
        refreshSignedLabel(signedLabel, rdpPath.get());
        if (rdpPath.get() != null) {
            Path output = RdpFileSigner.resolveSignedOutputPath(rdpPath.get(), ui);
            signedLabel.setText(
                    signedLabel.getText()
                            + "\n出力先（リポジトリルート）: "
                            + output
                            + "（元ファイル "
                            + rdpPath.get()
                            + " は変更しません）");
        }

        Button btnBrowse = new Button("参照...");
        btnBrowse.getStyleClass().add("btn-reload");
        btnBrowse.setOnAction(
                e -> {
                    FileChooser chooser = new FileChooser();
                    chooser.setTitle("署名する RDP プロファイル (.rdp)");
                    chooser.getExtensionFilters()
                            .add(new FileChooser.ExtensionFilter("リモートデスクトップ接続 (*.rdp)", "*.rdp"));
                    Path current = rdpPath.get();
                    if (current != null) {
                        Path parent = current.getParent();
                        if (parent != null && Files.isDirectory(parent)) {
                            chooser.setInitialDirectory(parent.toFile());
                        }
                    }
                    java.io.File chosen = chooser.showOpenDialog(stage);
                    if (chosen == null) {
                        return;
                    }
                    Path abs = chosen.toPath().toAbsolutePath().normalize();
                    rdpPath.set(abs);
                    updatePathField(pathField, abs);
                    refreshSignedLabel(signedLabel, abs);
                    if (statusConsumer != null) {
                        statusConsumer.accept("署名対象 RDP: " + abs);
                    }
                });

        HBox pathRow = new HBox(8, pathField, btnBrowse);
        pathRow.setAlignment(Pos.CENTER_LEFT);
        host.getChildren().addAll(intro, pathRow, signedLabel);
    }

    private static void buildStepCertificate(
            Stage stage,
            VBox host,
            AtomicReference<RdpFileSigner.CertificateListResult> certQuery,
            AtomicReference<SigningCertificate> selectedCert) {
        RdpFileSigner.CertificateListResult query =
                certQuery.get() != null ? certQuery.get() : new RdpFileSigner.CertificateListResult(List.of(), 0);
        List<SigningCertificate> eligible = query.eligible();

        Label intro =
                new Label(
                        "rdpsign.exe が使える証明書のみ表示します（コード署名 EKU、"
                                + "または TLS サーバー認証以外のデジタル署名）。"
                                + " CN=localhost など開発用証明書だけ成功していた場合、"
                                + " 他の証明書は SSL/TLS 用で署名できません。"
                                + " 一覧が空のときは下のボタンで RDP 署名専用証明書を作成してください。");
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);

        ComboBox<SigningCertificate> combo = new ComboBox<>();
        combo.setMaxWidth(Double.MAX_VALUE);
        combo.getItems().setAll(eligible);
        if (!eligible.isEmpty()) {
            combo.getSelectionModel().selectFirst();
            selectedCert.set(combo.getValue());
        } else {
            selectedCert.set(null);
        }
        combo.setOnAction(e -> selectedCert.set(combo.getValue()));

        Label emptyHint = new Label(formatCertHint(query));
        emptyHint.setWrapText(true);
        emptyHint.setMaxWidth(Double.MAX_VALUE);

        Button btnCreate = new Button("RDP署名用証明書を作成");
        btnCreate.getStyleClass().add("btn-reload");
        btnCreate.setTooltip(
                new Tooltip(
                        "コード署名 EKU 付きの自己署名証明書を CurrentUser\\My に作成します（有効 3 年）。"));
        btnCreate.setOnAction(
                e ->
                        createSigningCertificateAsync(
                                stage,
                                certQuery,
                                selectedCert,
                                combo,
                                emptyHint,
                                "湖南工場 RDP Signing"));

        Button btnRefresh = new Button("証明書を再読込");
        btnRefresh.getStyleClass().add("btn-reload");
        btnRefresh.setOnAction(
                e ->
                        loadCertificatesAsync(
                                stage,
                                certQuery,
                                () -> applyCertQueryToCombo(certQuery, selectedCert, combo, emptyHint),
                                combo));

        HBox actions = new HBox(8, btnCreate, btnRefresh);
        actions.setAlignment(Pos.CENTER_LEFT);
        host.getChildren().addAll(intro, combo, emptyHint, actions);
    }

    private static String formatCertHint(RdpFileSigner.CertificateListResult query) {
        if (query == null || query.eligible().isEmpty()) {
            int skipped = query != null ? query.skippedIneligibleCount() : 0;
            if (skipped > 0) {
                return "RDP署名に使える証明書がありません（"
                        + skipped
                        + " 件は SSL/TLS・メール等のため除外）。"
                        + " 「RDP署名用証明書を作成」を実行してください。";
            }
            return "RDP署名に使える証明書がありません。「RDP署名用証明書を作成」を実行してください。";
        }
        String base = "RDP署名に使える証明書 " + query.eligible().size() + " 件";
        if (query.skippedIneligibleCount() > 0) {
            base += "（SSL/TLS 等 " + query.skippedIneligibleCount() + " 件は一覧から除外）";
        }
        return base;
    }

    private static void applyCertQueryToCombo(
            AtomicReference<RdpFileSigner.CertificateListResult> certQuery,
            AtomicReference<SigningCertificate> selectedCert,
            ComboBox<SigningCertificate> combo,
            Label emptyHint) {
        applyCertQueryToCombo(certQuery, selectedCert, combo, emptyHint, null);
    }

    private static void applyCertQueryToCombo(
            AtomicReference<RdpFileSigner.CertificateListResult> certQuery,
            AtomicReference<SigningCertificate> selectedCert,
            ComboBox<SigningCertificate> combo,
            Label emptyHint,
            String preferredThumbprintSha1) {
        RdpFileSigner.CertificateListResult query =
                certQuery.get() != null ? certQuery.get() : new RdpFileSigner.CertificateListResult(List.of(), 0);
        combo.getItems().setAll(query.eligible());
        SigningCertificate pick = null;
        if (preferredThumbprintSha1 != null && !preferredThumbprintSha1.isBlank()) {
            String want = preferredThumbprintSha1.toUpperCase(Locale.ROOT);
            for (SigningCertificate c : query.eligible()) {
                if (want.equals(c.thumbprintSha1())) {
                    pick = c;
                    break;
                }
            }
        }
        if (pick == null && !query.eligible().isEmpty()) {
            pick = query.eligible().getFirst();
        }
        if (pick != null) {
            combo.getSelectionModel().select(pick);
            selectedCert.set(pick);
        } else {
            selectedCert.set(null);
        }
        emptyHint.setText(formatCertHint(query));
    }

    private static void createSigningCertificateAsync(
            Stage stage,
            AtomicReference<RdpFileSigner.CertificateListResult> certQuery,
            AtomicReference<SigningCertificate> selectedCert,
            ComboBox<SigningCertificate> combo,
            Label emptyHint,
            String commonName) {
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                SigningCertificate created =
                                        RdpFileSigner.createRdpSigningCertificate(commonName);
                                RdpFileSigner.CertificateListResult refreshed =
                                        RdpFileSigner.listSigningCertificates();
                                RdpFileSigner.CertificateListResult merged =
                                        refreshed.withEnsuredEligible(created);
                                Platform.runLater(
                                        () -> {
                                            certQuery.set(merged);
                                            applyCertQueryToCombo(
                                                    certQuery,
                                                    selectedCert,
                                                    combo,
                                                    emptyHint,
                                                    created.thumbprintSha1());
                                            showAlert(
                                                    Alert.AlertType.INFORMATION,
                                                    "証明書作成",
                                                    "RDP 署名用証明書を作成しました:\n"
                                                            + created.subject()
                                                            + "\n一覧に追加済みです。");
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                showAlert(
                                                        Alert.AlertType.ERROR,
                                                        "証明書作成失敗",
                                                        ex.getMessage() != null ? ex.getMessage() : ex.toString()));
                            }
                        },
                        "rdp-cert-create");
        worker.setDaemon(true);
        worker.start();
    }

    private static String formatRdpsignWorkDirLabel(Map<String, String> ui) {
        try {
            return RdpFileSigner.resolveRdpsignWorkDir(ui).toString();
        } catch (IOException ex) {
            return "（リポジトリルート）";
        }
    }

    private static String formatElevatedWorkDirLabel() {
        try {
            return RdpFileSigner.resolveRdpsignElevatedWorkDir().toString();
        } catch (IOException ex) {
            return "%ProgramData%\\PM-AI\\rdp-sign";
        }
    }

    private static void buildStepSign(
            VBox host,
            AtomicReference<Path> rdpPath,
            Map<String, String> ui,
            AtomicReference<SigningCertificate> selectedCert,
            Consumer<String> statusConsumer,
            Consumer<String> profileChangeHandler) {
        SigningCertificate cert = selectedCert.get();
        Path rdp = rdpPath.get();
        String signingNote = "";
        String rdpsignWorkDir = formatRdpsignWorkDirLabel(ui);
        String elevatedWorkDir = formatElevatedWorkDirLabel();
        if (rdp != null) {
            signingNote =
                    " ステージング: "
                            + rdpsignWorkDir
                            + "\\"
                            + RdpFileSigner.RDPSIGN_WORK_FILENAME
                            + " / UAC rdpsign: "
                            + elevatedWorkDir
                            + " / 出力（リポジトリ）: "
                            + RdpFileSigner.resolveSignedOutputPath(rdp, ui);
        }
        Label intro =
                new Label(
                        "テスト署名（/l）で確認してから本署名を実行します。"
                                + " 本署名は UAC 昇格 PowerShell が ProgramData\\PM-AI\\rdp-sign 上で rdpsign し、"
                                + " 成功後にリポジトリルートへコピーします。"
                                + " 対象: "
                                + (rdp != null ? rdp : "（未選択）")
                                + signingNote
                                + " / 証明書: "
                                + (cert != null ? cert.subject() : "（未選択）"));
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);

        CheckBox backupCheck =
                new CheckBox(
                        "再署名時、既存の "
                                + RdpFileSigner.SIGNED_OUTPUT_SUFFIX
                                + " を .unsigned-タイムスタンプ.bak として保存する");
        backupCheck.setSelected(true);

        TextArea logArea = new TextArea();
        logArea.setEditable(false);
        logArea.setPrefRowCount(8);
        logArea.setWrapText(true);
        VBox.setVgrow(logArea, Priority.ALWAYS);

        Button btnTest = new Button("テスト署名");
        btnTest.getStyleClass().add("btn-reload");
        Button btnSign = new Button("本署名を実行（UAC）");
        btnSign.getStyleClass().add("btn-reload");

        btnTest.setOnAction(
                e ->
                        runSignAction(
                                rdpPath,
                                ui,
                                selectedCert.get(),
                                true,
                                false,
                                logArea,
                                statusConsumer,
                                profileChangeHandler));

        btnSign.setOnAction(
                e ->
                        runSignAction(
                                rdpPath,
                                ui,
                                selectedCert.get(),
                                false,
                                backupCheck.isSelected(),
                                logArea,
                                statusConsumer,
                                profileChangeHandler));

        HBox actions = new HBox(8, btnTest, btnSign);
        actions.setAlignment(Pos.CENTER_LEFT);
        host.getChildren().addAll(intro, backupCheck, actions, logArea);
    }

    private static void buildStepTrust(
            VBox host,
            AtomicReference<Path> rdpPath,
            Map<String, String> ui,
            AtomicReference<SigningCertificate> selectedCert,
            Consumer<String> statusConsumer) {
        SigningCertificate cert = selectedCert.get();
        String thumb = cert != null ? cert.thumbprintSha1() : "";

        Label intro =
                new Label(
                        "署名だけでは警告が残ります。"
                                + " mstsc は HKCU と HKLM の両方を参照するため、"
                                + " この PC では HKCU + HKLM への登録を推奨します。");
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);

        TextArea diagnoseArea = new TextArea();
        diagnoseArea.setEditable(false);
        diagnoseArea.setPrefRowCount(7);
        diagnoseArea.setWrapText(true);
        Runnable refreshDiagnosis =
                () -> {
                    Path profile = rdpPath != null ? rdpPath.get() : null;
                    if (profile == null) {
                        diagnoseArea.setText("（署名済み .rdp が未選択です）");
                        return;
                    }
                    try {
                        Path preferred =
                                RdpFileSigner.resolvePreferredSignedProfilePath(profile, ui);
                        RdpFileSigner.ProfileTrustDiagnosis diagnosis =
                                RdpFileSigner.diagnoseProfileTrust(preferred, thumb);
                        diagnoseArea.setText(diagnosis.summary());
                    } catch (Exception ex) {
                        diagnoseArea.setText(
                                ex.getMessage() != null ? ex.getMessage() : ex.toString());
                    }
                };
        refreshDiagnosis.run();

        TextField thumbField = new TextField(thumb);
        thumbField.setEditable(false);
        HBox.setHgrow(thumbField, Priority.ALWAYS);

        Button btnCopyThumb = new Button("サムプリントをコピー");
        btnCopyThumb.getStyleClass().add("btn-reload");
        btnCopyThumb.setOnAction(e -> copyToClipboard(thumb, "SHA-1 サムプリント"));

        TextArea scriptArea = new TextArea();
        if (!thumb.isBlank()) {
            scriptArea.setText(RdpFileSigner.buildTrustedPublisherRegistryScript(thumb, true));
        }
        scriptArea.setEditable(false);
        scriptArea.setPrefRowCount(5);
        scriptArea.setWrapText(true);

        Button btnCopyScript = new Button("管理者 PowerShell スクリプトをコピー");
        btnCopyScript.getStyleClass().add("btn-reload");
        btnCopyScript.setOnAction(e -> copyToClipboard(scriptArea.getText(), "PowerShell スクリプト"));

        Button btnApplyBoth = new Button("信頼設定を適用（HKCU + HKLM・推奨）");
        btnApplyBoth.getStyleClass().add("btn-reload");
        btnApplyBoth.setDisable(thumb.isBlank());
        btnApplyBoth.setOnAction(
                e -> runTrustPolicyAction(thumb, true, true, refreshDiagnosis, statusConsumer));

        Button btnApplyUser = new Button("現在ユーザー（HKCU）のみ");
        btnApplyUser.getStyleClass().add("btn-reload");
        btnApplyUser.setDisable(thumb.isBlank());
        btnApplyUser.setOnAction(
                e -> runTrustPolicyAction(thumb, false, false, refreshDiagnosis, statusConsumer));

        Button btnRefreshDiag = new Button("診断を更新");
        btnRefreshDiag.getStyleClass().add("btn-reload");
        btnRefreshDiag.setOnAction(e -> refreshDiagnosis.run());

        Label gpoHint =
                new Label(
                        "GPO パス: コンピュータの構成 → 管理用テンプレート → Windows コンポーネント → "
                                + "Remote Desktop Services → Remote Desktop Connection Client → "
                                + "「信頼できる .rdp 発行元を表す証明書の SHA1 サムプリントを指定する」");
        gpoHint.setWrapText(true);
        gpoHint.setMaxWidth(Double.MAX_VALUE);

        HBox thumbRow = new HBox(8, thumbField, btnCopyThumb);
        thumbRow.setAlignment(Pos.CENTER_LEFT);
        HBox applyRow = new HBox(8, btnApplyBoth, btnApplyUser, btnRefreshDiag);
        applyRow.setAlignment(Pos.CENTER_LEFT);
        host.getChildren()
                .addAll(intro, thumbRow, applyRow, diagnoseArea, gpoHint, scriptArea, btnCopyScript);
    }

    private static void runTrustPolicyAction(
            String thumbprint,
            boolean machineWide,
            boolean allScopes,
            Runnable afterSuccess,
            Consumer<String> statusConsumer) {
        if (thumbprint == null || thumbprint.isBlank()) {
            showAlert(Alert.AlertType.WARNING, "未選択", "署名に使った証明書がありません。");
            return;
        }
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                RdpFileSigner.CommandResult result =
                                        allScopes
                                                ? RdpFileSigner.applyTrustedPublisherPolicyAllScopes(
                                                        thumbprint)
                                                : RdpFileSigner.applyTrustedPublisherPolicy(
                                                        thumbprint, machineWide);
                                String scope =
                                        allScopes
                                                ? "HKCU + HKLM"
                                                : machineWide ? "全ユーザー（HKLM）" : "現在ユーザー（HKCU）";
                                String msg =
                                        result.success()
                                                ? "RDP 信頼設定を適用しました: " + scope
                                                : "RDP 信頼設定の適用に失敗しました（"
                                                        + scope
                                                        + " 終了コード "
                                                        + result.exitCode()
                                                        + "）"
                                                        + (result.output().isBlank()
                                                                ? ""
                                                                : "\n" + result.output());
                                Platform.runLater(
                                        () -> {
                                            if (result.success()) {
                                                if (afterSuccess != null) {
                                                    afterSuccess.run();
                                                }
                                                if (statusConsumer != null) {
                                                    statusConsumer.accept(msg);
                                                }
                                                showAlert(Alert.AlertType.INFORMATION, "信頼設定", msg);
                                            } else if (result.exitCode() == RdpFileSigner.UAC_CANCELLED_EXIT_CODE) {
                                                showAlert(
                                                        Alert.AlertType.WARNING,
                                                        "UAC キャンセル",
                                                        "管理者権限の確認がキャンセルされました。");
                                            } else {
                                                showAlert(Alert.AlertType.ERROR, "信頼設定失敗", msg);
                                            }
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                showAlert(
                                                        Alert.AlertType.ERROR,
                                                        "信頼設定失敗",
                                                        ex.getMessage() != null
                                                                ? ex.getMessage()
                                                                : ex.toString()));
                            }
                        },
                        "rdp-trust-policy");
        worker.setDaemon(true);
        worker.start();
    }

    private static void runSignAction(
            AtomicReference<Path> rdpPathRef,
            Map<String, String> ui,
            SigningCertificate cert,
            boolean testOnly,
            boolean backup,
            TextArea logArea,
            Consumer<String> statusConsumer,
            Consumer<String> profileChangeHandler) {
        Path rdp = rdpPathRef != null ? rdpPathRef.get() : null;
        if (rdp == null || cert == null) {
            showAlert(Alert.AlertType.WARNING, "未選択", "RDP ファイルと証明書を選んでください。");
            return;
        }
        logArea.setText("実行中...");
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                RdpFileSigner.SignAttemptResult attempt =
                                        RdpFileSigner.attemptSign(
                                                rdp, cert.thumbprintSha1(), testOnly, backup, ui);
                                RdpFileSigner.CommandResult result = attempt.result();
                                StringBuilder msg =
                                        new StringBuilder(
                                                (testOnly ? "【テスト署名】" : "【本署名】")
                                                        + " 終了コード "
                                                        + result.exitCode()
                                                        + "\n");
                                msg.append("作業ファイル: ").append(attempt.target().signingPath()).append('\n');
                                msg.append("新規出力: ")
                                        .append(attempt.target().effectiveProfilePath())
                                        .append('\n');
                                msg.append("元ファイル（不変）: ")
                                        .append(attempt.target().sourcePath())
                                        .append('\n');
                                msg.append(RdpFileSigner.explainSignFailure(result));
                                Platform.runLater(
                                        () -> {
                                            logArea.setText(msg.toString());
                                            if (result.success()) {
                                                Path effective = attempt.target().effectiveProfilePath();
                                                if (!testOnly) {
                                                    rdpPathRef.set(effective);
                                                    if (profileChangeHandler != null) {
                                                        profileChangeHandler.accept(effective.toString());
                                                    }
                                                }
                                                String ok =
                                                        (testOnly ? "テスト署名 OK: " : "RDP 新規署名ファイル作成: ")
                                                                + effective;
                                                if (!testOnly && attempt.target().createsNewProfileFile()) {
                                                    ok += "（元 .rdp は未変更）";
                                                }
                                                if (statusConsumer != null) {
                                                    statusConsumer.accept(ok);
                                                }
                                            } else {
                                                showAlert(
                                                        Alert.AlertType.ERROR,
                                                        testOnly ? "テスト署名失敗" : "署名失敗",
                                                        msg.length() > 900
                                                                ? msg.substring(0, 900) + "…"
                                                                : msg.toString());
                                            }
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            String msg = ex.getMessage() != null ? ex.getMessage() : ex.toString();
                                            logArea.setText(msg);
                                            showAlert(Alert.AlertType.ERROR, "エラー", msg);
                                        });
                            }
                        },
                        "rdp-sign-wizard");
        worker.setDaemon(true);
        worker.start();
    }

    /** {@link #show(Window, Optional, Consumer, Consumer)} の profileChangeHandler 省略版。 */
    public static void show(Window owner, Optional<Path> initialRdp, Consumer<String> statusConsumer) {
        show(owner, initialRdp, statusConsumer, null);
    }

    private static void loadCertificatesAsync(
            Stage stage,
            AtomicReference<RdpFileSigner.CertificateListResult> certQuery,
            Runnable onDone,
            ComboBox<SigningCertificate> comboToRefresh) {
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                RdpFileSigner.CertificateListResult list =
                                        RdpFileSigner.listSigningCertificates();
                                Platform.runLater(
                                        () -> {
                                            certQuery.set(list);
                                            if (comboToRefresh != null) {
                                                comboToRefresh.getItems().setAll(list.eligible());
                                                if (!list.eligible().isEmpty()) {
                                                    comboToRefresh.getSelectionModel().selectFirst();
                                                }
                                            }
                                            if (onDone != null) {
                                                onDone.run();
                                            }
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            if (onDone != null) {
                                                onDone.run();
                                            }
                                            showAlert(
                                                    Alert.AlertType.WARNING,
                                                    "証明書一覧",
                                                    "証明書の取得に失敗しました: "
                                                            + (ex.getMessage() != null ? ex.getMessage() : ex));
                                        });
                            }
                        },
                        "rdp-cert-list");
        worker.setDaemon(true);
        worker.start();
    }

    private static void updatePathField(TextField field, Path path) {
        field.setText(path != null ? path.toString() : "");
    }

    private static void refreshSignedLabel(Label label, Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            label.setText("");
            return;
        }
        try {
            boolean signed = RdpFileSigner.isSigned(path);
            label.setText(signed ? "状態: 署名済み（signature:s: あり）" : "状態: 未署名");
        } catch (IOException ex) {
            label.setText("状態: 確認失敗 — " + ex.getMessage());
        }
    }

    private static void copyToClipboard(String text, String label) {
        if (text == null || text.isBlank()) {
            return;
        }
        ClipboardContent content = new ClipboardContent();
        content.putString(text);
        Clipboard.getSystemClipboard().setContent(content);
        showAlert(Alert.AlertType.INFORMATION, "コピー", label + " をクリップボードへコピーしました。");
    }

    private static void showAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    /** 依頼書入力タブから起動する際の初期 .rdp（環境変数プロファイル）。 */
    public static Optional<Path> initialProfileFromUi(java.util.Map<String, String> uiEnv) {
        return AppPaths.resolveRequestFormRdpProfile(uiEnv != null ? uiEnv : java.util.Map.of());
    }
}
