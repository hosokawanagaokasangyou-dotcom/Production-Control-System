package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;

/**
 * ポータブル配布の起動時自動バージョンアップ（配台 PMD / RDP ランチャー共通フロー）。
 */
public final class PortableBundleSelfUpdateService {

    private PortableBundleSelfUpdateService() {}

    /**
     * 起動時に正本の版が新しいとき確認ダイアログのあと ZIP 取得・本体ステージング・終了後適用を行う。
     *
     * @param profile 配布種別
     * @param bundleSourceEnvKey 正本パスを読む環境変数キー（空ならスキップ）
     * @param ui 環境変数マップ
     * @param dialogOwner ダイアログ親（null 可）
     * @param log ログ出力
     */
    public static void maybeSelfUpdateOnStartup(
            PortableBundleProfile profile,
            String bundleSourceEnvKey,
            Map<String, String> ui,
            Stage dialogOwner,
            Consumer<String> log) {
        Objects.requireNonNull(profile, "profile");
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        if (!profile.isPortableBundleLayout(cwd)) {
            logLine(
                    log,
                    "[startup] 自動バージョンアップは対象外（"
                            + profile.desktopExeName()
                            + " 直下のポータブル配布レイアウトがありません）。user.dir="
                            + PortableBundleSelfUpdater.safePathForLog(cwd));
            return;
        }
        logLine(
                log,
                "[startup] 自動バージョンアップ: ポータブル配布を検出 user.dir="
                        + PortableBundleSelfUpdater.safePathForLog(cwd));
        if (PortableBundleUpdateLauncher.tryApplyStagedBundleOnColdStart(profile, cwd, log)) {
            logLine(log, "[startup] 前回保留のデスクトップ本体更新を適用しました。");
        }

        String raw = ui != null ? ui.getOrDefault(bundleSourceEnvKey, "").strip() : "";
        if (raw.isEmpty()) {
            logLine(
                    log,
                    "[startup] 自動バージョンアップはスキップ（"
                            + bundleSourceEnvKey
                            + " が未設定）。正本フォルダまたは ZIP を環境変数に指定してください。");
            return;
        }
        Path canonical = Path.of(raw).toAbsolutePath().normalize();
        if (!PortableBundleSelfUpdater.isValidPortableBundleCanonical(canonical)) {
            logLine(
                    log,
                    "[startup] 自動バージョンアップはスキップ（正本を開けません）: "
                            + PortableBundleSelfUpdater.safePathForLog(canonical));
            return;
        }

        Path localData = profile.syncsPmAiData() ? cwd.resolve("pm-ai-data") : cwd;
        if (!PortableBundleSelfUpdater.shouldUpdateBundle(profile, canonical, cwd, localData)) {
            Optional<BigDecimal> cv =
                    PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(profile, canonical);
            Optional<BigDecimal> lv =
                    profile.syncsPmAiData()
                            ? PortableBundleSelfUpdater.readLocalBundleVersion(cwd, localData)
                            : PortableBundleSelfUpdater.readLocalBundleVersion(cwd, cwd);
            logLine(
                    log,
                    "[startup] 自動バージョンアップはスキップ（更新不要） 正本="
                            + cv.map(BigDecimal::toPlainString).orElse("（なし）")
                            + " ローカル="
                            + lv.map(BigDecimal::toPlainString).orElse("（なし・0扱い）"));
            return;
        }

        Optional<BigDecimal> cv = PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(profile, canonical);
        Optional<BigDecimal> lv =
                profile.syncsPmAiData()
                        ? PortableBundleSelfUpdater.readLocalBundleVersion(cwd, localData)
                        : PortableBundleSelfUpdater.readLocalBundleVersion(cwd, cwd);
        String canonVerStr = cv.map(BigDecimal::toPlainString).orElse("?");
        String localVerStr = lv.map(BigDecimal::toPlainString).orElse("（なし・初回）");
        Optional<Path> upgradeZip = PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(profile, canonical);
        if (upgradeZip.isEmpty()) {
            logLine(log, "[startup] 自動バージョンアップはスキップ（アップグレード ZIP が見つかりません）。");
            return;
        }

        Alert confirm = new Alert(AlertType.CONFIRMATION);
        if (dialogOwner != null) {
            confirm.initOwner(dialogOwner);
        }
        confirm.setTitle("自動バージョンアップ");
        confirm.setHeaderText(null);
        confirm.setContentText(
                "正本の更新があります（版 "
                        + canonVerStr
                        + "、ローカル "
                        + localVerStr
                        + "）。\n"
                        + "ZIP を展開し、"
                        + profile.desktopExeName()
                        + "・app・runtime を更新します。本体更新後は自動的にアプリを再起動します。\n"
                        + "実行してよいですか？");
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            logLine(log, "[startup] 自動バージョンアップはユーザーによりスキップ（版 " + canonVerStr + "）。");
            return;
        }

        runZipUpgradeTask(profile, cwd, canonical, upgradeZip.get(), canonVerStr, dialogOwner, log);
    }

    private static void runZipUpgradeTask(
            PortableBundleProfile profile,
            Path cwd,
            Path canonical,
            Path remoteZip,
            String canonVerStr,
            Stage dialogOwner,
            Consumer<String> log) {
        Stage wait = new Stage();
        wait.initModality(Modality.APPLICATION_MODAL);
        if (dialogOwner != null) {
            wait.initOwner(dialogOwner);
        }
        wait.setTitle("自動バージョンアップ");
        wait.setMinWidth(520);
        wait.setMinHeight(360);
        Label msg =
                new Label(
                        profile.desktopExeName()
                                + " を正本 ZIP から更新しています…\n詳細はログに記録されます。");
        msg.setWrapText(true);
        msg.setMaxWidth(480);
        Label downloadCaption = new Label("① 正本ZIPの取得");
        ProgressBar downloadBar = new ProgressBar(-1);
        downloadBar.setMaxWidth(Double.MAX_VALUE);
        Label extractCaption = new Label("② ZIPの展開");
        ProgressBar extractBar = new ProgressBar(0);
        extractBar.setMaxWidth(Double.MAX_VALUE);
        Label desktopCaption = new Label("③ デスクトップ本体のステージング");
        ProgressBar desktopBar = new ProgressBar(0);
        desktopBar.setMaxWidth(Double.MAX_VALUE);
        Label detail = new Label("準備中…");
        detail.setWrapText(true);
        detail.setMaxWidth(480);
        VBox root =
                new VBox(
                        12,
                        msg,
                        new VBox(6, downloadCaption, downloadBar),
                        new VBox(6, extractCaption, extractBar),
                        new VBox(6, desktopCaption, desktopBar),
                        detail);
        root.setAlignment(Pos.CENTER_LEFT);
        root.setStyle("-fx-padding: 20;");
        Scene waitScene = new Scene(root, 520, 360);
        wait.setScene(waitScene);
        if (dialogOwner != null && dialogOwner.getScene() != null) {
            waitScene.getStylesheets().setAll(dialogOwner.getScene().getStylesheets());
        }
        wait.show();

        PortableBundleUpgradeProgress.Listener progress =
                (phase, done, total, detailLine) ->
                        Platform.runLater(
                                () -> {
                                    ProgressBar bar;
                                    Label cap;
                                    switch (phase) {
                                        case DOWNLOAD -> {
                                            bar = downloadBar;
                                            cap = downloadCaption;
                                        }
                                        case EXTRACT -> {
                                            bar = extractBar;
                                            cap = extractCaption;
                                        }
                                        case SYNC_DESKTOP -> {
                                            bar = desktopBar;
                                            cap = desktopCaption;
                                        }
                                        default -> {
                                            bar = desktopBar;
                                            cap = desktopCaption;
                                        }
                                    }
                                    if (total > 0) {
                                        bar.setProgress((double) done / (double) total);
                                        cap.setText(
                                                cap.getText().split("（")[0]
                                                        + "（"
                                                        + done
                                                        + "/"
                                                        + total
                                                        + "）");
                                    } else if (done <= 0) {
                                        bar.setProgress(-1);
                                    }
                                    if (detailLine != null && !detailLine.isBlank()) {
                                        detail.setText(detailLine);
                                    }
                                });

        final Path[] localZipHolder = new Path[1];
        final Path[] extractedHolder = new Path[1];
        final AtomicBoolean deferredRelaunch = new AtomicBoolean();
        Consumer<String> syncLog = line -> logLine(log, line);

        Task<Void> task =
                new Task<>() {
                    @Override
                    protected Void call() throws Exception {
                        localZipHolder[0] =
                                PortableBundleSelfUpdater.copyUpgradeZipToLocal(
                                        remoteZip, syncLog, progress);
                        Path tmp =
                                PortableBundleSelfUpdater.extractUpgradeZipToTempDirectory(
                                        localZipHolder[0], syncLog, progress);
                        extractedHolder[0] = tmp;
                        Optional<Path> desktopRoot =
                                PortableBundleSelfUpdater.resolveDesktopBundleRoot(profile, tmp);
                        if (desktopRoot.isEmpty()) {
                            throw new IOException(
                                    "ZIP 内に "
                                            + profile.desktopExeName()
                                            + " / app のレイアウトがありません: "
                                            + remoteZip);
                        }
                        Path staging = PortableBundlePendingUpdate.defaultStagingDirectory(profile);
                        PortableBundleSelfUpdater.stageDesktopBundleForRelaunch(
                                profile, desktopRoot.get(), staging, syncLog);
                        deferredRelaunch.set(true);
                        return null;
                    }
                };

        task.setOnSucceeded(
                e -> {
                    if (localZipHolder[0] != null) {
                        try {
                            Files.deleteIfExists(localZipHolder[0]);
                        } catch (IOException ignored) {
                            /* best-effort */
                        }
                    }
                    if (extractedHolder[0] != null) {
                        PortableBundleSelfUpdater.deleteDirectoryRecursive(extractedHolder[0], syncLog);
                    }
                    wait.close();
                    if (!deferredRelaunch.get()) {
                        return;
                    }
                    try {
                        Path staging = PortableBundlePendingUpdate.defaultStagingDirectory(profile);
                        long pid = ProcessHandle.current().pid();
                        PortableBundlePendingUpdate.write(profile, cwd, staging, canonVerStr, pid, canonical);
                        PortableBundleUpdateLauncher.launchDeferredDesktopApply(
                                profile, cwd, staging, pid, canonVerStr, canonical, log);
                        logLine(
                                log,
                                "[startup] 自動バージョンアップ: 終了後に本体を適用して再起動します。");
                        Platform.runLater(
                                () -> {
                                    Alert info = new Alert(AlertType.INFORMATION);
                                    if (dialogOwner != null) {
                                        info.initOwner(dialogOwner);
                                    }
                                    info.setTitle("自動バージョンアップ");
                                    info.setHeaderText(null);
                                    info.setContentText(
                                            "版 "
                                                    + canonVerStr
                                                    + " の更新を適用するため、アプリを終了します。");
                                    info.showAndWait();
                                    if (dialogOwner != null) {
                                        dialogOwner.close();
                                    }
                                });
                    } catch (IOException ex) {
                        logLine(log, "[startup] 終了後更新の起動に失敗: " + ex.getMessage());
                    }
                });
        task.setOnFailed(
                e -> {
                    wait.close();
                    Throwable t = task.getException();
                    logLine(
                            log,
                            "[startup] 自動バージョンアップ失敗: "
                                    + (t != null ? t.getMessage() : "不明"));
                });
        new Thread(task, "portable-bundle-self-update").start();
    }

    private static void logLine(Consumer<String> log, String line) {
        if (log != null && line != null && !line.isBlank()) {
            log.accept(line);
        }
    }
}
