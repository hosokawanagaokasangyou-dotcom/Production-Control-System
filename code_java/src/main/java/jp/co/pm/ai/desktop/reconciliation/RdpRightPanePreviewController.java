package jp.co.pm.ai.desktop.reconciliation;

import java.awt.image.BufferedImage;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.image.Image;
import javafx.scene.image.WritableImage;

import jp.co.pm.ai.desktop.io.RdpMstscPreviewSupport;
import jp.co.pm.ai.desktop.io.win32.MstscCaptureTarget;
import jp.co.pm.ai.desktop.io.win32.MstscWindowCapture;

/** 右ペインへの mstsc 読み取り専用低 fps プレビュー。 */
public final class RdpRightPanePreviewController implements AutoCloseable {

    private static final double[] THREE_PANE_DIVIDERS = {0.45, 0.70};
    private static final Duration START_TIMEOUT = Duration.ofSeconds(90);
    private static final Duration FRAME_INTERVAL = Duration.ofMillis(333);
    private static final int MAX_BLANK_FRAMES = 10;

    private final javafx.scene.control.SplitPane rightPaneSplit;
    private RdpRightPanePreviewPane previewPane;
    private double[] savedDividers;
    private Thread worker;
    private WritableImage fxImage;

    public RdpRightPanePreviewController(javafx.scene.control.SplitPane rightPaneSplit) {
        this.rightPaneSplit = rightPaneSplit;
    }

    public boolean isSupported() {
        return RdpMstscPreviewSupport.isCaptureSupported();
    }

    public boolean isVisible() {
        return previewPane != null;
    }

    public void showPreviewPane() {
        if (previewPane != null) {
            return;
        }
        savedDividers = rightPaneSplit.getDividerPositions();
        previewPane = new RdpRightPanePreviewPane();
        rightPaneSplit.getItems().addFirst(previewPane);
        rightPaneSplit.setDividerPositions(THREE_PANE_DIVIDERS);
        javafx.scene.control.SplitPane.setResizableWithParent(previewPane, Boolean.TRUE);
    }

    public void removePreviewPane() {
        stopWorker();
        if (previewPane == null) {
            return;
        }
        rightPaneSplit.getItems().remove(previewPane);
        previewPane = null;
        fxImage = null;
        if (savedDividers != null && savedDividers.length == 1) {
            rightPaneSplit.setDividerPositions(savedDividers);
        } else if (rightPaneSplit.getItems().size() == 2) {
            rightPaneSplit.setDividerPositions(1.0 / 3.0);
        }
        savedDividers = null;
    }

    /**
     * mstsc HWND を待ち、低 fps でキャプチャする。
     *
     * @param onStopped FX スレッドで呼ぶ。成功継続中は呼ばない。失敗・中断時のみ非 null メッセージ。
     */
    public void previewWhenReady(
            OptionalLong mstscPidHint,
            Optional<Path> pidMarkerFile,
            Path rdpProfile,
            Consumer<String> onStopped) {
        if (!isSupported() || previewPane == null) {
            Platform.runLater(
                    () -> onStopped.accept(previewPane == null ? "プレビュー領域未初期化" : "プレビュー未対応"));
            return;
        }
        stopWorker();
        worker =
                new Thread(
                        () -> runPreviewLoop(mstscPidHint, pidMarkerFile, rdpProfile, onStopped),
                        "rdp-preview-capture");
        worker.setDaemon(true);
        worker.start();
    }

    private void runPreviewLoop(
            OptionalLong mstscPidHint,
            Optional<Path> pidMarkerFile,
            Path rdpProfile,
            Consumer<String> onStopped) {
        AtomicLong resolvedPid = new AtomicLong(mstscPidHint.orElse(-1L));
        Path marker = pidMarkerFile.orElse(null);
        long deadline = System.nanoTime() + START_TIMEOUT.toNanos();
        java.util.Optional<MstscCaptureTarget> captureTarget = java.util.Optional.empty();
        while (System.nanoTime() < deadline && captureTarget.isEmpty()) {
            if (Thread.currentThread().isInterrupted()) {
                notifyStopped(onStopped, "プレビュー待機が中断されました");
                return;
            }
            if (resolvedPid.get() <= 0) {
                long pid =
                        RdpMstscPreviewSupport.tryResolveMstscPid(
                                rdpProfile, OptionalLong.empty(), marker);
                if (pid > 0) {
                    resolvedPid.set(pid);
                }
            }
            long pidHint = resolvedPid.get();
            captureTarget = RdpMstscPreviewSupport.findCaptureTargetOnce(pidHint > 0 ? pidHint : -1L);
            if (captureTarget.isEmpty()) {
                showLoadingOnFx(
                        resolvedPid.get() > 0 ? "mstsc ウィンドウ探索中…" : "mstsc 起動待ち…");
                sleepQuiet(FRAME_INTERVAL);
            }
        }
        if (captureTarget.isEmpty()) {
            notifyStopped(onStopped, "mstsc ウィンドウを特定できずプレビューを中止しました（別ウィンドウで操作してください）");
            return;
        }

        MstscCaptureTarget target = captureTarget.get();
        showLoadingOnFx("プレビュー取得中…");
        int blankStreak = 0;
        int publishedFrames = 0;
        while (!Thread.currentThread().isInterrupted()) {
            if (blankStreak > 0 && blankStreak % 5 == 0) {
                captureTarget =
                        RdpMstscPreviewSupport.findCaptureTargetOnce(
                                resolvedPid.get() > 0 ? resolvedPid.get() : -1L);
                if (captureTarget.isPresent()) {
                    target = captureTarget.get();
                }
            }
            Optional<BufferedImage> frame = MstscWindowCapture.captureTarget(target);
            boolean blank = frame.isPresent() && MstscWindowCapture.isLikelyBlank(frame.get());
            if (frame.isEmpty()) {
                blankStreak++;
            } else if (blank) {
                blankStreak++;
            } else {
                blankStreak = 0;
                publishedFrames++;
                publishFrame(frame.get());
            }
            if (blankStreak >= MAX_BLANK_FRAMES) {
                if (publishedFrames == 0) {
                    notifyStopped(
                            onStopped,
                            "プレビューが取得できませんでした（黒画面または非対応）。別ウィンドウで操作してください");
                    return;
                }
                blankStreak = 0;
            }
            sleepQuiet(FRAME_INTERVAL);
        }
        notifyStopped(onStopped, "プレビュー待機が中断されました");
    }

    private void publishFrame(BufferedImage buffered) {
        Platform.runLater(
                () -> {
                    if (previewPane == null) {
                        return;
                    }
                    if (fxImage == null
                            || fxImage.getWidth() != buffered.getWidth()
                            || fxImage.getHeight() != buffered.getHeight()) {
                        fxImage = new WritableImage(buffered.getWidth(), buffered.getHeight());
                    }
                    Image converted = SwingFXUtils.toFXImage(buffered, fxImage);
                    previewPane.imageView().setImage(converted);
                    previewPane.showFrame();
                });
    }

    private void showLoadingOnFx(String message) {
        Platform.runLater(
                () -> {
                    if (previewPane != null) {
                        previewPane.showLoading(message);
                    }
                });
    }

    private static void sleepQuiet(Duration duration) {
        try {
            Thread.sleep(duration.toMillis());
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
        }
    }

    private void notifyStopped(Consumer<String> onStopped, String message) {
        Platform.runLater(
                () -> {
                    removePreviewPane();
                    if (onStopped != null) {
                        onStopped.accept(message);
                    }
                });
    }

    private void stopWorker() {
        Thread t = worker;
        worker = null;
        if (t != null) {
            t.interrupt();
        }
    }

    @Override
    public void close() {
        removePreviewPane();
    }
}
