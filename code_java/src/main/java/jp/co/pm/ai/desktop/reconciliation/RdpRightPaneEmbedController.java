package jp.co.pm.ai.desktop.reconciliation;

import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.control.SplitPane;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.win32.MstscWindowEmbedder;
import jp.co.pm.ai.desktop.io.win32.MstscWindowLocator;

/** 右ペイン {@link SplitPane} への mstsc 埋め込み Pane の追加・削除・HWND 同期。 */
public final class RdpRightPaneEmbedController {

    private static final double[] THREE_PANE_DIVIDERS = {0.45, 0.70};

    private final SplitPane rightPaneSplit;
    private final MstscWindowEmbedder embedder = new MstscWindowEmbedder();
    private RdpEmbedHostPane embedHostPane;
    private double[] savedDividers;
    private Runnable boundsSyncListener;

    public RdpRightPaneEmbedController(SplitPane rightPaneSplit) {
        this.rightPaneSplit = rightPaneSplit;
    }

    public boolean isSupported() {
        return embedder.isSupported();
    }

    public boolean isVisible() {
        return embedHostPane != null;
    }

    public MstscWindowEmbedder embedder() {
        return embedder;
    }

    public RdpEmbedHostPane embedHostPane() {
        return embedHostPane;
    }

    /** 右ペイン index 0 に埋め込み Pane を追加する。 */
    public void showEmbedPane(int width, int height) {
        if (embedHostPane != null) {
            embedHostPane.setPrefWidth(Math.max(width, embedHostPane.getMinWidth()));
            embedHostPane.embedSurface().setPrefSize(width, height);
            return;
        }
        savedDividers = rightPaneSplit.getDividerPositions();
        embedHostPane = new RdpEmbedHostPane(width, height);
        rightPaneSplit.getItems().addFirst(embedHostPane);
        rightPaneSplit.setDividerPositions(THREE_PANE_DIVIDERS);
        SplitPane.setResizableWithParent(embedHostPane, Boolean.TRUE);
    }

    /** 埋め込み Pane を削除し 2 段構成へ戻す。 */
    public void removeEmbedPane() {
        embedder.detach();
        if (embedHostPane == null) {
            return;
        }
        if (boundsSyncListener != null && embedHostPane.embedSurface().getScene() != null) {
            embedHostPane
                    .embedSurface()
                    .layoutBoundsProperty()
                    .removeListener(obs -> boundsSyncListener.run());
        }
        boundsSyncListener = null;
        rightPaneSplit.getItems().remove(embedHostPane);
        embedHostPane = null;
        if (savedDividers != null && savedDividers.length == 1) {
            rightPaneSplit.setDividerPositions(savedDividers);
        } else if (rightPaneSplit.getItems().size() == 2) {
            rightPaneSplit.setDividerPositions(1.0 / 3.0);
        }
        savedDividers = null;
    }

    /**
     * バックグラウンドで mstsc HWND を待ち、FX スレッドで attach する。
     *
     * @param onAttached 成功/失敗を FX スレッドで通知（message は失敗時のみ非空可）
     */
    public void attachWhenReady(Window window, long mstscPid, java.util.function.Consumer<String> onAttached) {
        if (!embedder.isSupported() || embedHostPane == null || mstscPid <= 0) {
            Platform.runLater(() -> onAttached.accept("埋め込み未対応"));
            return;
        }
        Thread worker =
                new Thread(
                        () -> {
                            var hwnd = MstscWindowLocator.findClientWindow(mstscPid);
                            Platform.runLater(
                                    () -> {
                                        if (hwnd.isEmpty()) {
                                            onAttached.accept("mstsc ウィンドウを特定できませんでした");
                                            return;
                                        }
                                        Node surface = embedHostPane.embedSurface();
                                        if (embedder.attach(window, surface, hwnd.getAsLong())) {
                                            installBoundsSync();
                                            embedHostPane.showEmbedded();
                                            onAttached.accept(null);
                                        } else {
                                            onAttached.accept("mstsc の埋め込みに失敗しました");
                                        }
                                    });
                        },
                        "rdp-embed-attach");
        worker.setDaemon(true);
        worker.start();
    }

    private void installBoundsSync() {
        if (embedHostPane == null) {
            return;
        }
        boundsSyncListener = embedder::syncBounds;
        embedHostPane.embedSurface().layoutBoundsProperty().addListener(obs -> boundsSyncListener.run());
        embedHostPane.widthProperty().addListener((obs, w, n) -> embedder.syncBounds());
        embedHostPane.heightProperty().addListener((obs, h, n) -> embedder.syncBounds());
        rightPaneSplit.getDividers().forEach(div -> div.positionProperty().addListener((obs, p, n) -> embedder.syncBounds()));
        Platform.runLater(embedder::syncBounds);
    }

    public void dispose() {
        removeEmbedPane();
        embedder.close();
    }
}
