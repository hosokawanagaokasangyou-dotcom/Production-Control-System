package jp.co.pm.ai.desktop.reconciliation;

import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings;

/** 右ペイン最上部の mstsc 埋め込み表示領域。 */
public final class RdpEmbedHostPane extends VBox {

    private final StackPane embedSurface = new StackPane();
    private final ProgressIndicator loadingIndicator = new ProgressIndicator();
    private final Label statusLabel = new Label("接続中…");

    public RdpEmbedHostPane(int prefWidth, int prefHeight) {
        super(8);
        getStyleClass().add("pm-rdp-embed-host-pane");
        setFillWidth(true);
        setMinWidth(RdpLaunchDisplaySettings.MIN_WIDTH);
        setMinHeight(RdpLaunchDisplaySettings.MIN_HEIGHT + 40);
        setPrefWidth(Math.max(RdpLaunchDisplaySettings.MIN_WIDTH, prefWidth));
        setPrefHeight(Math.max(RdpLaunchDisplaySettings.MIN_HEIGHT + 40, prefHeight + 40));

        Label title = new Label("リモートデスクトップ（接続中）");
        title.getStyleClass().add("pm-rdp-embed-host-title");

        embedSurface.getStyleClass().add("pm-rdp-embed-surface");
        embedSurface.setMinSize(RdpLaunchDisplaySettings.MIN_WIDTH, RdpLaunchDisplaySettings.MIN_HEIGHT);
        embedSurface.setPrefSize(
                Math.max(RdpLaunchDisplaySettings.MIN_WIDTH, prefWidth),
                Math.max(RdpLaunchDisplaySettings.MIN_HEIGHT, prefHeight));
        VBox.setVgrow(embedSurface, Priority.ALWAYS);

        StackPane loadingOverlay = new StackPane(loadingIndicator, statusLabel);
        loadingOverlay.setAlignment(Pos.CENTER);
        loadingOverlay.getStyleClass().add("pm-rdp-embed-loading");
        embedSurface.getChildren().add(loadingOverlay);

        getChildren().addAll(title, embedSurface);
    }

    public StackPane embedSurface() {
        return embedSurface;
    }

    public void showEmbedded() {
        embedSurface.getChildren().clear();
    }

    public void showLoading(String message) {
        if (embedSurface.getChildren().isEmpty()
                || !(embedSurface.getChildren().getFirst() instanceof StackPane)) {
            StackPane loadingOverlay = new StackPane(loadingIndicator, statusLabel);
            loadingOverlay.setAlignment(Pos.CENTER);
            loadingOverlay.getStyleClass().add("pm-rdp-embed-loading");
            embedSurface.getChildren().setAll(loadingOverlay);
        }
        statusLabel.setText(message != null && !message.isBlank() ? message : "接続中…");
    }
}
