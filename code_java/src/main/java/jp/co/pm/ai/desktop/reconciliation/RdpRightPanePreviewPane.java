package jp.co.pm.ai.desktop.reconciliation;

import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.image.ImageView;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.io.RdpLaunchDisplaySettings;

/** 右ペイン最上部の mstsc 読み取り専用プレビュー領域。 */
public final class RdpRightPanePreviewPane extends VBox {

    private final StackPane previewSurface = new StackPane();
    private final ImageView imageView = new ImageView();
    private final ProgressIndicator loadingIndicator = new ProgressIndicator();
    private final Label statusLabel = new Label("接続中…");

    public RdpRightPanePreviewPane() {
        super(8);
        getStyleClass().add("pm-rdp-preview-host-pane");
        setFillWidth(true);
        setMinWidth(RdpLaunchDisplaySettings.MIN_WIDTH);
        setMinHeight(RdpLaunchDisplaySettings.MIN_HEIGHT + 40);
        setPrefWidth(RdpLaunchDisplaySettings.MIN_WIDTH);
        setPrefHeight(RdpLaunchDisplaySettings.MIN_HEIGHT + 48);

        Label title = new Label("リモートデスクトップ（プレビュー・操作用は別ウィンドウ）");
        title.getStyleClass().add("pm-rdp-preview-host-title");
        title.setWrapText(true);

        imageView.setPreserveRatio(true);
        imageView.fitWidthProperty().bind(previewSurface.widthProperty());
        imageView.fitHeightProperty().bind(previewSurface.heightProperty());
        previewSurface.getChildren().add(imageView);
        previewSurface.setMinSize(
                RdpLaunchDisplaySettings.MIN_WIDTH, RdpLaunchDisplaySettings.MIN_HEIGHT);
        previewSurface.getStyleClass().add("pm-rdp-preview-surface");
        VBox.setVgrow(previewSurface, Priority.ALWAYS);

        StackPane loadingOverlay = new StackPane(loadingIndicator, statusLabel);
        loadingOverlay.setAlignment(Pos.CENTER);
        loadingOverlay.getStyleClass().add("pm-rdp-preview-loading");
        previewSurface.getChildren().add(loadingOverlay);

        getChildren().addAll(title, previewSurface);
    }

    public ImageView imageView() {
        return imageView;
    }

    public void showFrame() {
        previewSurface.getChildren().removeIf(node -> node != imageView);
    }

    public void showLoading(String message) {
        if (previewSurface.getChildren().stream().noneMatch(n -> n instanceof ProgressIndicator)) {
            StackPane loadingOverlay = new StackPane(loadingIndicator, statusLabel);
            loadingOverlay.setAlignment(Pos.CENTER);
            loadingOverlay.getStyleClass().add("pm-rdp-preview-loading");
            previewSurface.getChildren().add(loadingOverlay);
        }
        statusLabel.setText(message != null && !message.isBlank() ? message : "接続中…");
    }
}
