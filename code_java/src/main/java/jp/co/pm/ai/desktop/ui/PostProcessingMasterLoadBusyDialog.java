package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;
import javafx.stage.WindowEvent;

/**
 * 後加工商品マスタの参照・連携マスタ読込中に表示する進捗モーダル。
 *
 * <p>必須処理のためキャンセル不可。呼び出し側で必ず {@link #close()} する。
 * タイトルバーの閉じるボタンを出さないため {@link StageStyle#UNDECORATED}。
 */
public final class PostProcessingMasterLoadBusyDialog {

    public static final String TITLE = "後加工商品マスタ";
    public static final String HEADER = "マスタデータを読み込んでいます";
    public static final String STATUS_LOADING = "参照マスタ・連携マスタを読み込んでいます…";

    /** 明るいダイアログ本文用。テーマ CSS の白文字を載せない。 */
    public static final String BODY_TEXT_FILL = "#1a1a1a";
    private static final String BODY_TEXT_STYLE =
            "-fx-text-fill: " + BODY_TEXT_FILL + "; -fx-font-size: 13px;";
    private static final String HEADER_TEXT_STYLE =
            "-fx-font-weight: bold; -fx-font-size: 14px; -fx-text-fill: " + BODY_TEXT_FILL + ";";

    private final Stage stage;
    private final Label statusLabel;

    private PostProcessingMasterLoadBusyDialog(Stage stage, Label statusLabel) {
        this.stage = stage;
        this.statusLabel = statusLabel;
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ（{@code null} 可）
     * @param initialStatus 最初の状況文言
     */
    public static PostProcessingMasterLoadBusyDialog show(Window owner, String initialStatus) {
        Stage stage = new Stage(StageStyle.UNDECORATED);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(TITLE);
        stage.setResizable(false);
        stage.setOnCloseRequest(WindowEvent::consume);

        Label titleBar = new Label(TITLE);
        titleBar.setMaxWidth(Double.MAX_VALUE);
        titleBar.setPadding(new Insets(6, 12, 6, 12));
        titleBar.setStyle(
                "-fx-background-color: #3d3d3d; -fx-text-fill: #f2f2f2; -fx-font-size: 12px;");

        Label headerLabel = new Label(HEADER);
        headerLabel.setStyle(HEADER_TEXT_STYLE);
        headerLabel.setWrapText(true);
        headerLabel.setMaxWidth(380);

        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        progress.setMaxSize(48, 48);

        Label status = new Label(initialStatus != null ? initialStatus : STATUS_LOADING);
        status.setStyle(BODY_TEXT_STYLE);
        status.setWrapText(true);
        status.setMaxWidth(380);

        VBox content = new VBox(14, headerLabel, progress, status);
        content.setAlignment(Pos.CENTER);
        content.setPadding(new Insets(20, 24, 24, 24));
        content.setPrefWidth(400);

        VBox windowRoot = new VBox(titleBar, content);
        windowRoot.setStyle(
                "-fx-background-color: #f0f0f0; -fx-border-color: #6b6b6b; -fx-border-width: 1;");

        stage.setScene(new Scene(windowRoot));
        stage.sizeToScene();
        stage.show();
        positionOverOwner(stage, owner);
        stage.toFront();
        return new PostProcessingMasterLoadBusyDialog(stage, status);
    }

    static void positionOverOwner(Stage stage, Window owner) {
        if (owner == null) {
            return;
        }
        stage.setX(
                FactorySiteSwitchBusySupport.centerX(
                        owner.getX(), owner.getWidth(), stage.getWidth()));
        stage.setY(
                FactorySiteSwitchBusySupport.centerY(
                        owner.getY(), owner.getHeight(), stage.getHeight()));
    }

    public void setStatus(String text) {
        if (text != null) {
            statusLabel.setText(text);
        }
    }

    public boolean isShowing() {
        return stage.isShowing();
    }

    /** テスト用。タイトルバーの閉じるボタンを出さないため {@link StageStyle#UNDECORATED}。 */
    StageStyle stageStyle() {
        return stage.getStyle();
    }

    /** テスト用。明るい背景に対する本文色。 */
    String bodyTextStyle() {
        return statusLabel.getStyle() != null ? statusLabel.getStyle() : "";
    }

    public Scene scene() {
        return stage.getScene();
    }

    public void close() {
        stage.setOnCloseRequest(null);
        if (stage.isShowing()) {
            stage.close();
        }
    }
}
