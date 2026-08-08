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
 * 起動時の工場ワークスペース復元・環境変数照合中に表示する進捗モーダル。
 *
 * <p>必須処理のためキャンセル不可。呼び出し側で必ず {@link #close()} する。
 */
public final class EnvVarsStartupCheckBusyDialog {

    public static final String TITLE = "起動時チェック";
    public static final String HEADER = "環境変数を確認しています";
    public static final String HEADER_BACKGROUND_LOAD = "タブデータを読み込んでいます";

    public static final String STATUS_RESTORE_WORKSPACE = "工場ワークスペースを復元しています…";
    public static final String STATUS_STABILIZE = "環境変数の値を安定化しています…";
    public static final String STATUS_MATCH = "初期値と照合しています…";
    public static final String STATUS_FACTORY_SWITCH = "工場ワークスペースを切り替えています…";
    public static final String STATUS_BACKGROUND_LOAD = "タブデータを読み込んでいます…";
    public static final String STATUS_DONE = "完了しました";

    private final Stage stage;
    private final Label headerLabel;
    private final Label statusLabel;

    private EnvVarsStartupCheckBusyDialog(Stage stage, Label headerLabel, Label statusLabel) {
        this.stage = stage;
        this.headerLabel = headerLabel;
        this.statusLabel = statusLabel;
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ（{@code null} 可）
     * @param initialStatus 最初の状況文言
     */
    public static EnvVarsStartupCheckBusyDialog show(Window owner, String initialStatus) {
        Stage stage = new Stage(StageStyle.UTILITY);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(TITLE);
        stage.setResizable(false);
        // ユーザ操作での閉じを抑止（必須処理）
        stage.setOnCloseRequest(WindowEvent::consume);

        Label header = new Label(HEADER);
        header.getStyleClass().add("dialog-header");
        header.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");

        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        progress.setMaxSize(48, 48);

        Label status = new Label(initialStatus != null ? initialStatus : STATUS_STABILIZE);
        status.setWrapText(true);
        status.setMaxWidth(340);

        VBox root = new VBox(14, header, progress, status);
        root.setAlignment(Pos.CENTER);
        root.setPadding(new Insets(20, 24, 24, 24));
        root.setPrefWidth(380);

        stage.setScene(new Scene(root));
        stage.sizeToScene();
        stage.show();
        if (owner != null) {
            stage.centerOnScreen();
        }
        return new EnvVarsStartupCheckBusyDialog(stage, header, status);
    }

    public void setHeader(String text) {
        if (text != null && headerLabel != null) {
            headerLabel.setText(text);
        }
    }

    public void setStatus(String text) {
        if (text != null) {
            statusLabel.setText(text);
        }
    }

    public boolean isShowing() {
        return stage.isShowing();
    }

    public void close() {
        stage.setOnCloseRequest(null);
        if (stage.isShowing()) {
            stage.close();
        }
    }
}
