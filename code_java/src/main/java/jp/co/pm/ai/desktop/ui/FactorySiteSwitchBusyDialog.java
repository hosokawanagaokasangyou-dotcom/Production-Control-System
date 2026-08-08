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
 * 工場切替（湖南／国分）実行中に表示する進捗モーダル。
 *
 * <p>必須処理のためキャンセル不可。呼び出し側で必ず {@link #close()} する。
 */
public final class FactorySiteSwitchBusyDialog {

    public static final String TITLE = "工場切替";

    public static final String STATUS_SAVING = "現在の工場設定を保存しています…";
    public static final String STATUS_LOADING = "切替先の工場ワークスペースを読み込んでいます…";
    public static final String STATUS_ENV = "環境変数を切り替えています…";
    public static final String STATUS_REFRESH = "関連画面を更新しています…";
    public static final String STATUS_OPERATOR = "操作者を確認しています…";
    public static final String STATUS_STABILIZE = "環境変数の値を安定化しています…";
    public static final String STATUS_MATCH = "初期値と照合しています…";
    public static final String STATUS_DONE = "完了しました";

    private final Stage stage;
    private final Label statusLabel;

    private FactorySiteSwitchBusyDialog(Stage stage, Label statusLabel) {
        this.stage = stage;
        this.statusLabel = statusLabel;
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ（{@code null} 可）
     * @param header 見出し（例: 湖南工場 → 国分工場）
     * @param initialStatus 最初の状況文言
     */
    public static FactorySiteSwitchBusyDialog show(Window owner, String header, String initialStatus) {
        Stage stage = new Stage(StageStyle.UTILITY);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(TITLE);
        stage.setResizable(false);
        stage.setOnCloseRequest(WindowEvent::consume);

        Label headerLabel = new Label(header != null ? header : "工場を切り替えています");
        headerLabel.getStyleClass().add("dialog-header");
        headerLabel.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");
        headerLabel.setWrapText(true);
        headerLabel.setMaxWidth(380);

        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        progress.setMaxSize(48, 48);

        Label status = new Label(initialStatus != null ? initialStatus : STATUS_SAVING);
        status.setWrapText(true);
        status.setMaxWidth(380);

        VBox root = new VBox(14, headerLabel, progress, status);
        root.setAlignment(Pos.CENTER);
        root.setPadding(new Insets(20, 24, 24, 24));
        root.setPrefWidth(400);

        stage.setScene(new Scene(root));
        stage.sizeToScene();
        stage.show();
        if (owner != null) {
            stage.centerOnScreen();
        }
        return new FactorySiteSwitchBusyDialog(stage, status);
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
