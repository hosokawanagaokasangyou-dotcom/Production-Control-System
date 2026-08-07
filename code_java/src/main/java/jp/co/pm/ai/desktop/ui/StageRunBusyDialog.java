package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;
import javafx.stage.WindowEvent;

/**
 * 段階1／2 実行中に表示する進捗モーダル。
 *
 * <p>フェーズ（大項目）と詳細（直近ログ要約）を表示し、中断ボタンで子プロセスを終了できる。
 */
public final class StageRunBusyDialog {

    private final Stage stage;
    private final Label phaseLabel;
    private final Label detailLabel;

    private StageRunBusyDialog(Stage stage, Label phaseLabel, Label detailLabel) {
        this.stage = stage;
        this.phaseLabel = phaseLabel;
        this.detailLabel = detailLabel;
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ
     * @param title ウィンドウタイトル（例: 段階1 実行中）
     * @param header 見出し（例: 段階1 を実行しています）
     * @param initialPhase 最初のフェーズ文言
     * @param onCancel 中断ボタン押下時（{@code null} 可）
     */
    public static StageRunBusyDialog show(
            Window owner,
            String title,
            String header,
            String initialPhase,
            Runnable onCancel) {
        Stage stage = new Stage(StageStyle.UTILITY);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(title != null ? title : "実行中");
        stage.setResizable(false);
        stage.setOnCloseRequest(WindowEvent::consume);

        Label headerLabel = new Label(header != null ? header : "処理を実行しています");
        headerLabel.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");
        headerLabel.setWrapText(true);
        headerLabel.setMaxWidth(400);

        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        progress.setMaxSize(48, 48);

        Label phase = new Label(initialPhase != null ? initialPhase : "準備中…");
        phase.setStyle("-fx-font-weight: bold;");
        phase.setWrapText(true);
        phase.setMaxWidth(400);

        Label detail = new Label("");
        detail.setWrapText(true);
        detail.setMaxWidth(400);
        detail.getStyleClass().add("muted");

        VBox content = new VBox(10, headerLabel, progress, phase, detail);
        content.setAlignment(Pos.CENTER_LEFT);
        content.setPadding(new Insets(20, 24, 12, 24));
        content.setPrefWidth(440);

        VBox root;
        if (onCancel != null) {
            Button cancel = new Button("中断");
            cancel.setOnAction(e -> onCancel.run());
            VBox actions = new VBox(cancel);
            actions.setAlignment(Pos.CENTER_RIGHT);
            actions.setPadding(new Insets(0, 24, 20, 24));
            root = new VBox(content, actions);
        } else {
            root = content;
            root.setPadding(new Insets(0, 0, 12, 0));
        }

        stage.setScene(new Scene(root));
        stage.sizeToScene();
        stage.show();
        if (owner != null) {
            stage.centerOnScreen();
        }
        return new StageRunBusyDialog(stage, phase, detail);
    }

    public void setPhase(String text) {
        if (text != null && !text.isBlank()) {
            phaseLabel.setText(text.strip());
        }
    }

    public void setDetail(String text) {
        if (text == null || text.isBlank()) {
            detailLabel.setText("");
            detailLabel.setManaged(false);
            detailLabel.setVisible(false);
            return;
        }
        detailLabel.setText(text.strip());
        detailLabel.setManaged(true);
        detailLabel.setVisible(true);
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
