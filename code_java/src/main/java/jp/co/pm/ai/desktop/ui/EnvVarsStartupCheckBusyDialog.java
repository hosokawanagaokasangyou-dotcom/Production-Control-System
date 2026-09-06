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

/**
 * 起動時の工場ワークスペース復元・環境変数照合中に表示する進捗モーダル。
 *
 * <p>ワークスペース復元・環境変数照合は必須処理のためキャンセル不可。タブ読込段階（起動後読込）だけは
 * 「バックグラウンドで続行」が有効になり、ダイアログを閉じて読込を低優先度で継続しつつ工場切替などの
 * 操作へ進める。呼び出し側で必ず {@link #close()} する。
 */
public final class EnvVarsStartupCheckBusyDialog {

    public static final String TITLE = "起動時チェック";
    public static final String HEADER = "環境変数を確認しています";
    public static final String HEADER_BACKGROUND_LOAD = "タブデータを読み込んでいます";

    public static final String CANCEL_TEXT = "バックグラウンドで続行";
    public static final String CANCEL_HINT =
            "ダイアログを閉じ、読込は優先度を下げてバックグラウンドで継続します。工場切替などの操作ができます";
    public static final String CANCEL_HINT_BLOCKED = "必須チェック中はキャンセルできません";

    public static final String STATUS_RESTORE_WORKSPACE = "工場ワークスペースを復元しています…";
    public static final String STATUS_STABILIZE = "環境変数の値を安定化しています…";
    public static final String STATUS_MATCH = "初期値と照合しています…";
    public static final String STATUS_FACTORY_SWITCH = "工場ワークスペースを切り替えています…";
    public static final String STATUS_BACKGROUND_LOAD = "タブデータを読み込んでいます…";
    public static final String STATUS_DONE = "完了しました";

    public static final String STEP_RESTORE_WORKSPACE = "1/3 ワークスペース復元";
    public static final String STEP_ENV_MATCH = "2/3 環境変数照合";
    public static final String STEP_TAB_LOAD = "3/3 タブ読込";

    private final Stage stage;
    private final Label headerLabel;
    private final Label stepLabel;
    private final Label statusLabel;
    private final Button cancelButton;
    private final Label cancelHintLabel;

    private EnvVarsStartupCheckBusyDialog(
            Stage stage,
            Label headerLabel,
            Label stepLabel,
            Label statusLabel,
            Button cancelButton,
            Label cancelHintLabel) {
        this.stage = stage;
        this.headerLabel = headerLabel;
        this.stepLabel = stepLabel;
        this.statusLabel = statusLabel;
        this.cancelButton = cancelButton;
        this.cancelHintLabel = cancelHintLabel;
    }

    /** タブ読込段階（起動後読込）の状況文言か。見出し・ステップ・キャンセル可否の判定に使う。 */
    public static boolean isTabLoadStatus(String status) {
        return status != null
                && (STATUS_BACKGROUND_LOAD.equals(status) || status.startsWith("起動後読込"));
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ（{@code null} 可）
     * @param initialStatus 最初の状況文言
     */
    public static EnvVarsStartupCheckBusyDialog show(Window owner, String initialStatus) {
        return show(owner, initialStatus, null);
    }

    /**
     * モーダル Stage を表示する。
     *
     * @param owner 親ウィンドウ（{@code null} 可）
     * @param initialStatus 最初の状況文言
     * @param onCancel キャンセル押下時（{@code null} でキャンセル不可）。既定は無効で、
     *     {@link #setCancelEnabled(boolean)} で有効化する
     */
    public static EnvVarsStartupCheckBusyDialog show(
            Window owner, String initialStatus, Runnable onCancel) {
        Stage stage = new Stage(StageStyle.UTILITY);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(TITLE);
        stage.setResizable(false);

        Label header = new Label(HEADER);
        header.getStyleClass().add("dialog-header");
        header.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");

        Label step = new Label("");
        step.setStyle("-fx-font-size: 12px; -fx-text-fill: -fx-text-inner-color;");

        ProgressIndicator progress = new ProgressIndicator();
        progress.setPrefSize(48, 48);
        progress.setMaxSize(48, 48);

        Label status = new Label(initialStatus != null ? initialStatus : STATUS_STABILIZE);
        status.setWrapText(true);
        status.setMaxWidth(340);

        VBox root = new VBox(14, header, step, progress, status);
        root.setAlignment(Pos.CENTER);
        root.setPadding(new Insets(20, 24, 24, 24));
        root.setPrefWidth(380);

        Button cancel = null;
        Label cancelHint = null;
        if (onCancel != null) {
            cancelHint = new Label(CANCEL_HINT_BLOCKED);
            cancelHint.setWrapText(true);
            cancelHint.setMaxWidth(340);
            cancelHint.setStyle("-fx-font-size: 11px;");
            cancelHint.getStyleClass().add("muted");
            cancel = new Button(CANCEL_TEXT);
            cancel.setDisable(true);
            cancel.setOnAction(e -> onCancel.run());
            root.getChildren().addAll(cancelHint, cancel);
        }

        // ユーザ操作での閉じは抑止し、キャンセル可能な段階だけキャンセルへ振り替える
        Button cancelForClose = cancel;
        stage.setOnCloseRequest(
                event -> {
                    event.consume();
                    if (cancelForClose != null && !cancelForClose.isDisabled()) {
                        onCancel.run();
                    }
                });

        stage.setScene(new Scene(root));
        stage.sizeToScene();
        stage.show();
        if (owner != null) {
            stage.centerOnScreen();
        }
        return new EnvVarsStartupCheckBusyDialog(stage, header, step, status, cancel, cancelHint);
    }

    /** キャンセルボタンの有効・無効。タブ読込段階（{@link #isTabLoadStatus(String)}）でのみ有効にする。 */
    public void setCancelEnabled(boolean enabled) {
        if (cancelButton == null) {
            return;
        }
        if (cancelButton.isDisabled() == !enabled) {
            return;
        }
        cancelButton.setDisable(!enabled);
        if (cancelHintLabel != null) {
            cancelHintLabel.setText(enabled ? CANCEL_HINT : CANCEL_HINT_BLOCKED);
        }
        stage.sizeToScene();
    }

    boolean isCancelEnabled() {
        return cancelButton != null && !cancelButton.isDisabled();
    }

    /** テスト用。キャンセルボタンを押した状態にする。 */
    void fireCancelForTest() {
        if (cancelButton != null) {
            cancelButton.fire();
        }
    }

    /** 起動シーケンスの大項目（例: {@link #STEP_RESTORE_WORKSPACE}）。空文字で非表示。 */
    public void setStep(String text) {
        if (stepLabel == null) {
            return;
        }
        boolean show = text != null && !text.isBlank();
        stepLabel.setText(show ? text : "");
        stepLabel.setVisible(show);
        stepLabel.setManaged(show);
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
