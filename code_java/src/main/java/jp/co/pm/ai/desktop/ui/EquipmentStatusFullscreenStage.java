package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.function.Function;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;

/** ダッシュボード全画面表示用 Stage。 */
public final class EquipmentStatusFullscreenStage {

    private final Stage stage = new Stage(StageStyle.UNDECORATED);
    private final FlowPane cardPane = new FlowPane(12.0, 12.0);
    private final Label metaLabel = new Label();
    private Runnable onClose;

    public EquipmentStatusFullscreenStage() {
        BorderPane root = new BorderPane();
        root.getStyleClass().add("pm-equipment-status-fullscreen-root");

        HBox top = new HBox(12.0);
        top.setAlignment(Pos.CENTER_LEFT);
        top.setPadding(new Insets(8, 12, 8, 12));
        Label title = new Label("ダッシュボード");
        title.getStyleClass().add("pm-equipment-status-fullscreen-title");
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        metaLabel.getStyleClass().add("pm-equipment-status-fullscreen-meta");
        Button close = new Button("閉じる");
        close.setOnAction(e -> hide());
        top.getChildren().addAll(title, spacer, metaLabel, close);

        cardPane.setPadding(new Insets(12.0));
        cardPane.getStyleClass().add("pm-equipment-status-card-flow");
        ScrollPane scroll = new ScrollPane(cardPane);
        scroll.setFitToWidth(true);
        scroll.getStyleClass().add("pm-equipment-status-scroll");

        root.setTop(top);
        root.setCenter(scroll);

        Scene scene = new Scene(root);
        stage.setScene(scene);
        stage.initModality(Modality.NONE);
        scene.setOnKeyPressed(
                e -> {
                    if (e.getCode() == KeyCode.ESCAPE) {
                        hide();
                    }
                });
        stage.setOnHidden(e -> {
            if (onClose != null) {
                onClose.run();
            }
        });
    }

    public void setOnClose(Runnable onClose) {
        this.onClose = onClose;
    }

    public boolean isShowing() {
        return stage.isShowing();
    }

    public void show(
            Window owner,
            List<EquipmentMachineStatus> statuses,
            EquipmentStatusCardFactory.DisplayOptions opts,
            Function<String, PersonBadgeStyle> badgeStyleResolver,
            String metaText) {
        stage.initOwner(owner);
        metaLabel.setText(metaText != null ? metaText : "");
        rebuildCards(statuses, opts, badgeStyleResolver);
        stage.setFullScreen(true);
        if (!stage.isFullScreen()) {
            stage.setMaximized(true);
        }
        stage.show();
        stage.toFront();
    }

    public void rebuildCards(
            List<EquipmentMachineStatus> statuses,
            EquipmentStatusCardFactory.DisplayOptions opts,
            Function<String, PersonBadgeStyle> badgeStyleResolver) {
        cardPane.getChildren().clear();
        if (statuses != null) {
            for (EquipmentMachineStatus s : statuses) {
                cardPane.getChildren()
                        .add(
                                EquipmentStatusCardFactory.createCard(
                                        s, opts, badgeStyleResolver, true));
            }
        }
    }

    public void hide() {
        if (stage.isShowing()) {
            stage.setFullScreen(false);
            stage.hide();
        }
    }
}
