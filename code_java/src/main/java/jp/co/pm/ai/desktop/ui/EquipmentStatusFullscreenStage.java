package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.function.Function;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.ScrollPane;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;

/** ダッシュボード全画面表示用 Stage。 */
public final class EquipmentStatusFullscreenStage {

    private static final String FULLSCREEN_THEME_CLASS_PREFIX =
            "pm-equipment-status-fullscreen-theme-";

    private final Stage stage = new Stage(StageStyle.UNDECORATED);
    private final BorderPane root = new BorderPane();
    private final FlowPane cardPane = new FlowPane();
    private final ScrollPane scrollPane = new ScrollPane();
    private final VBox emptyStateHost = new VBox();
    private final VBox loadingHost = new VBox(12.0);
    private final Label metaLabel = new Label();
    private Runnable onClose;
    private boolean ownerInitialized;
    private EquipmentStatusDashboardAppearancePrefs appearance =
            EquipmentStatusDashboardAppearancePrefs.defaults();

    public EquipmentStatusFullscreenStage() {
        root.getStyleClass().add("pm-equipment-status-fullscreen-root");
        applyFullscreenTheme(EquipmentStatusDashboardAppearancePrefs.defaults());

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

        cardPane.getStyleClass().add("pm-equipment-status-card-flow");
        scrollPane.setContent(cardPane);
        scrollPane.setFitToWidth(true);
        scrollPane.getStyleClass().add("pm-equipment-status-scroll");
        scrollPane.viewportBoundsProperty()
                .addListener(
                        (o, a, b) ->
                                EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                                        cardPane, appearance, true, b.getWidth()));

        emptyStateHost.setAlignment(Pos.CENTER);
        emptyStateHost.getStyleClass().add("pm-equipment-status-empty-host");
        emptyStateHost.setVisible(false);
        emptyStateHost.setManaged(false);

        loadingHost.setAlignment(Pos.CENTER);
        loadingHost.getStyleClass().add("pm-equipment-status-loading-overlay");
        loadingHost.setVisible(false);
        loadingHost.setManaged(false);
        ProgressIndicator busy = new ProgressIndicator();
        busy.setPrefSize(56, 56);
        Label loadingLbl = new Label("データ読込中…");
        loadingLbl.getStyleClass().add("pm-equipment-status-fullscreen-meta");
        loadingHost.getChildren().addAll(busy, loadingLbl);

        StackPane center = new StackPane(scrollPane, emptyStateHost, loadingHost);

        root.setTop(top);
        root.setCenter(center);

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
            EquipmentStatusDashboardAppearancePrefs appearancePrefs,
            Function<String, PersonBadgeStyle> badgeStyleResolver,
            String metaText,
            String actualDateLabel,
            String planDateLabel,
            boolean sourcesLoaded) {
        if (!ownerInitialized && owner != null) {
            stage.initOwner(owner);
            ownerInitialized = true;
        }
        applyAppearance(appearancePrefs);
        metaLabel.setText(metaText != null ? metaText : "");
        rebuildCards(
                statuses,
                opts,
                appearancePrefs,
                badgeStyleResolver,
                actualDateLabel,
                planDateLabel,
                sourcesLoaded);
        stage.setFullScreen(true);
        if (!stage.isFullScreen()) {
            stage.setMaximized(true);
        }
        stage.show();
        stage.toFront();
    }

    public void applyAppearance(EquipmentStatusDashboardAppearancePrefs appearancePrefs) {
        this.appearance =
                appearancePrefs != null
                        ? appearancePrefs
                        : EquipmentStatusDashboardAppearancePrefs.defaults();
        applyFullscreenTheme(this.appearance);
    }

    public void rebuildCards(
            List<EquipmentMachineStatus> statuses,
            EquipmentStatusCardFactory.DisplayOptions opts,
            EquipmentStatusDashboardAppearancePrefs appearancePrefs,
            Function<String, PersonBadgeStyle> badgeStyleResolver,
            String actualDateLabel,
            String planDateLabel,
            boolean sourcesLoaded) {
        applyAppearance(appearancePrefs);
        EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                cardPane,
                appearance,
                true,
                scrollPane.getViewportBounds().getWidth());
        cardPane.getChildren().clear();
        boolean empty = statuses == null || statuses.isEmpty();
        emptyStateHost.getChildren().clear();
        emptyStateHost.setVisible(empty);
        emptyStateHost.setManaged(empty);
        if (empty) {
            emptyStateHost
                    .getChildren()
                    .add(
                            EquipmentStatusCardFactory.createEmptyState(
                                    actualDateLabel, planDateLabel, sourcesLoaded, true));
            return;
        }
        for (EquipmentMachineStatus s : statuses) {
            cardPane.getChildren()
                    .add(
                            EquipmentStatusCardFactory.createCard(
                                    s, opts, appearance, badgeStyleResolver, true));
        }
    }

    public void setMetaText(String text) {
        metaLabel.setText(text != null ? text : "");
    }

    public void setLoadingVisible(boolean on) {
        loadingHost.setVisible(on);
        loadingHost.setManaged(on);
        scrollPane.setOpacity(on ? 0.45 : 1.0);
        if (on) {
            metaLabel.setText("データ読込中…");
        }
    }

    public void hide() {
        if (stage.isShowing()) {
            stage.setFullScreen(false);
            stage.hide();
        }
    }

    private void applyFullscreenTheme(EquipmentStatusDashboardAppearancePrefs prefs) {
        root.getStyleClass()
                .removeIf(c -> c.startsWith(FULLSCREEN_THEME_CLASS_PREFIX));
        root.getStyleClass().add(prefs.fullscreenThemeStyleClass());
    }
}
