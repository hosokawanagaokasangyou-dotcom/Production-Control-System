package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.function.IntConsumer;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tooltip;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
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

    private static final String DESKTOP_CSS =
            Objects.requireNonNull(
                            EquipmentStatusFullscreenStage.class.getResource(
                                    "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css"),
                            "pm-ai-desktop.css")
                    .toExternalForm();

    private final Stage stage = new Stage(StageStyle.UNDECORATED);
    private final BorderPane root = new BorderPane();
    private final HBox cardHost = new HBox();
    private final FlowPane cardPane = new FlowPane();
    private final ScrollPane scrollPane = new ScrollPane();
    private final VBox emptyStateHost = new VBox();
    private final VBox loadingHost = new VBox(12.0);
    private final Label actualDateLabel = new Label();
    private final Label planDateLabel = new Label();
    private final Label metaLabel = new Label();
    private Runnable onClose;
    private IntConsumer onAdjustActualDateDays;
    private IntConsumer onAdjustPlanDateDays;
    private boolean ownerInitialized;
    private EquipmentStatusDashboardAppearancePrefs appearance =
            EquipmentStatusDashboardAppearancePrefs.defaults();

    public EquipmentStatusFullscreenStage() {
        root.getStyleClass().add("pm-equipment-status-fullscreen-root");
        applyFullscreenTheme(EquipmentStatusDashboardAppearancePrefs.defaults());

        HBox top = new HBox(16.0);
        top.setAlignment(Pos.CENTER_LEFT);
        top.setPadding(new Insets(8, 12, 8, 12));
        top.getStyleClass().add("pm-equipment-status-fullscreen-top");
        Label title = new Label("ダッシュボード");
        title.getStyleClass().add("pm-equipment-status-fullscreen-title");
        actualDateLabel.getStyleClass().add("pm-equipment-status-fullscreen-date");
        planDateLabel.getStyleClass().add("pm-equipment-status-fullscreen-date");
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        metaLabel.getStyleClass().add("pm-equipment-status-fullscreen-meta");
        Button close = new Button("閉じる");
        close.setOnAction(e -> hide());
        top.getChildren()
                .addAll(title, actualDateLabel, planDateLabel, spacer, metaLabel, close);

        cardPane.getStyleClass().add("pm-equipment-status-card-flow");
        cardHost.getChildren().add(cardPane);
        scrollPane.setContent(cardHost);
        scrollPane.getStyleClass().add("pm-equipment-status-scroll");
        scrollPane.viewportBoundsProperty()
                .addListener(
                        (o, a, b) ->
                                applyFlowLayout(b.getWidth()));

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

        root.setFocusTraversable(true);
        root.setTop(top);
        root.setCenter(center);
        root.setBottom(buildOperationGuide());

        Scene scene = new Scene(root);
        if (!scene.getStylesheets().contains(DESKTOP_CSS)) {
            scene.getStylesheets().add(DESKTOP_CSS);
        }
        scene.addEventFilter(KeyEvent.KEY_PRESSED, this::onSceneKeyPressed);
        stage.setScene(scene);
        stage.initModality(Modality.NONE);
        stage.setOnHidden(e -> {
            if (onClose != null) {
                onClose.run();
            }
        });
    }

    private void onSceneKeyPressed(KeyEvent e) {
        if (e.getCode() == KeyCode.ESCAPE) {
            hide();
            e.consume();
            return;
        }
        if (e.isControlDown() || e.isAltDown() || e.isMetaDown()) {
            return;
        }
        int shift =
                switch (e.getCode()) {
                    case LEFT -> -1;
                    case RIGHT -> 1;
                    default -> 0;
                };
        if (shift == 0) {
            return;
        }
        if (e.isShiftDown()) {
            if (onAdjustPlanDateDays != null) {
                onAdjustPlanDateDays.accept(shift);
                e.consume();
            }
            return;
        }
        if (onAdjustActualDateDays != null) {
            onAdjustActualDateDays.accept(shift);
            e.consume();
        }
    }

    public void setOnClose(Runnable onClose) {
        this.onClose = onClose;
    }

    /** 全画面表示中に ← / → で実績日を前後させる（日数: 負=前日、正=翌日）。 */
    public void setOnAdjustActualDateDays(IntConsumer onAdjustActualDateDays) {
        this.onAdjustActualDateDays = onAdjustActualDateDays;
    }

    /** 全画面表示中に Shift + ← / → で予定日を前後させる。 */
    public void setOnAdjustPlanDateDays(IntConsumer onAdjustPlanDateDays) {
        this.onAdjustPlanDateDays = onAdjustPlanDateDays;
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
            String actualDateLabelText,
            String planDateLabelText,
            boolean sourcesLoaded) {
        if (!ownerInitialized && owner != null) {
            stage.initOwner(owner);
            ownerInitialized = true;
        }
        applyAppearance(appearancePrefs);
        setHeaderDates(actualDateLabelText, planDateLabelText);
        metaLabel.setText(metaText != null ? metaText : "");
        rebuildCards(
                statuses,
                opts,
                appearancePrefs,
                badgeStyleResolver,
                actualDateLabelText,
                planDateLabelText,
                sourcesLoaded);
        stage.setFullScreen(true);
        if (!stage.isFullScreen()) {
            stage.setMaximized(true);
        }
        stage.show();
        stage.toFront();
        javafx.application.Platform.runLater(
                () -> {
                    root.requestFocus();
                    applyFlowLayout(scrollPane.getViewportBounds().getWidth());
                });
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
            String actualDateLabelText,
            String planDateLabelText,
            boolean sourcesLoaded) {
        applyAppearance(appearancePrefs);
        applyFlowLayout(scrollPane.getViewportBounds().getWidth());
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
                                    actualDateLabelText, planDateLabelText, sourcesLoaded, true));
            return;
        }
        for (EquipmentMachineStatus s : statuses) {
            cardPane.getChildren()
                    .add(
                            EquipmentStatusCardFactory.createCard(
                                    s, opts, appearance, badgeStyleResolver, true));
        }
    }

    public void setHeaderDates(String actualDateLabelText, String planDateLabelText) {
        actualDateLabel.setText(formatHeaderDate("実績", actualDateLabelText));
        planDateLabel.setText(formatHeaderDate("予定", planDateLabelText));
    }

    public void setMetaText(String text) {
        metaLabel.setText(text != null ? text : "");
    }

    public void setLoadingVisible(boolean on) {
        loadingHost.setVisible(on);
        loadingHost.setManaged(on);
        if (on) {
            metaLabel.setText("データ読込中…");
        }
    }

    private void applyFlowLayout(double viewportWidth) {
        boolean fillViewport =
                EquipmentStatusDashboardAppearanceApplier.configureFlowPane(
                        cardPane, appearance, true, viewportWidth);
        EquipmentStatusDashboardAppearanceApplier.applyFlowHostLayout(cardHost, cardPane, fillViewport);
        scrollPane.setFitToWidth(
                EquipmentStatusDashboardAppearanceApplier.scrollShouldFitToWidth(appearance));
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

    private static String formatHeaderDate(String prefix, String dateLabel) {
        if (dateLabel == null || dateLabel.isBlank()) {
            return prefix + " —";
        }
        return prefix + " " + dateLabel.strip();
    }

    private HBox buildOperationGuide() {
        HBox bar = new HBox(24.0);
        bar.setAlignment(Pos.CENTER);
        bar.setPadding(new Insets(6, 12, 10, 12));
        bar.getStyleClass().add("pm-equipment-status-fullscreen-guide");

        Label title = new Label("操作");
        title.getStyleClass().add("pm-equipment-status-fullscreen-guide-title");
        bar.getChildren()
                .addAll(
                        title,
                        guideRowActual(),
                        guideRowPlan(),
                        guideRowClose());
        return bar;
    }

    private HBox guideRowActual() {
        HBox keysInner = new HBox(2.0);
        keysInner.setAlignment(Pos.CENTER);
        keysInner
                .getChildren()
                .addAll(
                        guideButton("←", "実績日を1日前", () -> adjustActual(-1)),
                        guideSepLabel("/"),
                        guideButton("→", "実績日を1日後", () -> adjustActual(1)));
        return wrapGuideItem(keysInner, "実績日を前後");
    }

    private HBox guideRowPlan() {
        HBox keysInner = new HBox(2.0);
        keysInner.setAlignment(Pos.CENTER);
        keysInner
                .getChildren()
                .addAll(
                        guideModLabel("Shift +"),
                        guideButton("←", "予定日を1日前", () -> adjustPlan(-1)),
                        guideSepLabel("/"),
                        guideButton("→", "予定日を1日後", () -> adjustPlan(1)));
        return wrapGuideItem(keysInner, "予定日を前後");
    }

    private HBox guideRowClose() {
        HBox keysInner = new HBox(guideButton("Esc", "全画面を閉じる", this::hide));
        keysInner.setAlignment(Pos.CENTER);
        return wrapGuideItem(keysInner, "全画面を閉じる");
    }

    private HBox wrapGuideItem(HBox keysInner, String description) {
        HBox keysBox = new HBox(keysInner);
        keysBox.setAlignment(Pos.CENTER);
        keysBox.getStyleClass().add("pm-equipment-status-fullscreen-guide-keys");
        Label descLabel = new Label(description);
        descLabel.getStyleClass().add("pm-equipment-status-fullscreen-guide-desc");
        HBox item = new HBox(8.0, keysBox, descLabel);
        item.setAlignment(Pos.CENTER_LEFT);
        return item;
    }

    private Button guideButton(String text, String tooltipText, Runnable action) {
        Button button = new Button(text);
        button.getStyleClass().add("pm-equipment-status-fullscreen-guide-btn");
        button.setFocusTraversable(false);
        if (tooltipText != null && !tooltipText.isBlank()) {
            Tooltip.install(button, new Tooltip(tooltipText));
        }
        if (action != null) {
            button.setOnAction(e -> action.run());
        }
        return button;
    }

    private Label guideSepLabel(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("pm-equipment-status-fullscreen-guide-sep");
        return label;
    }

    private Label guideModLabel(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("pm-equipment-status-fullscreen-guide-mod");
        return label;
    }

    private void adjustActual(int days) {
        if (onAdjustActualDateDays != null) {
            onAdjustActualDateDays.accept(days);
        }
    }

    private void adjustPlan(int days) {
        if (onAdjustPlanDateDays != null) {
            onAdjustPlanDateDays.accept(days);
        }
    }
}
