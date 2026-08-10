package jp.co.pm.ai.desktop;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URL;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import javafx.animation.FadeTransition;
import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.shape.Rectangle;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.stage.WindowEvent;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.AppVersionInfo;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.StartupFactorySiteResolver;
import jp.co.pm.ai.desktop.ui.AppWindowIconSupport;

/**
 * Premium startup splash until main window FXML is loaded and initialized.
 *
 * <p>Visual assets: {@code css/startup-splash.css}, {@code images/splash-background.png}.
 */
final class StartupSplashStage {

    /** スプラッシュが表示されたとみなしてから、本体ロード等の後続処理を始めるまでの待ち（ナノ秒）。 */
    private static final long SPLASH_NEXT_LOGIC_DELAY_NANOS = 3_000_000_000L;

    private static final String SPLASH_CSS_RESOURCE = "/jp/co/pm/ai/desktop/css/startup-splash.css";
    private static final String SPLASH_BACKGROUND_RESOURCE =
            "/jp/co/pm/ai/desktop/images/splash-background.png";

    private StartupSplashStage() {}

    private static URL resolveClasspathResource(String absoluteResourcePath) {
        for (Class<?> anchor : new Class<?>[] {StartupSplashStage.class, PmAiFxApp.class}) {
            URL url = anchor.getResource(absoluteResourcePath);
            if (url != null) {
                return url;
            }
        }
        String relative =
                absoluteResourcePath.startsWith("/")
                        ? absoluteResourcePath.substring(1)
                        : absoluteResourcePath;
        ClassLoader context = Thread.currentThread().getContextClassLoader();
        if (context != null) {
            URL url = context.getResource(relative);
            if (url != null) {
                return url;
            }
        }
        return ClassLoader.getSystemResource(relative);
    }

    private static String requireClasspathResourceUrl(String absoluteResourcePath) {
        return Objects.requireNonNull(
                        resolveClasspathResource(absoluteResourcePath),
                        () ->
                                "classpath resource missing: "
                                        + absoluteResourcePath
                                        + " (run mvn compile; check target/classes"
                                        + absoluteResourcePath
                                        + ")")
                .toExternalForm();
    }

    private static Image loadSplashBackgroundImage() {
        return loadSplashBackgroundImage(SPLASH_BACKGROUND_RESOURCE);
    }

    private static Image loadSplashBackgroundImage(String resourcePath) {
        String path =
                resourcePath != null && !resourcePath.isBlank()
                        ? resourcePath.strip()
                        : SPLASH_BACKGROUND_RESOURCE;
        if (!path.startsWith("/")) {
            path = "/" + path;
        }
        for (Class<?> anchor : new Class<?>[] {StartupSplashStage.class, PmAiFxApp.class}) {
            try (InputStream in = anchor.getResourceAsStream(path)) {
                if (in != null) {
                    byte[] bytes = in.readAllBytes();
                    return new Image(new ByteArrayInputStream(bytes), true);
                }
            } catch (IOException e) {
                throw new UncheckedIOException("failed to read splash background: " + path, e);
            }
        }
        String relative = path.startsWith("/") ? path.substring(1) : path;
        ClassLoader context = Thread.currentThread().getContextClassLoader();
        if (context != null) {
            try (InputStream in = context.getResourceAsStream(relative)) {
                if (in != null) {
                    byte[] bytes = in.readAllBytes();
                    return new Image(new ByteArrayInputStream(bytes), true);
                }
            } catch (IOException e) {
                throw new UncheckedIOException("failed to read splash background: " + path, e);
            }
        }
        try (InputStream in = ClassLoader.getSystemResourceAsStream(relative)) {
            if (in != null) {
                byte[] bytes = in.readAllBytes();
                return new Image(new ByteArrayInputStream(bytes), true);
            }
        } catch (IOException e) {
            throw new UncheckedIOException("failed to read splash background: " + path, e);
        }
        throw new IllegalStateException(
                "classpath resource missing: "
                        + path
                        + " (run mvn compile; check target/classes"
                        + path
                        + ")");
    }

    /**
     * Creates and shows the splash. Must run on the JavaFX application thread.
     *
     * <p>{@code outVisibleSinceNanos} が非 null のとき、最初にウィンドウが表示されたとみなせる時刻（ナノ秒）を
     * 一度だけ格納する。{@link javafx.stage.WindowEvent#WINDOW_SHOWN} またはその直後のパルスで設定する。
     *
     * <p>{@code afterSplashFullyDisplayed} が非 null のとき、{@link javafx.stage.WindowEvent#WINDOW_SHOWN} のあと 2
     * パルス経過後（レイアウト・初回描画後）、かつスプラッシュ表示開始から {@value #SPLASH_NEXT_LOGIC_DELAY_NANOS}
     * ナノ秒経過後に一度だけ呼ぶ。本体ロード等はこのコールバック内で開始する。
     *
     * @param outVisibleSinceNanos 表示開始時刻を格納するコンテナ（必要なければ {@code null}）
     * @param afterSplashFullyDisplayed スプラッシュの画面表示完了後に実行する処理（不要なら {@code null}）
     * @return the stage; close it when the main window is ready
     */
    static Stage createAndShow(
            AtomicLong outVisibleSinceNanos, Consumer<Stage> afterSplashFullyDisplayed) {
        return createAndShow(StartupSplashBranding.PMD, outVisibleSinceNanos, afterSplashFullyDisplayed);
    }

    static Stage createAndShow(
            StartupSplashBranding branding,
            AtomicLong outVisibleSinceNanos,
            Consumer<Stage> afterSplashFullyDisplayed) {
        StartupSplashBranding b = branding != null ? branding : StartupSplashBranding.PMD;
        boolean rdpLauncher = isRemoteDesktopLauncher(b);
        FactorySite factorySite =
                b.showFactorySite() ? StartupFactorySiteResolver.resolveForSplash() : null;

        Stage stage = new Stage();
        stage.initStyle(StageStyle.TRANSPARENT);
        stage.initModality(Modality.APPLICATION_MODAL);
        stage.setAlwaysOnTop(true);
        stage.setTitle("起動中");
        AppWindowIconSupport.applyTo(
                stage,
                rdpLauncher
                        ? AppWindowIconSupport.Variant.RDP_LAUNCHER
                        : AppWindowIconSupport.Variant.DESKTOP);

        ImageView background = null;
        if (hasBackgroundResource(b.backgroundResource())) {
            background = new ImageView(loadSplashBackgroundImage(b.backgroundResource()));
            background.setPreserveRatio(false);
            background.setSmooth(true);
            background.getStyleClass().add("splash-background-image");
        }

        Region overlay = new Region();
        overlay.getStyleClass().add("splash-overlay");

        Region accentBar = new Region();
        accentBar.getStyleClass().add("splash-accent-bar");
        javafx.scene.Node brandMark = rdpLauncher ? StartupSplashRdpGraphics.createBrandIcon() : accentBar;

        Label company = new Label("NAGAOKA SANGYOU");
        company.getStyleClass().add("splash-company-name");

        Label title = new Label(b.title());
        title.getStyleClass().add("splash-title");

        Label subtitleJa = new Label(b.subtitleJa());
        subtitleJa.getStyleClass().add("splash-subtitle-ja");

        Label subtitleEn = new Label(b.subtitleEn());
        subtitleEn.getStyleClass().add("splash-subtitle-en");

        VBox titleBlock;
        if (factorySite != null) {
            Label factoryBadge = new Label(factorySite.displayLabelJa());
            factoryBadge.getStyleClass().add("splash-factory-badge");
            titleBlock = new VBox(4, company, factoryBadge, title, subtitleJa, subtitleEn);
        } else {
            titleBlock = new VBox(6, company, title, subtitleJa, subtitleEn);
        }

        HBox brandRow = new HBox(12, brandMark, titleBlock);
        brandRow.getStyleClass().add("splash-brand-row");
        brandRow.setAlignment(Pos.CENTER_LEFT);

        Label status = new Label(b.statusText());
        status.getStyleClass().add("splash-status");

        ProgressIndicator busy = new ProgressIndicator();
        busy.setPrefSize(42, 42);
        busy.setMaxSize(42, 42);
        busy.getStyleClass().add("splash-progress");

        String versionText =
                "v"
                        + AppVersionInfo.resolveDisplayedVersion(
                                Path.of(System.getProperty("user.dir", ".")), Map.of());
        Label version = new Label(versionText);
        version.getStyleClass().add("splash-version");

        Region footerSpacer = new Region();
        HBox.setHgrow(footerSpacer, Priority.ALWAYS);

        HBox footer = new HBox(footerSpacer, version);
        footer.getStyleClass().add("splash-footer");
        footer.setAlignment(Pos.CENTER_RIGHT);

        VBox content = new VBox(10, brandRow, status, busy, footer);
        content.getStyleClass().add("splash-content");
        content.setAlignment(Pos.CENTER_LEFT);
        VBox.setMargin(busy, new Insets(2, 0, 0, 17));

        StackPane root = new StackPane();
        if (background != null) {
            root.getChildren().add(background);
        }
        root.getChildren().add(overlay);
        if (rdpLauncher) {
            root.getChildren().add(StartupSplashRdpGraphics.createBackgroundDecor());
        }
        root.getChildren().add(content);
        root.getStyleClass().add("splash-root");
        if (factorySite != null) {
            root.getStyleClass().add(factoryStyleClass(factorySite));
        }
        if (b.rootStyleClass() != null && !b.rootStyleClass().isBlank()) {
            root.getStyleClass().add(b.rootStyleClass().strip());
        }
        root.setPrefSize(520, 320);
        root.setMinSize(520, 320);
        root.setMaxSize(520, 320);

        Scene scene = new Scene(root);
        scene.setFill(null);
        scene.getStylesheets().add(requireClasspathResourceUrl(SPLASH_CSS_RESOURCE));

        if (background != null) {
            background.fitWidthProperty().bind(root.widthProperty());
            background.fitHeightProperty().bind(root.heightProperty());
        }

        Rectangle clip = new Rectangle();
        clip.widthProperty().bind(root.widthProperty());
        clip.heightProperty().bind(root.heightProperty());
        clip.setArcWidth(24);
        clip.setArcHeight(24);
        root.setClip(clip);

        stage.setScene(scene);
        stage.setResizable(false);
        stage.centerOnScreen();

        root.setOpacity(0.0);
        FadeTransition fadeIn = new FadeTransition(Duration.millis(420), root);
        fadeIn.setFromValue(0.0);
        fadeIn.setToValue(1.0);

        if (outVisibleSinceNanos != null) {
            stage.addEventHandler(
                    WindowEvent.WINDOW_SHOWN,
                    e -> outVisibleSinceNanos.compareAndSet(0L, System.nanoTime()));
        }
        AtomicBoolean bootstrapStarted = new AtomicBoolean(false);
        Runnable startNextLogic =
                () -> {
                    if (afterSplashFullyDisplayed == null) {
                        return;
                    }
                    if (!bootstrapStarted.compareAndSet(false, true)) {
                        return;
                    }
                    long since =
                            outVisibleSinceNanos != null
                                    ? outVisibleSinceNanos.get()
                                    : 0L;
                    if (since == 0L) {
                        since = System.nanoTime();
                    }
                    long deadlineNanos = since + SPLASH_NEXT_LOGIC_DELAY_NANOS;
                    long waitNs = deadlineNanos - System.nanoTime();
                    Runnable runBootstrap = () -> afterSplashFullyDisplayed.accept(stage);
                    if (waitNs <= 0) {
                        runBootstrap.run();
                        return;
                    }
                    PauseTransition pause =
                            new PauseTransition(Duration.millis(waitNs / 1_000_000.0));
                    pause.setOnFinished(e -> runBootstrap.run());
                    pause.play();
                };
        if (afterSplashFullyDisplayed != null) {
            stage.addEventHandler(
                    WindowEvent.WINDOW_SHOWN,
                    e ->
                            Platform.runLater(
                                    () -> Platform.runLater(startNextLogic)));
        }
        stage.setOnShown(e -> fadeIn.play());
        stage.show();
        raiseToFront(stage);
        if (outVisibleSinceNanos != null) {
            Platform.runLater(() -> outVisibleSinceNanos.compareAndSet(0L, System.nanoTime()));
        }
        if (afterSplashFullyDisplayed != null) {
            Platform.runLater(() -> Platform.runLater(startNextLogic));
        }
        return stage;
    }

    private static boolean isRemoteDesktopLauncher(StartupSplashBranding branding) {
        return "splash-app-rdp-launcher".equals(branding.rootStyleClass());
    }

    private static boolean hasBackgroundResource(String resourcePath) {
        return resourcePath != null && !resourcePath.isBlank();
    }

    private static String factoryStyleClass(FactorySite site) {
        if (site == FactorySite.KOKUBU) {
            return "splash-factory-kokubu";
        }
        return "splash-factory-konan";
    }

    /** Moves splash forward after OS focus steal or other Stage creation. */
    static void raiseToFront(Stage splash) {
        if (splash == null) {
            return;
        }
        splash.toFront();
        splash.requestFocus();
    }
}
