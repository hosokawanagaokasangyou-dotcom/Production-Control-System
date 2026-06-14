package jp.co.pm.ai.desktop;

import javafx.geometry.Pos;
import javafx.scene.layout.Pane;
import javafx.scene.layout.StackPane;
import javafx.scene.shape.Circle;
import javafx.scene.shape.Line;
import javafx.scene.shape.Rectangle;

/** リモートデスクトップ RPA ランチャー向けスプラッシュ装飾（モニター・接続線）。 */
final class StartupSplashRdpGraphics {

    private StartupSplashRdpGraphics() {}

    /** 右側背景のモニター接続モチーフ。 */
    static Pane createBackgroundDecor() {
        Pane layer = new Pane();
        layer.getStyleClass().add("splash-rdp-decor");
        layer.setMouseTransparent(true);
        layer.setPickOnBounds(false);

        Rectangle grid = new Rectangle(220, 200);
        grid.getStyleClass().add("splash-rdp-grid");
        grid.setLayoutX(288);
        grid.setLayoutY(28);

        Pane remoteMonitor = buildMonitor("splash-rdp-monitor-remote", 128, 92);
        remoteMonitor.setLayoutX(352);
        remoteMonitor.setLayoutY(44);

        Pane localMonitor = buildMonitor("splash-rdp-monitor-local", 76, 56);
        localMonitor.setLayoutX(262);
        localMonitor.setLayoutY(108);

        Line linkMain = new Line(310, 136, 368, 92);
        linkMain.getStyleClass().add("splash-rdp-link");
        linkMain.getStyleClass().add("splash-rdp-link-main");

        Line linkGlow = new Line(310, 136, 368, 92);
        linkGlow.getStyleClass().add("splash-rdp-link");
        linkGlow.getStyleClass().add("splash-rdp-link-glow");

        Circle nodeLocal = new Circle(310, 136, 4);
        nodeLocal.getStyleClass().add("splash-rdp-node");
        nodeLocal.getStyleClass().add("splash-rdp-node-local");

        Circle nodeRemote = new Circle(368, 92, 4.5);
        nodeRemote.getStyleClass().add("splash-rdp-node");
        nodeRemote.getStyleClass().add("splash-rdp-node-remote");

        Circle pulse = new Circle(339, 114, 5);
        pulse.getStyleClass().add("splash-rdp-pulse");

        layer.getChildren().addAll(
                grid, linkGlow, linkMain, nodeLocal, pulse, nodeRemote, localMonitor, remoteMonitor);
        return layer;
    }

    /** ブランド行左の RDP アイコン（モニター＋接続）。 */
    static StackPane createBrandIcon() {
        StackPane icon = new StackPane();
        icon.getStyleClass().add("splash-rdp-brand-icon");
        icon.setMinSize(52, 56);
        icon.setPrefSize(52, 56);
        icon.setMaxSize(52, 56);
        icon.setAlignment(Pos.CENTER);

        Pane leftMonitor = buildMonitor("splash-rdp-brand-monitor-left", 22, 16);
        leftMonitor.setTranslateX(-10);
        leftMonitor.setTranslateY(6);

        Pane rightMonitor = buildMonitor("splash-rdp-brand-monitor-right", 30, 22);
        rightMonitor.setTranslateX(10);
        rightMonitor.setTranslateY(-4);

        Line brandLink = new Line(-2, 10, 12, 2);
        brandLink.getStyleClass().add("splash-rdp-brand-link");

        Circle brandNode = new Circle(5, 6, 2.5);
        brandNode.getStyleClass().add("splash-rdp-brand-node");

        icon.getChildren().addAll(leftMonitor, rightMonitor, brandLink, brandNode);
        return icon;
    }

    private static Pane buildMonitor(String rootStyleClass, double width, double height) {
        double bezel = 3;
        double screenH = height - bezel * 2 - 2;
        double standW = width * 0.34;
        double standH = 4;
        double totalHeight = height + standH;

        Pane monitor = new Pane();
        monitor.getStyleClass().add(rootStyleClass);
        monitor.setPrefSize(width, totalHeight);
        monitor.setMinSize(width, totalHeight);
        monitor.setMaxSize(width, totalHeight);
        monitor.setMouseTransparent(true);

        Rectangle bezelRect = new Rectangle(0, 0, width, height);
        bezelRect.getStyleClass().add("splash-rdp-monitor-bezel");

        Rectangle screenRect = new Rectangle(bezel, bezel, width - bezel * 2, screenH);
        screenRect.getStyleClass().add("splash-rdp-monitor-screen");

        Rectangle standRect = new Rectangle((width - standW) / 2.0, height, standW, standH);
        standRect.getStyleClass().add("splash-rdp-monitor-stand");

        monitor.getChildren().addAll(bezelRect, screenRect, standRect);
        return monitor;
    }
}
