package jp.co.pm.ai.desktop.print;

import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.transform.Scale;

/**
 * オペレーターカード印刷用の組み立て。
 *
 * <p>画面プレビューは {@link OperatorCardPreviewFactory#A4_PREF_WIDTH}（794px）で組み立て、
 * {@link javafx.scene.control.ScrollPane#setFitToWidth(boolean)} で枠に合わせて縮小表示する。
 * 印刷は {@link javafx.print.PrinterJob#printPage} が Node 寸法をそのまま用紙へ描画するため、
 * ルート幅を可印刷幅へ直接縮めると列幅が潰れてヘッダーが「当日…」のように切れる。
 * 本クラスはプレビューと同じ 794px レイアウトを維持し、印刷直前に均一スケールで可印刷幅へ収める。
 */
public final class OperatorCardPrintCompositor {

    private OperatorCardPrintCompositor() {}

    /** 可印刷幅に対するレイアウト幅（794px）の縮小率。 */
    public static double printScaleForWidth(double printableWidth) {
        if (!Double.isFinite(printableWidth) || printableWidth <= 0) {
            return 1.0;
        }
        return printableWidth / OperatorCardPreviewFactory.A4_PREF_WIDTH;
    }

    /**
     * 794px で組み立て済みの 1 ページ分 {@link VBox} を、可印刷幅へ均一スケールした印刷ルートへ包む。
     */
    public static Parent wrapScaledPrintPage(VBox layoutRoot, double printableWidth) {
        if (layoutRoot == null) {
            return new StackPane();
        }
        double scale = printScaleForWidth(printableWidth);
        OperatorCardPreviewFactory.prepareForLayoutMeasure(
                layoutRoot, OperatorCardPreviewFactory.A4_PREF_WIDTH);
        double layoutHeight = layoutRoot.getBoundsInLocal().getHeight();
        double scaledHeight = Math.max(1.0, layoutHeight * scale);

        StackPane sheet = new StackPane();
        sheet.setAlignment(Pos.TOP_LEFT);
        sheet.setPrefWidth(printableWidth);
        sheet.setMinWidth(printableWidth);
        sheet.setMaxWidth(printableWidth);
        sheet.setPrefHeight(scaledHeight);
        sheet.setMinHeight(scaledHeight);
        sheet.setMaxHeight(scaledHeight);
        sheet.setStyle("-fx-background-color: white;");

        layoutRoot.getTransforms().setAll(new Scale(scale, scale, 0, 0));
        sheet.getChildren().add(layoutRoot);
        return sheet;
    }

    /**
     * 印刷用 {@link Scene} を生成し CSS を適用してレイアウトを確定する。
     */
    public static Scene createPrintScene(Parent printRoot, double printableWidth, double printableHeight) {
        double sceneW = Math.max(1.0, printableWidth);
        double sceneH =
                Math.max(
                        printableHeight,
                        printRoot.prefHeight(sceneW) > 0
                                ? printRoot.prefHeight(sceneW)
                                : printRoot.getBoundsInLocal().getHeight());
        Scene scene = new Scene(printRoot, sceneW, sceneH, Color.WHITE);
        OperatorCardPreviewFactory.attachDesktopStylesheet(scene);
        printRoot.applyCss();
        printRoot.layout();
        return scene;
    }
}
