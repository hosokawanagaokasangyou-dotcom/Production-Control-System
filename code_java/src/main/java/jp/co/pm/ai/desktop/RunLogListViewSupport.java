package jp.co.pm.ai.desktop;

import java.util.Locale;

import javafx.geometry.Pos;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Region;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.scene.text.Font;
import javafx.scene.text.Text;

/**
 * 実行ログ ListView の1行描画。{@code TextFlow} は幅制約で折り返すため使わず、
 * {@code HBox}+{@code Text} で単一行にし、セル高を超えた描画はクリップする。
 */
final class RunLogListViewSupport {

    /** CSS {@code .list-cell} の上下 padding（2+2）に余裕を足した値。 */
    static final double CELL_VERTICAL_PADDING_PX = 8.0;

    private RunLogListViewSupport() {}

    static double measureLineHeightPx(Font font) {
        Text probe = new Text("Áyあ|");
        if (font != null) {
            probe.setFont(font);
        }
        double h = probe.getLayoutBounds().getHeight();
        if (!Double.isFinite(h) || h <= 0) {
            double size = font != null ? font.getSize() : 14.0;
            return Math.ceil(size * 1.35);
        }
        return h;
    }

    static double fixedCellSizePx(Font font) {
        double line = Math.ceil(measureLineHeightPx(font));
        return Math.clamp(line + CELL_VERTICAL_PADDING_PX, 22.0, 72.0);
    }

    static void installOverflowClip(Region node) {
        if (node == null || node.getClip() instanceof Rectangle) {
            return;
        }
        Rectangle clip = new Rectangle();
        clip.widthProperty().bind(node.widthProperty());
        clip.heightProperty().bind(node.heightProperty());
        node.setClip(clip);
    }

    static HBox buildLineGraphic(String item, Font font, Color baseFill, String search) {
        HBox row = new HBox();
        row.getStyleClass().add("pm-log-line");
        row.setAlignment(Pos.CENTER_LEFT);
        row.setFillHeight(false);
        double line = measureLineHeightPx(font);
        row.setMinHeight(line);
        row.setPrefHeight(line);
        row.setMaxHeight(line);
        row.setMinWidth(0);
        row.setPrefWidth(0);
        String text = item == null ? "" : item;
        String needle = search == null ? "" : search;
        if (needle.isEmpty()) {
            row.getChildren().add(plainText(text, font, baseFill));
            return row;
        }
        String lowerItem = text.toLowerCase(Locale.ROOT);
        String searchLower = needle.toLowerCase(Locale.ROOT);
        int from = 0;
        while (from < text.length()) {
            int idx = lowerItem.indexOf(searchLower, from);
            if (idx < 0) {
                row.getChildren().add(plainText(text.substring(from), font, baseFill));
                break;
            }
            if (idx > from) {
                row.getChildren().add(plainText(text.substring(from, idx), font, baseFill));
            }
            Text hit = plainText(text.substring(idx, idx + needle.length()), font, baseFill);
            hit.getStyleClass().add("pm-log-search-hit");
            row.getChildren().add(hit);
            from = idx + needle.length();
        }
        return row;
    }

    private static Text plainText(String value, Font font, Color fill) {
        Text text = new Text(value);
        text.setWrappingWidth(0);
        if (font != null) {
            text.setFont(font);
        }
        if (fill != null) {
            text.setFill(fill);
        }
        return text;
    }
}
